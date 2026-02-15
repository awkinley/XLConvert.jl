# flat_match.jl
#
# Provides @flat_match: a pattern-matching macro for FlatExpr that lets you
# write nested ExcelExpr patterns as if the expression were still a recursive
# ExcelExpr tree.  Any ExcelExpr(...) that appears in an argument position of
# another ExcelExpr pattern is automatically resolved through the FlatExpr's
# parts array via the corresponding FlatIdx.
#
# Requires: Match.jl  (MacroTools is optional / not needed here)

using Match

# ── internal helpers ────────────────────────────────────────────────────────

"""
    _is_excel_pattern(ex) -> Bool

True when `ex` is a Julia Expr of the form `ExcelExpr(...)`.
"""
_is_excel_pattern(ex) =
    ex isa Expr && Meta.isexpr(ex, :call) &&
    !isempty(ex.args) && ex.args[1] === :ExcelExpr

# ─────────────────────────────────────────────────────────────────────────────
"""
    _extract_nested!(ex, extractions) -> new_ex

Walk a pattern expression.  Every ExcelExpr(...) found *inside* an argument
position of another ExcelExpr is replaced with `FlatIdx(<gensym>)` and the
pair `(gensym, original_sub_pattern)` is appended to `extractions`.

Recursion stops at the replaced node; the sub-pattern itself is processed
lazily when `build_body` later builds the inner @match for that node.
"""
function _extract_nested!(ex, extractions)
    ex isa Expr || return ex                  # literals / symbols: leave as-is

    if _is_excel_pattern(ex)
        new_args = Any[ex.args[1]]            # keep the :ExcelExpr head
        for arg in ex.args[2:end]
            push!(new_args, _process_arg!(arg, extractions))
        end
        return Expr(:call, new_args...)
    elseif Meta.isexpr(ex, :vect)
        # top-level vector literal – treat each element as an arg
        return Expr(:vect, [_process_arg!(a, extractions) for a in ex.args]...)
    else
        # Other expression forms (tuples, etc.): recurse uniformly
        return Expr(ex.head, [_extract_nested!(a, extractions) for a in ex.args]...)
    end
end

"""
    _process_arg!(arg, extractions) -> new_arg

Handle a single argument inside an ExcelExpr pattern:
- If `arg` is itself an ExcelExpr pattern → replace with FlatIdx(gensym).
- If `arg` is a vector literal `[...]` → recurse element-wise.
- Otherwise → recurse normally.
"""
function _process_arg!(arg, extractions)
    if _is_excel_pattern(arg)
        var = gensym("fi")                    # unique FlatIdx binding variable
        push!(extractions, var => arg)        # record for later inner match
        return :(FlatIdx($var))
    elseif Meta.isexpr(arg, :vect)
        new_elems = Any[]
        for elem in arg.args
            if _is_excel_pattern(elem)
                var = gensym("fi")
                push!(extractions, var => elem)
                push!(new_elems, :(FlatIdx($var)))
            else
                push!(new_elems, _extract_nested!(elem, extractions))
            end
        end
        return Expr(:vect, new_elems...)
    else
        return _extract_nested!(arg, extractions)
    end
end

# ─────────────────────────────────────────────────────────────────────────────
"""
    build_body(flat_sym, body, extractions) -> expr

Given:
- `flat_sym`    – the local symbol bound to the FlatExpr value
- `body`        – the original arm body
- `extractions` – vector of `(idx_var, sub_pattern)` pairs

Returns code that wraps `body` in nested `@match` calls (innermost first)
so that each `FlatIdx` variable is resolved to its ExcelExpr part and
matched against its sub-pattern before `body` runs.

If any inner pattern fails to match, a descriptive error is thrown.
"""
function build_body(flat_sym, body, extractions)
    result = body
    # Wrap from the last extraction inward so the ordering in code is natural
    # (first extraction in extractions → outermost wrapping)
    for (idx_var, sub_pattern) in Iterators.reverse(extractions)
        # The sub_pattern itself may contain further nested ExcelExpr patterns
        inner_ext = Pair{Symbol,Any}[]
        new_sub   = _extract_nested!(sub_pattern, inner_ext)
        inner_body = build_body(flat_sym, result, inner_ext)

        pat_str = string(sub_pattern)       # for the error message
        result = quote
            if @ismatch $flat_sym.parts[$idx_var] $new_sub 
                # $new_sub => $inner_body
                $inner_body
                # _ => error(
                #     "@flat_match: nested pattern failed to match.\n" *
                #     "  Expected: " * $pat_str * "\n" *
                #     "  Got:      " * string($flat_sym.parts[$idx_var])
                # )
                # _ => @match_fail
            else
                @match_fail
            end
        end
    end
    return result
end

# ── public macro ─────────────────────────────────────────────────────────────

"""
    @flat_match flat_expr subject begin
        pattern1 => body1
        pattern2 => body2
        ...
    end

Pattern-match `subject` (typically `flat_expr.parts[i]`) against a set of
ExcelExpr patterns that may contain **nested** ExcelExpr sub-patterns.
Any `ExcelExpr(...)` appearing in an argument position of another `ExcelExpr`
pattern is automatically resolved by looking up the matching `FlatIdx` in
`flat_expr.parts`.

Variables in patterns are bound by Match.jl and are in scope for the
corresponding body.

## Comparison

**Before** (manual FlatIdx dereferencing):
```julia
@match part begin
    ExcelExpr(:range, [FlatIdx(lhs_i), FlatIdx(rhs_i)]) => begin
        lhs_expr = flat.parts[lhs_i]
        rhs_expr = flat.parts[rhs_i]
        lhs_expr.head == :cell_ref || throw("unexpected")
        rhs_expr.head == :cell_ref || throw("unexpected")
        lhs = lhs_expr.args[1]; sheet = lhs_expr.args[2]
        rhs = rhs_expr.args[1]
        # ... actual logic ...
    end
end
```

**After** (nested pattern matching via @flat_match):
```julia
@flat_match flat parts[i] begin
    ExcelExpr(:range, [ExcelExpr(:cell_ref, [lhs, sheet]),
                       ExcelExpr(:cell_ref, [rhs, sheet])]) => begin
        # lhs, sheet, rhs all bound; sheet equality handled by Match.jl
        # ... actual logic ...
    end
    ExcelExpr(:cell_ref, [cell, sheet]) => begin
        # ...
    end
    _ => continue
end
```

## Notes
- `flat_expr` is evaluated once and stored in a `let` binding.
- If an outer arm matches but an inner sub-pattern fails, an error is thrown
  with the expression that did not match.  Structure outer arms defensively
  (or add a `_` fall-through arm) if the sub-expressions may vary.
- Nesting is arbitrarily deep; each level is expanded recursively.
- Any pattern that does **not** contain nested ExcelExpr calls is passed
  through to Match.jl unchanged (e.g. `ExcelExpr(:cell_ref, [c, s])`,
  wildcard `_`, or a bare variable).
"""
macro flat_match(flat_expr, subject, cases)
    Meta.isexpr(cases, :block) ||
        error("@flat_match: third argument must be a begin...end block")

    flat_sym = gensym("flat")   # hygienic binding for the FlatExpr value
    new_arms = Any[]

    for stmt in cases.args
        stmt isa LineNumberNode && continue

        (Meta.isexpr(stmt, :call) && stmt.args[1] === :(=>)) ||
            error("@flat_match: expected `pattern => body`, got: $(stmt)")

        pattern = stmt.args[2]
        body    = stmt.args[3]

        # Extract any nested ExcelExpr patterns from this arm's pattern
        extractions = Pair{Symbol,Any}[]
        new_pat = _extract_nested!(pattern, extractions)

        # Wrap the body in the chain of inner @match calls
        new_body = build_body(flat_sym, body, extractions)

        push!(new_arms, :($new_pat => $new_body))
    end

    # Emit:  let flat_sym = flat_expr
    #            @match subject begin  arm1; arm2; ...  end
    #        end
    esc(quote
        let $flat_sym = $flat_expr
            @match $subject begin
                $(new_arms...)
            end
        end
    end)
end


# ── usage example ─────────────────────────────────────────────────────────────
# The convert_to_broadcasted function from the problem statement rewritten with
# @flat_match.  Compare to the fully manual version in the original code.

#= Example (requires the rest of your library to be in scope):

function convert_to_broadcasted(expr::FlatExpr, row_offset, col_offset)
    new_expr = copy(expr)
    handled  = Set{Int}()

    for (i, part) in enumerate(new_expr.parts)
        i in handled && continue

        @flat_match new_expr part begin

            # ── range whose both ends are absolute cell references ──────────
            ExcelExpr(:range, [ExcelExpr(:cell_ref, [lhs, sheet]),
                               ExcelExpr(:cell_ref, [rhs, sheet])]) => begin
                function is_fixed(cell)
                    m = match(r"([$]?[A-Z]+)([$]?[0-9]+)", cell)
                    @assert m.match == cell "Cell didn't parse properly"
                    m[1][1] == '$' && m[2][1] == '$'
                end
                if is_fixed(lhs) && is_fixed(rhs)
                    # Re-attach the two cell_ref parts under a broadcast_protect
                    push!(new_expr.parts, ExcelExpr(:cell_ref, lhs, sheet))
                    i1 = length(new_expr.parts); push!(handled, i1)
                    push!(new_expr.parts, ExcelExpr(:cell_ref, rhs, sheet))
                    i2 = length(new_expr.parts); push!(handled, i2)
                    push!(new_expr.parts, ExcelExpr(:range, FlatIdx(i1), FlatIdx(i2)))
                    ir = length(new_expr.parts); push!(handled, ir)
                    new_expr.parts[i] = ExcelExpr(:broadcast_protect, FlatIdx(ir))
                else
                    throw("Converting range to broadcasted is complicated: $(part)")
                end
            end

            # ── free-standing cell reference ────────────────────────────────
            ExcelExpr(:cell_ref, [cell, sheet]) => begin
                range_start = CellDependency(sheet, cell)
                stop_cell   = offset_cell_str(cell, row_offset, col_offset)
                range_stop  = CellDependency(sheet, stop_cell)

                if range_start != range_stop
                    push!(new_expr.parts, ExcelExpr(:cell_ref, range_start.cell, sheet))
                    i1 = length(new_expr.parts)
                    push!(new_expr.parts, ExcelExpr(:cell_ref, range_stop.cell, sheet))
                    i2 = length(new_expr.parts)
                    push!(handled, i1); push!(handled, i2)
                    new_expr.parts[i] = ExcelExpr(:range, FlatIdx(i1), FlatIdx(i2))
                end
            end

            # ── anything else: skip ─────────────────────────────────────────
            _ => nothing
        end
    end
    new_expr
end
=#