"""
As part of converting an excel workbook into code (just julia code at the moment), we want to try and improve the generated code as much as possible.

Part of this is converting repeated formulas (that might arrise from copy and
pasting a formula between rows of a table for example), into a format that is
more like a person would write.

This function aims to very generically identify these situations on each sheet.
It has to ensure all the required properties for the conversion to be valid at
met, this includes the statements not being inter-depdenent, having the same
formula, parameters acting accordingly to how broadcasting should work, etc.


"""

"""
Partition `nodes` into antichains under full-graph reachability.

Arguments:
- g     : global DAG
- topo  : topological_sort(g), precomputed once
- nodes : subset to partition

Returns:
Vector of antichains, each a Vector{Int}.
"""
function partition_to_antichains(
    g::DiGraph,
    topo::Vector{Int},
    nodes::Vector{Int},
)
    n = nv(g)

    # Mark S
    inS = falses(n)
    for u in nodes
        inS[u] = true
    end

    # DP ranks (0 = unreachable from S)
    rank = zeros(Int, n)

    # Activate frontier
    active = falses(n)
    for u in nodes
        active[u] = true
        rank[u] = 1
    end

    # Propagate only from active nodes
    for u in topo
        if !active[u]
            continue
        end

        # @show u

        ru = rank[u]
        for v in outneighbors(g, u)
            # If we improve v, propagate
            if rank[v] <= ru
                rank[v] = ru + 1
                active[v] = true
            end
        end
    end

    # Collect antichains
    layers = Dict{Int, Vector{Int}}()
    for u in nodes
        push!(get!(layers, rank[u], Int[]), u)
    end

    return collect(values(layers))
end


function new_broadcast(statements::Vector{AbstractStatement}; debug::Bool=true)
    stmt_graph = make_statement_graph(statements)
    new_broadcast(statements, stmt_graph, debug=debug)
end

struct ReductionResult
    consumed::Vector{Int}
    produced::Vector{AbstractStatement}
end

"""
Find antichains within `stmt_nodes` and keep only chains that are large enough to
attempt broadcast reduction. Returned chains are sorted by set cell.
"""
function filter_independent_chains(
    statements::Vector{AbstractStatement},
    stmt_graph::DiGraph,
    topo::Vector{Int},
    stmt_nodes::Vector{Int};
    min_chain_len::Int = 3,
    debug::Bool = false,
)
    isempty(stmt_nodes) && return Vector{Vector{Int}}()

    stmt_lhs = [get_set_cells(statements[s])[1] for s in stmt_nodes]
    seed_order = sortperm(stmt_lhs)
    sorted_nodes = stmt_nodes[seed_order]

    stmt_antichains = partition_to_antichains(stmt_graph, topo, sorted_nodes)
    debug && println("Number of antichains = $(length(stmt_antichains))")

    chains = Vector{Vector{Int}}()

    for chain in stmt_antichains
        if length(chain) < min_chain_len
            if debug
                println("A chain of length $(length(chain)) was skipped over for broadcasting")
                @show [get_set_cells(statements[s])[1] for s in chain]
            end
            continue
        end

        chain_lhs = [get_set_cells(statements[s])[1] for s in chain]
        chain_order = sortperm(chain_lhs)
        push!(chains, chain[chain_order])
    end

    sort!(chains; by = c -> get_set_cells(statements[first(c)])[1])
    return chains
end

"""
Reduce one independent chain into grouped/broadcasted statements.

`params` is expected to have one row per statement in `chain`, in the same order
as `chain`.
"""
function reduce_chain(
    statements::Vector{AbstractStatement},
    chain::Vector{Int},
    func,
    params::AbstractMatrix{ExcelExpr};
    debug::Bool = false,
)::ReductionResult
    isempty(chain) && return ReductionResult(Int[], AbstractStatement[])

    chain_lhs = [get_set_cells(statements[s])[1] for s in chain]
    sort_order = sortperm(chain_lhs)

    sorted_chain = chain[sort_order]
    sorted_params = params[sort_order, :]

    reduced = apply_broadcast_stages(statements[sorted_chain], func, sorted_params; debug = debug)
    return ReductionResult(sorted_chain, reduced)
end

function new_broadcast(statements::Vector{AbstractStatement}, stmt_graph::DiGraph; debug::Bool = true, test_sheet = nothing, test_cell = nothing)

    new_statements = copy(statements)

    statements_by_sheet = group_to_dict(1:length(statements), s -> (c -> c.sheet_name).(get_set_cells(statements[s])))

    topo = topological_sort(stmt_graph)

    for (sheets, sheet_statements) in statements_by_sheet
        # @show sheet
        length(sheets) == 1 || continue
        if !isnothing(test_sheet)
            sheets[1] == test_sheet || continue
        end
        # Some debugging lines
        # sheets[1] == "growth- cohort group 1"  || continue
        # sheets[1] == "growth- cohort group 2" || continue
        # startswith(sheets[1], "growth- cohort") || continue
        # sheets[1] == "Structure Calcs" || continue
        # sheets[1] == "Operations" || continue
        # sheets[1] == "oyster Husbandry model" || continue

        sheet = sheets[1]

        println("$sheet has $(length(sheet_statements)) statements")

        func_numbering = ObjectNumbering(Vector{Any}())
        func_params = Vector{Matrix{ExcelExpr}}(undef, length(sheet_statements))
        statement_funcs = Vector{Int64}(undef, length(sheet_statements))

        stmt_to_sheet_i = Dict(n => i for (i, n) in enumerate(sheet_statements))

        for (i, stmt_i) in enumerate(sheet_statements)
            stmt = statements[stmt_i]
            func, params = functionalize(stmt.rhs_expr)

            num = get_num!(func_numbering, func isa FlatExpr ? FastHashedFlatExpr(func) : func)
            statement_funcs[i] = num

            func_params[i] = params
        end

        uses_same_function = group_to_dict(sheet_statements, c -> statement_funcs[stmt_to_sheet_i[c]])

        @show length(uses_same_function)


        for (func_i, stmt_indices) in uses_same_function
            func = get_obj(func_numbering, func_i)
            func isa FastHashedFlatExpr || continue
            length(stmt_indices) <= 1 && continue

            # @display func.expr

            stmt_lhs = [get_set_cells(statements[s])[1] for s in stmt_indices]
            sheet_stmt_i = [stmt_to_sheet_i[i] for i in stmt_indices]

            # test_cell = CellDependency("growth- cohort group 2", "AB64")
            # test_cell = CellDependency("growth- cohort group 2", "G12")
            # test_cell = CellDependency("Structure Calcs", "BR6")
            # test_cell = CellDependency("oyster Husbandry model", "A32")

            if !isnothing(test_cell)
                if !(CellDependency(test_sheet, test_cell) in stmt_lhs)
                    # println("Has $(test_cell)!!")
                    continue
                end
            end


            if debug
                println("\nFunction has $(length(stmt_indices)) usages")
            end
            # if test_cell in stmt_lhs
            #     println("Has $(test_cell)!!")
            # else
            #     continue
            # end

            sort_order = sortperm(stmt_lhs)

            stmt_lhs = stmt_lhs[sort_order]
            stmt_nodes = stmt_indices[sort_order]
            params = reduce(vcat, func_params[sheet_stmt_i])
            params = params[sort_order, :]

            chain_candidates = filter_independent_chains(
                statements,
                stmt_graph,
                topo,
                stmt_nodes;
                min_chain_len = 3,
                debug = debug,
            )

            if isempty(chain_candidates)
                if debug
                    println("No broadcast-eligible independent chains were found")
                end
                continue
            end

            node_to_param_i = Dict(n => i for (i, n) in enumerate(stmt_nodes))

            for chain in chain_candidates
                if debug
                    println("Chain of length $(length(chain))")
                end

                chain_lhs = [get_set_cells(statements[s])[1] for s in chain]
                # if test_cell in chain_lhs
                #     println("chain has test_cell!!")
                # else
                #     continue
                # end

                if debug
                    if length(chain_lhs) > 3
                        @show chain_lhs[begin:3]
                    else
                        @show chain_lhs
                    end
                end

                chain_param_rows = [node_to_param_i[n] for n in chain]
                chain_params = params[chain_param_rows, :]

                length_before = length(new_statements)
                statement_set = Set(statements[chain])
                filter!(!in(statement_set), new_statements)

                reduction = reduce_chain(statements, chain, func, chain_params; debug = debug)
                append!(new_statements, reduction.produced)
                length_after = length(new_statements)

                if debug
                    println("Reduced the number of statements by $(length_before - length_after) by broadcasting, which gave $(length(reduction.produced)) statements")
                end
            end



        end
    end

    println("Net removed $(length(statements) - length(new_statements)) total statements")

    new_statements

end


struct BroadcastCtx
    stmts::Vector{AbstractStatement}
    params::Matrix{ExcelExpr}
end

abstract type AbstractChunk end

struct Chunk <: AbstractChunk
    idxs::Vector{Int}      # indices into ctx.stmts
    kind::Symbol           # :candidate | :passthrough | :group
    reason::Symbol
end

struct Chunk2D <: AbstractChunk
    idxs::Matrix{Int}      # indices into ctx.stmts
    kind::Symbol           # :candidate | :passthrough | :group
    reason::Symbol
end

struct DiagEvent
    stage::Symbol
    reason::Symbol
    idxs::Vector{Int}
    meta::NamedTuple
end

# generic stage runner
function run_stage(chunks::Vector{<:AbstractChunk}, ctx::BroadcastCtx, stage::Symbol, f, diag::Vector{DiagEvent})
    out = AbstractChunk[]
    for ch in chunks
        if ch.kind != :candidate
            push!(out, ch)
            continue
        end
        produced, events = f(ch, ctx)
        append!(out, produced)
        append!(diag, events)
    end
    out
end

function split_non_table(ch::AbstractChunk, ctx::BroadcastCtx)
    idxs = ch.idxs
    isempty(idxs) && return Chunk[], DiagEvent[]

    # Keep order stable.
    # table_mask = ctx.is_table[idxs]
    table_mask = isa.(ctx.stmts[idxs], [TableStatement])

    # Fast path: all are table statements, keep candidate unchanged.
    if all(table_mask)
        return [ch], DiagEvent[]
    end

    table_idxs = idxs[table_mask]
    non_table_idxs = idxs[.!table_mask]

    out = Chunk[]
    events = DiagEvent[]

    # Non-table statements are emitted as passthrough.
    if !isempty(non_table_idxs)
        push!(out, Chunk(non_table_idxs, :passthrough, :non_table_statement))
    end

    # Table statements continue only if they can still form a group.
    if length(table_idxs) >= 2
        push!(out, Chunk(table_idxs, :candidate, :table_only))
    elseif length(table_idxs) == 1
        push!(out, Chunk(table_idxs, :passthrough, :single_table_statement))
    end

    push!(events, DiagEvent(
        :table_only,
        :mixed_table_non_table,
        copy(vec(idxs)),
        (
            total = length(idxs),
            table_count = length(table_idxs),
            non_table_count = length(non_table_idxs),
            pct_table = 100 * length(table_idxs) / length(idxs),
        ),
    ))

    return out, events
end

function split_single_cell(ch::Chunk, ctx::BroadcastCtx)
    idxs = ch.idxs
    isempty(idxs) && return Chunk[], DiagEvent[]

    single_mask = map(i -> length(get_set_cells(ctx.stmts[i])) == 1, idxs)

    if all(single_mask)
        return [Chunk(idxs, :candidate, :single_cell_only)], DiagEvent[]
    end

    single_idxs = idxs[single_mask]
    non_single_idxs = idxs[.!single_mask]

    out = Chunk[]
    events = DiagEvent[]

    if !isempty(non_single_idxs)
        push!(out, Chunk(non_single_idxs, :passthrough, :not_single_cell_target))
    end
    if !isempty(single_idxs)
        push!(out, Chunk(single_idxs, :candidate, :single_cell_only))
    end

    push!(events, DiagEvent(
        :single_cell,
        :split_single_and_multi_cell,
        copy(idxs),
        (
            total = length(idxs),
            single_count = length(single_idxs),
            multi_count = length(non_single_idxs),
        ),
    ))

    return out, events
end

function split_regions(ch::Chunk2D, ctx::BroadcastCtx)
    return [ch], DiagEvent[]
end
function split_regions(ch::Chunk, ctx::BroadcastCtx)
    idxs = ch.idxs
    isempty(idxs) && return Chunk[], DiagEvent[]

    set_cells = [get_set_cells(ctx.stmts[i])[1] for i in idxs]
    coords = [(colnum(c), rownum(c)) for c in set_cells]
    coord_to_stmt_i = Dict(c => i for (c, i) in zip(coords, idxs))
    sort!(coords)
    regions = get_2d_regions(coords)

    out = AbstractChunk[]
    events = DiagEvent[]

    push!(events, DiagEvent(
        :regions,
        :region_count,
        copy(idxs),
        (num_regions = length(regions),),
    ))

    for (cols, rows) in regions
        region_idxs = [coord_to_stmt_i[(c, r)] for r in rows, c in cols]
        region_area = length(cols) * length(rows)

        firstcol = XLSX.encode_column_number(cols[1])
        lastcol = XLSX.encode_column_number(cols[end])

        push!(events, DiagEvent(
            :regions,
            :region,
            copy(vec(region_idxs)),
            (
                rows = rows[1]:rows[end],
                cols = "$firstcol:$lastcol",
                area = region_area,
            ),
        ))

        if region_area == 1
            push!(out, Chunk(vec(region_idxs), :passthrough, :region_size_one))
        else
            push!(out, Chunk2D(region_idxs, :candidate, :region_candidate))
        end
    end

    return out, events
end

function get_broadcast_end(stmts::Vector{AbstractStatement}, run_statement_idxs, params; debug::Bool = true)
    if length(run_statement_idxs) == 1
        return 1, "single statement"
    end

    start_stmt_i = run_statement_idxs[1]
    params1 = params[start_stmt_i, :]
    stmt_i2 = run_statement_idxs[2]
    params2 = params[stmt_i2, :]


    equal_params = Vector{Int}()
    changing_params = Vector{Int}()
    for (i, (p1, p2)) in enumerate(zip(params1, params2))
        if p1 == p2
            # push!(equal_params, i)
            push!(changing_params, i)
        else
            push!(changing_params, i)
        end
    end

    start_cell = get_set_cells(stmts[run_statement_idxs[1]])[1]

    for (i, s_i) in enumerate(run_statement_idxs[2:end])
        cell = get_set_cells(stmts[s_i])[1]

        row_offset = rownum(cell) - rownum(start_cell)
        col_offset = colnum(cell) - colnum(start_cell)
        for p_i in equal_params
            if params[s_i, p_i] != params1[p_i]
                if debug
                    # println("Changing params not equal")
                    # @show params[s_i, p_i]
                    # @show params1[p_i]
                end
                return i, "equal params not equal $(params[s_i, p_i]), $(params1[p_i])"
            end
        end

        for p_i in changing_params
            if !exprs_equal_with_offset(params[s_i, p_i], params1[p_i], row_offset, col_offset)
                if debug
                    # println("Changing params not equal with offset")
                    # @show p_i findfirst(==(p_i), changing_params)
                    # @show params[s_i, p_i]
                    # @show params1[p_i]
                    # @show row_offset col_offset
                end
                return i, "changing params not equal with offset $(params[s_i, p_i]), $(params1[p_i])"
                # return i
            end
        end
    end

    return length(run_statement_idxs), ""
end

function get_param_cell_coords(expr::ExcelExpr)
    @match expr begin
        ExcelExpr(:cell_ref, [cell, sheet]) => reverse(parse_cell(cell))
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => (row_idx, col_idx)
        # We don't have enough context to determine the actual cell a named range belongs to
        # so instead we'll just hash the name to handle the equal case
        # This doesn't protect against hash collisions, but given the context it
        # seems pretty unlikely that two different named ranges will be
        # referenced in a broadcasted statement and also have names that hash
        # collide
        ExcelExpr(:named_range, [name]) => (hash(name), hash(name))
        _ => (missing, missing)
    end
end

function all_equal(vals)
    all(==(vals[1]), vals)
end

"""
Returns the offset (row_diff, col_diff) between elements in vals

Returns nothing if no such pattern can be identified
"""
function identify_broadcast_behavior(vals)
    @assert length(size(vals)) == 2

    if all_equal(vals)
        return (0, 0)
    end

    try
        # all columns are equal
        if (all(all_equal, eachcol(vals)))
            row_vals = @view vals[1, :]
            row_diff = diff(row_vals)
            if all_equal(row_diff)
                return (0, row_diff[1])
            end
        end

        if (all(all_equal, eachrow(vals)))
            col_vals = @view vals[:, 1]
            col_diff = diff(col_vals)
            if all_equal(col_diff)
                return (col_diff[1], 0)
            end
        end
    catch
        return nothing
    end



    return nothing
end

function split_broadcast_runs(ch::Chunk2D, ctx::BroadcastCtx)
    idxs = ch.idxs
    isempty(idxs) && return Chunk[], DiagEvent[]
    length(idxs) == 1 && return [Chunk(vec(idxs), :passthrough, :run_size_one)], DiagEvent[]

    out = Chunk[]
    events = DiagEvent[]

    stmts_mat = ctx.stmts[idxs]
    params_mat = ctx.params[idxs, :]
    # @show size(ctx.params) size(idxs) size(params_mat)

    # first index is (row_coord, col_coord)
    if !isempty(params_mat)
        # shape = 2 x rows x cols x param_num
        param_coords = stack(get_param_cell_coords, params_mat)
        # @display param_coords

        # @show size(param)
        param_broadcast_behavior = Matrix{Any}(nothing, size(ctx.params, 2), 2)
        for i in axes(param_coords, 4)
            # println("Param $i")
            coords = @view param_coords[:, :, :, i]
            # @display coords
            # if i == 4
            #     # @display params_mat[1:2, 1:2, i]
            #     @display params_mat[:, 1, i]
            # end

            if any(ismissing, coords)
                continue
            end

            rows = @view coords[1, :, :]
            cols = @view coords[2, :, :]
            # @display rows
            # @display cols

            row_behavior = identify_broadcast_behavior(rows)
            col_behavior = identify_broadcast_behavior(cols)
            # if i == 4
            #     # @display rows[1:2, 1:2]
            #     # @display cols[1:2, 1:2]
            #     @display rows
            #     @display cols
            #     @show row_behavior col_behavior
            # end

            param_broadcast_behavior[i, :] .= (row_behavior, col_behavior)
        end

        if !any(isnothing, param_broadcast_behavior)
            push!(events, DiagEvent(
                :broadcast_runs,
                :range_count,
                copy(vec(idxs)),
                (num_ranges = 1,),
            ))

            # push!(out, ch)
            return [ch], events
        else
            failed_params = findall(r -> any(isnothing, r), eachrow(param_broadcast_behavior))
            push!(events, DiagEvent(
                :broadcast_runs,
                :full_broadcast_fail,
                copy(vec(idxs)),
                (failed_params = failed_params, example_params=ctx.params[1:2, failed_params]),
            ))
        end
    end


    local_stmts = vec(ctx.stmts[idxs])
    # local_params = reshape(ctx.params[vec(idxs), :], length(idxs) size(ctx.params, 2))
    local_params = ctx.params[vec(idxs), :]
    local_run_idxs = collect(eachindex(local_stmts))
    # @show size(local_params)

    function try_stmt_order(stmts, params)
        # out = Chunk[]
        # events = DiagEvent[]
        try_ranges = Vector{UnitRange{Int}}()
        try_reasons = String[]

        last_i = 1
        while last_i <= length(local_run_idxs)
            run_len, reason = get_broadcast_end(stmts, @view(local_run_idxs[last_i:end]), params; debug = false)
            push!(try_reasons, reason)
            push!(try_ranges, last_i:(last_i+run_len-1))
            last_i += run_len
        end

        try_ranges, try_reasons
    end

    ranges, reasons = try_stmt_order(local_stmts, local_params)

    if length(ranges) >= 2
        function sort_by(s)
            c = get_set_cells(s)[1]
            return (rownum(c), colnum(c))
        end
        sort_order = sortperm(local_stmts, by = sort_by)

        new_ranges, new_reasons = try_stmt_order(local_stmts[sort_order], local_params[sort_order, :])
        if length(new_ranges) < sqrt(length(ranges))
            ranges = [sort_order[r] for r in new_ranges]
            reasons = new_reasons
        end
    end

    # ranges = Vector{UnitRange{Int}}()
    # reasons = String[]

    # last_i = 1
    # while last_i <= length(local_run_idxs)
    #     run_len, reason = get_broadcast_end(local_stmts, @view(local_run_idxs[last_i:end]), local_params; debug = false)
    #     push!(reasons, reason)
    #     push!(ranges, last_i:(last_i+run_len-1))
    #     last_i += run_len
    # end

    push!(events, DiagEvent(
        :broadcast_runs,
        :range_count,
        copy(vec(idxs)),
        (num_ranges = length(ranges),),
    ))

    for (r, reason) in zip(ranges, reasons)
        run_idxs = idxs[r]
        run_len = length(r)

        push!(events, DiagEvent(
            :broadcast_runs,
            :range,
            copy(run_idxs),
            (length = run_len, reason = reason),
        ))

        if run_len == 1
            push!(out, Chunk(run_idxs, :passthrough, :run_size_one))
        else
            push!(out, Chunk(run_idxs, :candidate, :broadcast_run))
        end
    end

    return out, events
end

to_single_index(idx::Integer) = Int(idx)
to_single_index(idx::AbstractRange{<:Integer}) = length(idx) == 1 ? Int(first(idx)) : nothing

parse_single_table_ref(expr) = nothing
function parse_single_table_ref(expr::ExcelExpr)
    if expr.head != :table_ref || length(expr.args) != 5
        return nothing
    end

    table, row_idx, col_idx, fixed_row, fixed_col = expr.args
    if !(fixed_row isa Tuple{Bool, Bool}) || !(fixed_col isa Tuple{Bool, Bool})
        return nothing
    end

    # row = to_single_index(row_idx)
    # col = to_single_index(col_idx)
    row = row_idx
    col = col_idx
    if isnothing(row) || isnothing(col)
        return nothing
    end

    (
        table = table,
        row = row,
        col = col,
        fixed_row = fixed_row,
        fixed_col = fixed_col,
    )
end

function materialize_chunk(chunk::AbstractChunk, ctx::BroadcastCtx, func)
    stmts = ctx.stmts[chunk.idxs]
    params = ctx.params[chunk.idxs, :]
    # @display params

    function fallback()
        # println("\nMaking grouped statement!!")
        # println("Falling back to grouped statement")
        GroupedStatement(vec(stmts))
    end

    function fallback(msg)
        # println("\nMaking grouped statement!!")
        # println("Fallback because: ", msg)
        GroupedStatement(vec(stmts))
    end



    # We want to check if we can broadcast a single function over the whole chunk

    # The requirements for this are:
    # - the statements are all TableStatements
    # - the lhs_expr's form a rectangular region
    # - the non-fixed params behave correctly with broadcasting (expr_equal_with_offset)
    # - the non-fixed params are all of type :table_ref, and only reference a single value
    #
    # of these, only the last one isn't implied by the current chunking stages

    # Look at convert_to_broadcasted in src/expr_utils.jl for a previous implementation of handling broadcasting

    if !all(s -> s isa TableStatement, stmts)
        return fallback("not all table statements")
    end

    # Vector{ExcelExpr}, if they're all table statements, they should all have
    # :table_ref as the head
    lhs_exprs = (s -> s.lhs_expr).(stmts)
    if !all(e -> (e isa ExcelExpr) && e.head == :table_ref, lhs_exprs)
        return fallback("lhs_expr not all table_ref")
    end

    tables = get_set_table.(stmts)
    if length(unique(tables)) != 1
        return fallback("num set tables not 1")
    end
    table = tables[1]

    set_cells = [get_set_cells(s)[1] for s in stmts]
    row_vals = rownum.(set_cells)
    col_vals = colnum.(set_cells)
    unique_rows = sort(unique(row_vals))
    unique_cols = sort(unique(col_vals))

    expected_area = length(unique_rows) * length(unique_cols)
    if expected_area != length(stmts)
        return fallback("expected area != length(stmts)")
    end

    if unique_rows != collect(first(unique_rows):last(unique_rows))
        return fallback("unique rows not right")
    end
    if unique_cols != collect(first(unique_cols):last(unique_cols))
        return fallback("unique cols not right")
    end

    row_offset = last(unique_rows) - first(unique_rows)
    col_offset = last(unique_cols) - first(unique_cols)

    if size(params)[begin:end-1] != size(stmts)
        return fallback("size params not correct")
    end

    base_func = func isa FastHashedFlatExpr ? func.expr : func
    # param_replacements = Dict{Int, ExcelExpr}()
    num_params = size(params, 3)

    for p_i in 1:num_params
        col_params = params[:, :, p_i]

        coords = stack(get_param_cell_coords, col_params)
        rows = @view coords[1, :, :]
        cols = @view coords[2, :, :]
        if any(ismissing, rows) || any(ismissing, cols)
            return fallback()
        end
        row_behavior = identify_broadcast_behavior(rows)
        col_behavior = identify_broadcast_behavior(cols)

        if isnothing(row_behavior) || isnothing(col_behavior)
            return fallback("Broadcast behavior was nothing")
        end

        # @show row_behavior col_behavior

        first_param = col_params[1]
        # is_changing = !all(v -> isequal(v, first_param), col_params)
        is_changing = row_behavior != (0, 0) || col_behavior != (0, 0)

        if !is_changing
            if length(rows[1]) > 1 || length(cols[1]) > 1
                # param_replacements[p_i] = ExcelExpr(:broadcast_protect, first_param)
                continue
            else
                # param_replacements[p_i] = first_param
                continue
            end
        end

        first_ref = parse_single_table_ref(first_param)
        if isnothing(first_ref)
            @show first_param
            return fallback("first ref not single")
        end

        for p in col_params[2:end]
            parsed = parse_single_table_ref(p)
            if isnothing(parsed)
                return fallback("param not single")
            end
            if parsed.table != first_ref.table || parsed.fixed_row != first_ref.fixed_row || parsed.fixed_col != first_ref.fixed_col
                return fallback("param not same table and behavior")
            end
        end

        # fixed_row = first_ref.fixed_row
        # fixed_col = first_ref.fixed_col

        row_idx = first_ref.row
        col_idx = first_ref.col

        # row_idx = @match fixed_row begin
        #     (true, true) => row_idx
        #     (false, false) => begin
        #         if length(row_idx) == 1
        #             row_idx[1]:(row_idx[1]+row_offset)
        #         else
        #             # throw("Broadcasting table row range ref is complicated")
        #             return fallback()
        #         end
        #     end
        #     (true, false) => return fallback()
        #     (false, true) => return fallback()
        # end
        # col_idx = @match fixed_col begin
        #     (true, true) => col_idx
        #     (false, false) => begin
        #         if length(col_idx) == 1
        #             col_idx[1]:(col_idx[1]+col_offset)
        #         else
        #             return fallback()
        #             # throw("Broadcasting table col range ref is complicated")
        #         end
        #     end
        #     (true, false) => return fallback()
        #     (false, true) => return fallback()
        # end

        # row_idx = first_ref.row:offset_table_idx(first_ref.row, first_ref.fixed_row, row_offset)
        # col_idx = first_ref.col:offset_table_idx(first_ref.col, first_ref.fixed_col, col_offset)
        # @show first_ref row_offset row_idx
        # @display col_params
        # @show row_idx col_idx row_offset col_offset
        if !any(==(0), row_behavior) || !any(==(0), col_behavior)
            return fallback("complex row or col behavior")
        end
        # if any(s -> !(s isa Int), row_behavior)
        #     @display rows
        #     @show row_behavior
        # end

        # if any(s -> !(s isa Int), col_behavior)
        #     @show col_behavior
        # end
        # row_idx = row_idx:(row_idx + first(row_behavior[1]) * row_offset + first(row_behavior[2]) * col_offset)
        # col_idx = col_idx:(col_idx + first(col_behavior[1]) * row_offset + first(col_behavior[2]) * col_offset)

        # @assert first(row_idx) >= 1
        # @assert last(row_idx) <= size(first_ref.table)[1]
        # @assert first(col_idx) >= 1
        # @assert last(col_idx) <= size(first_ref.table)[2]
        # param_replacements[p_i] = ExcelExpr(:table_ref, first_ref.table, row_idx, col_idx, first_ref.fixed_row, first_ref.fixed_col)
    end
    # @display param_replacements

    # rhs_expr = try
    #     replace_func_params(base_func, param_replacements)
    # catch
    #     return fallback()
    # end

    # If we can broadcast everything, then we have to generate the broadcasted expression
    # look in src/grouped_statements.jl and specifically at the replace_func_params function for 
    # an example of putting values back into a functionalized statement
    # Additionally, look at src/transform/table_broadcast.jl for a previous example of this kind of transform

    # If we can't broadcast everything, just return a grouped statement

    lhs_row_idx = (first(unique_rows)-startrow(table)+1):(last(unique_rows)-startrow(table)+1)
    lhs_col_idx = (first(unique_cols)-startcol(table)+1):(last(unique_cols)-startcol(table)+1)

    @assert first(lhs_row_idx) >= 1
    @assert last(lhs_row_idx) <= size(table)[1]
    @assert first(lhs_col_idx) >= 1
    @assert last(lhs_col_idx) <= size(table)[2]
    lhs_expr = ExcelExpr(:table_ref, table, lhs_row_idx, lhs_col_idx, (true, true), (true, true))
    assigned_vars = reduce(vcat, get_set_cells.(vec(stmts)))
    rhs_dependencies = reduce(vcat, get_cell_deps.(vec(stmts))) |> unique |> collect

    # println("\nMaking table statement!!\n")
    # TableStatement(lhs_expr, assigned_vars, rhs_expr, rhs_dependencies, true)
    BroadcastedStatement(lhs_expr, assigned_vars, base_func, params, rhs_dependencies)
end

function materialize_chunks(chunks::Vector{AbstractChunk}, ctx::BroadcastCtx, func)
    out = Vector{AbstractStatement}()
    sizehint!(out, length(chunks))

    for ch in chunks
        stmts = ctx.stmts[ch.idxs]
        if ch.kind == :passthrough
            append!(out, stmts)
        elseif length(stmts) >= 2
            push!(out, materialize_chunk(ch, ctx, func))
        else
            append!(out, stmts)
        end
    end
    return out
end

function _diag_idx_preview(idxs::Vector{Int}; max_items::Int = 6)
    if length(idxs) <= max_items
        return string(idxs)
    end

    head = join(idxs[1:max_items], ", ")
    return "[$head, ...] (n=$(length(idxs)))"
end

function _diag_meta_string(meta::NamedTuple)
    # names = sort!(collect(keys(meta)); by = string)
    names = keys(meta)
    isempty(names) && return "{}"

    parts = String[]
    for name in names
        push!(parts, "$(name)=$(getfield(meta, name))")
    end
    return "{" * join(parts, ", ") * "}"
end

function print_diag(diag::Vector{DiagEvent})
    if isempty(diag)
        println("Broadcast diagnostics: no events")
        return nothing
    end

    by_stage = Dict{Symbol, Int}()
    by_stage_reason = Dict{Tuple{Symbol, Symbol}, Int}()
    total_idx_refs = 0

    for ev in diag
        by_stage[ev.stage] = get(by_stage, ev.stage, 0) + 1
        key = (ev.stage, ev.reason)
        by_stage_reason[key] = get(by_stage_reason, key, 0) + 1
        total_idx_refs += length(ev.idxs)
    end

    println("Broadcast diagnostics:")
    println("  events = $(length(diag))")
    println("  total_idx_refs = $total_idx_refs")
    println("  by_stage:")

    stages = sort!(collect(keys(by_stage)); by = string)
    for stage in stages
        println("    $(stage): $(by_stage[stage])")
        reasons = [r for (s, r) in keys(by_stage_reason) if s == stage]
        sort!(reasons; by = string)
        for reason in reasons
            println("      $(reason): $(by_stage_reason[(stage, reason)])")
        end
    end

    println("  event_details:")
    for (i, ev) in enumerate(diag)
        println(
            "    [$i] stage=$(ev.stage) reason=$(ev.reason) idxs=$(_diag_idx_preview(ev.idxs)) meta=$(_diag_meta_string(ev.meta))",
        )
    end

    return nothing
end


function apply_broadcast_stages(statements::Vector{AbstractStatement}, func, params; debug::Bool = true)
    # set_cells = [get_set_cells(s)[1] for s in statements]
    ctx = BroadcastCtx(statements, params)
    diag = DiagEvent[]

    chunks = [Chunk(collect(eachindex(statements)), :candidate, :initial)]
    chunks = run_stage(chunks, ctx, :single_cell, split_single_cell, diag)
    chunks = run_stage(chunks, ctx, :regions, split_regions, diag)
    chunks = run_stage(chunks, ctx, :broadcast_runs, split_broadcast_runs, diag)
    chunks = run_stage(chunks, ctx, :table_only, split_non_table, diag)
    chunks = run_stage(chunks, ctx, :regions, split_regions, diag)

    out = materialize_chunks(chunks, ctx, func)  # :group => GroupedStatement, :passthrough => raw stmts

    debug && print_diag(diag)
    return out
end

"""
Expects a list of statements that all use functionalized expression func.
The statements should probably also be sorted by set_cell.
The statements should set the same sheet.
Statements should not be inter-dependent. (Can be achived with partition_to_antichains).

Attempts to both do the broadcasting, as well as provide meaningful diagnositics for why it couldn't.
"""
function do_broadcast_reduction(statements::Vector{AbstractStatement}, func, params; debug::Bool = true)
    new_statements = Vector{AbstractStatement}()

    set_a_single_cell = [length(get_set_cells(s)) == 1 for s in statements]
    stmts = statements[set_a_single_cell]
    append!(new_statements, statements[.!set_a_single_cell])
    masked_params = params[set_a_single_cell, :]

    set_cells = [get_set_cells(s)[1] for s in stmts]
    row_nums = rownum.(set_cells)
    col_nums = colnum.(set_cells)
    coords = zip(col_nums, row_nums) |> collect
    coord_to_statement_i = Dict(c => s for (c, s) in zip(coords, eachindex(stmts)))
    sort!(coords)
    regions = get_2d_regions(coords)

    if debug
        println("Num regions = $(length(regions))")
    end

    for region in regions
        cols, rows = region
        region_area = (length(cols) * length(rows))

        if debug
            firstcol = XLSX.encode_column_number(cols[1])
            lastcol = XLSX.encode_column_number(cols[end])
            println("rows = $rows, cols = $firstcol:$lastcol")
        end

        run_statement_idxs = vec([coord_to_statement_i[(c, r)] for r in rows, c in cols])
        # run_statement_idxs = vec([coord_to_statement_i[(c, r)] for c in cols, r in rows])

        if region_area == 1
            append!(new_statements, stmts[run_statement_idxs])
            continue
        end

        # @show get_set_cells.(stmts[run_statement_idxs[begin:2]])

        last_i = 1
        ranges = Vector{UnitRange{Int}}()

        while last_i <= length(run_statement_idxs)
            i = get_broadcast_end(stmts, @view(run_statement_idxs[last_i:end]), masked_params)
            # @show last_i i

            push!(ranges, last_i:(last_i+i-1))
            last_i += i
        end

        for range in ranges
            if length(range) == 1
                if debug
                    println("Range has length of 1")
                end
                append!(new_statements, stmts[run_statement_idxs[range]])
                continue
            end

            if debug
                println("Range has $(length(range)) statements")
                # @show stmts[run_statement_idxs[range]][begin:3]
            end

            range_statements = stmts[run_statement_idxs[range]]

            mask = isa.(range_statements, [TableStatement])

            if sum(mask) != length(range)
                if debug
                    println("A range seemed to broadcast but $(100 * sum(mask) / length(range))% of the statements were table statements")
                    not_tbl_stmts = range_statements[.!mask]
                    if length(not_tbl_stmts) >= 3
                        @show get_set_cells.(not_tbl_stmts[begin:3])
                    else
                        @show get_set_cells.(not_tbl_stmts)
                    end
                end

                append!(new_statements, range_statements[.!mask])
                range_statements = range_statements[mask]

            end

            if !isempty(range_statements)
                push!(new_statements, GroupedStatement(range_statements))
            end
        end

    end

    return new_statements
end

