@auto_hash_equals struct FlatIdx
    i::Int32
end

@auto_hash_equals struct FlatExpr
    parts::Array{ExcelExpr, 1}
end

function Base.deepcopy(e::FlatExpr)
    FlatExpr(deepcopy(e.parts))
end
function Base.copy(e::FlatExpr)
    FlatExpr(copy(e.parts))
end

function Base.show(io::IO, expr::FlatExpr)
    # arg_str = join(map(repr, expr.args), ", ")
    # arg_str = join(map(repr, expr.args), ", ")
    # print(io, "ExcelExpr(:$(expr.head), $(arg_str))")
    print(io, "FlatExpr($(expr.parts))")
end

function Base.show(io::IO, ::MIME"text/plain", expr::FlatExpr)
    println(io, "FlatExpr")
    for i in eachindex(expr.parts)
        println(io, "\t", i, ":", expr.parts[i])
    end
end

function pretty_string(expr::FlatExpr)
    io = IOBuffer()
    show(io, "text/plain", expr)

    String(take!(io))
end

convert_to_flat_expr(val) = val

function convert_to_flat_expr(expr::ExcelExpr, start_idx::Int)
    parts = Vector{ExcelExpr}(undef, 1)
    next_idx = 2

    new_expr_args = similar(expr.args)
    for (i, arg) in enumerate(expr.args)
        if arg isa ExcelExpr
            flat_expr = convert_to_flat_expr(arg, start_idx + next_idx - 1)
            append!(parts, flat_expr.parts)
            # push!(new_expr_args, FlatIdx(start_idx + next_idx - 1))
            new_expr_args[i] = FlatIdx(start_idx + next_idx - 1)
            next_idx += length(flat_expr.parts)
        else
            new_expr_args[i] = arg
            # push!(new_expr_args, arg)
        end
    end
    parts[1] = ExcelExpr(expr.head, new_expr_args)
    # insert!(parts, 1, ExcelExpr(expr.head, new_expr_args))
    FlatExpr(parts)
end

convert_to_flat_expr(expr::ExcelExpr) = convert_to_flat_expr(expr, 1)

function convert_to_expr(expr::ExcelExpr, flat_expr::FlatExpr)
    has_flat_idx = false
    for arg in expr.args
        if arg isa FlatIdx
            has_flat_idx = true
            break
        end
    end
    !has_flat_idx && return expr

    new_args = similar(expr.args)
    for (i, arg) in enumerate(expr.args)
        if arg isa FlatIdx
            new_args[i] = convert_to_expr(flat_expr.parts[arg.i], flat_expr)
        else
            new_args[i] = arg
        end
    end

    ExcelExpr(expr.head, new_args)
end
convert_to_expr(expr::FlatExpr) = convert_to_expr(expr.parts[1], expr)

function offset(flat_expr::FlatExpr, rows::Int, cols::Int)
    new_expr = FlatExpr(copy(flat_expr.parts))
    for i in eachindex(new_expr.parts)
        part = new_expr.parts[i]
        @match part begin
            ExcelExpr(:cell_ref, args) => begin
                new_args = copy(args)
                new_args[1] = offset_cell_str(new_args[1], rows, cols)

                new_expr.parts[i] = ExcelExpr(:cell_ref, new_args)
            end
            ExcelExpr(:table_ref, [table, row_idx, col_idx, fixed_row, fixed_col]) => begin
                row_idx = @match fixed_row begin
                    (true, true) => row_idx
                    (false, false) => row_idx .+ rows
                    (true, false) => first(row_idx):(last(row_idx)+rows)
                    (false, true) => (first(row_idx)+rows):last(row_idx)
                end
                col_idx = @match fixed_col begin
                    (true, true) => col_idx
                    (false, false) => col_idx .+ cols
                    (true, false) => first(col_idx):(last(col_idx)+cols)
                    (false, true) => (first(col_idx)+cols):last(col_idx)
                end

                new_expr.parts[i] = ExcelExpr(:table_ref, table, row_idx, col_idx, fixed_row, fixed_col)
            end
            _ => ()
        end

    end

    new_expr
end


"""
    insert_expr_front!(expr::FlatExpr, new_part::ExcelExpr)


Insert a new part at the front of a FlatExpr, then does the required FlatIdx renumbering.

The inserted argument also has renumbering applied, so it can reference existing parts using their pre-insertion index.
"""
function insert_expr_front!(expr::FlatExpr, new_part::ExcelExpr)
    pushfirst!(expr.parts, new_part)

    for part in expr.parts
        for i in eachindex(part.args)
            arg = part.args[i]
            if arg isa FlatIdx
                part.args[i] = FlatIdx(arg.i + 1)
            end
        end
    end

    expr
end



"""
    insert_expr_front(expr::FlatExpr, new_part::ExcelExpr)


Insert a new part at the front of a FlatExpr, then does the required FlatIdx renumbering.

The inserted argument also has renumbering applied, so it can reference existing parts using their pre-insertion index.
"""
function insert_expr_front(expr::FlatExpr, new_part::ExcelExpr)
    insert_expr_front!(deepcopy(expr), new_part)
end

"""
Modify the expression by inserting a (recursively defined) ExcelExpr at a given part index.

This new expression replaces the existing expression tree from that location. 

The given new_value can be recursively defined, or can include FlatIdx
references to existing parts of expr. These must be to parts after part_i.

The index of any parts afterwards is not guarenteed, as things may be shuffled around an renumbered
"""
function change_expr_part!(expr::FlatExpr, part_i::Int, new_value::ExcelExpr)
    old_parts = expr.parts
    old_n = length(old_parts)
    (part_i < 1 || part_i > old_n) && throw(BoundsError(old_parts, part_i))

    local_parts = Vector{ExcelExpr}()
    local_internal_arg = Vector{BitVector}()

    function flatten_local!(node::ExcelExpr)
        idx = length(local_parts) + 1
        push!(local_parts, ExcelExpr(:missing, Any[]))
        push!(local_internal_arg, falses(length(node.args)))

        new_args = Vector{Any}(undef, length(node.args))
        for (arg_i, arg) in enumerate(node.args)
            if arg isa ExcelExpr
                child_i = flatten_local!(arg)
                new_args[arg_i] = FlatIdx(child_i)
                local_internal_arg[idx][arg_i] = true
            else
                new_args[arg_i] = arg
            end
        end

        local_parts[idx] = ExcelExpr(node.head, new_args)
        idx
    end
    flatten_local!(new_value)

    added_n = length(local_parts)
    idx_shift = added_n - 1
    interim_n = old_n + idx_shift
    interim_parts = Vector{ExcelExpr}(undef, interim_n)

    function remap_old_part(part::ExcelExpr)
        has_flat_idx = false
        for arg in part.args
            if arg isa FlatIdx
                has_flat_idx = true
                break
            end
        end
        !has_flat_idx && return part

        new_args = copy(part.args)
        for arg_i in eachindex(new_args)
            arg = new_args[arg_i]
            if arg isa FlatIdx
                new_i = if arg.i == part_i
                    part_i
                elseif arg.i > part_i
                    arg.i + idx_shift
                else
                    arg.i
                end
                new_args[arg_i] = FlatIdx(new_i)
            end
        end
        ExcelExpr(part.head, new_args)
    end

    for old_i in 1:(part_i - 1)
        interim_parts[old_i] = remap_old_part(old_parts[old_i])
    end

    for local_i in eachindex(local_parts)
        part = local_parts[local_i]
        new_args = copy(part.args)

        for arg_i in eachindex(new_args)
            arg = new_args[arg_i]
            arg isa FlatIdx || continue

            if local_internal_arg[local_i][arg_i]
                new_args[arg_i] = FlatIdx(part_i + arg.i - 1)
            else
                arg.i > part_i || throw(ArgumentError("FlatIdx in replacement must reference a part after part_i. part_i=$part_i, ref=$(arg.i)"))
                arg.i <= old_n || throw(BoundsError(old_parts, arg.i))
                new_args[arg_i] = FlatIdx(arg.i + idx_shift)
            end
        end

        interim_parts[part_i + local_i - 1] = ExcelExpr(part.head, new_args)
    end

    for old_i in (part_i + 1):old_n
        interim_parts[old_i + idx_shift] = remap_old_part(old_parts[old_i])
    end

    reachable = falses(interim_n)
    stack = Int[1]

    while !isempty(stack)
        i = pop!(stack)
        reachable[i] && continue
        reachable[i] = true

        for arg in interim_parts[i].args
            if arg isa FlatIdx
                (arg.i < 1 || arg.i > interim_n) && throw(BoundsError(interim_parts, arg.i))
                push!(stack, arg.i)
            end
        end
    end

    old_to_new = zeros(Int, interim_n)
    new_parts = Vector{ExcelExpr}()
    sizehint!(new_parts, count(reachable))
    for i in eachindex(interim_parts)
        reachable[i] || continue
        push!(new_parts, interim_parts[i])
        old_to_new[i] = length(new_parts)
    end

    for i in eachindex(new_parts)
        part = new_parts[i]
        has_flat_idx = false
        for arg in part.args
            if arg isa FlatIdx
                has_flat_idx = true
                break
            end
        end
        !has_flat_idx && continue

        new_args = copy(part.args)
        for arg_i in eachindex(new_args)
            arg = new_args[arg_i]
            if arg isa FlatIdx
                remapped_i = old_to_new[arg.i]
                remapped_i == 0 && throw(ArgumentError("Internal FlatIdx remapping failed for index $(arg.i)"))
                new_args[arg_i] = FlatIdx(remapped_i)
            end
        end
        new_parts[i] = ExcelExpr(part.head, new_args)
    end

    resize!(expr.parts, length(new_parts))
    for i in eachindex(new_parts)
        expr.parts[i] = new_parts[i]
    end
    expr
end

function change_expr_part(expr::FlatExpr, part_i::Int, new_value::ExcelExpr)
    change_expr_part!(deepcopy(expr), part_i, new_value)
end