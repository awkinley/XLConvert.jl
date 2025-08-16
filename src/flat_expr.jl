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
