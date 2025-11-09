
function get_expr_dependencies(expr, key_values::Dict)
    return []
end

function get_expr_dependencies(expr::FlatExpr, key_values::Dict)
    deps = Vector{CellDependency}()
    handled = Set{Int}()
    # @display expr
    for (i, part) in enumerate(expr.parts)
        i in handled && continue
        # @show part

        @match part begin
            ExcelExpr(:cell_ref, [cell, sheet]) => push!(deps, CellDependency(sheet, cell))
            # ExcelExpr(:sheet_ref, (sheet_name, ref)) => get_expr_dependencies(ref, key_values)
            ExcelExpr(:named_range, [name]) => append!(deps, get_expr_dependencies(key_values[name], key_values))
            ExcelExpr(:structured_reference, args) => begin
                throw("Don't know how to handle structured references!")
            end
            ExcelExpr(:range, [FlatIdx(lhs_i), FlatIdx(rhs_i)]) => begin
                lhs_expr = expr.parts[lhs_i]
                rhs_expr = expr.parts[rhs_i]
                if !((lhs_expr.head == :cell_ref) && (rhs_expr.head == :cell_ref))
                    # throw("Don't know how to get dependencies because of a range expression without cell refs. i = $i. lhs_expr = $(lhs_expr.head), rhs_expr = $(rhs_expr.head)")
                    continue
                end
                sheet = lhs_expr.args[2]
                if (sheet != rhs_expr.args[2])
                    throw("Don't know how to get dependencies because of a range expression that doesn't share a cell. i = $i")
                end

                lhs = lhs_expr.args[1]
                rhs = rhs_expr.args[1]

                start_col, start_row = parse_cell(lhs)
                end_col, end_row = parse_cell(rhs)

                @assert end_row >= start_row
                @assert end_col >= start_col

                for col in start_col:end_col, r in start_row:end_row
                    push!(deps, CellDependency(sheet, index_to_cellname(col, r)))
                end

                push!(handled, lhs_i)
                push!(handled, rhs_i)
            end
            _ => continue
        end
    end

    deps
end


function get_expr_dependencies(expr::ExcelExpr, key_values::Dict)::Vector{CellDependency}
    @match expr begin
        ExcelExpr(:cell_ref, [cell, sheet]) => [CellDependency(sheet, cell)]
        # ExcelExpr(:sheet_ref, (sheet_name, ref)) => get_expr_dependencies(ref, key_values)
        ExcelExpr(:named_range, [name]) => begin
            value = key_values[name]
            if value isa ExcelExpr
                get_expr_dependencies(value, key_values)
            else
                Vector{CellDependency}()
            end
        end
        ExcelExpr(:range, [ExcelExpr(:cell_ref, [lhs, sheet]), ExcelExpr(:cell_ref, [rhs, sheet])]) => begin
            start_col, start_row = parse_cell(lhs)
            end_col, end_row = parse_cell(rhs)

            @assert end_row >= start_row
            @assert end_col >= start_col
            [CellDependency(sheet, index_to_cellname(c, r)) for c ∈ start_col:end_col for r ∈ start_row:end_row]
        end
        # TODO: Handle?
        ExcelExpr(:range, [lhs, rhs]) => throw("Don't know how to get dependencies for $(expr)")
        ExcelExpr(op, args) => begin
            res = Vector{CellDependency}()
            for a in args
                if a isa ExcelExpr
                    append!(res, get_expr_dependencies(a, key_values))
                end
            end
            res
        end
    end
end

insert_table_refs(expr, tables) = expr

function cell_fixed_col(cell::T) where {T <: AbstractString}
    cell[1] == '$'
end

function cell_fixed_row(cell::T) where {T <: AbstractString}
    first_num = findfirst(isdigit, cell)

    cell[first_num-1] == '$'
end

function insert_table_refs(expr::FlatExpr, tables)
    # @info "In primary insert_table_refs" expr

    new_expr = copy(expr)

    deps = Vector{CellDependency}()
    handled = Set{Int}()
    to_remove = Set{Int}()

    for (i, part) in enumerate(new_expr.parts)
        i in handled && continue
        @match part begin
            ExcelExpr(:cell_ref, [cell, sheet::AbstractString]) => begin
                c, r = parse_cell(cell)
                for table in tables
                    if (sheet == table.sheet_name
                        &&
                        (r, c) in table)
                        row_idx = r - startrow(table) + 1
                        col_idx = c - startcol(table) + 1
                        fixed_row = cell_fixed_row(cell)
                        fixed_col = cell_fixed_col(cell)
                        # fixed_row = offset(expr, 1, 0) == expr
                        # fixed_col = offset(expr, 0, 1) == expr

                        # println("Inserting table ref into flatexpr at idx $i")
                        # @show cell c r table row_idx col_idx
                        new_expr.parts[i] = ExcelExpr(:table_ref, Any[table, row_idx, col_idx, (fixed_row, fixed_row), (fixed_col, fixed_col)])
                    end
                end
            end
            # ExcelExpr(:sheet_ref, (sheet_name, ref)) => ExcelExpr(:sheet_ref, sheet_name, insert_table_refs(ref, tables))
            # ExcelExpr(:named_range, (name,)) => convert_to_broadcasted(get_key_value(ctx, name), ctx, row_offset, col_offset)

            ExcelExpr(:range, [FlatIdx(lhs_i), FlatIdx(rhs_i)]) => begin
                lhs_expr = new_expr.parts[lhs_i]
                rhs_expr = new_expr.parts[rhs_i]
                if !((lhs_expr.head == :cell_ref) && (rhs_expr.head == :cell_ref))
                    continue
                    # throw("Don't know how to get dependencies for $(expr)")
                end
                sheet = lhs_expr.args[2]
                if (sheet != rhs_expr.args[2])
                    throw("Don't know how to get dependencies for $(expr)")
                end

                lhs = lhs_expr.args[1]
                rhs = rhs_expr.args[1]

                start_c, start_r = parse_cell(lhs)
                stop_c, stop_r = parse_cell(rhs)

                for table in tables
                    if (sheet == table.sheet_name
                        && (start_r, start_c) in table
                        && (stop_r, stop_c) in table)
                        col_idx = (start_c:stop_c) .- startcol(table) .+ 1
                        row_start_idx = start_r - startrow(table) + 1
                        row_stop_idx = stop_r - startrow(table) + 1

                        # is_fixed_row(expr) = offset(expr, 1, 0) == expr
                        # is_fixed_col(expr) = offset(expr, 0, 1) == expr
                        fixed_row = (cell_fixed_row(lhs), cell_fixed_row(rhs))
                        fixed_col = (cell_fixed_col(lhs), cell_fixed_col(rhs))
                        # fixed_row = tuple(is_fixed_row.(part.args)...)
                        # fixed_col = tuple(is_fixed_col.(part.args)...)

                        new_expr.parts[i] = ExcelExpr(:table_ref, Any[table, row_start_idx:row_stop_idx, col_idx, fixed_row, fixed_col])

                        # push!(handled, lhs_i)
                        # push!(handled, rhs_i)
                        push!(to_remove, lhs_i)
                        push!(to_remove, rhs_i)

                        break
                    end
                end
                push!(handled, lhs_i)
                push!(handled, rhs_i)
            end
            _ => continue
        end
    end

    for (part_i, part) in enumerate(new_expr.parts)
        new_args = copy(part.args)
        for (i, arg) in enumerate(new_args)
            if arg isa FlatIdx
                offset = 0
                for t in to_remove
                    if t < arg.i
                        offset += 1
                    end
                end

                new_i = arg.i - offset
                new_args[i] = FlatIdx(new_i)
            end

        end
        new_expr.parts[part_i] = ExcelExpr(part.head, new_args)
        # part.args = new_args
    end
    deleteat!(new_expr.parts, sort(collect(to_remove)))
    new_expr
end
function insert_table_refs(expr::ExcelExpr, tables)
    # @info "In primary insert_table_refs" expr
    @match expr begin
        ExcelExpr(:cell_ref, [cell, sheet]) => begin
            c, r = parse_cell(cell)
            for table in tables
                if (sheet == table.sheet_name
                    && (r, c) in table)
                    # && c >= startcol(table)
                    # && c <= endcol(table)
                    # && r >= startrow(table)
                    # && r <= endrow(table))
                    row_idx = r - startrow(table) + 1
                    col_idx = c - startcol(table) + 1
                    fixed_row = offset(expr, 1, 0) == expr
                    fixed_col = offset(expr, 0, 1) == expr

                    return ExcelExpr(:table_ref, table, row_idx, col_idx, (fixed_row, fixed_row), (fixed_col, fixed_col))
                end
            end

            expr
            # range_start = CellDependency(ctx.current_sheet, cell)
            # range_stop = offset(range_start, row_offset, col_offset)
            # # offset 
            # # [CellDependency(ctx.current_sheet, cell)]
            # ExcelExpr(:range, (ExcelExpr(:cell_ref, (range_start.cell,)), ExcelExpr(:cell_ref, (range_stop.cell,))))
        end
        # ExcelExpr(:sheet_ref, (sheet_name, ref)) => ExcelExpr(:sheet_ref, sheet_name, insert_table_refs(ref, tables))
        # ExcelExpr(:named_range, (name,)) => convert_to_broadcasted(get_key_value(ctx, name), ctx, row_offset, col_offset)
        ExcelExpr(:range, [ExcelExpr(:cell_ref, [lhs, sheet]), ExcelExpr(:cell_ref, [rhs, sheet])]) => begin
            start_c, start_r = parse_cell(lhs)
            stop_c, stop_r = parse_cell(rhs)
            # @info "found range" start_c start_r stop_c stop_r

            for table in tables
                # @info "Checking table" ctx.current_sheet == table.sheet_name start_c >= startcol(table) stop_c <= endcol(table) start_r >= startrow(table) stop_r <= endrow(table)
                if (sheet == table.sheet_name
                    && (start_r, start_c) in table
                    && (stop_r, stop_c) in table)
                    # && start_c >= startcol(table)
                    # && stop_c <= endcol(table)
                    # && start_r >= startrow(table)
                    # && stop_r <= endrow(table))
                    # @info "Found a matching table"
                    # col_names = [column_name(table, c - startcol(table) + 1) for c in start_c:stop_c]
                    # col_idx = [c - startcol(table) + 1 for c in start_c:stop_c]
                    col_idx = (start_c:stop_c) .- startcol(table) .+ 1
                    row_start_idx = start_r - startrow(table) + 1
                    row_stop_idx = stop_r - startrow(table) + 1

                    is_fixed_row(expr) = offset(expr, 1, 0) == expr
                    is_fixed_col(expr) = offset(expr, 0, 1) == expr
                    fixed_row = tuple(is_fixed_row.(expr.args)...)
                    fixed_col = tuple(is_fixed_col.(expr.args)...)
                    # fixed_row = offset(expr.args[1], 1, 0) == expr.args[1]
                    # fixed_col = offset(expr, 0, 1) == expr

                    return ExcelExpr(:table_ref, table, row_start_idx:row_stop_idx, col_idx, fixed_row, fixed_col)
                end
            end
            expr
        end
        ExcelExpr(:range, [lhs, rhs]) => throw("Don't know how to get dependencies for $(expr)")
        ExcelExpr(op, args) => begin
            new_args = similar(args)
            map!(a -> insert_table_refs(a, tables), new_args, args)
            # mapped = map(a -> insert_table_refs(a, tables), args)
            # ExcelExpr(op, mapped...)
            ExcelExpr(op, new_args)
            # reduce(vcat, mapped)
        end
    end
end

iscontiguous(values::Vector{T}) where {T <: Integer} = values == minimum(values):maximum(values)
function get_contiguous_runs(values::Vector{T}) where {T <: Integer}
    if isempty(values)
        return values
    end
    if length(values) == 1
        return [values]
    end

    result = Vector{Vector{T}}()
    current = Vector{T}()
    last_value = values[1]
    push!(current, last_value)

    for next_value in values[2:end]
        if next_value == last_value + 1
            push!(current, next_value)
        else
            push!(result, current)
            current = [next_value]
        end
        last_value = next_value
    end
    push!(result, current)

    result
end

contains_if(expr) = false
function contains_if(expr::ExcelExpr)
    @match expr begin
        ExcelExpr(:call, ["IF", args...]) => begin
            true
        end
        ExcelExpr(op, args) => begin
            any(contains_if(a) for a in args)
        end
    end
end

function contains_if(expr::FlatExpr)
    for part in expr.parts
        if part.head == :call && part.args[1] == "IF"
            return true
        end
    end
    return false
end

convert_to_broadcasted(expr, row_offset, col_offset) = expr

function convert_to_broadcasted(expr::ExcelExpr, row_offset, col_offset)
    @match expr begin
        ExcelExpr(:cell_ref, [cell, sheet]) => begin
            range_start = CellDependency(sheet, cell)
            stop_cell = offset(expr, row_offset, col_offset).args[1]
            range_stop = CellDependency(sheet, stop_cell)
            # range_stop = offset(range_start, row_offset, col_offset)
            # offset 
            # [CellDependency(ctx.current_sheet, cell)]
            if range_start != range_stop
                ExcelExpr(:range, ExcelExpr(:cell_ref, range_start.cell, sheet), ExcelExpr(:cell_ref, range_stop.cell, sheet))
            else
                expr
            end
        end
        # ExcelExpr(:sheet_ref, (sheet_name, ref)) => ExcelExpr(:sheet_ref, sheet_name, convert_to_broadcasted(ref, sheet_name, row_offset, col_offset))
        # ExcelExpr(:named_range, (name,)) => convert_to_broadcasted(get_key_value(ctx, name), ctx, row_offset, col_offset)
        ExcelExpr(:range, [ExcelExpr(:cell_ref, [lhs, sheet]), ExcelExpr(:cell_ref, [rhs, sheet])]) => begin
            function is_fixed(cell)
                offset_cell_parse_rgx = r"([$]?[A-Z]+)([$]?[0-9]+)"
                cell_match = match(offset_cell_parse_rgx, cell)
                @assert cell_match.match == cell "Cell didn't parse properly"
                cell_match[1][1] == '$' && cell_match[2][1] == '$'
            end
            if is_fixed(lhs) && is_fixed(rhs)
                # I think this isn't strictly correct, but it might be good enough
                mapped = map(a -> convert_to_broadcasted(a, row_offset, col_offset), expr.args)
                sub_expr = ExcelExpr(expr.head, mapped...)

                ExcelExpr(:broadcast_protect, sub_expr)
            else
                throw("Converting ranged to broadcasted is complicated $(expr)")
            end
        end

        ExcelExpr(:table_ref, [table, row_idx, col_idx, fixed_row, fixed_col]) => begin
            if fixed_row == (true, true) && fixed_col == (true, true)
                return ExcelExpr(:broadcast_protect, expr)
            end

            row_idx = @match fixed_row begin
                (true, true) => row_idx
                (false, false) => begin
                    if length(row_idx) == 1
                        row_idx[1]:(row_idx[1]+row_offset)
                    else
                        throw("Broadcasting table row range ref is complicated")
                    end
                end
                (true, false) => throw("Broadcasting partially fixed table refs is complicated")
                (false, true) => throw("Broadcasting partially fixed table refs is complicated")
            end
            col_idx = @match fixed_col begin
                (true, true) => col_idx
                (false, false) => begin
                    if length(col_idx) == 1
                        col_idx[1]:(col_idx[1]+col_offset)
                    else
                        throw("Broadcasting table col range ref is complicated")
                    end
                end
                (true, false) => throw("Broadcasting partially fixed table refs is complicated")
                (false, true) => throw("Broadcasting partially fixed table refs is complicated")
            end

            ExcelExpr(:table_ref, table, row_idx, col_idx, fixed_row, fixed_col)
        end

        # ExcelExpr(:call, ("IF", args...)) => begin
        #     throw("Broadcasting if doesn't work!")
        # end
        ExcelExpr(:range, [lhs, rhs]) => throw("Don't know how to get dependencies for $(expr)")
        ExcelExpr(op, args) => begin
            mapped = map(a -> convert_to_broadcasted(a, row_offset, col_offset), args)
            ExcelExpr(op, mapped...)
            # reduce(vcat, mapped)
        end
    end
end

function convert_to_broadcasted(expr::FlatExpr, row_offset, col_offset)
    new_expr = copy(expr)

    deps = Vector{CellDependency}()
    handled = Set{Int}()
    for (i, part) in enumerate(new_expr.parts)
        i in handled && continue
        @match part begin
            ExcelExpr(:cell_ref, [cell, sheet]) => begin

                range_start = CellDependency(sheet, cell)
                stop_cell = offset_cell_str(cell, row_offset, col_offset)
                range_stop = CellDependency(sheet, stop_cell)

                if range_start != range_stop
                    push!(new_expr.parts, ExcelExpr(:cell_ref, range_start.cell, sheet))
                    i1 = length(new_expr.parts)
                    push!(new_expr.parts, ExcelExpr(:cell_ref, range_stop.cell, sheet))
                    i2 = length(new_expr.parts)
                    push!(handled, i1)
                    push!(handled, i2)

                    new_expr.parts[i] = ExcelExpr(:range, FlatIdx(i1), FlatIdx(i2))

                end
            end

            ExcelExpr(:range, [FlatIdx(lhs_i), FlatIdx(rhs_i)]) => begin
                lhs_expr = new_expr.parts[lhs_i]
                rhs_expr = new_expr.parts[rhs_i]
                if !((lhs_expr.head == :cell_ref) && (rhs_expr.head == :cell_ref))
                    throw("Don't know how to get dependencies for $(expr)")
                end
                sheet = lhs_expr.args[2]
                if (sheet != rhs_expr.args[2])
                    throw("Don't know how to get dependencies for $(expr)")
                end

                lhs = lhs_expr.args[1]
                rhs = rhs_expr.args[1]

                start_c, start_r = parse_cell(lhs)
                stop_c, stop_r = parse_cell(rhs)

                function is_fixed(cell)
                    offset_cell_parse_rgx = r"([$]?[A-Z]+)([$]?[0-9]+)"
                    cell_match = match(offset_cell_parse_rgx, cell)
                    @assert cell_match.match == cell "Cell didn't parse properly"
                    cell_match[1][1] == '$' && cell_match[2][1] == '$'
                end
                if is_fixed(lhs) && is_fixed(rhs)
                    push!(handled, lhs_i)
                    push!(handled, rhs_i)
                    push!(new_expr.parts, part)
                    idx = length(new_expr.parts)
                    push!(handled, idx)
                    new_expr.parts[i] = ExcelExpr(:broadcast_protect, FlatIdx(idx))
                else
                    throw("Converting ranged to broadcasted is complicated")
                end
            end

            ExcelExpr(:table_ref, [table, row_idx, col_idx, fixed_row, fixed_col]) => begin
                if fixed_row == (true, true) && fixed_col == (true, true)
                    push!(new_expr.parts, part)
                    idx = length(new_expr.parts)
                    push!(handled, idx)
                    new_expr.parts[i] = ExcelExpr(:broadcast_protect, FlatIdx(idx))
                end

                row_idx = @match fixed_row begin
                    (true, true) => row_idx
                    (false, false) => begin
                        if length(row_idx) == 1
                            row_idx[1]:(row_idx[1]+row_offset)
                        else
                            throw("Broadcasting table row range ref is complicated")
                        end
                    end
                    (true, false) => throw("Broadcasting partially fixed table refs is complicated")
                    (false, true) => throw("Broadcasting partially fixed table refs is complicated")
                end
                col_idx = @match fixed_col begin
                    (true, true) => col_idx
                    (false, false) => begin
                        if length(col_idx) == 1
                            col_idx[1]:(col_idx[1]+col_offset)
                        else
                            throw("Broadcasting table col range ref is complicated")
                        end
                    end
                    (true, false) => throw("Broadcasting partially fixed table refs is complicated")
                    (false, true) => throw("Broadcasting partially fixed table refs is complicated")
                end

                new_expr.parts[i] = ExcelExpr(:table_ref, table, row_idx, col_idx, fixed_row, fixed_col)
            end
            _ => continue
        end
    end
    new_expr
end

function functionalize!(expr, previous_params::Vector{ExcelExpr})
    expr
end

function functionalize_at!(exprs::Vector{Any}, index::Int, previous_params::Vector{ExcelExpr})
    expr = exprs[index]
    if expr isa ExcelExpr
        if expr.head in (:cell_ref, :sheet_ref, :named_range, :range, :table_ref)
            push!(previous_params, expr)
            exprs[index] = ExcelExpr(:func_param, length(previous_params))
        else
            for i in eachindex(expr.args)
                functionalize_at!(expr.args, i, previous_params)
            end

        end
    end
end

function functionalize!(expr::ExcelExpr, previous_params::Vector{ExcelExpr})
    if expr.head in (:cell_ref, :sheet_ref, :named_range, :range, :table_ref)
        push!(previous_params, expr)
        return ExcelExpr(:func_param, length(previous_params))
    end

    for i in eachindex(expr.args)
        functionalize_at!(expr.args, i, previous_params)
    end

    expr

    # @match expr begin
    #     ExcelExpr(:cell_ref, [cell, sheet]) => begin
    #         push!(previous_params, expr)
    #         ExcelExpr(:func_param, length(previous_params))
    #     end
    #     ExcelExpr(:sheet_ref, [sheet_name, ref]) => begin
    #         push!(previous_params, expr)
    #         ExcelExpr(:func_param, length(previous_params))
    #     end
    #     ExcelExpr(:named_range, [name,]) => begin 
    #         push!(previous_params, expr)
    #         ExcelExpr(:func_param, length(previous_params))
    #     end
    #     ExcelExpr(:range, args) => begin
    #         push!(previous_params, expr)
    #         ExcelExpr(:func_param, length(previous_params))
    #     end
    #     ExcelExpr(:table_ref, args) => begin
    #         push!(previous_params, expr)
    #         ExcelExpr(:func_param, length(previous_params))
    #     end

    #     ExcelExpr(op, args) => begin
    #         mapped_exprs = similar(args)
    #         # mapped_exprs = []
    #         # for a in args
    #         for i in eachindex(args)
    #             if args[i] isa ExcelExpr
    #                 e, previous_params = functionalize(args[i], previous_params)
    #                 mapped_exprs[i] = e
    #             else
    #                 mapped_exprs[i] = args[i]
    #             end
    #             # push!(mapped_exprs, e)
    #         end
    #         # mapped = map(a -> functionalize(a, ctx, previous_params), args)
    #         # (ExcelExpr(op, mapped_exprs...), previous_params)
    #         (ExcelExpr(op, mapped_exprs), previous_params)
    #         # reduce(vcat, mapped)
    #     end
    # end
end

function functionalize(expr, previous_params::Vector{ExcelExpr})
    expr
end
function functionalize(expr::ExcelExpr, previous_params::Vector{ExcelExpr})
    if expr.head in (:cell_ref, :sheet_ref, :named_range, :range, :table_ref)
        # (ExcelExpr(:func_param, Any[length(previous_params) + 1]), [previous_params... expr])
        push!(previous_params, expr)
        ExcelExpr(:func_param, length(previous_params))
    else
        op = expr.head
        args = expr.args
        mapped_exprs = similar(args)
        for i in eachindex(args)
            if typeof(args[i]) == ExcelExpr
                mapped_exprs[i] = functionalize(args[i], previous_params)
            else
                mapped_exprs[i] = args[i]
            end
        end
        ExcelExpr(op, mapped_exprs)
    end
end

function functionalize(expr, previous_params::Matrix{ExcelExpr})
    (expr, previous_params)
end

function functionalize(expr::FlatExpr, previous_params::Vector{ExcelExpr})
    to_remove = Vector{Int}()
    for i in eachindex(expr.parts)
        part = expr.parts[i]
        if part.head in (:cell_ref, :sheet_ref, :named_range, :range, :table_ref)
            for arg in part.args
                if arg isa FlatIdx
                    push!(to_remove, arg.i)
                end
            end
        end
    end
    new_parts = Vector{ExcelExpr}(undef, length(expr.parts) - length(to_remove))
    add_i = 1

    function add_param(part::ExcelExpr)
        param_expr = convert_to_expr(part, expr)
        param_idx = findfirst(isequal(param_expr), previous_params)
        if isnothing(param_idx)
            push!(previous_params, param_expr)
            ExcelExpr(:func_param, Any[length(previous_params)])
        else
            ExcelExpr(:func_param, Any[param_idx])
        end
    end
    # new_expr = copy(expr)
    # to_remove = Vector{Int}()
    for i in eachindex(expr.parts)
        i in to_remove && continue

        part = expr.parts[i]
        if part.head in (:cell_ref, :sheet_ref, :named_range, :range, :table_ref)
            new_parts[add_i] = add_param(part)
        else
            # new_args = similar(part.args)
            if all(v -> !(v isa FlatIdx), part.args)
                new_parts[add_i] = part
                # push!(new_parts, part)
            else
                new_args = copy(part.args)
                # for (i, arg) in enumerate(part.args)
                for (i, arg) in enumerate(new_args)
                    if arg isa FlatIdx
                        offset = 0
                        for t in to_remove
                            if t < arg.i
                                offset += 1
                            end
                        end

                        new_i = arg.i - offset
                        new_args[i] = FlatIdx(new_i)
                    end
                end
                new_parts[add_i] = ExcelExpr(part.head, new_args)
                # push!(new_parts, ExcelExpr(part.head, new_args))
            end
        end
        add_i += 1
    end

    # new_parts = Vector{ExcelExpr}()
    # for (i, part) in enumerate(new_expr.parts)
    #     i in to_remove && continue

    #     new_args = similar(part.args)

    #     for (i, arg) in enumerate(part.args)
    #         if arg isa FlatIdx
    #             new_i = arg.i - sum(to_remove .< arg.i)
    #             new_args[i] = FlatIdx(new_i)
    #         else
    #             new_args[i] = arg
    #         end
    #     end
    #     push!(new_parts, ExcelExpr(part.head, new_args))
    # end
    # FlatExpr(filter(e -> e.head != :missing, new_expr.parts))
    FlatExpr(new_parts)
    # new_expr
end

function functionalize(expr)
    params = Vector{ExcelExpr}()
    new_expr = functionalize(expr, params)
    (new_expr, permutedims(params))

    # params = Vector{ExcelExpr}()

    # new_expr = functionalize!(deepcopy(expr), params)
    # (new_expr, permutedims(params))
end