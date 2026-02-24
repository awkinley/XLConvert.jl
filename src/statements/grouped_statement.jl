mutable struct GroupedStatement <: AbstractStatement
    sub_statements::Vector{AbstractStatement}
end

get_cell_deps(stmt::GroupedStatement) = reduce(vcat, get_cell_deps.(stmt.sub_statements))
get_set_cells(stmt::GroupedStatement) = reduce(vcat, get_set_cells.(stmt.sub_statements))
function apply_expr_transform!(stmt::GroupedStatement, transform)
    for s in stmt.sub_statements
        apply_expr_transform!(s, transform)
    end
end

function to_string(exporter, statement::GroupedStatement)

    children = if length(statement.sub_statements) > 5
        join(to_string.((exporter,), statement.sub_statements[begin:5]), ", ") * "..."
    else
        join(to_string.((exporter,), statement.sub_statements), ", ")
    end
    # children = join(to_string.((exporter,), statement.sub_statements), ", ")
    "GroupedStatement($children)"
end


function offset_with_fixed(idx::Int, fixed::Tuple{Bool, Bool}, offset::Int)
    @match fixed begin
        (true, true) => idx
        (false, false) => idx + offset
        (true, false) => idx:(idx+offset)
        (false, true) => (idx+offset):idx
    end
end

function offset_with_fixed(idx::UnitRange{Int}, fixed::Tuple{Bool, Bool}, offset::Int)
    (first(idx)+(!fixed[1])*offset):(last(idx)+(!fixed[2])*offset)
end

function can_be_for_looped(expressions, row_offset, col_offset)
    fixed_row = false
    fixed_col = false
    start_row_idx = nothing
    start_col_idx = nothing
    # @show expressions
    first_val = expressions[1]
    if first_val.head == :table_ref
        start_row_idx = first_val.args[2]
        start_col_idx = first_val.args[3]
        fixed_row = first_val.args[4]
        fixed_col = first_val.args[5]
    else
        # println("Failed because a changing param wasn't a table ref")
        return CantLoop("A changing param wasn't a table ref, $(first_val)")
        # return false
    end

    for j in eachindex(expressions)[2:end]
        val = expressions[j]
        if val.head == :table_ref
            row_idx = val.args[2]
            col_idx = val.args[3]

            # good_row = row_idx == start_row_idx .+ ((j - 1) * row_offset)
            # good_col = col_idx == start_col_idx .+ ((j - 1) * col_offset)
            offset_row = offset_with_fixed(start_row_idx, fixed_row, ((j - 1) * row_offset))
            good_row = row_idx == offset_row
            good_col = col_idx == offset_with_fixed(start_col_idx, fixed_col, ((j - 1) * col_offset))
            if !(good_row && good_col)
                # println("Failed because an expression wasn't offset correctly")

                return CantLoop("An expression wasn't offset correctly, $(good_row), $(good_col), $(row_idx), $(offset_row)")
                # return false
            end
        else
            # println("Failed because a changing param wasn't a table ref")
            return CantLoop("a changing param wasn't a table ref, $(val)")
            # return false
        end
    end

    true
end

function make_loop_idx_str(::JuliaExporter, base_idx::Int, offset, fixed)
    if offset == 0
        "$base_idx"
    elseif offset == 1
        "$base_idx + i"
    else
        "$base_idx + (i * $offset)"
    end
end

function make_loop_idx_str(::PythonExporter, base_idx::Int, offset, fixed)
    a = base_idx
    if offset == 0
        "$a"
    elseif offset == 1
        "$a + i"
    else
        "$a + (i * $offset)"
    end
end

function make_loop_idx_str(::JuliaExporter, base_idx::UnitRange{Int}, offset, fixed)
    if offset == 0
        "$base_idx"
    elseif fixed == (false, false)
        if offset == 1
            "($base_idx) .+ i"
        else
            "($base_idx) .+ (i * $offset)"
        end
    else
        left = string(first(base_idx))
        if !fixed[1]
            if offset == 1
                left *= " + i"
            else
                left *= " + (i * $offset)"
            end
        end

        right = string(last(base_idx))
        if !fixed[2]
            if offset == 1
                right *= " + i"
            else
                right *= " + (i * $offset)"
            end
        end
        "($left):($right)"
    end
end

function make_loop_idx_str(e::PythonExporter, base_idx::UnitRange{Int}, offset, fixed)
    a = first(base_idx)
    b = last(base_idx)
    if length(base_idx) == 1 && fixed == (false, false)
        return make_loop_idx_str(e, a, offset, fixed)
    end

    if offset == 0
        "$a:$(b + 1)"
        # "$base_idx"
    elseif fixed == (false, false)
        if offset == 1
            "($a + i):($b + i)"
        else
            "($a + i * $offset):($b + i * $offset)"
        end
    else
        left = string(a)
        if !fixed[1]
            if offset == 1
                left *= " + i"
            else
                left *= " + (i * $offset)"
            end
        end

        right = string(last(base_idx) + 1)
        if !fixed[2]
            if offset == 1
                right *= " + i"
            else
                right *= " + (i * $offset)"
            end
        end
        "($left):($right)"
    end
end

struct CustomFuncParamHandler
    value_getter::Any
end

function handle(handler::CustomFuncParamHandler, expr, exporter::JuliaExporter, ctx)
    @match expr begin
        ExcelExpr(:func_param, [param_num]) => begin
            handler.value_getter(param_num, exporter, ctx)
        end
        ExcelExpr(:func_param, [param_num, type]) => begin
            handler.value_getter(param_num, exporter, ctx)
        end
        _ => missing
    end
end

function handle(handler::CustomFuncParamHandler, expr, exporter::PythonExporter, ctx)
    @match expr begin
        ExcelExpr(:func_param, [param_num]) => begin
            handler.value_getter(param_num, exporter, ctx)
        end
        ExcelExpr(:func_param, [param_num, type]) => begin
            handler.value_getter(param_num, exporter, ctx)
        end
        _ => missing
    end
end

replace_func_params(expr, params_dict) = expr
function replace_func_params(expr::ExcelExpr, params_dict)
    @match expr begin
        ExcelExpr(:func_param, [param_num]) => get(params_dict, param_num, expr)
        # ExcelExpr(head, args) => ExcelExpr(head, map(e -> replace_func_params(e, params_dict), args)...)
        _ => expr
    end
end

function replace_func_params(expr::FlatExpr, params_dict)
    new_expr = FlatExpr(copy(expr.parts))
    for i in eachindex(expr.parts)
        part = expr.parts[i]
        new_expr.parts[i] = @match part begin
            ExcelExpr(:func_param, [param_num]) => get(params_dict, param_num, part)
            _ => part
        end

    end

    new_expr
end

struct CantLoop
    reason::String
end

function can_loop_stmts(statements::AbstractArray{AbstractStatement}, functionalized)
    if any(!(s isa TableStatement) for s in statements)
        # They need to all be TableStatements
        # println("Failed because not every statement was a table statement")
        return CantLoop("not every statement was a table statement")
    end

    if length(statements) == 0
        return CantLoop("No statements")
        # return false
    end

    all_params = []
    # func, params = functionalize(statements[1].rhs_expr, [])
    func, params = functionalized[1]
    push!(all_params, params)
    # params = []
    for (new_func, func_params) in functionalized[2:end]
        # for s in statements[2:end]
        #     new_func, func_params = functionalize(s.rhs_expr, [])
        # if func != new_func
        if !isequal(func, new_func)
            # println("Failed because not all the functions are the same")
            return CantLoop("Not all the functions are the same")
            # return false
        end

        push!(all_params, func_params)
        # append!(params, func_params)

    end

    # funcs_and_params = [functionalize(s.rhs_expr, []) for s in statements]
    # funcs = [a[1] for a in funcs_and_params]
    # # @show funcs
    # unique_funcs = unique(funcs)
    # if length(unique_funcs) != 1
    #     # Different functions
    #     println("Failed because not all the functions are the same")
    #     return false
    # end

    # params = reduce(vcat, [a[2] for a in funcs_and_params])
    params = reduce(vcat, all_params)
    param_sets = unique.(eachcol(params))
    fixed_params = findall(length.(param_sets) .== 1)
    changing_params = findall(length.(param_sets) .!= 1)


    set_cells = get_set_cells.(statements)
    row_nums = map(s -> rownum(s[1]), set_cells)
    col_nums = map(s -> colnum(s[1]), set_cells)
    drow = diff(row_nums)
    dcol = diff(col_nums)
    if !all(drow .== drow[1]) || !all(dcol .== dcol[1])
        # Non-constant offset
        # println("Failed because there is not a consistent offset")
        return CantLoop("There is not a consistent offset")
        # return false
    end
    row_offset = drow[1]
    col_offset = dcol[1]

    # if !all(can_be_for_looped.(eachcol(params[:, changing_params]), row_offset, col_offset))
    if !all(x -> can_be_for_looped(x, row_offset, col_offset) == true, eachcol(params[:, changing_params]))
        # println("Failed because a changing parameter couldn't be for looped")
        reasons = map(x -> can_be_for_looped(x, row_offset, col_offset), eachcol(params[:, changing_params]))
        return CantLoop("A changing parameter couldn't be for looped, $(reasons)")
        # return false
    end

    lhs_exprs = map(s -> s.lhs_expr, statements)

    if can_be_for_looped(lhs_exprs, row_offset, col_offset) != true
        # println("Failed because a lhs couldn't be for looped")
        return CantLoop("A changing lhs couldn't be for looped")
        # return false
    end

    true
end


function export_looped(exporter::JuliaExporter, wb::ExcelWorkbook, statements)
    funcs_and_params = [functionalize(s.rhs_expr) for s in statements]
    # @show funcs_and_params
    func = funcs_and_params[1][1]
    # funcs = [a[1] for a in funcs_and_params]
    # unique_funcs = unique(funcs)

    params = reduce(vcat, [a[2] for a in funcs_and_params])
    # @show size(params)
    # @show params
    param_sets = unique.(eachcol(params))
    fixed_params = findall(length.(param_sets) .== 1)
    changing_params = findall(length.(param_sets) .!= 1)


    set_cells = get_set_cells.(statements)
    row_nums = map(s -> rownum(s[1]), set_cells)
    col_nums = map(s -> colnum(s[1]), set_cells)
    drow = diff(row_nums)
    dcol = diff(col_nums)
    row_offset = drow[1]
    col_offset = dcol[1]
    lhs_exprs = map(s -> s.lhs_expr, statements)

    lhs_expr = lhs_exprs[1]

    table, row_idx, col_idx, row_fixed, col_fixed = lhs_expr.args

    row_str = make_loop_idx_str(exporter, row_idx, row_offset, row_fixed)
    col_str = make_loop_idx_str(exporter, col_idx, col_offset, col_fixed)
    lhs_str = "$(getname(table))[$row_str, $col_str]"
    # println("\t$str")

    function get_param_str(param_num, exporter, ctx)
        if param_num in changing_params
            param_expr = params[1, param_num]
            table, row_idx, col_idx, row_fixed, col_fixed = param_expr.args
            row_str = make_loop_idx_str(exporter, row_idx, row_offset, row_fixed)
            col_str = make_loop_idx_str(exporter, col_idx, col_offset, col_fixed)
            # @show col_idx col_offset, col_fixed
            "$(getname(table))[$row_str, $col_str]"
        else
            throw("Tried to get param_str for param_num $param_num, but it wasn't a changing param")
        end
    end

    fixed_params_dict = Dict(fixed_params .=> map(v -> v[1], param_sets[fixed_params]))
    # println("Inserting fixed params")
    # @show fixed_params_dict

    # println("Original expr")
    # show(stdout, "text/plain", statements[1].rhs_expr)
    # println("Before replacing func params")
    # show(stdout, "text/plain", func)
    rhs_expr = replace_func_params(func, fixed_params_dict)
    # println("After replacing func params")
    # show(stdout, "text/plain", func)

    typed_params = Dict{Int64, ExcelExpr}()
    # @show changing_params
    for param_num in changing_params
        all_exprs = params[:, param_num]
        # @show all_exprs
        param_type = reduce(union_types, map(e -> get_type(e, table.sheet_name, exporter.cell_types, exporter.named_values), all_exprs))
        typed_params[param_num] = ExcelExpr(:func_param, param_num, param_type)
    end
    # println("Before replacing func params")
    # show(stdout, "text/plain", rhs_expr)
    rhs_expr = replace_func_params(rhs_expr, typed_params)
    # println("After replacing func params")
    # show(stdout, "text/plain", rhs_expr)

    custom_handler = CustomFuncParamHandler(get_param_str)
    custom_exporter = JuliaExporter(exporter.wb, exporter.var_names, exporter.tables, exporter.named_values, [custom_handler, exporter.handlers...], exporter.cell_types)
    rhs_str = convert(custom_exporter, rhs_expr, table.sheet_name)
    """
    for i in 0:$(length(statements) - 1)
    \t$lhs_str = $rhs_str
    end
    """
end

part_to_workbook_range(expr::XLConvert.FlatExpr, i::FlatIdx) = part_to_workbook_range(expr, i.i)
function part_to_workbook_range(expr::XLConvert.FlatExpr, i::Int32)
    part = expr.parts[i]

    if @ismatch part ExcelExpr(:sheet_ref, [sheet, sub_i])
        return part_to_workbook_range(expr, sub_i)
    end

    if @ismatch part ExcelExpr(:broadcast_protect, [protected])
        part = protected
        # return part_to_workbook_range(expr, sub_i)
    end

    @match part begin
        ExcelExpr(:range, [FlatIdx(lhs_i), FlatIdx(rhs_i)]) => begin
            lhs_expr = expr.parts[lhs_i]
            rhs_expr = expr.parts[rhs_i]
            if !((lhs_expr.head == :cell_ref) && (rhs_expr.head == :cell_ref))
                # throw("Don't know how to get dependencies because of a range expression without cell refs. i = $i. lhs_expr = $(lhs_expr.head), rhs_expr = $(rhs_expr.head)")
                @show lhs_expr rhs_expr
                return nothing
            end
            sheet = lhs_expr.args[2]
            if (sheet != rhs_expr.args[2])
                throw("Don't know how to get dependencies because of a range expression that doesn't share a cell. i = $i")
            end

            lhs = lhs_expr.args[1]
            rhs = rhs_expr.args[1]

            WorkbookRegion(CellDependency(sheet, lhs), CellDependency(sheet, rhs))
        end
        ExcelExpr(:table_ref, [table, row_idx, col_idx, fixed_row, fixed_col]) => begin
            start_cell = CellDependency(table.sheet_name, startcol(table) + first(col_idx) - 1, startrow(table) + first(row_idx) - 1)
            end_cell = CellDependency(table.sheet_name, startcol(table) + last(col_idx) - 1, startrow(table) + last(row_idx) - 1)

            WorkbookRegion(start_cell, end_cell) 
        end
        _ => nothing
    end
end


function export_looped(exporter::PythonExporter, wb::ExcelWorkbook, statements)
    funcs_and_params = [functionalize(s.rhs_expr) for s in statements]
    func = funcs_and_params[1][1]

    params = reduce(vcat, [a[2] for a in funcs_and_params])
    param_sets = unique.(eachcol(params))
    fixed_params = findall(length.(param_sets) .== 1)
    changing_params = findall(length.(param_sets) .!= 1)


    set_cells = get_set_cells.(statements)
    row_nums = map(s -> rownum(s[1]), set_cells)
    col_nums = map(s -> colnum(s[1]), set_cells)
    drow = diff(row_nums)
    dcol = diff(col_nums)
    row_offset = drow[1]
    col_offset = dcol[1]
    lhs_exprs = map(s -> s.lhs_expr, statements)

    lhs_expr = lhs_exprs[1]

    table, row_idx, col_idx, row_fixed, col_fixed = lhs_expr.args

    row_str = make_loop_idx_str(exporter, row_idx .- 1, row_offset, row_fixed)
    col_str = make_loop_idx_str(exporter, col_idx .- 1, col_offset, col_fixed)
    lhs_str = "$(getname(table)).iloc[$row_str, $col_str]"
    # println("\t$str")

    function get_param_str(param_num, exporter, ctx)
        if param_num in changing_params
            param_expr = params[1, param_num]
            table, row_idx, col_idx, row_fixed, col_fixed = param_expr.args
            row_str = make_loop_idx_str(exporter, row_idx .- 1, row_offset, row_fixed)
            col_str = make_loop_idx_str(exporter, col_idx .- 1, col_offset, col_fixed)
            "$(getname(table)).iloc[$row_str, $col_str]"
        else
            throw("Tried to get param_str for param_num $param_num, but it wasn't a changing param")
        end
    end

    fixed_params_dict = Dict(fixed_params .=> map(v -> v[1], param_sets[fixed_params]))
    # println("Inserting fixed params")
    # @show fixed_params_dict

    # println("Original expr")
    # show(stdout, "text/plain", statements[1].rhs_expr)
    # println("Before replacing func params")
    # show(stdout, "text/plain", func)
    rhs_expr = replace_func_params(func, fixed_params_dict)
    # println("After replacing func params")
    # show(stdout, "text/plain", func)

    typed_params = Dict{Int64, ExcelExpr}()
    # @show changing_params
    for param_num in changing_params
        all_exprs = params[:, param_num]
        # @show all_exprs
        param_type = reduce(union_types, map(e -> get_type(e, table.sheet_name, exporter.cell_types, exporter.named_values), all_exprs))
        typed_params[param_num] = ExcelExpr(:func_param, param_num, param_type)
    end
    # println("Before replacing func params")
    # show(stdout, "text/plain", rhs_expr)
    rhs_expr = replace_func_params(rhs_expr, typed_params)
    # look_for_xlookups(rhs_expr)
    # xlookup_to_indexing!(rhs_expr)
    # println("After replacing func params")
    # show(stdout, "text/plain", rhs_expr)

    custom_handler = CustomFuncParamHandler(get_param_str)
    custom_exporter = PythonExporter(exporter.wb, exporter.var_names, exporter.tables, exporter.named_values, [custom_handler, exporter.handlers...], exporter.cell_types)
    rhs_str = convert(custom_exporter, rhs_expr, table.sheet_name)
    """
    for i in range($(length(statements))):
    \t$lhs_str = $rhs_str

    """
end

function get_loop_end(statements_in::AbstractArray{AbstractStatement}, functionalized_in, start_i)
    statements = @view statements_in[start_i:end]
    functionalized = @view functionalized_in[start_i:end]

    @assert length(statements) == length(functionalized)
    @assert length(statements) > 0

    if length(statements) == 1
        return start_i
    end

    if length(statements) == 2
        if can_loop_stmts(statements, functionalized) == true
            return start_i + 1
        else
            return start_i
        end
    end


    if any(isnothing, functionalized[1:2])
        return start_i
    end
    func, params = functionalized[1]
    func2, params2 = functionalized[2]

    # funcs_equal = func == func2
    # if func != func2
    # if ismissing(funcs_equal) || !funcs_equal
    if !isequal(func, func2)
        return start_i
    end

    equal_params = Vector{Int}()
    changing_params = Vector{Int}()
    for (i, (p1, p2)) in enumerate(zip(params, params2))
        if p1 == p2
            push!(equal_params, i)
        else
            if p1.head != :table_ref
                # return CantLoop("A changing param wasn't a table ref, $(first_val)")
                return start_i
            end
            push!(changing_params, i)
        end
    end


    cells1, cells2 = get_set_cells.(statements[1:2])
    row_offset = rownum(cells2[1]) - rownum(cells1[1])
    col_offset = colnum(cells2[1]) - colnum(cells1[1])

    start_lhs = statements[1].lhs_expr

    for i in 2:length(statements)

        lhs = statements[i].lhs_expr

        if lhs.head != :table_ref
            return start_i + i - 2
        end

        row_idx = lhs.args[2]
        col_idx = lhs.args[3]

        start_row_idx = start_lhs.args[2]
        start_col_idx = start_lhs.args[3]
        fixed_row = start_lhs.args[4]
        fixed_col = start_lhs.args[5]
        offset_row = offset_with_fixed(start_row_idx, fixed_row, ((i - 1) * row_offset))
        offset_col = offset_with_fixed(start_col_idx, fixed_col, ((i - 1) * col_offset))

        if (row_idx != offset_row) || (col_idx != offset_col)
            return start_i + i - 2
        end


        func_i, params_i = functionalized[i]

        # if func_i != func
        if !isequal(func_i, func)
            return start_i + i - 2
        end

        for p_i in equal_params
            if params_i[p_i] != params[p_i]
                return start_i + i - 2
            end
        end

        for p_i in changing_params
            p = params_i[p_i]

            if p.head != :table_ref
                return start_i + i - 2
            end

            start_p = params[p_i]
            row_idx = p.args[2]
            col_idx = p.args[3]

            start_row_idx = start_p.args[2]
            start_col_idx = start_p.args[3]
            fixed_row = start_p.args[4]
            fixed_col = start_p.args[5]
            offset_row = offset_with_fixed(start_row_idx, fixed_row, ((i - 1) * row_offset))
            offset_col = offset_with_fixed(start_col_idx, fixed_col, ((i - 1) * col_offset))

            if (row_idx != offset_row) || (col_idx != offset_col)
                return start_i + i - 2
            end
        end
    end

    return length(statements_in)
end

function export_with_for_loops(exporter, wb::ExcelWorkbook, statement::GroupedStatement)
    sub_statements = statement.sub_statements
    lines = Vector{String}()

    can_loop_ranges = []

    last_idx = 1
    last_stmt = sub_statements[last_idx]

    function maybe_functionalize(stmt)
        if stmt isa TableStatement
            functionalize(stmt.rhs_expr)
        else
            nothing
        end
    end

    functionalized = maybe_functionalize.(sub_statements)
    while last_idx <= length(sub_statements)
        i = get_loop_end(sub_statements, functionalized, last_idx)
        # @show last_idx i length(sub_statements)
        if i > last_idx
            should_loop = can_loop_stmts(@view(sub_statements[last_idx:i]), @view(functionalized[last_idx:i]))
            if should_loop != true
                println("get_loop_end went too far!")
                println("With last_idx = $(last_idx) i = $i, should_loop = $(should_loop)")

                for j in last_idx:i
                    stmt = sub_statements[j]
                    @show get_set_cells(stmt)
                    @show functionalized[j][2]
                end


                @assert should_loop == true
            end
        end
        if i < length(sub_statements)
            shouldnt_loop = can_loop_stmts(@view(sub_statements[last_idx:(i+1)]), @view(functionalized[last_idx:(i+1)]))
            if shouldnt_loop == true
                println("get_loop_end didn't go far enough too far!")
                println("With last_idx = $(last_idx) i = $i, should_loop = $(should_loop)")
                @assert shouldnt_loop != true
            end
        end

        push!(can_loop_ranges, last_idx:i)
        last_idx = i + 1
    end


    if length(can_loop_ranges) == length(sub_statements)
        msg = can_loop_stmts(@view(sub_statements[1:2]), @view(functionalized[1:2])).reason
        push!(lines, "# Can't loop because $msg\n")
    end

    unique_funcs = unique(map(s -> isnothing(s) ? s : s[1], functionalized))
    # if !any(isnothing, functionalized) && length(unique_funcs) == 1
    #     params = reduce(vcat, map(s -> s[2], functionalized))
    #     params_str = join(["param_$i" for i in axes(params, 2)], ", ")
    #     push!(lines, "def func($params_str):\n")
    #     push!(lines, "\treturn " * convert(exporter, first(collect(unique_funcs)), "") * "\n\n")

    #     name = get_function_name(exporter, statement)
    #     for stmt_indices in can_loop_ranges
    #         if length(stmt_indices) == 1
    #             # params_strings = [convert(exporter, param_expr, sheet) for param_expr in params[first(stmt_indices), :]]
    #             params_strings = [convert(exporter, param_expr, "") for param_expr in params[first(stmt_indices), :]]
    #             func_params = join(params_strings, ", ")
    #             rhs = "func($(func_params))"

    #             lhs = try
    #                 convert(exporter, sub_statements[stmt_indices[1]].lhs_expr, "")
    #             catch e
    #                 @show statement.lhs_expr
    #                 throw(e)
    #             end
    #             push!(lines, "$lhs = $rhs\n")
    #             # push!(lines, export_statement(exporter, wb, sub_statements[first(stmt_indices)]))
    #         else
    #             println("$name: looping $(length(stmt_indices)) statements")
    #             push!(lines, export_looped(exporter, wb, sub_statements[stmt_indices]))
    #         end
    #     end

    # else

    name = get_function_name(exporter, statement)
    for stmt_indices in can_loop_ranges
        if length(stmt_indices) == 1

            push!(lines, export_statement(exporter, wb, sub_statements[first(stmt_indices)]))
        else
            println("$name: looping $(length(stmt_indices)) statements")
            push!(lines, export_looped(exporter, wb, sub_statements[stmt_indices]))
        end
    end
    # end


    # @show lines

    reduce(*, lines)
end

function try_make_for_loop(exporter::T, wb::ExcelWorkbook, statement::GroupedStatement) where T
    sub_statements = statement.sub_statements

    function_usages = Dict()
    for s in sub_statements
        func, params = functionalize(s.rhs_expr, [])
        if func in keys(function_usages)
            push!(function_usages[func], (s, params))
        else
            function_usages[func] = [(s, params)]
        end
    end
    usage_counts = collect(zip(keys(function_usages), length.(values(function_usages))))
    sort!(usage_counts, rev = true, by = v -> v[2])
    # @show usage_counts

    most_used_func = usage_counts[1][1]
    # @show most_used_func

    usage_list = function_usages[most_used_func]
    params = reduce(vcat, map(s -> s[2], usage_list))
    # @show size(params)
    # @show params[1]
    param_sets = unique.(eachcol(params))
    fixed_params = findall(length.(param_sets) .== 1)
    changing_params = findall(length.(param_sets) .!= 1)

    usage_stmts = map(s -> s[1], usage_list)
    if any(!(s isa TableStatement) for s in usage_stmts)
        println("Failed because not every statement was a table statement")
        return missing
    end

    set_cells = get_set_cells.(usage_stmts)
    row_nums = map(s -> rownum(s[1]), set_cells)
    col_nums = map(s -> colnum(s[1]), set_cells)
    drow = diff(row_nums)
    dcol = diff(col_nums)
    if !all(drow .== drow[1]) || !all(dcol .== dcol[1])
        println("Failed because there is not a consistent offset")
        return missing
    end
    row_offset = drow[1]
    col_offset = dcol[1]

    # if !all(can_be_for_looped.(eachcol(params[:, changing_params]), row_offset, col_offset))
    if !all(x -> can_be_for_looped(x, row_offset, col_offset), eachcol(params[:, changing_params]))
        println("Failed because a changing parameter couldn't be for looped")
        for x in eachcol(params[:, changing_params])
            println(x)
        end
        return missing
    end

    lhs_exprs = map(s -> s.lhs_expr, usage_stmts)

    if !(can_be_for_looped(lhs_exprs, row_offset, col_offset))
        println("Failed because the lhs_expr couldn't be for looped")
        return missing
    end

    lhs_expr = lhs_exprs[1]

    table, row_idx, col_idx, _, _ = lhs_expr.args

    row_str = make_loop_idx_str(row_idx, row_offset)
    col_str = make_loop_idx_str(col_idx, col_offset)
    lhs_str = "$(getname(table))[$row_str, $col_str]"
    # println("\t$str")

    function get_param_str(param_num, exporter, ctx)
        if param_num in changing_params
            param_expr = params[1, param_num]
            table, row_idx, col_idx, _, _ = param_expr.args
            row_str = make_loop_idx_str(row_idx, row_offset)
            col_str = make_loop_idx_str(col_idx, col_offset)
            "$(getname(table))[$row_str, $col_str]"
        else
            throw("Tried to get param_str for param_num $param_num, but it wasn't a changing param")
        end
    end

    fixed_params_dict = Dict(fixed_params .=> map(v -> v[1], param_sets[fixed_params]))
    rhs_expr = replace_func_params(most_used_func, fixed_params_dict)
    typed_params = Dict{Int64, ExcelExpr}()
    for param_num in changing_params
        all_exprs = params[:, param_num]
        param_type = reduce(union_types, map(e -> get_type(e, table.sheet_name, exporter.cell_types, exporter.named_values), all_exprs))
        typed_params[param_num] = ExcelExpr(:func_param, param_num, param_type)
    end
    rhs_expr = replace_func_params(rhs_expr, typed_params)

    custom_handler = CustomFuncParamHandler(get_param_str)
    custom_exporter = T(exporter.wb, exporter.var_names, exporter.tables, exporter.named_values, [custom_handler, exporter.handlers...], exporter.cell_types)
    rhs_str = convert(custom_exporter, rhs_expr, table.sheet_name)
    """
    for i in 0:$(length(usage_stmts) - 1)
    \t$lhs_str = $rhs_str
    end
    """
end

function get_grouped_statement_body(exporter, wb::ExcelWorkbook, statement::GroupedStatement)
    export_with_for_loops(exporter, wb, statement)
    # try
    #     body = try_make_for_loop(exporter, wb, statement)
    # catch exception
    #     @warn "Failed to make grouped statement into for loop because of exception: " exception
    #     # throw(exception)
    #     body = missing
    # end
    # if ismissing(body)
    #     body = reduce(*, [export_statement(exporter, wb, s) for s in statement.sub_statements])
    # end
    # body
end

function get_function_name(exporter, statement::GroupedStatement)
    cell_for_naming = get_set_cells(statement)[end]
    function_name = "group_calculate_$(normalize_var_name(cell_for_naming.sheet_name))_$(cell_for_naming.cell)"

    function_name
end

function get_params_str(exporter, statement::GroupedStatement)
    needed_vars = get_cell_deps(statement)

    # table_sub_stmts = filter(s -> s isa TableStatement, statement.sub_statements)
    table_vars = [get_set_cells(stmt)[1] for stmt in statement.sub_statements if stmt isa TableStatement]
    append!(needed_vars, table_vars)

    scope_vars = get_required_scope_vars(exporter.tables, exporter.var_names, needed_vars)
    params_str = join(scope_vars, ", ")

    params_str
end


function get_function_string(exporter::JuliaExporter, wb::ExcelWorkbook, statement::GroupedStatement)
    table_sub_stmts = filter(s -> s isa TableStatement, statement.sub_statements)
    if length(table_sub_stmts) != length(statement.sub_statements)
        return nothing
    end

    if length(unique(get_set_table.(table_sub_stmts))) != 1
        return nothing
    end

    # needed_vars = get_cell_deps(statement)

    # intermediate_vars = unique(reduce(vcat, get_set_cells.(statement.sub_statements)))
    # filter!(v -> !(v in intermediate_vars), needed_vars)


    # scope_vars = get_required_scope_vars(exporter.tables, exporter.var_names, needed_vars)
    # cell_for_naming = get_set_cells(statement)[end]
    # function_name = "group_calculate_$(normalize_var_name(cell_for_naming.sheet_name))_$(cell_for_naming.cell)"
    function_name = get_function_name(exporter, statement)
    # params_str = join(scope_vars, ", ")
    params_str = get_params_str(exporter, statement)

    middle_lines = get_grouped_statement_body(exporter, wb, statement)
    # middle_lines = reduce(*, [export_statement(exporter, wb, s) for s in statement.sub_statements])

    lines = split(middle_lines, "\n")
    function_inner = join(["\t" * l for l in lines], "\n")

    """
    function $function_name($params_str)
    $function_inner
    end"""
end

function get_function_string(exporter::PythonExporter, wb::ExcelWorkbook, statement::GroupedStatement)
    table_sub_stmts = filter(s -> s isa TableStatement, statement.sub_statements)
    if length(table_sub_stmts) != length(statement.sub_statements)
        return nothing
    end

    if length(unique(get_set_table.(table_sub_stmts))) != 1
        return nothing
    end

    function_name = get_function_name(exporter, statement)
    params_str = get_params_str(exporter, statement)

    middle_lines = get_grouped_statement_body(exporter, wb, statement)

    lines = split(middle_lines, "\n")
    function_inner = join(["\t" * l for l in lines], "\n")

    """
    def $function_name($params_str):
    $function_inner

    """
end

function export_statement(exporter::JuliaExporter, wb::ExcelWorkbook, statement::GroupedStatement)
    set_cells = get_set_cells(statement)
    out_var_exprs = [ExcelExpr(:cell_ref, Any[cell.cell, cell.sheet_name]) for cell in set_cells]
    out_var_exprs = map(Base.Fix2(insert_table_refs, exporter.tables), out_var_exprs)

    variable_names = [convert(exporter, e, "") for e in out_var_exprs]

    xf = wb.xf
    assert_lines = ""
    # for (cell_ref, name) in zip(set_cells, variable_names)
    #     assert_lines *= "@assert xl_compare($name, $(repr(xf[string(cell_ref.sheet_name)][cell_ref.cell]))) # $(to_string(cell_ref))\n"
    # end


    table_sub_stmts = filter(s -> s isa TableStatement, statement.sub_statements)
    if length(table_sub_stmts) != length(statement.sub_statements) || length(unique(get_set_table.(table_sub_stmts))) != 1
        # middle_lines = reduce(*, [export_statement(exporter, wb, s) for s in statement.sub_statements])
        middle_lines = get_grouped_statement_body(exporter, wb, statement)
        """
        # Group of $(length(statement.sub_statements)) statements
        begin
        $middle_lines\
        end
        """
    else
        # needed_vars = get_cell_deps(statement)

        # intermediate_vars = unique(reduce(vcat, get_set_cells.(statement.sub_statements)))
        # filter!(v -> !(v in intermediate_vars), needed_vars)

        # scope_vars = get_required_scope_vars(exporter.tables, exporter.var_names, needed_vars)
        # cell_for_naming = get_set_cells(statement)[end]
        # function_name = "group_calculate_$(normalize_var_name(cell_for_naming.sheet_name))_$(cell_for_naming.cell)"
        function_name = get_function_name(exporter, statement)
        # params_str = join(scope_vars, ", ")
        params_str = get_params_str(exporter, statement)

        """
        $function_name($params_str)
        $(assert_lines)
        """
    end
end


function export_statement(exporter::PythonExporter, wb::ExcelWorkbook, statement::GroupedStatement)
    set_cells = get_set_cells(statement)
    out_var_exprs = [ExcelExpr(:cell_ref, Any[cell.cell, cell.sheet_name]) for cell in set_cells]
    out_var_exprs = map(Base.Fix2(insert_table_refs, exporter.tables), out_var_exprs)

    variable_names = [convert(exporter, e, "") for e in out_var_exprs]

    xf = wb.xf
    assert_lines = ""
    for (cell_ref, name) in zip(set_cells, variable_names)
        cell_value = xf[string(cell_ref.sheet_name)][cell_ref.cell]
        value_str = convert(exporter, cell_value, cell_ref.sheet_name)
        assert_lines *= "assert xl.compare($name, $value_str) # $(to_string(cell_ref))\n"
    end


    table_sub_stmts = filter(s -> s isa TableStatement, statement.sub_statements)
    if length(table_sub_stmts) != length(statement.sub_statements) || length(unique(get_set_table.(table_sub_stmts))) != 1
        # middle_lines = reduce(*, [export_statement(exporter, wb, s) for s in statement.sub_statements])
        middle_lines = get_grouped_statement_body(exporter, wb, statement)
        """
        # Group of $(length(statement.sub_statements)) statements
        $middle_lines
        $assert_lines
        """
    else
        # needed_vars = get_cell_deps(statement)

        # intermediate_vars = unique(reduce(vcat, get_set_cells.(statement.sub_statements)))
        # filter!(v -> !(v in intermediate_vars), needed_vars)

        # scope_vars = get_required_scope_vars(exporter.tables, exporter.var_names, needed_vars)
        # cell_for_naming = get_set_cells(statement)[end]
        # function_name = "group_calculate_$(normalize_var_name(cell_for_naming.sheet_name))_$(cell_for_naming.cell)"
        function_name = get_function_name(exporter, statement)
        # params_str = join(scope_vars, ", ")
        params_str = get_params_str(exporter, statement)

        """
        $function_name($params_str)
        $(assert_lines)
        """
    end
end
