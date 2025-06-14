mutable struct GroupedStatement <: AbstractStatement
    sub_statements::Vector{AbstractStatement}
end

get_cell_deps(stmt::GroupedStatement) = reduce(vcat, get_cell_deps.(stmt.sub_statements))
get_set_cells(stmt::GroupedStatement) = reduce(vcat, get_set_cells.(stmt.sub_statements))
function apply_expr_transform!(stmt::GroupedStatement, transform)
    for s in stmt.sub_statments
        apply_expr_transform!(s, transform)
    end
end

function to_string(exporter, statement::GroupedStatement)

    children = join(to_string.((exporter,), statement.sub_statements), ", ")
    "GroupedStatement($children)"
end

function can_be_for_looped(expressions, row_offset, col_offset)
    start_row_idx = nothing
    start_col_idx = nothing
    # @show expressions
    first_val = expressions[1]
    if first_val.head == :table_ref
        start_row_idx = first_val.args[2]
        start_col_idx = first_val.args[3]
    else
        println("Failed because a changing param wasn't a table ref")
        return false
    end

    for j in eachindex(expressions)[2:end]
        val = expressions[j]
        if val.head == :table_ref
            row_idx = val.args[2]
            col_idx = val.args[3]

            good_row = row_idx == start_row_idx .+ ((j - 1) * row_offset)
            good_col = col_idx == start_col_idx .+ ((j - 1) * col_offset)
            if !(good_row && good_col)
                println("Failed because an expression wasn't offset correctly")
                return false
            end
        else
            println("Failed because a changing param wasn't a table ref")
            return false
        end
    end

    true
end

function make_loop_idx_str(base_idx, offset)
    if offset == 0
        "$base_idx"
    else
        if offset == 1
            "$base_idx .+ i"
        else
            "$base_idx .+ (i * $offset)"
        end
    end
end

struct CustomFuncParamHandler
    value_getter
end

function handle(handler::CustomFuncParamHandler, expr, exporter::JuliaExporter, ctx)
    @match expr begin
        ExcelExpr(:func_param, [param_num,]) => begin
            handler.value_getter(param_num, exporter, ctx)
        end
        ExcelExpr(:func_param, [param_num, type]) => begin
            handler.value_getter(param_num, exporter, ctx)
        end
        _ => missing
    end
end

function replace_func_params(expr, params_dict)
    @match expr begin
        ExcelExpr(:func_param, [param_num,]) => get(params_dict, param_num, expr)
        ExcelExpr(head, args) => ExcelExpr(head, map(e -> replace_func_params(e, params_dict), args)...)
        _ => expr
    end
end

function can_loop_stmts(statements::AbstractArray{AbstractStatement}, functionalized)
    if any(!(s isa TableStatement) for s in statements)
        # They need to all be TableStatements
        println("Failed because not every statement was a table statement")
        return false
    end

    if length(statements) == 0
        return false
    end

    all_params = []
    # func, params = functionalize(statements[1].rhs_expr, [])
    func, params = functionalized[1]
    push!(all_params, params)
    # params = []
    for (new_func, func_params) in functionalized[2:end]
        # for s in statements[2:end]
        #     new_func, func_params = functionalize(s.rhs_expr, [])
        if func != new_func
            println("Failed because not all the functions are the same")
            return false
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
        println("Failed because there is not a consistent offset")
        return false
    end
    row_offset = drow[1]
    col_offset = dcol[1]

    # if !all(can_be_for_looped.(eachcol(params[:, changing_params]), row_offset, col_offset))
    if !all(x -> can_be_for_looped(x, row_offset, col_offset), eachcol(params[:, changing_params]))
        println("Failed because a changing parameter couldn't be for looped")
        return false
    end

    lhs_exprs = map(s -> s.lhs_expr, statements)

    if !(can_be_for_looped(lhs_exprs, row_offset, col_offset))
        println("Failed because a lhs couldn't be for looped")
        return false
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
    rhs_expr = replace_func_params(func, fixed_params_dict)
    typed_params = Dict{Int64,ExcelExpr}()
    # @show changing_params
    for param_num in changing_params
        all_exprs = params[:, param_num]
        # @show all_exprs
        param_type = reduce(union_types, map(e -> get_type(e, table.sheet_name, exporter.cell_types, exporter.named_values), all_exprs))
        typed_params[param_num] = ExcelExpr(:func_param, param_num, param_type)
    end
    rhs_expr = replace_func_params(rhs_expr, typed_params)

    custom_handler = CustomFuncParamHandler(get_param_str)
    custom_exporter = JuliaExporter(exporter.wb, exporter.var_names, exporter.tables, exporter.named_values, [custom_handler, exporter.handlers...], exporter.cell_types)
    rhs_str = convert(custom_exporter, rhs_expr, table.sheet_name)
    """
    for i in 0:$(length(statements) - 1)
    \t$lhs_str = $rhs_str
    end
    """
end

function export_with_for_loops(exporter::JuliaExporter, wb::ExcelWorkbook, statement::GroupedStatement)
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

    for i in 2:length(sub_statements)
        # @show i
        if !can_loop_stmts(@view(sub_statements[last_idx:i]), @view(functionalized[last_idx:i]))
            push!(can_loop_ranges, last_idx:i-1)
            last_idx = i
        end
    end
    push!(can_loop_ranges, last_idx:length(sub_statements))
    # @show last_idx
    # @show can_loop_ranges

    for stmt_indices in can_loop_ranges
        if length(stmt_indices) == 1
            push!(lines, export_statement(exporter, wb, sub_statements[first(stmt_indices)]))
        else
            push!(lines, export_looped(exporter, wb, sub_statements[stmt_indices]))
        end
    end

    # @show lines

    reduce(*, lines)
end

function try_make_for_loop(exporter::JuliaExporter, wb::ExcelWorkbook, statement::GroupedStatement)
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
    sort!(usage_counts, rev=true, by=v -> v[2])
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
    typed_params = Dict{Int64,ExcelExpr}()
    for param_num in changing_params
        all_exprs = params[:, param_num]
        param_type = reduce(union_types, map(e -> get_type(e, table.sheet_name, exporter.cell_types, exporter.named_values), all_exprs))
        typed_params[param_num] = ExcelExpr(:func_param, param_num, param_type)
    end
    rhs_expr = replace_func_params(rhs_expr, typed_params)

    custom_handler = CustomFuncParamHandler(get_param_str)
    custom_exporter = JuliaExporter(exporter.wb, exporter.var_names, exporter.tables, exporter.named_values, [custom_handler, exporter.handlers...], exporter.cell_types)
    rhs_str = convert(custom_exporter, rhs_expr, table.sheet_name)
    """
    for i in 0:$(length(usage_stmts) - 1)
    \t$lhs_str = $rhs_str
    end
    """
end

function get_grouped_statement_body(exporter::JuliaExporter, wb::ExcelWorkbook, statement::GroupedStatement)
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


function get_function_string(exporter::JuliaExporter, wb::ExcelWorkbook, statement::GroupedStatement)
    table_sub_stmts = filter(s -> s isa TableStatement, statement.sub_statements)
    if length(table_sub_stmts) != length(statement.sub_statements)
        return nothing
    end

    if length(unique(get_set_table.(table_sub_stmts))) != 1
        return nothing
    end

    needed_vars = get_cell_deps(statement)

    intermediate_vars = unique(reduce(vcat, get_set_cells.(statement.sub_statements)))
    filter!(v -> !(v in intermediate_vars), needed_vars)


    scope_vars = get_required_scope_vars(exporter.tables, exporter.var_names, needed_vars)
    cell_for_naming = get_set_cells(statement)[end]
    function_name = "group_calculate_$(normalize_var_name(cell_for_naming.sheet_name))_$(cell_for_naming.cell)"
    params_str = join(scope_vars, ", ")

    middle_lines = get_grouped_statement_body(exporter, wb, statement)
    # middle_lines = reduce(*, [export_statement(exporter, wb, s) for s in statement.sub_statements])

    lines = split(middle_lines, "\n")
    function_inner = join(["\t" * l for l in lines], "\n")

    """
    function $function_name($params_str)
    $function_inner
    end"""
end

function export_statement(exporter::JuliaExporter, wb::ExcelWorkbook, statement::GroupedStatement)

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
        needed_vars = get_cell_deps(statement)

        intermediate_vars = unique(reduce(vcat, get_set_cells.(statement.sub_statements)))
        filter!(v -> !(v in intermediate_vars), needed_vars)

        scope_vars = get_required_scope_vars(exporter.tables, exporter.var_names, needed_vars)
        cell_for_naming = get_set_cells(statement)[end]
        function_name = "group_calculate_$(normalize_var_name(cell_for_naming.sheet_name))_$(cell_for_naming.cell)"
        params_str = join(scope_vars, ", ")

        """
        $function_name($params_str)
        """
    end
end