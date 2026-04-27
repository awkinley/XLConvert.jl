
"""
While a normal FunctionStatement returns a value, a generic function statement 
does not have to. Instead it may just modify its input tables and such.

This means it behaves more like a grouped statement in terms of being considered
to set all of its intermediate values.
"""
struct GenericFunctionStatement <: AbstractStatement
    name::AbstractString
    sub_statements::Vector{AbstractStatement}
end


get_set_cells(stmt::GenericFunctionStatement) = reduce(vcat, get_set_cells.(stmt.sub_statements))
get_cell_deps(stmt::GenericFunctionStatement) = unique(reduce(vcat, get_cell_deps.(stmt.sub_statements)))

function apply_expr_transform!(stmt::GenericFunctionStatement, transform)
    for s in stmt.sub_statements
        apply_expr_transform!(s, transform)
    end
end

function get_func_name(exporter, statement::GenericFunctionStatement)
    last_cell = first(get_set_cells(statement.sub_statements[end]))
    cell_str = normalize_var_name(string(last_cell.sheet_name ,"_", last_cell.cell))
    "calculate_$(cell_str)"
end

function to_string(exporter, statement::GenericFunctionStatement)
    name = statement.name
    # lhs = exporter.var_names[cell_ref]

    "GenericFunctionStatement($name)"
end


function get_scope_vars(exporter, statement::GenericFunctionStatement)
    needed_vars = get_cell_deps(statement)
    intermediate_vars = Set(reduce(vcat, get_set_cells.(statement.sub_statements)))
    filter!(v -> !(v in intermediate_vars), needed_vars)
    pushfirst!(needed_vars, get_set_cells(statement.sub_statements[1])[1])

    get_required_scope_vars(exporter.tables, exporter.var_names, needed_vars)
end

function export_statement(exporter::JuliaExporter, wb::ExcelWorkbook, statement::GenericFunctionStatement)
    scope_vars = get_scope_vars(exporter, statement)

    function_name = get_func_name(exporter, statement)
    # function_name = "calculate_$(exporter.var_names[statement.assigned_var])"

    params_str = join(scope_vars, ", ")
    """
    $lhs = $function_name($params_str)
    """
end

function get_function_string(exporter::JuliaExporter, wb::ExcelWorkbook, statement::GenericFunctionStatement)
    # needed_vars = get_cell_deps(statement)
    # intermediate_vars = Set(reduce(vcat, get_set_cells.(statement.sub_statements)))
    # filter!(v -> !(v in intermediate_vars), needed_vars)


    # scope_vars = get_required_scope_vars(exporter.tables, exporter.var_names, needed_vars)
    scope_vars = get_scope_vars(exporter, statement)

    # function_name = "calculate_$(exporter.var_names[statement.assigned_var])"
    function_name = get_func_name(exporter, statement)

    params_str = join(scope_vars, ", ")

    function_lines = Vector{String}()

    for child in statement.sub_statements
        line = export_statement(exporter, wb, child)
        # if occursin("ExcelExpr", line)
        #     @show line
        #     @show child
        #     throw("ExcelExpr seemed to get exported?")
        # end
        push!(function_lines, line)
    end

    function_inner = reduce(*, function_lines)
    lines = split(function_inner, "\n")
    function_inner = join(["\t" * l for l in lines], "\n")

    """
    function $function_name($params_str)
    $function_inner
    \t$(exporter.var_names[statement.assigned_var])
    end"""
end

function export_statement(exporter::PythonExporter, wb::ExcelWorkbook, statement::GenericFunctionStatement)
    # needed_vars = get_cell_deps(statement)
    # intermediate_vars = Set(reduce(vcat, get_set_cells.(statement.sub_statements)))
    # filter!(v -> !(v in intermediate_vars), needed_vars)

    # scope_vars = get_required_scope_vars(exporter.tables, exporter.var_names, needed_vars)
    scope_vars = get_scope_vars(exporter, statement)

    # function_name = "calculate_$(exporter.var_names[statement.assigned_var])"
    function_name = get_func_name(exporter, statement)

    params_str = join(scope_vars, ", ")
    """
    $function_name($params_str)
    """
end

function get_function_string(exporter::PythonExporter, wb::ExcelWorkbook, statement::GenericFunctionStatement)
    # needed_vars = get_cell_deps(statement)
    # intermediate_vars = Set(reduce(vcat, get_set_cells.(statement.sub_statements)))
    # filter!(v -> !(v in intermediate_vars), needed_vars)

    # scope_vars = get_required_scope_vars(exporter.tables, exporter.var_names, needed_vars)
    scope_vars = get_scope_vars(exporter, statement)

    function_name = get_func_name(exporter, statement)

    params_str = join(scope_vars, ", ")

    function_lines = Vector{String}()

    for child in statement.sub_statements
        line = export_statement(exporter, wb, child)
        # if occursin("ExcelExpr", line)
        #     @show line
        #     @show child
        #     throw("ExcelExpr seemed to get exported?")
        # end
        push!(function_lines, line)
    end

    function_inner = reduce(*, function_lines)
    lines = split(function_inner, "\n")
    function_inner = join(["\t" * l for l in lines], "\n")

    dependent_funcs = ""
    for stmt in statement.sub_statements
        if stmt isa GroupedStatement
            func_str = get_function_string(exporter, wb, stmt)
            if !isnothing(func_str)
                dependent_funcs *= func_str * "\n"
            end
        end
    end


    """
    $dependent_funcs
    def $function_name($params_str):
    $function_inner
    """
end