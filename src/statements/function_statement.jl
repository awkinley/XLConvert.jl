struct FunctionStatement <: AbstractStatement
    assigned_var::CellDependency
    inputs::Vector{AbstractStatement}
    intermediates::Vector{AbstractStatement}
end


get_set_cells(stmt::FunctionStatement) = [stmt.assigned_var]
get_cell_deps(stmt::FunctionStatement) = reduce(vcat, get_set_cells.(stmt.inputs))

function to_string(exporter, statement::FunctionStatement)
    cell_ref = statement.assigned_var
    lhs = exporter.var_names[cell_ref]

    "FunctionStatement(lhs = $lhs)"
end

function get_required_scope_vars(tables::Vector{ExcelTable}, var_names, cells::Vector{CellDependency})
    vars = Vector{String}()
    for cell in cells
        tbl = find_table_containing_cell(cell, tables)
        var = if isnothing(tbl)
            var_names[cell]
        else
            getname(tables[tbl])
        end

        if !(var in vars)
            push!(vars, var)
        end
    end

    vars
end

function export_statement(exporter::JuliaExporter, wb::ExcelWorkbook, statement::FunctionStatement)
    cell_ref = statement.assigned_var
    lhs = exporter.var_names[cell_ref]

    needed_vars = copy(get_cell_deps(statement))
    for child in statement.intermediates
        append!(needed_vars, get_cell_deps(child))
    end

    intermediate_vars = unique(reduce(vcat, get_set_cells.(filter(s -> s isa StandardStatement, statement.intermediates))))
    filter!(v -> !(v in intermediate_vars), needed_vars)


    scope_vars = get_required_scope_vars(exporter.tables, exporter.var_names, needed_vars)

    function_name = "calculate_$(exporter.var_names[statement.assigned_var])"

    params_str = join(scope_vars, ", ")
    """
    $lhs = $function_name($params_str)
    """
end

function get_function_string(exporter::JuliaExporter, wb::ExcelWorkbook, statement::FunctionStatement)
    needed_vars = copy(get_cell_deps(statement))
    for child in statement.intermediates
        append!(needed_vars, get_cell_deps(child))
    end

    intermediate_vars = unique(reduce(vcat, get_set_cells.(filter(s -> s isa StandardStatement, statement.intermediates))))
    filter!(v -> !(v in intermediate_vars), needed_vars)


    scope_vars = get_required_scope_vars(exporter.tables, exporter.var_names, needed_vars)

    function_name = "calculate_$(exporter.var_names[statement.assigned_var])"

    params_str = join(scope_vars, ", ")

    function_lines = Vector{String}()

    for child in statement.intermediates
        line = export_statement(exporter, wb, child)
        if occursin("ExcelExpr", line)
            @show line
            @show child
            throw("ExcelExpr seemed to get exported?")
        end
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