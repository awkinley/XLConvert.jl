mutable struct OutputStatement <: AbstractStatement
    output_vars::Vector{CellDependency}
end

get_set_cells(::OutputStatement) = []
get_cell_deps(stmt::OutputStatement) = stmt.output_vars
apply_expr_transform!(::OutputStatement, transform) = nothing

function to_string(exporter, statement::OutputStatement)
    "OutputStatement"
end

function make_outupt_struct(exporter::JuliaExporter, wb::ExcelWorkbook, statement::OutputStatement)
    variable_names = [exporter.var_names[v] for v in statement.output_vars]
    make_struct(exporter, "Outputs", variable_names)
end

function export_statement(exporter::JuliaExporter, wb::ExcelWorkbook, statement::OutputStatement)
    out_vars = statement.output_vars
    variable_names = [exporter.var_names[v] for v in out_vars]

    var_str = join(variable_names, ",\n\t")

    """
    Outputs(
        $var_str    
    )
    """
end