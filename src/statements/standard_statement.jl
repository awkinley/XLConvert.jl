mutable struct StandardStatement <: AbstractStatement
    assigned_var::CellDependency
    rhs_expr
    rhs_dependencies::Vector{CellDependency}
end

get_cell_deps(stmt::StandardStatement) = stmt.rhs_dependencies
get_set_cells(stmt::StandardStatement) = [stmt.assigned_var]
apply_expr_transform!(stmt::StandardStatement, transform) = stmt.rhs_expr = transform(stmt, stmt.rhs_expr)

function to_string(exporter, statement::StandardStatement)
    cell_ref = statement.assigned_var
    lhs = exporter.var_names[cell_ref]

    "StandardStatement(lhs = $lhs)"
end

function export_statement(exporter::JuliaExporter, wb::ExcelWorkbook, statement::StandardStatement)
    cell_ref = statement.assigned_var
    lhs = exporter.var_names[cell_ref]
    expr = statement.rhs_expr
    rhs = convert(exporter, expr, cell_ref.sheet_name)

    xf = wb.xf
    cell = getcell(xf, statement.assigned_var)
    formula_str = ""
    if !(cell.formula isa XLSX.FormulaReference)
        formula_str = replace(cell.formula.formula, "\n" => "\n# ")
    end

    """
    # =$formula_str
    $lhs = $rhs # $(cell_ref.sheet_name) $(cell_ref.cell)
    @assert xl_compare($lhs, $(repr(xf[string(cell_ref.sheet_name)][cell_ref.cell]))) # $(to_string(cell_ref))
    """
    # """
    # # =$formula_str
    # $lhs = $rhs # $(cell_ref.sheet_name) $(cell_ref.cell)
    # """
end