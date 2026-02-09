mutable struct StandardStatement <: AbstractStatement
    assigned_var::CellDependency
    rhs_expr::Any
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

    xf = wb.xf
    cell = getcell(xf, statement.assigned_var)
    # @show cell
    formula_str = ""
    if !isempty(cell) && !(cell.formula isa XLSX.FormulaReference)
        formula_str = replace(cell.formula.formula, "\n" => "\n# ")
    end

    # sub_exprs, new_expr = common_subexpression_elimination(expr)
    if false #(2 * length(sub_exprs) + length(new_expr.parts)) < length(expr.parts)

        res = """
        # =$formula_str
        $lhs = begin
        """
        for (i, sub) in enumerate(sub_exprs)
            rhs = convert(exporter, sub, cell_ref.sheet_name)
            res *= "\tparam_$(i) = $rhs\n"
        end
        res *= "\n"

        rhs = convert(exporter, new_expr, cell_ref.sheet_name)
        res *= "\t" * rhs * "\n"
        res *= "end\n"
        res *= "@assert xl_compare($lhs, $(repr(xf[string(cell_ref.sheet_name)][cell_ref.cell]))) # $(to_string(cell_ref))\n"

        res
    else

        rhs = convert(exporter, expr, cell_ref.sheet_name)

        # """
        # # =$formula_str
        # $lhs = $rhs # $(cell_ref.sheet_name) $(cell_ref.cell)
        # @assert xl_compare($lhs, $(repr(xf[string(cell_ref.sheet_name)][cell_ref.cell]))) # $(to_string(cell_ref))
        # """
        """
        # =$formula_str
        $lhs = $rhs # $(cell_ref.sheet_name) $(cell_ref.cell)
        """

    end

end