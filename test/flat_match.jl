using Test
using Match

using XLConvert
using XLConvert: toexpr, lower_sheet_names!, get_expr_dependencies, convert_to_flat_expr, FlatExpr, FlatIdx

function test_flat_match()
    expr = toexpr("SUM(A1:B20, B2)")
    lower_sheet_names!(expr, "Sheet")
    flat_expr = convert_to_flat_expr(expr)



    for (i, part) in enumerate(flat_expr.parts)
        @show part
        @flat_match flat_expr part begin
            ExcelExpr(:range, [ExcelExpr(:cell_ref, [lhs, sheet]), ExcelExpr(:cell_ref, [rhs, sheet])]) => begin
                println("Found range on sheet $sheet from cell $lhs to $rhs")
            end
            _ => continue
        end
    end
end