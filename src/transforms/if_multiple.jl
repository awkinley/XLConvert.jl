convert_if_multiple(expr) = expr
function convert_if_multiple(expr::ExcelExpr)
    if @ismatch expr ExcelExpr(:call, ["IF", ExcelExpr(:eq, [ExcelExpr(:call, ["MOD", dividend, divisor]), 0.0]), result_value, 0.0])
        ExcelExpr(:call, "IF_MULTIPLE", dividend, divisor, result_value)
    else
        expr
    end
end

function if_multiple_transform!(statements::AbstractArray{AbstractStatement})
    for s in statements
        apply_expr_transform!(s, (_, expr) -> convert_if_multiple(expr))
    end
end