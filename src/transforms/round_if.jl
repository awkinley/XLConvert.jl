
function convert_round_if(expr)
    if @ismatch expr ExcelExpr(:call, ("IF", cond, ExcelExpr(:call, ("ROUNDUP", t, 0)), f))
        if t == f
            return ExcelExpr(:call, "ROUNDUP_IF", cond, convert_round_if(t))
        end
    end

    if @ismatch expr ExcelExpr(op, args)
        ExcelExpr(op, convert_round_if.(args)...)
    else
        expr
    end
end

function round_if_transform!(statements::AbstractArray{AbstractStatement})
    for s in statements
        apply_expr_transform!(s, (_, expr) -> convert_round_if(expr))
    end
end