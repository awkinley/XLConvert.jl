
struct ToggleCondHandler end

function handle(::ToggleCondHandler, expr::ExcelExpr, exporter::JuliaExporter, ctx)
    func = a -> convert(exporter, a, ctx)
    @match expr begin
        ExcelExpr(:call, ["IF", ExcelExpr(:eq, [ExcelExpr(:named_range, [toggle_var,]), 1.0]), t, f]) => begin
            # toggle_var_name = exporter.var_names[CellDependency(ctx, toggle_cell)]
            toggle_var_name = convert(exporter, exporter.named_values[toggle_var], ctx)
            if !contains(toggle_var_name, "toggle")
                return missing
            end

            "(Bool($toggle_var_name) ? $(func(t)) : $(func(f)))"
        end
        _ => missing
    end
end

function convert_if_toggle(expr)
    postwalk(expr) do expr
        @match expr begin
            ExcelExpr(:call, ["IF", ExcelExpr(:eq, [ExcelExpr(:named_range, [toggle_var,]), 1.0]), t, f]) => begin
                if !(contains(toggle_var, "toggle") || contains(toggle_var, "flag"))
                    return expr
                end

                ExcelExpr(:call, Any["IF", ExcelExpr(:named_range, toggle_var), convert_if_toggle(t), convert_if_toggle(f)])
            end
            ExcelExpr(:call, ["IF", ExcelExpr(:eq, [ExcelExpr(:named_range, [toggle_var,]), 0.0]), t, f]) => begin
                if !(contains(toggle_var, "toggle") || contains(toggle_var, "flag"))
                    return expr
                end

                ExcelExpr(:call, Any["IF", ExcelExpr(:call, Any["NOT", ExcelExpr(:named_range, toggle_var)]), convert_if_toggle(t), convert_if_toggle(f)])
            end
            _ => expr
        end
    end

    # end
    # if @ismatch expr ExcelExpr(:call, ["IF", ExcelExpr(:eq, [ExcelExpr(:named_range, [toggle_var,]), 1.0]), t, f])
    #     if !(contains(toggle_var, "toggle") || contains(toggle_var, "flag"))
    #         return expr
    #     end

    #     ExcelExpr(:call, "IF", ExcelExpr(:named_range, toggle_var), convert_if_toggle(t), convert_if_toggle(f))
    # elseif @ismatch expr ExcelExpr(:call, ["IF", ExcelExpr(:eq, [ExcelExpr(:named_range, [toggle_var,]), 0.0]), t, f])
    #     if !(contains(toggle_var, "toggle") || contains(toggle_var, "flag"))
    #         return expr
    #     end

    #     ExcelExpr(:call, "IF", ExcelExpr(:call, "NOT", ExcelExpr(:named_range, toggle_var)), convert_if_toggle(t), convert_if_toggle(f))
    # elseif @ismatch expr ExcelExpr(op, args)
    #     ExcelExpr(op, convert_if_toggle.(args)...)
    # else
    #     expr
    # end
end

function if_toggle_transform!(statements::AbstractArray{AbstractStatement})
    for s in statements
        apply_expr_transform!(s, (_, expr) -> convert_if_toggle(expr))
    end
end