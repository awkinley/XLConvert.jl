@auto_hash_equals struct ExcelExpr
    head::Symbol
    args::Tuple

    ExcelExpr(head::Symbol, args...) = new(head, args)
end


function Base.show(io::IO, expr::ExcelExpr)
    arg_str = join(map(repr, expr.args), ", ")
    print(io, "ExcelExpr(:$(expr.head), $(arg_str))")
end