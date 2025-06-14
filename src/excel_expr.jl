@auto_hash_equals struct ExcelExpr
    head::Symbol
    # args::Tuple
    args::Array{Any,1}

    ExcelExpr(head::Symbol, arg::Any) = new(head, Any[arg])
    ExcelExpr(head::Symbol, args...) = new(head, collect(Any, args))
    # ExcelExpr(head::Symbol, args::Tuple) = new(head, collect(Any, args))
    ExcelExpr(head::Symbol, args::Array{Any,1}) = new(head, args)
end

function Base.deepcopy(e::ExcelExpr)
    ExcelExpr(e.head, deepcopy(e.args))
end

function Base.show(io::IO, expr::ExcelExpr)
    # arg_str = join(map(repr, expr.args), ", ")
    # arg_str = join(map(repr, expr.args), ", ")
    # print(io, "ExcelExpr(:$(expr.head), $(arg_str))")
    print(io, "ExcelExpr(:$(expr.head), $(expr.args))")
end

walk(x, inner, outer) = outer(x)
function walk(x::ExcelExpr, inner, outer)
    new_args = similar(x.args)
    map!(inner, new_args, x.args)
    outer(ExcelExpr(x.head, new_args))
end

"""
    postwalk(f, expr)

Applies `f` to each node in the given expression tree, returning the result.
`f` sees expressions *after* they have been transformed by the walk.

See also: [`prewalk`](@ref).
"""
postwalk(f, x) = walk(x, x -> postwalk(f, x), f)

"""
    prewalk(f, expr)

Applies `f` to each node in the given expression tree, returning the result.
`f` sees expressions *before* they have been transformed by the walk, and the
walk will be applied to whatever `f` returns.

This makes `prewalk` somewhat prone to infinite loops; you probably want to try
[`postwalk`](@ref) first.
"""
prewalk(f, x) = walk(f(x), x -> prewalk(f, x), identity)