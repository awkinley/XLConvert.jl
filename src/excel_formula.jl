# using Test
# using PrettyPrint
# using AutoHashEquals
# using Match
# using Dates
# using Statistics
# using XLSX
# using DataFrames
# using .XLExpr
include("./functions/arithmetic.jl")
include("./functions/comparison.jl")


flatten(arr::AbstractArray) = reduce(vcat, arr)
flatten(arr) = arr

function offset_cell_str(cell::T, rows::Int, cols::Int) where {T<:AbstractString}
    first_let = findfirst(c -> 'A' <= c <= 'Z', cell)
    first_num = findfirst(isdigit, cell)

    col_str = if cell[first_num-1] == '$'
        @view cell[first_let:(first_num-2)]
    else
        @view cell[first_let:(first_num-1)]
    end
    row_str = @view cell[first_num:end]

    if cell[1] == '$'
        new_col = col_str
    else
        new_col_num = XLSX.decode_column_number(col_str) + cols
        # This allows for (invalid) offsetting of expressions into negative columns
        if new_col_num >= 1
            new_col = XLSX.encode_column_number(new_col_num)
        else
            new_col = string(new_col_num)
        end
    end
    if cell[first_num-1] == '$'
        new_row = row_str
    else
        new_row = string(parse(Int, row_str) + rows)
    end

    new_col * new_row
end


function offset(value, ::Int, ::Int)
    value
end
offset_cell_parse_rgx = r"([$]?[A-Z]+)([$]?[0-9]+)"
function offset(expr::ExcelExpr, rows::Int, cols::Int)
    head = expr.head
    args = expr.args

    @match expr begin
        ExcelExpr(:cell_ref, args) => begin
            new_args = copy(args)
            # new_cell = offset_cell_str(args[1], rows, cols)
            new_args[1] = offset_cell_str(new_args[1], rows, cols)
            # ExcelExpr(:cell_ref, new_cell, args[2:end]...)
            ExcelExpr(:cell_ref, new_args)
            # cell_match = match(offset_cell_parse_rgx, args[1])
            # @assert cell_match.match == args[1] "Cell didn't parse properly"
            # col_str = cell_match[1]
            # row_str = cell_match[2]

            # if col_str[1] == '$'
            #     new_col = col_str
            # else
            #     new_col_num = XLSX.decode_column_number(col_str[1:end]) + cols
            #     # This allows for (invalid) offsetting of expressions into negative columns
            #     if new_col_num >= 1
            #         new_col = XLSX.encode_column_number(new_col_num)
            #     else
            #         new_col = string(new_col_num)
            #     end
            # end
            # if row_str[1] == '$'
            #     new_row = row_str
            # else
            #     new_row = string(parse(Int, row_str) + rows)
            # end
            # ExcelExpr(:cell_ref, String(new_col * new_row), args[2:end]...)
        end
        ExcelExpr(:table_ref, [table, row_idx, col_idx, fixed_row, fixed_col]) => begin
            row_idx = @match fixed_row begin
                (true, true) => row_idx
                (false, false) => row_idx .+ rows
                (true, false) => first(row_idx):(last(row_idx)+rows)
                (false, true) => (first(row_idx)+rows):last(row_idx)
            end
            col_idx = @match fixed_col begin
                (true, true) => col_idx
                (false, false) => col_idx .+ cols
                (true, false) => first(col_idx):(last(col_idx)+cols)
                (false, true) => (first(col_idx)+cols):last(col_idx)
            end

            ExcelExpr(:table_ref, Any[table, row_idx, col_idx, fixed_row, fixed_col])
        end
        _ => begin
            new_args = similar(args)
            for i in eachindex(args)
                new_args[i] = offset(args[i], rows, cols)
            end
            # map!(x -> offset(x, rows, cols), new_args, args)
            # ExcelExpr(head, map(x -> offset(x, rows, cols), args)...)
            ExcelExpr(head, new_args)
        end
    end
end


struct SheetValues1
    cell_values::Any
end
SheetValues = SheetValues1
# abstract type XLContext end
struct ExcelContext5
    current_sheet::Any
    sheets::Any
    key_values::Any
end
ExcelContext = ExcelContext5
function get_cell_value(ctx::SheetValues, cell)
    ctx.cell_values[cell]
end

function get_cell_value(ctx::ExcelContext, cell)
    get_cell_value(ctx.sheets[ctx.current_sheet], cell)
end

function with_current_sheet(ctx::ExcelContext, sheet)
    ExcelContext(sheet, ctx.sheets, ctx.key_values)
end

function get_key_value(ctx::ExcelContext, key)
    ctx.key_values[key]
end
function exec(expr::T, ctx::ExcelContext) where {T <: Number}
    expr
end
function exec(expr::String, ctx::ExcelContext)
    expr
end


cell_parse_rgx = r"\$?([A-Z]+)\$?([0-9]+)"
function old_parse_cell(cell)
    cell_match = match(cell_parse_rgx, cell)
    @assert cell_match.match == cell "Cell didn't parse properly"
    col_str = cell_match[1]
    row_str = cell_match[2]
    (XLSX.decode_column_number(col_str), parse(Int, row_str))
end

function parse_cell(cell::T) where {T <: AbstractString}
    first_let = findfirst(c -> 'A' <= c <= 'Z', cell)
    first_num = findfirst(isdigit, cell)
    # cell_match = match(cell_parse_rgx, cell)
    col_str = if cell[first_num-1] == '$'
        @view cell[first_let:(first_num-2)]
    else
        @view cell[first_let:(first_num-1)]
    end
    # @assert cell_match.match == cell "Cell didn't parse properly"
    # col_str = cell_match[1]
    # row_str = cell_match[2]
    row_str = @view cell[first_num:end]
    (XLSX.decode_column_number(col_str), parse(Int, row_str))
end


function clean_cell(cell)
    replace(cell, '$' => "")
end
function index_to_cellname(col, row)
    XLSX.encode_column_number(col) * string(row)
end
function rangetomatrix(ctx::ExcelContext, lhs::String, rhs::String)
    start_col, start_row = parse_cell(lhs)
    end_col, end_row = parse_cell(rhs)

    @assert end_row >= start_row
    @assert end_col >= start_col

    output = []
    for col ∈ start_col:end_col
        row_values = (r -> get_cell_value(ctx, index_to_cellname(col, r))).(start_row:end_row)
        push!(output, row_values)
    end
    reduce(hcat, output)
end

missing_to(value, default = 0.0) = ismissing(value) ? default : value

function eval(value, ctx::ExcelContext)
    value
end
function eval(expr::ExcelExpr, ctx::ExcelContext)
    missing_to(exec(expr, ctx))
end
function exec(expr::ExcelExpr, ctx::ExcelContext)
    result = @match expr begin
        ExcelExpr(:cell_ref, [cell]) => get_cell_value(ctx, clean_cell(cell))
        ExcelExpr(:cell_ref, [cell, sheet]) => get_cell_value(ctx, clean_cell(cell))
        ExcelExpr(:+, [unary]) => exec(unary, ctx)
        ExcelExpr(:+, [lhs, rhs]) => xl_add(missing_to(exec(lhs, ctx)), missing_to(exec(rhs, ctx)))
        ExcelExpr(:-, [unary]) => -1 * exec(unary, ctx)
        ExcelExpr(:-, [lhs, rhs]) => xl_sub(missing_to(exec(lhs, ctx)), missing_to(exec(rhs, ctx)))
        ExcelExpr(:*, [lhs, rhs]) => missing_to(exec(lhs, ctx)) * missing_to(exec(rhs, ctx))
        ExcelExpr(:/, [lhs, rhs]) => missing_to(exec(lhs, ctx)) / missing_to(exec(rhs, ctx))
        ExcelExpr(:^, [lhs, rhs]) => exec(lhs, ctx)^exec(rhs, ctx)
        ExcelExpr(:&, [lhs, rhs]) => string(exec(lhs, ctx), exec(rhs, ctx))
        ExcelExpr(:eq, [lhs, rhs]) => xl_eq(exec(lhs, ctx), exec(rhs, ctx))
        ExcelExpr(:neq, [lhs, rhs]) => !xl_eq(exec(lhs, ctx), exec(rhs, ctx))
        ExcelExpr(:leq, [lhs, rhs]) => xl_leq(exec(lhs, ctx), exec(rhs, ctx))
        ExcelExpr(:geq, [lhs, rhs]) => xl_geq(exec(lhs, ctx), exec(rhs, ctx))
        ExcelExpr(:lt, [lhs, rhs]) => xl_lt(exec(lhs, ctx), exec(rhs, ctx))
        ExcelExpr(:gt, [lhs, rhs]) => xl_gt(exec(lhs, ctx), exec(rhs, ctx))
        ExcelExpr(:range, [lhs, rhs]) => rangetomatrix(ctx, exec_to_cell(lhs), exec_to_cell(rhs))
        ExcelExpr(:sheet_ref, [sheet_name, ref]) => exec(ref, with_current_sheet(ctx, sheet_name))
        ExcelExpr(:named_range, [name]) => exec(get_key_value(ctx, name), ctx)
        ExcelExpr(:call, ["IF", args...]) => xl_logical(eval(args[1], ctx)) ? exec(args[2], ctx) : exec(args[3], ctx)
        ExcelExpr(:call, [fn_name, args...]) => eval_function(ctx, fn_name, args)
    end
    # println("$(expr) = $(result)")
    result
end

xl_logical(v::Bool) = v
xl_logical(v::Number) = v == 1
xl_logical(v::String) = v == "1"

function eval_function(ctx::ExcelContext, fn_name, raw_args)
    args = map((a -> exec(a, ctx)), raw_args)
    @match fn_name begin
        "ABS" => abs(args...)
        "AVERAGE" => xl_average(args...)
        "SUM" => xl_sum(args...)
        "SUMPRODUCT" => xl_sum_product(args...)
        "SQRT" => sqrt(args...)
        "AND" => all(args)
        "RAND" => rand()
        "ROUND" => round(args[1], RoundNearestTiesUp, digits = Int(args[2]))
        "ROUNDUP" => round(args[1], RoundFromZero, digits = Int(args[2]))
        "ROUNDDOWN" => round(args[1], RoundToZero, digits = Int(args[2]))
        "FLOOR" => xl_floor(args...)
        "MAX" => xl_max(ctx, raw_args...)
        "MIN" => xl_min(args...)
        "MEDIAN" => xl_median(args...)
        "PMT" => xl_pmt(args...)
        "PRODUCT" => xl_product(args...)
        "_xlfn.STDEV.S" => xl_stdev(args...)
        "_xlfn.XLOOKUP" => xl_xlookup(args...)
        "INDIRECT" => exec(toexpr(parse_formula(exec(args[1], ctx))), ctx)
        "COUNTIFS" => xl_countifs(ctx, args...)
        "EXP" => exp(args...)
        "DATE" => xl_date(args...)
        "EDATE" => xl_edate(args...)
        "MOD" => xl_mod(args...)
        "PI" => π
        "LINEST" => xl_linest(args...)
    end
end

asarray(x) = [x]

function asarray(a::T) where {T <: Number}
    [a]
end
function asarray(x::Vector{Vector})
    reduce(vcat, x)
end
function asarray(x::AbstractArray)
    x
end

function xl_sum_inner(val::Number)
    val
end
function xl_sum_inner(val::Missing)
    0.0
end

function xl_sum_inner(arr::AbstractArray)
    sum(xl_sum_inner, arr)
end

function xl_sum_inner(arr::AbstractDataFrame)
    xl_sum_inner(xl_sum_inner.(eachcol(arr)))
    # xl_sum_inner(Matrix(arr))
end

function xl_sum_inner(arr::DataFrameRow)
    sum(map(xl_sum_inner, arr))
    # xl_sum_inner(Matrix(arr))
end
function xl_sum(args...)
    predicate = x -> !ismissing(x)
    sum(xl_sum_inner, args)
    # sum(map(xl_sum_inner, args))
    # each_arg = (a -> xl_sum(filter(predicate, asarray(a)))).(args)
    # # @show each_arg typeof(each_arg)
    # flattened = reduce(vcat, each_arg)
    # if ismissing(flattened)
    #     0.0
    # else
    #     sum(flattened)
    # end
end

function xl_product(args...)
    predicate = x -> !ismissing(x)
    each_arg = (a -> filter(predicate, asarray(a))).(args)
    reduce(*, (reduce(vcat, each_arg)))
end

function xl_sum_product(args...)
    sum(reduce(.*, args))
end

function xl_average(args...)
    predicate = x -> !ismissing(x)
    filtered_args = reduce(vcat, (a -> filter(predicate, asarray(a))).(args))
    sum(filtered_args) / length(filtered_args)
end

xl_mod(a, b) = mod(a, b)
xl_mod(a::AbstractArray, b::T) where {T <: AbstractFloat} = mod.(a, b)
xl_mod(a::T, b::AbstractArray) where {T <: AbstractFloat} = mod.(a, b)

function xl_pmt(rate, nper, pv)
    -1 * sign(pv) * (pv * rate) / (1 - (1 + rate)^(-nper))
end

function xl_median(args...)
    predicate = x -> !ismissing(x)
    filtered_args = reduce(vcat, (a -> filter(predicate, asarray(a))).(args))
    median(filtered_args)
end

function xl_min(v::T) where {T <: Number}
    v
end
function xl_min(arr::AbstractArray)
    minimum(arr)
end

function xl_min(args...)
    minimum(map(xl_min, args))
end
function xl_max(v::T) where {T <: Number}
    v
end
notmissing(x) = !ismissing(x)
# function xl_max(arr::AbstractArray)
#     maximum(filter(notmissing, arr))
# end


xl_max2(a::Number, b::Number) = max(a, b)
xl_max2(a::Number, b::Date) = max(xl_num_to_date(a), b)
xl_max2(a::Date, b::Number) = xl_max(b, a)
function xl_max(args...)
    res = 0
    for a in args
        if a isa Bool
            res = max(res, a)
        elseif a isa String
            try
                v = parse(Float64, a)
                res = max(v, res)
            catch
                continue
            end
        else
            values = asarray(a)
            predicate = x -> ((x isa Number) && !(x isa Bool)) || (x isa Date)
            filtered = filter(predicate, collect(Base.Flatten(values)))
            # @show filtered
            max_values = maximum(filtered)
            res = xl_max2(res, max_values)
        end
    end
    res

end

function xl_max(ctx::ExcelContext, raw_args...)
    raw_max = 0
    for a in raw_args
        if a isa Bool
            raw_max = max(raw_max, a)
        elseif a isa String
            try
                v = parse(Float64, a)
                raw_max = max(v, raw_max)
            catch
                continue
            end
        elseif a isa ExcelExpr
            values = asarray(exec(a, ctx))
            predicate = x -> ((x isa Number) && !(x isa Bool)) || (x isa Date)
            max_values = maximum(filter(predicate, values))
            raw_max = xl_max(raw_max, max_values)
        end
    end
    raw_max
end


function xl_stdev(args...)
    predicate = x -> !ismissing(x)
    filtered_args = reduce(vcat, (a -> filter(predicate, asarray(a))).(args))
    std(filtered_args)
end

function xl_countifs(ctx, args...)
    # @info "xl_countifs" args
    @assert iseven(length(args))
    values = args[1:2:end]
    predicates = args[2:2:end]

    results = []
    for (list, pred) in zip(values, predicates)
        mask = [exec(toexpr(parse_formula(string(v, pred))), ctx) for v in list]
        push!(results, mask)
        # for v in list
        #     formula = string(v, pred)
        #     result = exec(toexpr(parse_formula(formula)), ctx)

        # end
    end

    sum(reduce(.&, results))
end

function xl_date_to_num(date::Date)
    date - Dates.Date(1900, 1, 1) + Dates.Day(2)
end

function xl_num_to_date(a::Number)
    # This is only corerct because excel thinks 1900 is a leap year.
    # So any dates before march 1st 1900 wil be off by one
    Date(1900) + Dates.Day(a)
end

function xl_date(year, month, day)
    xl_date_to_num(Dates.Date(year, month, day))
end

function xl_edate(start::Dates.Date, offset)
    start + Dates.Month(offset)
end

function xl_xlookup(lookup_value::AbstractArray, lookup_array, return_array)
    map(v -> xl_xlookup(v, lookup_array, return_array), lookup_value)
end

function xl_xlookup(lookup_value, lookup_array, return_array)
    for i ∈ eachindex(lookup_array)
        if xl_eq(lookup_value, lookup_array[i])
            return return_array[i]
        end
    end
    missing
end

function xl_floor(x, step)
    floor(x / step) * step
end
function xl_ceiling(x, step)
    ceil(x / step) * step
end

function xl_linest(y, x)
    # y = convert(Matrix{Float64}, y)
    y = reshape(y, 1, :)
    # x = convert(Matrix{Float64}, x)
    x = reshape(x, 1, :)

    X = hcat(x', ones(size(x))')
    ((X'*X)\(X'*y'))[1]
end

function exec_to_cell(expr::ExcelExpr)
    @match expr begin
        ExcelExpr(:cell_ref, [cell]) => cell
    end
end


xl_compare(v1, v2) = v1 == v2
xl_compare(v1::Number, v2::Number) = isapprox(v1, v2, atol = 1e-10, rtol = √eps())
xl_compare(v1::Number, v2::Dates.Day) = v1 == Dates.value(v2)
xl_compare(v1::Dates.Day, v2::Number) = Dates.value(v1) == v2
xl_compare(v1::String, v2::String) = v1 == v2
xl_compare(::Missing, v::String) = v == ""
xl_compare(v::String, ::Missing) = v == ""
xl_compare(v1::Missing, v2::Missing) = true
xl_compare(v1::Missing, v2) = false
xl_compare(v1, v2::Missing) = false

xl_atan(a) = atan(a)
xl_sin(a) = sin(a)
xl_asin(a) = asin(a)
xl_cos(a) = cos(a)
xl_tan(a) = tan(a)
xl_radians(a) = deg2rad(a)