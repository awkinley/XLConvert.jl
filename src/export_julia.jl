function getdatatype(ws::XLSX.Worksheet, cell::XLSX.Cell)

    if XLSX.iserror(cell)
        return Any
    end

    if isempty(cell.value)
        return Missing
    end

    if cell.datatype == "inlineStr"
        return String
    end

    if cell.datatype == "s"
        return String
    elseif (isempty(cell.datatype) || cell.datatype == "n")
        if !isempty(cell.style) && XLSX.styles_is_datetime(ws, cell.style)
            # datetime
            return Dates.Date
        else
            # float
            return Float64
        end
    elseif cell.datatype == "b"
        return Bool
    elseif cell.datatype == "str"
        return String
    end

    error("Couldn't parse type for $cell.")
end

abstract type AbstractHandler end

struct JuliaExporter3
    wb::Any
    var_names::Dict{CellDependency, String}
    tables::Vector{ExcelTable}
    named_values::Any
    handlers::Any
    cell_types::Any
end

JuliaExporter = JuliaExporter3

function handle(handler::AbstractHandler, expr, exporter::JuliaExporter, ctx)
    missing
end


struct BasicOpHandler <: AbstractHandler end

is_number_type(expr, ctx) = false
is_number_type(expr::Float64, ctx) = true
is_number_type(expr::Int64, ctx) = true
is_number_type(exporter::JuliaExporter, expr, ctx) = get_type(expr, sheetname(ctx), exporter.cell_types, exporter.named_values) == Float64

function op_binding_affinity(op::Symbol)
    @match op begin
        :+ => (2, 3)
        :- => (2, 3)
        :* => (4, 5)
        :/ => (4, 5)
        :^ => (6, 7)
    end
end

function handle(::BasicOpHandler, expr::ExcelExpr, exporter::JuliaExporter, ctx)
    c = e -> convert(exporter, e, withbindingaffinity(ctx, -1))
    @match expr begin

        ExcelExpr(op, [lhs, rhs]), if op ∈ (:+, :-, :*, :/)
        end => begin

            lhs_number = false
            rhs_number = false
            try
                lhs_number = is_number_type(exporter, lhs, ctx)
                rhs_number = is_number_type(exporter, rhs, ctx)
            catch
            end


            if !(lhs_number && rhs_number)
                func_name = @match op begin
                    :+ => "xl_add"
                    :- => "xl_sub"
                    :* => "xl_mul"
                    :/ => "xl_div"
                end

                return "$func_name($(c(lhs)), $(c(rhs)))"
            end

            left_affinity, right_affinity = op_binding_affinity(op)
            wrap_parens = left_affinity < bindingaffinity(ctx)

            lhs_str = convert(exporter, lhs, withbindingaffinity(ctx, left_affinity))
            rhs_str = convert(exporter, rhs, withbindingaffinity(ctx, right_affinity))
            binop_str = "$lhs_str $(string(op)) $rhs_str"

            if wrap_parens
                "(" * binop_str * ")"
            else
                binop_str
            end
        end
        ExcelExpr(:+, [unary]) => c(unary)
        ExcelExpr(:-, [unary]) => "(-1 * " * c(unary) * ")"
        ExcelExpr(:^, [lhs, rhs]) => "(($(c(lhs))) ^ ($(c(rhs))))"
        ExcelExpr(:%, [unary]) => "(($(c(unary))) / 100.0)"
        ExcelExpr(:&, [lhs, rhs]) => "($(c(lhs))* $(c(rhs)))"
        ExcelExpr(:eq, [lhs, rhs]) => "xl_eq($(c(lhs)), $(c(rhs)))"
        ExcelExpr(:neq, [lhs, rhs]) => "!xl_eq($(c(lhs)), $(c(rhs)))"
        ExcelExpr(:leq, [lhs, rhs]) => "xl_leq($(c(lhs)), $(c(rhs)))"
        ExcelExpr(:geq, [lhs, rhs]) => "xl_geq($(c(lhs)), $(c(rhs)))"
        ExcelExpr(:lt, [lhs, rhs]) => "xl_lt($(c(lhs)), $(c(rhs)))"
        ExcelExpr(:gt, [lhs, rhs]) => "xl_gt($(c(lhs)), $(c(rhs)))"
        _ => missing
    end
end


struct TableRefHandler <: AbstractHandler end


function handle(::TableRefHandler, expr::ExcelExpr, exporter::JuliaExporter, ctx)
    @match expr begin
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => begin
            # row_idx_str = if row_idx == 1:size(table)(1)
            #     ":"
            # elseif row_idx isa Int && length(col_idx) > 1
            #     # Want to avoid slicing into a DataFrameRow, becaue that doesn't broadcast
            #     repr(row_idx:row_idx)
            # else
            #     repr(row_idx)
            # end

            # col_idx_str = if col_idx == 1:size(table)[1]
            #     ":"
            # elseif col_idx isa Int && length(col_idx) > 1
            #     # Want to avoid slicing into a DataFramecol, becaue that doesn't broadcast
            #     repr(col_idx:col_idx)
            # else
            #     repr(col_idx)
            # end

            # "$(getname(table))[$row_idx_str, $col_idx_str]"
            col_name = [string(column_name(table, c)) for c in col_idx]
            # if length(col_name) == 1
            #     col_name = col_name[1]
            # end

            # if row_idx isa Int
            #     row_idx = row_idx:row_idx
            # end

            # row_idx_str = if row_idx isa UnitRange{Int} && row_idx.start == 1 && row_idx.stop == size(table)[1]
            row_idx_str = if row_idx == 1:size(table)[1]
                "!"
            elseif row_idx isa Int && length(col_idx) > 1
                # Want to avoid slicing into a DataFrameRow, becaue that doesn't broadcast
                repr(row_idx:row_idx)
                # elseif row_idx isa UnitRange{Int} && row_idx.start == row_idx.stop
                #     repr(row_idx.start)
            else
                repr(row_idx)
            end

            col_idx_str = if col_name isa AbstractArray && length(col_name) == 1
                repr(string(col_name[1]))
            elseif col_idx isa UnitRange
                "Between($(repr(col_name[begin])), $(repr(col_name[end])))"
            elseif col_name isa AbstractArray
                repr(col_name)
            elseif length(col_name) == 1
                repr(string(col_name[1]))
            else
                repr(string(col_name))
            end

            "$(getname(table))[$row_idx_str, $col_idx_str]"
        end
        _ => missing
    end
end





# function xl_expr_to_julia(expr::String, ctx, var_names, tables)
#     "\"" * expr * "\""
# end

# function xl_expr_to_julia(expr, ctx, var_names, tables)
#     string(expr)
#     # "xl($(string(expr)))"
# end

function xl_call_to_julia(fn_name, args)
    complex = @match fn_name begin
        # "AND" => "all(($(join(args, ", "))))"
        "RAND" => "rand()"
        "ROUND" => "round($(args[1]), RoundNearestTiesUp, digits=Int($(args[2])))"
        "ROUNDUP" => "round($(args[1]), RoundFromZero, digits=Int($(args[2])))"
        "ROUNDDOWN" => "round($(args[1]), RoundToZero, digits=Int($(args[2])))"
        # "MAX" => "xl_max(ctx, $(args)...)"
        "INDIRECT" => "exec(toexpr(parse_formula(exec($(args[1]), ctx))), ctx)"
        "COUNTIFS" => "xl_countifs(ctx, $(args)...)"
        "AND" => begin
            wrap_logical = s -> "xl_logical($s)"
            "all([" * join(map(wrap_logical, args), ", ") * "])"
        end
        "OR" => begin
            wrap_logical = s -> "xl_logical($s)"
            "any([" * join(map(wrap_logical, args), ", ") * "])"
        end
        "NOT" => "(!($(args[1])))"
        _ => missing
    end

    if !ismissing(complex)
        return complex
    end

    params = "($(join(args, ", ")))"
    @match fn_name begin
        "MAX" => "xl_max" * params
        "ABS" => "abs" * params
        "AVERAGE" => "xl_average" * params
        "SUM" => "xl_sum" * params
        "SUMPRODUCT" => "xl_sum_product" * params
        "SQRT" => "sqrt" * params
        # "AND" => "all" * params
        # "OR" => "any(" * params * ")"
        "FLOOR" => "xl_floor" * params
        "CEILING" => "xl_ceiling" * params
        "MIN" => "xl_min" * params
        "MEDIAN" => "xl_median" * params
        "PMT" => "xl_pmt" * params
        "PRODUCT" => "xl_product" * params
        "_xlfn.STDEV.S" => "xl_stdev" * params
        "_xlfn.XLOOKUP" => "xl_xlookup" * params
        "EXP" => "exp" * params
        "DATE" => "xl_date" * params
        "EDATE" => "xl_edate" * params
        "MOD" => "xl_mod" * params
        "PI" => "π"
        "LINEST" => "xl_linest" * params
        "IF_MULTIPLE" => "if_multiple" * params
        "VLOOKUP" => "xl_vlookup" * params
        "TEXT" => "xl_text" * params
        "ATAN" => "xl_atan" * params
        "SIN" => "xl_sin" * params
        "ASIN" => "xl_asin" * params
        "COS" => "xl_cos" * params
        "TAN" => "xl_tan" * params
        "RADIANS" => "xl_radians" * params
        # "IF_ELSE_FALSE" => "if_else_false" * params
        "ROUNDUP_IF" => "xl_roundup_if" * params
        # TODO: implement these
        "MATCH" => "xl_match" * params
        "INDEX" => "xl_index" * params
        "LOOKUP" => "xl_lookup" * params
        "OFFSET" => "xl_offset" * params
        "COUNTA" => "xl_counta" * params
        "NPV" => "xl_npv" * params
        fn_name => begin
            println("Function $fn_name not handled!")
            "xl_" * lowercase(fn_name) * params
        end
    end
end

struct EverythingElseHandler <: AbstractHandler end

function handle(::EverythingElseHandler, expr::ExcelExpr, exporter::JuliaExporter, ctx)
    if @ismatch expr ExcelExpr(:range, [ExcelExpr(:cell_ref, [lhs, sheet]), ExcelExpr(:cell_ref, [rhs, sheet])])
        start_col, start_row = parse_cell(lhs)
        end_col, end_row = parse_cell(rhs)

        # @assert end_row >= start_row
        # @assert end_col >= start_col
        for table in exporter.tables
            if (sheetname(ctx) == table.sheet_name
                && start_col >= startcol(table)
                && end_col <= endcol(table)
                && start_row >= startrow(table)
                && end_row <= endrow(table))
                # println("Found table reference!")
                # @show expr table
            end
        end

        return if start_row == end_row
            col_values = (c -> exporter.var_names[CellDependency(sheet, index_to_cellname(c, start_row))]).(start_col:end_col)
            "[" * join(col_values, ", ") * "]"
        elseif start_col == end_col
            row_values = (r -> exporter.var_names[CellDependency(sheet, index_to_cellname(start_col, r))]).(start_row:end_row)
            "[" * join(row_values, ", ") * "]"
        else
            output = []
            for col ∈ start_col:end_col
                row_values = (r -> exporter.var_names[CellDependency(sheet, index_to_cellname(col, r))]).(start_row:end_row)
                push!(output, "[" * join(row_values, ", ") * "]")
            end
            "[" * join(output, " ") * "]"
        end
    end
    func = a -> convert(exporter, a, ctx)
    # @show expr

    @match expr begin
        ExcelExpr(:func_param, [param_num]) => "param_$param_num"
        ExcelExpr(:func_param, [param_num, _type]) => "param_$param_num"
        ExcelExpr(:cell_ref, [cell, sheet]) => exporter.var_names[CellDependency(sheet, cell)]
        ExcelExpr(:named_range, [name]) => convert(exporter, get(exporter.named_values, name, "undef_var_$name"), ctx)
        ExcelExpr(:call, ["IF", cond, t, f]) => begin
            cond_is_bool = get_type(cond, sheetname(ctx), exporter.cell_types, exporter.named_values) == Bool
            if cond_is_bool
                "($(func(cond)) ? $(func(t)) : $(func(f)))"
            else
                "(xl_logical($(func(cond))) ? $(func(t)) : $(func(f)))"
            end
        end
        ExcelExpr(:call, [fn_name, args...]) => xl_call_to_julia(fn_name, map(func, args))
        ExcelExpr(:broadcast_protect, [expr]) => "($(func(expr)),)"
        ExcelExpr(:cols, [sheet, columns]) => "columns($(repr(sheet)), $(repr(columns)))"
        _ => missing
    end
end

mutable struct JlExporterCtx
    current_sheet::String
    last_binding_affinity::Int
end

sheetname(ctx::JlExporterCtx) = ctx.current_sheet
function withsheet(ctx::JlExporterCtx, sheet::AbstractString)
    JlExporterCtx(
        sheet,
        ctx.last_binding_affinity,
    )
end

bindingaffinity(ctx::JlExporterCtx) = ctx.last_binding_affinity
function withbindingaffinity(ctx::JlExporterCtx, affinity::Int)
    JlExporterCtx(
        sheetname(ctx),
        affinity,
    )
end
# function setsheet!(ctx::JlExporterCtx, sheet::AbstractString) 
#     ctx.current_sheet = string(sheet)
#     ctx
# end



function convert(exporter::JuliaExporter, expr, ctx::JlExporterCtx)
    if expr isa String
        jl_str = escape_string(expr)
        jl_str = replace(jl_str, "\$" => "\\\$")
        "\"" * jl_str * "\""
    elseif expr isa Int
        repr(Float64(expr))
    else
        repr(expr)
    end
end

function convert(exporter::JuliaExporter, expr::FlatExpr, ctx::JlExporterCtx)
    convert(exporter, convert_to_expr(expr), ctx)
end
function convert(exporter::JuliaExporter, expr::ExcelExpr, ctx::JlExporterCtx)
    # @info "convert" ctx expr
    if expr.head == :sheet_ref
        sheet_name, ref = expr.args
        # setsheet!(ctx, sheet_name)
        # @info "After :sheet_ref" sheet_name sheetname(ctx)
        return convert(exporter, ref, withsheet(ctx, sheet_name))
    end

    for handler in exporter.handlers
        res = handle(handler, expr, exporter, ctx)
        if !ismissing(res)
            if occursin("ExcelExpr", res)
                @show res
            end
            return res
        end
    end


    @show expr
    # string(expr)
    throw("Failed to handle an excel expr: $expr")
end

convert(exporter::JuliaExporter, expr, current_sheet::AbstractString) = convert(exporter, expr, JlExporterCtx(current_sheet, -1))

function make_struct(exporter::JuliaExporter, struct_name::AbstractString, var_names::Vector{<:AbstractString}; var_types::Union{Nothing, Vector} = nothing, ismutable = false, default_values = nothing, var_comments::Union{Nothing, Vector} = nothing)
    lines = Vector{String}()
    sizehint!(lines, 2 + length(var_names))
    def_terms = Vector{String}()
    if !isnothing(default_values)
        push!(def_terms, "@kwdef")
    end
    if ismutable
        push!(def_terms, "mutable")
    end

    push!(def_terms, "struct $struct_name")
    push!(lines, join(def_terms, " "))

    for (i, name) in enumerate(var_names)
        line = "\t$name"
        if !isnothing(var_types)
            type = var_types[i]
            if !ismissing(type)
                line *= "::$type"
            end
        end

        if !isnothing(default_values)
            default_val = default_values[i]
            if !ismissing(default_val)
                line *= " = $default_val"
            end
        end

        if !isnothing(var_comments)
            comment = var_comments[i]
            if !isnothing(comment)
                push!(lines, "\t# " * comment)
            end
        end

        push!(lines, line)
    end
    push!(lines, "end\n")

    join(lines, "\n")
end
