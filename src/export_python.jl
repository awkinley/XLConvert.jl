struct PythonExporter
    wb::Any
    var_names::Dict{CellDependency, String}
    tables::Vector{ExcelTable}
    named_values::Any
    handlers::Any
    cell_types::Any
end

function with_handler(exporter::PythonExporter, new_handler)
    PythonExporter(exporter.wb, exporter.var_names, exporter.tables, exporter.named_values, [new_handler, exporter.handlers...], exporter.cell_types)
end




function handle(handler::AbstractHandler, expr, exporter::PythonExporter, ctx)
    missing
end

struct ColRowNameHandler <: AbstractHandler
    row_name::String
    column_name::String
end

function handle(handler::ColRowNameHandler, expr, exporter::PythonExporter, ctx)
    @match expr begin
        ExcelExpr(:column_name, []) => handler.column_name
        ExcelExpr(:row_name, []) => handler.row_name
        _ => missing
    end
end

get_type(expr, exporter::PythonExporter, ctx) = get_type(expr, sheetname(ctx), exporter.cell_types, exporter.named_values)
is_number_type(exporter::PythonExporter, expr, ctx) = get_type(expr, exporter, ctx) == Float64


# function op_binding_affinity(op::Symbol)
#     @match op begin
#         :+ => (2, 3)
#         :- => (2, 3)
#         :* => (4, 5)
#         :/ => (4, 5)
#         :^ => (6, 7)
#     end
# end

is_date_type(expr, ctx) = false
is_date_type(expr::Dates.Date, ctx) = true
is_date_type(expr::Dates.DateTime, ctx) = true
is_date_type(exporter::PythonExporter, expr, ctx) = get_type(expr, sheetname(ctx), exporter.cell_types, exporter.named_values) in (Dates.Date, Dates.DateTime)

wrap_parens(s::AbstractString) = string('(', s, ')')

function convert_wrapped(exporter::PythonExporter, expr, binding_strength, ctx)
    do_wrap = binding_strength < bindingaffinity(ctx)
    new_ctx = do_wrap ? withbindingaffinity(ctx, -1) : withbindingaffinity(ctx, binding_strength)
    str = convert(exporter, expr, new_ctx)
    if do_wrap
        wrap_parens(str)
    else
        str
    end
end


function handle(::BasicOpHandler, expr::ExcelExpr, exporter::PythonExporter, ctx)
    c = e -> convert(exporter, e, withbindingaffinity(ctx, -1))
    @match expr begin

        ExcelExpr(op, [lhs, rhs]), if op ∈ (:+, :-, :*, :/)
        end => begin
            left_affinity, right_affinity = op_binding_affinity(op)
            wrap_parens = left_affinity < bindingaffinity(ctx)

            lhs_str = convert(exporter, lhs, withbindingaffinity(ctx, left_affinity))
            rhs_str = convert(exporter, rhs, withbindingaffinity(ctx, right_affinity))

            if op in (:+, :-) && (is_date_type(exporter, lhs, ctx) || is_date_type(exporter, rhs, ctx))
                # rhs_str = "datetime.timedelta(days=$rhs_str)" 
                if op == :+
                    "xl.date_add($lhs_str, $rhs_str)" 
                else
                    "xl.date_sub($lhs_str, $rhs_str)" 
                end
            else
                binop_str = "$lhs_str $(string(op)) $rhs_str"

                if wrap_parens
                    "(" * binop_str * ")"
                else
                    binop_str
                end
            end
        end
        ExcelExpr(:+, [unary]) => c(unary)
        ExcelExpr(:-, [unary]) => "(-1 * " * c(unary) * ")"
        ExcelExpr(:^, [lhs, rhs]) => begin
            lhs_str = convert_wrapped(exporter, lhs, 6, ctx)
            rhs_str = convert_wrapped(exporter, rhs, 7, ctx)
            # "(($(c(lhs))) ** ($(c(rhs))))"
            "$lhs_str ** $rhs_str"
        end
        ExcelExpr(:%, [unary]) => "(($(c(unary))) / 100.0)"
        ExcelExpr(:&, [lhs, rhs]) => "xl.concat($(c(lhs)), $(c(rhs)))"
        ExcelExpr(cmp_op, [lhs, rhs]) where cmp_op ∈ (:eq, :neq, :leq, :geq, :lt, :gt) => begin
            lhs_number = false
            rhs_number = false
            try
                lhs_number = is_number_type(exporter, lhs, ctx)
                rhs_number = is_number_type(exporter, rhs, ctx)
            catch
            end

            # are_num = lhs_number && rhs_number
            are_num = false

            if are_num
                infix_op = @match cmp_op begin
                    :eq => "=="
                    :neq => "!="
                    :leq => "<="
                    :geq => ">="
                    :lt => "<"
                    :gt => ">"
                end

                lhs_str = convert_wrapped(exporter, lhs, 0, ctx)
                rhs_str = convert_wrapped(exporter, rhs, 0, ctx)

                "$lhs_str $infix_op $rhs_str"
            else
                func = @match cmp_op begin
                    :eq => "xl.eq"
                    :neq => "not xl.eq"
                    :leq => "xl.leq"
                    :geq => "xl.geq"
                    :lt => "xl.lt"
                    :gt => "xl.gt"
                end

                "$func($(c(lhs)), $(c(rhs)))"
            end
        end
        # ExcelExpr(:eq, [lhs, rhs]) => "xl.eq($(c(lhs)), $(c(rhs)))"
        # ExcelExpr(:neq, [lhs, rhs]) => "not xl.eq($(c(lhs)), $(c(rhs)))"
        # ExcelExpr(:leq, [lhs, rhs]) => "xl.leq($(c(lhs)), $(c(rhs)))"
        # ExcelExpr(:geq, [lhs, rhs]) => "xl.geq($(c(lhs)), $(c(rhs)))"
        # ExcelExpr(:lt, [lhs, rhs]) => "xl.lt($(c(lhs)), $(c(rhs)))"
        # ExcelExpr(:gt, [lhs, rhs]) => "xl.gt($(c(lhs)), $(c(rhs)))"
        _ => missing
    end
end



function handle(::TableRefHandler, expr::ExcelExpr, exporter::PythonExporter, ctx)
    @match expr begin
        ExcelExpr(:table_ref_col, [table, row_idx, col_idx]) => begin
            # col_name = [string(column_name(table, c)) for c in col_idx]

            row_names = row_name.(Ref(table), row_idx)
            row_idx_str = if row_idx == 1:size(table)[1]
                @show col_idx
                if length(col_idx) == 1
                    return "$(getname(table))[$col_idx_str]"
                else
                    ":"
                end
            elseif length(row_idx) > 1
                # Want to avoid slicing into a DataFrameRow, becaue that doesn't broadcast
                "$(repr(row_names[begin])):$(repr(row_names[end]))"
            else
                name = row_names isa AbstractString ? row_names : first(row_names)
                if name isa Integer
                    string(name)
                else
                    repr(name)
                end
            end

            col_val = convert(exporter, col_idx, ctx)
            "$(getname(table)).loc[$row_idx_str, str($col_val)]"
        end
        ExcelExpr(:table_ref_idx, [table, row_idx, col_idx]) => begin
            col_name = [string(column_name(table, c)) for c in col_idx]

            col_idx_str = if col_name isa AbstractArray && length(col_name) == 1
                repr(string(col_name[1]))
            elseif col_idx isa UnitRange
                throw("table_ref_idx can't handle a col_idx that is a unitrange")
            elseif col_name isa AbstractArray
                repr(col_name)
            elseif length(col_name) == 1
                repr(string(col_name[1]))
            else
                repr(string(col_name))
            end
            # @show row_idx
            need_values = false
            if @ismatch row_idx ExcelExpr(:table_ref, [inner_table, inner_row, inner_col, _, _])
                if length(inner_row) == size(inner_table)[1] && length(inner_col) == 1
                    inner_col_name = repr(column_name(inner_table, first(inner_col)))
                    return "$(getname(inner_table)).join($(getname(table)), on=$inner_col_name, rsuffix=\"_right\")[$col_idx_str]"
                end

                need_values = length(inner_row) != 1 || length(inner_col) != 1
            end
            row_val = convert(exporter, row_idx, ctx)
            if need_values
                "$(getname(table)).loc[$row_val, $col_idx_str].values"
            else
                "$(getname(table)).loc[$row_val, $col_idx_str]"
            end
        end
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => begin
            if is_transposed(table)
                (row_idx, col_idx) = (col_idx, row_idx)
            end

            # "$(getname(table))[$row_idx_str, $col_idx_str]"
            col_name = [string(column_name(table, c)) for c in col_idx]
            # if length(col_name) == 1
            #     col_name = col_name[1]
            # end

            # if row_idx isa Int
            #     row_idx = row_idx:row_idx
            # end

            # row_idx_str = if row_idx isa UnitRange{Int} && row_idx.start == 1 && row_idx.stop == size(table)[1]

            col_idx_str = if col_name isa AbstractArray && length(col_name) == 1
                repr(string(col_name[1]))
            elseif col_idx isa UnitRange
                # if length(col_idx) == 1 && size(table)[2] == 1
                #     repr(col_name[begin])
                if length(col_idx) == size(table)[2]
                    ":"
                else
                    "$(repr(col_name[begin])):$(repr(col_name[end]))"
                end
            elseif col_name isa AbstractArray
                repr(col_name)
            elseif length(col_name) == 1
                repr(string(col_name[1]))
            else
                repr(string(col_name))
            end

            row_names = row_name.(Ref(table), row_idx)
            # row_names = [string(row_name(table, r)) for r in row_idx]
                
            row_idx_str = if row_idx == 1:size(table)[1]
                if length(col_idx) == 1
                    return "$(getname(table))[$col_idx_str]"
                # elseif length(row_idx) == 1
                #     repr(row_names[begin])
                else
                    ":"
                end
            # elseif row_idx isa Int && length(row_idx) > 1
            elseif length(row_idx) > 1
                # Want to avoid slicing into a DataFrameRow, becaue that doesn't broadcast
                "$(repr(row_names[begin])):$(repr(row_names[end]))"
                # repr(row_idx:row_idx)
                # elseif row_idx isa UnitRange{Int} && row_idx.start == row_idx.stop
                #     repr(row_idx.start)
            else
                name = row_names isa AbstractString ? row_names : first(row_names)
                # name = first(row_names)
                if name isa Integer
                    string(name)
                else
                    repr(name)
                end
                # repr(row_names[1])
                # repr(row_idx)
            end

            indexer = if length(row_idx) == 1 && length(col_idx) == 1
                "at"
            else
                "loc"
            end

            "$(getname(table)).$indexer[$row_idx_str, $col_idx_str]"
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

function xl_call_to_python(fn_name, args)
    complex = @match fn_name begin
        # "AND" => "all(($(join(args, ", "))))"
        "RAND" => "rand()"
        # "ROUND" => "round($(args[1]), RoundNearestTiesUp, digits=Int($(args[2])))"
        # "ROUND" => "xl_round($(args[1]))"
        # "ROUNDUP" => "round($(args[1]), RoundFromZero, digits=Int($(args[2])))"
        # "ROUNDDOWN" => "round($(args[1]), RoundToZero, digits=Int($(args[2])))"
        # "MAX" => "xl_max(ctx, $(args)...)"
        # "INDIRECT" => "exec(toexpr(parse_formula(exec($(args[1]), ctx))), ctx)"
        "INDIRECT" => "locals()[$(args[1])]"
        "COUNTIFS" => "xl.countifs(ctx, $(args)...)"
        "AND" => begin
            wrap_logical = s -> "xl.logical($s)"
            "xl.xlall([" * join(map(wrap_logical, args), ", ") * "])"
        end
        "OR" => begin
            wrap_logical = s -> "xl.logical($s)"
            "xl.xlany([" * join(map(wrap_logical, args), ", ") * "])"
        end
        "NOT" => "(!($(args[1])))"
        _ => missing
    end

    if !ismissing(complex)
        return complex
    end

    params = "($(join(args, ", ")))"
    @match fn_name begin
        "MAX" => "xl.max" * params
        "ABS" => "abs" * params
        "AVERAGE" => "xl.average" * params
        "SUM" => "xl.xlsum" * params
        "SUMPRODUCT" => "xl.sum_product" * params
        "SQRT" => "sqrt" * params
        "_xlfn.CONCAT" => "\"\".join(map(str," * params * "))"
        # "AND" => "all" * params
        # "OR" => "any(" * params * ")"
        "FLOOR" => "xl.floor" * params
        "CEILING" => "xl.ceiling" * params
        "MIN" => "xl.min" * params
        "MEDIAN" => "xl.median" * params
        "PMT" => "xl.pmt" * params
        "PRODUCT" => "xl.product" * params
        "ROUND" => "xl.xlround" * params
        "ROUNDUP" => "xl.roundup" * params
        "ROUNDDOWN" => "xl.rounddown" * params
        "_xlfn.STDEV.S" => "xl.stdev" * params
        "_xlfn.XLOOKUP" => "xl.xlookup" * params
        "_xlfn.XMATCH" => "xl.xmatch" * params
        "EXP" => "np.exp" * params
        "_xlfn.DAYS" => "xl.days" * params
        "DATE" => "xl.date" * params
        "EDATE" => "xl.edate" * params
        "MOD" => "xl.mod" * params
        "PI" => "np.pi"
        "LINEST" => "xl.linest" * params
        "IF_MULTIPLE" => "if.multiple" * params
        "VLOOKUP" => "xl.vlookup" * params
        "TEXT" => "xl.text" * params
        "ATAN" => "xl.atan" * params
        "SIN" => "xl.sin" * params
        "ASIN" => "xl.asin" * params
        "COS" => "xl.cos" * params
        "TAN" => "xl.tan" * params
        "RADIANS" => "xl.radians" * params
        # "IF_ELSE_FALSE" => "if_else_false" * params
        "ROUNDUP_IF" => "xl.roundup_if" * params
        # TODO: implement these
        "MATCH" => "xl.match" * params
        "INDEX" => "xl.index" * params
        "LOOKUP" => "xl.lookup" * params
        "OFFSET" => "xl.offset" * params
        "COUNTA" => "xl.counta" * params
        "NPV" => "xl.npv" * params
        "ISBLANK" => "ismissing" * params
        "ISNUMBER" => "xl.isnumber" * params
        "TRANSPOSE" => "np.atleast_2d" * params * ".T"
        "CONVERT" => "xl.convert" * params
        "_xlfn.NORM.DIST" => "xl.norm_dist" * params
        fn_name => begin
            # println("Function $fn_name not handled!")
            "xl." * lowercase(fn_name) * params
        end
    end
end


function handle(::EverythingElseHandler, expr::ExcelExpr, exporter::PythonExporter, ctx)
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
            row_values = (r -> get(exporter.var_names, CellDependency(sheet, index_to_cellname(start_col, r)), missing)).(start_row:end_row)
            "[" * join(row_values, ", ") * "]"
        else
            output = []
            for col ∈ start_col:end_col
                row_values = (r -> exporter.var_names[CellDependency(sheet, index_to_cellname(col, r))]).(start_row:end_row)
                push!(output, "[" * join(row_values, ", ") * "]")
            end
            "[" * join(output, ", ") * "]"
        end
    end
    func = a -> convert(exporter, a, withbindingaffinity(ctx, -1))
    # @show expr

    @match expr begin
        ExcelExpr(:func_param, [param_num]) => "param_$param_num"
        ExcelExpr(:func_param, [param_num, _type]) => "param_$param_num"
        ExcelExpr(:spill_ref, [row, col, arr]) => "($(func(arr)))[$row, $col]"
        ExcelExpr(:cell_ref, [cell, sheet]) => exporter.var_names[CellDependency(sheet, cell)]
        ExcelExpr(:named_range, [name]) => convert(exporter, get(exporter.named_values, name, "undef_var_$name"), ctx)
        ExcelExpr(:call, ["IF", cond, t, f]) => begin
            cond_is_bool = get_type(cond, sheetname(ctx), exporter.cell_types, exporter.named_values) == Bool
            t_str = func(t)
            cond_str = func(cond)
            f_str = func(f)
            # "(($(func(t))) if ($(func(cond))) else ($(func(f))))"
            "(($t_str) if xl.logical($cond_str) else ($f_str))"
            # if cond_is_bool
            # else
            #     "(xl_logical($(func(cond))) ? $(func(t)) : $(func(f)))"
            # end
        end
        ExcelExpr(:call, ["IF", cond, t]) => begin
            cond_is_bool = get_type(cond, sheetname(ctx), exporter.cell_types, exporter.named_values) == Bool
            t_str = func(t)
            cond_str = func(cond)
            "(($t_str) if xl.logical($cond_str) else False)"
            # if cond_is_bool
            #     "($(func(cond)) ? $(func(t)) : missing)"
            # else
            #     "(xl_logical($(func(cond))) ? $(func(t)) : missing)"
            # end
        end
        ExcelExpr(:call, ["AND", args...]) => begin

            function wrap_logical(e)
                if get_type(e, sheetname(ctx), exporter.cell_types, exporter.named_values) == Bool
                    func(e)
                else
                    "xl.logical($(func(e)))"
                end
            end
            "xl.xlall([" * join(map(wrap_logical, args), ", ") * "])"
        end
        ExcelExpr(:call, ["OR", args...]) => begin

            function wrap_logical(e)
                if get_type(e, sheetname(ctx), exporter.cell_types, exporter.named_values) == Bool
                    func(e)
                else
                    "xl.logical($(func(e)))"
                end
            end
            "xl.xlany([" * join(map(wrap_logical, args), ", ") * "])"
        end
        ExcelExpr(:call, [fn_name, args...]) => xl_call_to_python(fn_name, map(func, args))
        ExcelExpr(:broadcast_protect, [expr]) => "($(func(expr)),)"
        ExcelExpr(:cols, [sheet, columns]) => "columns($(repr(sheet)), $(repr(columns)))"
        ExcelExpr(:array, args) => "[" * join(map(func, args), ",") * "]"
        _ => missing
    end
end

mutable struct PyExporterCtx
    current_sheet::String
    last_binding_affinity::Int
end

sheetname(ctx::PyExporterCtx) = ctx.current_sheet
function withsheet(ctx::PyExporterCtx, sheet::AbstractString)
    PyExporterCtx(
        sheet,
        ctx.last_binding_affinity,
    )
end

bindingaffinity(ctx::PyExporterCtx) = ctx.last_binding_affinity
function withbindingaffinity(ctx::PyExporterCtx, affinity::Int)
    PyExporterCtx(
        sheetname(ctx),
        affinity,
    )
end
# function setsheet!(ctx::PyExporterCtx, sheet::AbstractString) 
#     ctx.current_sheet = string(sheet)
#     ctx
# end



function convert(exporter::PythonExporter, expr, ctx::PyExporterCtx)
    if expr isa String
        jl_str = escape_string(expr)
        jl_str = replace(jl_str, "\$" => "\\\$")
        "\"" * jl_str * "\""
    elseif expr isa Int
        repr(expr)
    elseif expr isa Dates.Date
        y = year(expr)
        m = lpad(month(expr), 2, '0')
        d = lpad(day(expr), 2, '0')
        # "datetime.date($y, $m, $d)"
        "pd.Timestamp(\"$y-$m-$d\")"
    elseif expr isa Dates.DateTime
        # replace(repr(expr), "Dates.DateTime" => "datetime.datetime.fromisoformat")
        date_str = Dates.format(expr, "yyyy-mm-ddTHH:MM:SS.sss")
        "pd.Timestamp(\"$date_str\")"
        # replace(repr(expr), "Dates.DateTime" => "pd.Timestamp")
    elseif expr isa Dates.Time
        repr(expr.instant.value)
    elseif ismissing(expr)
        "pd.NA"
    elseif expr isa Bool
        expr ? "True" : "False"
    else
        repr(expr)
    end
end

function convert(exporter::PythonExporter, expr::FlatExpr, ctx::PyExporterCtx)
    convert(exporter, convert_to_expr(expr), ctx)
end
function convert(exporter::PythonExporter, expr::ExcelExpr, ctx::PyExporterCtx)
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
                # @show res
            end
            return res
        end
    end


    # @show expr
    string(expr)
    # throw("Failed to handle an excel expr: $expr")
end

convert(exporter::PythonExporter, expr, current_sheet::AbstractString) = convert(exporter, expr, PyExporterCtx(current_sheet, -1))

function make_struct(exporter::PythonExporter, struct_name::AbstractString, var_names::Vector{<:AbstractString}; var_types::Union{Nothing, Vector} = nothing, ismutable = false, default_values = nothing, var_comments::Union{Nothing, Vector} = nothing)
    lines = Vector{String}()
    sizehint!(lines, 2 + length(var_names))
    def_terms = Vector{String}()
    # if !isnothing(default_values)
    #     push!(def_terms, "@kwdef")
    # end
    # if ismutable
    #     push!(def_terms, "mutable")
    # end
    type_to_str = Dict(
        Missing => "Any",
        Float64 => "float",
        String => "str",
        Bool => "bool",
        Dates.Date => "pd.Timestamp",
    )

    push!(def_terms, "@dataclass\nclass $struct_name:")
    push!(lines, join(def_terms, " "))

    for (i, name) in enumerate(var_names)
        line = "\t$name"
        if !isnothing(var_types)
            type = var_types[i]
            if !ismissing(type)
                type_str = get(type_to_str, type, string(type))
                line *= ":$type_str"
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
    push!(lines, "\n")

    join(lines, "\n")
end

xlookup_to_indexing!(expr) = expr
function xlookup_to_indexing!(expr::FlatExpr)

    for (i, part) in enumerate(expr.parts)
        # i in handled && continue

        @match part begin
            ExcelExpr(:call, ["_xlfn.XLOOKUP", FlatIdx(val), FlatIdx(ref_range), FlatIdx(value_range)]) => begin
                ref_part = expr.parts[ref_range]
                value_part = expr.parts[value_range]

                if value_part.head == :broadcast_protect
                    value_part = value_part.args[1]
                    # @show value_part
                end

                if ref_part.head == :broadcast_protect
                    ref_part = ref_part.args[1]
                    # @show ref_part
                end

                if value_part.head != :table_ref
                    continue
                end

                value_row_idx, value_col_idx = value_part.args[2:3]

                ref_region = part_to_workbook_range(expr, ref_range)
                # @show ref_range
                # @show ref_region
                if isnothing(ref_region)
                    println("xlookup_to_indexing failed because ref_region is nothing")
                    @show ref_range
                    continue
                end

                if size(ref_region)[2] != 1 || length(value_col_idx) != 1
                    @show value_row_idx
                    if size(ref_region)[1] != 1
                        continue
                    end

                    println("Found an xlookup that could probably be a column lookup")

                    value_tbl = value_part.args[1]
                    if !(':' in value_tbl.column_names_range)
                        continue
                    end
                    col_start, col_end = split(value_tbl.column_names_range, ":")
                    table_cols_region = WorkbookRegion(value_tbl.sheet_name, col_start, col_end)

                    if ref_region in table_cols_region
                        println("Ref's a column!")
                        # @show expr.parts[val] ref_part value_part
                        expr.parts[i] = ExcelExpr(:table_ref_col, Any[value_tbl, value_row_idx, FlatIdx(val)])
                    end


                else
                    value_tbl = value_part.args[1]
                    if !(':' in value_tbl.row_names_range)
                        continue
                    end

                    row_start, row_end = split(value_tbl.row_names_range, ":")
                    table_rows_region = WorkbookRegion(value_tbl.sheet_name, row_start, row_end)

                    if table_rows_region == ref_region
                        println("Ref's on name row!")
                        expr.parts[i] = ExcelExpr(:table_ref_idx, Any[value_tbl, FlatIdx(val), first(value_col_idx)])
                    end
                end



            end
            _ => continue
        end
    end
end