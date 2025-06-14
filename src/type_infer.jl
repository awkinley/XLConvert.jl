abstract type AbstractExcelOp end
struct ExcelAdd <: AbstractExcelOp end
function get_op_type(::ExcelAdd, lhs_type::Type, rhs_type::Type)
    t = (lhs_type, rhs_type)
    if t == (Float64, Float64)
        Float64
    elseif t == (Float64, Missing)
        Float64
    elseif t == (Missing, Float64)
        Float64
    elseif t == (Missing, Missing)
        Missing
    elseif t == (Bool, Float64)
        Float64
    elseif t == (Float64, Bool)
        Float64
    elseif t == (Date, Float64)
        Date
    elseif t == (Float64, Date)
        Date
    elseif t == (Date, Date)
        # @info "get_type(::ExcelAdd) returning Any for $t"
        Any
    else
        # @info "get_type(::ExcelAdd) returning Any for $t"
        Any
    end
end

struct ExcelSub <: AbstractExcelOp end
function get_op_type(::ExcelSub, lhs_type, rhs_type)
    t = (lhs_type, rhs_type)
    if t == (Float64, Float64)
        Float64
    elseif t == (Float64, Missing)
        Float64
    elseif t == (Missing, Float64)
        Float64
    elseif t == (Missing, Missing)
        Missing
    elseif t == (Float64, Bool)
        Float64
    elseif t == (Date, Float64)
        Date
    elseif t == (Float64, Date)
        Date
    elseif t == (Date, Date)
        Float64
    else
        # @info "get_type(::ExcelSub) returning Any for $t"
        Any
    end
end

struct ExcelMul <: AbstractExcelOp end
function get_op_type(::ExcelMul, lhs_type, rhs_type)
    t = (lhs_type, rhs_type)
    if t == (Float64, Float64)
        Float64
    elseif t == (Float64, Missing)
        Float64
    elseif t == (Missing, Float64)
        Float64
    elseif t == (Missing, Missing)
        Missing
    else
        # @info "get_op_type(::ExcelMul) returning Any for $t"
        Any
    end
end

struct ExcelDiv <: AbstractExcelOp end
function get_op_type(::ExcelDiv, lhs_type, rhs_type)
    t = (lhs_type, rhs_type)
    if t == (Float64, Float64)
        Float64
    elseif t == (Float64, Missing)
        # @info "get_op_type(::ExcelDiv) returning Any for divide by missing for $t"
        Any
    elseif t == (Missing, Float64)
        Float64
    elseif t == (Missing, Missing)
        Missing
    else
        # @info "get_op_type(::ExcelDiv) returning Any for $t"
        Any
    end
end

function get_type(op::T, lhs_type::Type, rhs_type::Type) where {T<:AbstractExcelOp}
    get_op_type(op, lhs_type, rhs_type)
end

function get_type(op::T, lhs_type::Set{DataType}, rhs_type::Set{DataType}) where {T<:AbstractExcelOp}
    res_set = reduce(union_types, map(tup -> get_op_type(op, tup...), Iterators.product(lhs_type, rhs_type)))

    # if length(res_set) == 1
    #     first(res_set)
    # else
    #     res_set
    # end
    res_set
end

get_type(op::T, lhs_type::Set{DataType}, rhs_type::Type) where {T<:AbstractExcelOp} = get_type(op, lhs_type, Set{DataType}((rhs_type,)))
get_type(op::T, lhs_type::Type, rhs_type::Set{DataType}) where {T<:AbstractExcelOp} = get_type(op, Set{DataType}((lhs_type,)), rhs_type)

function union_types(a::Set{DataType}, b::Set{DataType})
    res = union(a, b)
    Any in res && return Any
    length(res) == 1 ? first(res) : res
end
union_types(a::Type, b::Type) = a == b ? a : union_types(Set{DataType}((a,)), Set{DataType}((b,)))
union_types(a::Type, b::Set{DataType}) = union_types(Set{DataType}((a,)), b)
union_types(a::Set{DataType}, b::Type) = union_types(a, Set{DataType}((b,)))


get_type(::Float64, current_sheet, cell_types, key_values) = Float64
get_type(::Int64, current_sheet, cell_types, key_values) = Float64
get_type(::String, current_sheet, cell_types, key_values) = String
get_type(::Bool, current_sheet, cell_types, key_values) = Bool
get_type(::Type{T}, current_sheet, cell_types, key_values) where {T} = T
get_type(::Missing, current_sheet, cell_types, key_values) = Missing
function get_type(expr::ExcelExpr, current_sheet, cell_types, key_values)
    c = e -> get_type(e, current_sheet, cell_types, key_values)

    @match expr begin
        ExcelExpr(:+, [lhs, rhs]) => get_type(ExcelAdd(), c(lhs), c(rhs))
        ExcelExpr(:-, [lhs, rhs]) => get_type(ExcelSub(), c(lhs), c(rhs))
        ExcelExpr(:*, [lhs, rhs]) => get_type(ExcelMul(), c(lhs), c(rhs))
        ExcelExpr(:/, [lhs, rhs]) => get_type(ExcelDiv(), c(lhs), c(rhs))
        # ExcelExpr(op, (lhs, rhs)), if op in (:+, :-, :*, :/)
        # end => begin
        #     if c(lhs) == Float64 && c(rhs) == Float64
        #         Float64
        #     else
        #         @info "get_type(::ExcelExpr) returning Any" op lhs rhs c(lhs) c(rhs)
        #         Any
        #     end
        # end
        ExcelExpr(:+, [unary,]) => c(unary)
        ExcelExpr(:-, [unary,]) => c(unary)
        ExcelExpr(:^, [lhs, rhs]) => Float64
        ExcelExpr(:%, [unary]) => Float64
        ExcelExpr(:&, [lhs, rhs]) => String
        ExcelExpr(:eq, [lhs, rhs]) => Bool
        ExcelExpr(:neq, [lhs, rhs]) => Bool
        ExcelExpr(:leq, [lhs, rhs]) => Bool
        ExcelExpr(:geq, [lhs, rhs]) => Bool
        ExcelExpr(:lt, [lhs, rhs]) => Bool
        ExcelExpr(:gt, [lhs, rhs]) => Bool
        ExcelExpr(:named_range, [name,]) => c(key_values[name])
        ExcelExpr(:sheet_ref, [sheet_name, ref]) => get_type(ref, sheet_name, cell_types, key_values)

        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => begin
            rows = startrow(table) .+ row_idx .- 1
            cols = startcol(table) .+ col_idx .- 1

            cell_deps = [CellDependency(table.sheet_name, index_to_cellname(c, r)) for c in cols for r in rows]

            types = reduce(union_types, [get(cell_types, c, Any) for c in cell_deps])
            # @info "get_type for table_ref" types

            types
        end
        ExcelExpr(:range, [ExcelExpr(:cell_ref, [lhs, sheet]), ExcelExpr(:cell_ref, [rhs, sheet])]) => begin
            start_col, start_row = parse_cell(lhs)
            end_col, end_row = parse_cell(rhs)

            @assert end_row >= start_row
            @assert end_col >= start_col

            cell_deps = [CellDependency(sheet, index_to_cellname(c, r)) for c in start_col:end_col for r in start_row:end_row]

            types = reduce(union_types, [get(cell_types, c, Any) for c in cell_deps])
            # @info "get_type for range" types
            types
        end
        # ExcelExpr(:range, (lhs, rhs)) => throw("Don't know how to get dependencies for $(expr)")
        ExcelExpr(:range, args) => begin
            # @info "get_type(::ExcelExpr) returning Any for range" args
            Any
        end
        ExcelExpr(:cell_ref, [cell, sheet]) => begin
            cell_dep = CellDependency(sheet, cell)
            get(cell_types, cell_dep, Any)
        end
        ExcelExpr(:call, ["IF", cond, t, f]) => begin
            t_type = c(t)
            f_type = c(f)
            res = union_types(t_type, f_type)
            if res == Any
                # @info "get_type(::ExcelExpr) returning Any from if" t_type f_type
            end
            res
            # if t_type == f_type
            #     t_type
            # else
            #     Set{Type}((t_type, f_type))
            # end
        end
        ExcelExpr(:call, [fn_name, args...]) => begin
            number_returning_funcs = [
                "SUM",
                "AVERAGE",
                "SUMPRODUCT",
                "PI",
                "MOD",
                "EXP",
                "RAND",
                "ROUND",
                "ROUNDUP",
                "ROUNDDOWN",
                "COUNTIFS",
                "MAX",
                "ABS",
                "SUMPRODUCT",
                "SQRT",
                "FLOOR",
                "CEILING",
                "MIN",
                "MEDIAN",
                "PMT",
                "PRODUCT",
                "EXP",
                "LINEST",
                "ATAN",
                "ASIN",
                "SIN",
                "COS",
                "TAN",
                "RADIANS",
                "ROUNDUP_IF"
            ]

            bool_returning_funcs = [
                "AND",
                "OR"
            ]

            if fn_name in number_returning_funcs
                Float64
            elseif fn_name == "_xlfn.XLOOKUP"
                c(args[3])
            elseif fn_name in bool_returning_funcs
                Bool
            else
                # @info "get_type(::ExcelExpr) returning Any for unknown function" fn_name args
                Any
            end
            # fn_name in number_returning_funcs ? Float64 : Any
        end
        ExcelExpr(:func_param, [idx, type]) => type
        default => begin
            # println("Type of $default is Any")
            Any
        end
    end
end