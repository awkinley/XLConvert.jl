part_to_cell_dep(expr::XLConvert.FlatExpr, i::FlatIdx) = part_to_cell_dep(expr, i.i)
function part_to_cell_dep(expr::XLConvert.FlatExpr, i::Int32)
    part = expr.parts[i]
    if @ismatch part ExcelExpr(:cell_ref, [cell, sheet])
        return CellDependency(sheet, cell)
    end

    if @ismatch part ExcelExpr(:sheet_ref, [sheet, sub_i])
        return part_to_cell_dep(expr, sub_i)
    end


    nothing
end

# part_to_workbook_range(expr::XLConvert.FlatExpr, i::FlatIdx) = part_to_workbook_range(expr, i.i)
# function part_to_workbook_range(expr::XLConvert.FlatExpr, i::Int32)
#     part = expr.parts[i]

#     if @ismatch part ExcelExpr(:sheet_ref, [sheet, sub_i])
#         return part_to_workbook_range(expr, sub_i)
#     end

#     @match part begin
#         ExcelExpr(:range, [FlatIdx(lhs_i), FlatIdx(rhs_i)]) => begin
#             lhs_expr = expr.parts[lhs_i]
#             rhs_expr = expr.parts[rhs_i]
#             if !((lhs_expr.head == :cell_ref) && (rhs_expr.head == :cell_ref))
#                 # throw("Don't know how to get dependencies because of a range expression without cell refs. i = $i. lhs_expr = $(lhs_expr.head), rhs_expr = $(rhs_expr.head)")
#                 @show lhs_expr rhs_expr
#                 return nothing
#             end
#             sheet = lhs_expr.args[2]
#             if (sheet != rhs_expr.args[2])
#                 throw("Don't know how to get dependencies because of a range expression that doesn't share a cell. i = $i")
#             end

#             lhs = lhs_expr.args[1]
#             rhs = rhs_expr.args[1]

#             WorkbookRegion(CellDependency(sheet, lhs), CellDependency(sheet, rhs))
#         end
#         _ => nothing
#     end
# end

xl_text_concat(a::AbstractString, b::AbstractString) = string(a, b)
xl_text_concat(a::AbstractString, b::Missing) = string(a)
xl_text_concat(a::Missing, b::AbstractString) = string(b)
xl_text_concat(a::Missing, b::Missing) = ""


function lookup_const_propagate!(wb::XLConvert.ExcelWorkbook2, expr_in::XLConvert.FlatExpr, const_regions::Vector{Any}; debug::Bool = false)
    # deps = Vector{Union{WorkbookRegion, CellDependency}}()
    expr = copy(expr_in)
    function is_const(cell_ref)
        for const_region in const_regions
            if cell_ref in const_region
                return true
            end
        end

        false
    end
    handled = Set{Int}()

    xf = wb.xf

    function get_workbook_values(cell::CellDependency)
        xf[cell.sheet_name][cell.cell]
    end
    function get_workbook_values(range::WorkbookRegion)
        first = range.first
        last = range.last

        xf[first.sheet_name]["$(first.cell):$(last.cell)"]
    end

    function insert_default!(idx, value)
        for i in 1:idx
            part = expr.parts[i]
            !(FlatIdx(idx) in part.args) && continue

            expr.parts[i] = ExcelExpr(part.head, replace(part.args, FlatIdx(idx) => value))
        end
    end

    """
        get_const_value(value)

        Given some value or FlatIdx, tries to get its value based on the constant parts of the sheet.
        If it's not constant or can't be evaluated, the value is nothing
    """
    function get_const_value(value)
        if debug
            println("get_const_value, value of unknown type $(value)")
        end
        nothing
    end
    get_const_value(value::Real) = value
    get_const_value(value::AbstractString) = value
    function get_const_value(value::FlatIdx)
        cell_dep = part_to_cell_dep(expr, value)
        if !isnothing(cell_dep) && is_const(cell_dep)
            return get_workbook_values(cell_dep)
        elseif debug
            # if isnothing(cell_dep)
            #     println("get_const_value value was not a cell ref")
            # else
            #     println("get_const_value value was not constant")
            # end
        end

        cell_range = part_to_workbook_range(expr, value)
        if !isnothing(cell_range) && is_const(cell_range)
            return get_workbook_values(cell_range)
        elseif debug
            if isnothing(cell_range)
                println("get_const_value value was not a cell range")
            else
                println("get_const_value cell range was not constant")
            end
        end

        part = expr.parts[value.i]
        @match part begin
            ExcelExpr(:&, [lhs, rhs]) => begin
                debug && println("Found ampersand expr")
                lhs_val = get_const_value(lhs)
                if isnothing(lhs_val)
                    debug && println("left hand is not a constant value")
                    return nothing
                end
                rhs_val = get_const_value(rhs)
                if isnothing(rhs_val)
                    debug && println("right hand is not a constant value")
                    return nothing
                end

                if lhs_val isa AbstractString && rhs_val isa AbstractString
                    lhs_val * rhs_val
                elseif lhs_val isa AbstractArray && rhs_val isa AbstractArray
                    if length(lhs_val) != length(rhs_val)
                        if debug
                            println("get_const_value: Ampersand expr arrays had different lengths")
                            @show lhs_val rhs_val
                        end
                        return nothing
                    end

                    map((a) -> xl_text_concat(a[1], a[2]), zip(lhs_val, rhs_val))
                else
                    if debug
                        println("get_const_value: Ampersand expr had const values, but they weren't strings or arrays")
                        @show lhs_val rhs_val
                    end
                    nothing
                end
            end
            _ => nothing
        end
    end

    for (i, part) in enumerate(expr.parts)
        i in handled && continue

        @match part begin
            ExcelExpr(:call, ["_xlfn.XLOOKUP", args...]) => begin
                if length(args) < 3
                    println("Invalid xlookup")
                    @show part
                end
                if !(args[1] isa FlatIdx)
                    # @display expr
                    continue
                end

                debug && println("Found xlookup")

                value = args[1]
                ref_range = args[2]
                result_range = args[3].i

                actual_value = get_const_value(value)
                if isnothing(actual_value)
                    if debug
                        println("Value cell can't get const value")
                        if value isa XLConvert.FlatIdx
                            println("value cell: $(expr.parts[value.i])")
                        else
                            println("value : $(value)")
                        end
                    end
                    continue
                end
                ref_values = get_const_value(ref_range)
                if isnothing(ref_values)
                    if debug
                        println("Can't turn ref range into workbook range")
                        if ref_range isa XLConvert.FlatIdx
                            println("ref_range cell: $(expr.parts[ref_range.i])")
                        else
                            println("ref_range : $(ref_range)")
                        end
                    end
                    continue
                end

                result_workbook_range = part_to_workbook_range(expr, result_range)
                if isnothing(result_workbook_range)
                    if debug
                        println("Can't const propagate because the result isn't a workbook range")
                        @show expr.parts[result_range]
                    end
                    continue
                end
                # @show result_workbook_range

                # actual_value = get_workbook_values(value_cell)
                # ref_values = get_workbook_values(ref_workbook_range)
                ref_size = size(ref_values)
                # @show ref_size
                if ref_size[1] != 1 && ref_size[2] != 1
                    if debug
                        println("The ref range isn't 1 dimensional")
                        @show ref_values
                        println("")
                    end
                    continue
                end
                # @show actual_value ref_values
                if ismissing(actual_value)
                    if length(args) >= 4
                        debug && println("Lookup value is missing, but default value is available")
                        insert_default!(i, args[4])
                        continue
                    else
                        # This is a weird hack to handle when we are looking for a missing value
                        debug && println("Lookup value is missing, no default available")
                        # insert_default!(i, nothing)
                        value_cell = part_to_cell_dep(expr, value)
                        expr.parts[i] = ExcelExpr(:cell_ref, Any[value_cell.cell, value_cell.sheet_name])
                        continue
                    end
                end

                result_index = findfirst(==(actual_value), vec(ref_values))
                # @show result_index
                if isnothing(result_index) && length(args) >= 4
                    println("Couldn't find the lookup, but default value is available")
                    continue
                elseif isnothing(result_index)
                    debug && println("Couldn't find the lookup, but no default value is available")
                    # @show value ref_range result_range
                    # @show ExcelExpr(:cell_ref, Any[value_cell.cell, value_cell.sheet_name])
                    value_cell = part_to_cell_dep(expr, value)
                    # @show value_cell
                    # @show expr.parts[value.i]
                    expr.parts[i] = ExcelExpr(:cell_ref, Any[value_cell.cell, value_cell.sheet_name])
                    continue
                end

                result_index -= 1
                if debug
                    @show size(ref_values)
                    @show result_workbook_range
                end

                result_size = size(result_workbook_range)
                if result_size[1] != 1 && result_size[2] != 1
                    # println("The result range isn't 1 dimensional")
                    # @show result_workbook_range
                    # println("")
                    if size(ref_values)[1] == 1
                        offset_rows = 0
                        offset_cols = result_index
                    else
                        offset_rows = result_index
                        offset_cols = 0
                    end

                    # Just steal some of the existing references to overwrite
                    expr.parts[i] = ExcelExpr(:range, Any[FlatIdx(value.i), FlatIdx(ref_range.i)])
                    new_first = XLConvert.offset(result_workbook_range.first, offset_rows, offset_cols)
                    # new_last = XLConvert.offset(result_workbook_range.last, offset_rows, offset_cols)
                    # new_last = XLConvert.offset(new_first, 0, result_size[2] - 1)
                    if size(ref_values)[1] == 1
                        new_last = XLConvert.offset(new_first, result_size[1] - 1, 0)
                    else
                        new_last = XLConvert.offset(new_first, 0, result_size[2] - 1)
                    end
                    if debug
                        println("Making range from $(new_first) to $(new_last)")
                    end

                    @assert new_first in result_workbook_range
                    @assert new_last in result_workbook_range

                    expr.parts[value.i] = ExcelExpr(:cell_ref, Any[new_first.cell, new_first.sheet_name])
                    expr.parts[ref_range.i] = ExcelExpr(:cell_ref, Any[new_last.cell, new_last.sheet_name])
                else
                    offset_rows = result_index
                    offset_cols = 0
                    if result_size[1] == 1
                        (offset_cols, offset_rows) = (offset_rows, offset_cols)
                    end

                    const_ref_cell = XLConvert.offset(result_workbook_range.first, offset_rows, offset_cols)
                    debug && println("New cell ref = $(const_ref_cell)")
                    if !(const_ref_cell in result_workbook_range)
                        @show result_workbook_range result_index offset_rows offset_cols
                        @show const_ref_cell
                        @display expr
                        @assert const_ref_cell in result_workbook_range
                    end

                    # fixed_row = false
                    # fixed_col = false
                    # if value isa XLConvert.FlatIdx
                    #     part = expr.parts[value.i]
                    #     if part.head == :cell_ref
                    #         part_cell = part.args[1]
                    #         fixed_row = XLConvert.cell_fixed_row(part_cell)
                    #         fixed_col = XLConvert.cell_fixed_col(part_cell)
                    #     end
                    # end

                    # cell_str = (fixed_col ? '$' : "") * XLSX.encode_column_number(colnum(const_ref_cell)) * (fixed_row ? '$' : "") * string(rownum(const_ref_cell))
                    cell_str = const_ref_cell.cell

                    expr.parts[i] = ExcelExpr(:cell_ref, Any[cell_str, const_ref_cell.sheet_name])
                end
            end
            _ => continue
        end
    end

    # This is a quick and dirty way to prune out now unused values
    # @display XLConvert.convert_to_expr(expr)
    expr = XLConvert.convert_to_flat_expr(XLConvert.convert_to_expr(expr))
end