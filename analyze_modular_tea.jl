if false
    include("./src/XLConvert.jl")
end
using AutoHashEquals
using XLConvert
using XLConvert: FlatExpr, FlatIdx
using XLSX
using Graphs
using Match

convert_is_blank!(expr) = expr
function convert_is_blank!(expr::FlatExpr)
    for (i, part) in enumerate(expr.parts)
        if @ismatch part ExcelExpr(:eq, [FlatIdx(idx), ""])
            expr.parts[i] = ExcelExpr(:call, Any["ISBLANK", FlatIdx(idx)])
        end
    end

    expr
end

function is_blank_transform!(statements::AbstractArray{AbstractStatement})
    for s in statements
        XLConvert.apply_expr_transform!(s, (_, expr) -> convert_is_blank!(expr))
    end
end

convert_if_one_zero!(expr) = expr
function convert_if_one_zero!(expr::FlatExpr)
    for (i, part) in enumerate(expr.parts)
        @flat_match expr part begin
            ExcelExpr(:call, ["IF", ExcelExpr(:eq, [a, b]), 1, 0]) => begin
                # println("Convert if one zero found something!")
                # @display expr
                expr.parts[i] = ExcelExpr(:eq, a, b)
            end
            _ => continue
        end
        # if @ismatch part ExcelExpr(:eq, [FlatIdx(idx), ""])
        #     expr.parts[i] = ExcelExpr(:call, Any["ISBLANK", FlatIdx(idx)])
        # end
    end

    expr
end

function if_one_zero_transform!(statements::AbstractArray{AbstractStatement})
    for s in statements
        XLConvert.apply_expr_transform!(s, (_, expr) -> convert_if_one_zero!(expr))
    end
end

convert_average_two_cell_range!(expr) = expr
function convert_average_two_cell_range!(expr::XLConvert.FlatExpr)
    did_work = false
    for (i, part) in enumerate(expr.parts)
        @flat_match expr part begin
            ExcelExpr(:call, ["AVERAGE", ExcelExpr(:range, [ExcelExpr(:cell_ref, [lhs, sheet]), ExcelExpr(:cell_ref, [rhs, sheet])])]) => begin
                region = WorkbookRegion(sheet, lhs, rhs)
                if prod(size(region)) != 2
                    @match_fail
                end
                # Don't match if the average is the whole expression, since it might be a date thing
                if i == 1
                    @match_fail
                end
                range_i = part.args[2]

                # @show sheet lhs rhs
                expr.parts[i] = ExcelExpr(:/, range_i, 2)
                expr.parts[range_i.i] = ExcelExpr(:+, expr.parts[range_i.i].args)
                did_work = true
            end
            _ => continue
        end
    end

    if did_work
        # @display expr
        expr = XLConvert.convert_to_flat_expr(XLConvert.convert_to_expr(expr))
        # @display expr
    end

    expr
end

function average_two_cell_range_transform!(statements::AbstractArray{AbstractStatement})
    for s in statements
        XLConvert.apply_expr_transform!(s, (_, expr) -> convert_average_two_cell_range!(expr))
    end
end

convert_indirects!(expr) = expr
function convert_indirects!(expr::FlatExpr)
    for (i, part) in enumerate(expr.parts)
        @flat_match expr part begin
            ExcelExpr(:call, [
                "AVERAGE", 
                ExcelExpr(:range, [
                    ExcelExpr(:call, [
                        "INDIRECT", 
                        ExcelExpr(:call, [
                            "CELL",
                            "address",
                            ExcelExpr(:call, ["_xlfn.XLOOKUP", lhs_value, lhs_refs, lhs_value_range]),
                        ])
                    ]),
                    ExcelExpr(:call, [
                        "INDIRECT", 
                        ExcelExpr(:call, [
                            "CELL",
                            "address",
                            ExcelExpr(:call, ["_xlfn.XLOOKUP", rhs_value, rhs_refs, rhs_value_range]),
                        ])
                    ])
                ])
            ]) => begin
                if expr.parts[lhs_refs.i] != expr.parts[rhs_refs.i] || expr.parts[lhs_value_range.i] != expr.parts[rhs_value_range.i]
                    @match_fail
                end

                new_part = ExcelExpr(:call, "AVERAGEIFS", lhs_value_range, lhs_refs, ExcelExpr(:&, ">=", lhs_value), rhs_refs, ExcelExpr(:&, "<=", rhs_value))
                # @display expr
                # @display change_expr_part(expr, i, new_part)
                change_expr_part!(expr, i, new_part)
            end
            _ => continue
        end
    end

    expr
end

function indirect_transform!(statements::AbstractArray{AbstractStatement})
    for s in statements
        XLConvert.apply_expr_transform!(s, (_, expr) -> convert_indirects!(expr))
    end
end

function table_col_row_name_transform!(statements::AbstractArray{AbstractStatement}, xf)
    handled = Set{Int64}()

    for s in statements
        s isa XLConvert.TableStatement || continue

        lhs = XLConvert.TableRef(s.lhs_expr)

        if length(XLConvert.get_rows(lhs)) != 1 || length(XLConvert.get_cols(lhs)) != 1
            continue
        end

        table = XLConvert.get_table(lhs)
        row_name_region = XLConvert.row_name_region(table)
        col_name_region = XLConvert.column_name_region(table)

        expr = s.rhs_expr
        if !(expr isa XLConvert.FlatExpr)
            continue
        end

        row_name_cell = if isnothing(row_name_region)
            nothing
        elseif all(v -> isequal(v, true), XLConvert.get_cell_values(row_name_region, xf) .== XLConvert.get_row_names(table))
            row_name_region[first(XLConvert.get_rows(lhs)), 1]
        else
            nothing
        end
        col_name_cell = if isnothing(col_name_region)
            nothing
        elseif all(v -> isequal(v, true), XLConvert.get_cell_values(col_name_region, xf) .== XLConvert.get_column_names(table))
            col_name_region[1, first(XLConvert.get_cols(lhs))]
        else
            nothing
        end

        empty!(handled)

        for (i, part) in enumerate(expr.parts)
            i in handled && continue

            @match part begin
                ExcelExpr(:range, [lhs, rhs]) => begin
                    # This is a pretty lazy way of trying to avoid ranges
                    push!(handled, lhs.i)
                    push!(handled, rhs.i)
                end
                ExcelExpr(:table_ref, [tbl, row_idx, col_idx, _, _]) => begin
                    if length(row_idx) != 1 || length(col_idx) != 1
                        @match_fail
                    end

                    cell = XLConvert.cell_dep(XLConvert.TableRef(part))
                    if cell == row_name_cell
                        # println("Statement setting $(XLConvert.cell_dep(lhs)) had a row name region ref $cell")
                        expr.parts[i] = ExcelExpr(:row_name, Any[])
                    elseif cell == col_name_cell
                        # println("Statement setting $(XLConvert.cell_dep(lhs)) had a column name region ref $cell")
                        expr.parts[i] = ExcelExpr(:column_name, Any[])
                    end
                end
                ExcelExpr(:cell_ref, [cell_str, sheet]) => begin
                    cell = CellDependency(sheet, cell_str)
                    if cell == row_name_cell
                        # println("Statement setting $(XLConvert.cell_dep(lhs)) had a row name region ref $cell")
                        expr.parts[i] = ExcelExpr(:row_name, Any[])
                    elseif cell == col_name_cell
                        # println("Statement setting $(XLConvert.cell_dep(lhs)) had a column name region ref $cell")
                        expr.parts[i] = ExcelExpr(:column_name, Any[])
                    end
                end
                _ => continue
            end

        end


    end
end

struct IndirectXlookupRangeHandler end

function XLConvert.handle(::IndirectXlookupRangeHandler, expr::ExcelExpr, exporter::PythonExporter, ctx::XLConvert.PyExporterCtx)
    func = a -> XLConvert.convert(exporter, a, ctx)
    @match expr begin
        ExcelExpr(
            :range,
            [
                ExcelExpr(:call, ["INDIRECT", ExcelExpr(:call, ["CELL", "address", ExcelExpr(:call, ["_xlfn.XLOOKUP", val, ref, value])])]),
                ExcelExpr(:call, ["INDIRECT", ExcelExpr(:call, ["CELL", "address", ExcelExpr(:call, ["_xlfn.XLOOKUP", val2, ref, value])])]),
            ],
        ) => begin
            # println("Convert convert_indirect_xlookup_range found something!")
            ref_str = func(ref)
            col_str = func(value)
            mask_str = "($(ref_str) >= $(func(val))) & ($(ref_str) <= $(func(val2)))"

            "$col_str[$mask_str]"
        end
        _ => missing
    end
end

struct EdgeCaseHandler end

function XLConvert.handle(::EdgeCaseHandler, expr::ExcelExpr, exporter::PythonExporter, ctx::XLConvert.PyExporterCtx)
    func = a -> XLConvert.convert(exporter, a, ctx)

    function make_ifs_mask(args)
        @assert length(args) % 2 == 0
        conds = String[]
        for i in 1:(length(args)÷2)
            test_val = args[2*i-1]
            test_val_str = func(test_val)
            test = args[2*i]

            cond = @match test begin
                "<>#N/A" => "~pd.isna($(func(test_val)))"
                "<>FALSE" => "($(func(test_val)) != False)"
                val::AbstractString => begin
                    if any(s -> startswith(val, s), [">", "<"])
                        "($test_val_str $val)"
                    elseif val == ""
                        "($test_val_str.fillna(\"\") == $(repr(val)))"
                    else
                        "($test_val_str == $(repr(val)))"
                    end

                end
                ExcelExpr(:&, [">", val::ExcelExpr]) => "($test_val_str > $(func(val)))"
                ExcelExpr(:&, ["<", val::ExcelExpr]) => "($test_val_str < $(func(val)))"
                ExcelExpr(:&, [">=", val::ExcelExpr]) => "($test_val_str >= $(func(val)))"
                ExcelExpr(:&, ["<=", val::ExcelExpr]) => "($test_val_str <= $(func(val)))"
                val::ExcelExpr => begin
                    val_type = get_type(val, XLConvert.sheetname(ctx), exporter.cell_types, exporter.named_values)
                    if val_type <: AbstractString || val.head in (:column_name, :row_name)
                        "xl.match_case_insensitive($(func(test_val)), $(func(val)))"
                    else get_type(test, XLConvert.sheetname(ctx), exporter.cell_types, exporter.named_values)
                        "xl.as_array($test_val_str == $(func(val)))"
                    end
                end
                val::Number => begin
                    "($test_val_str == $(func(val)))"
                end
                _ => missing
            end
            if ismissing(cond)
                println("Make ifs mask got a test value it didn't recognize!")
                @show test
                throw("make_ifs_mask error")
            else
                push!(conds, cond)
            end
        end

        join(conds, " & ")
    end

    @match expr begin
        ExcelExpr(:call, ["_xlfn.MAKEARRAY", val, rest...]) => begin
            func(val)
        end
        ExcelExpr(:error_ref, [error]) => begin
            # repr(error)
            "pd.NA"
        end
        ExcelExpr(:call, ["_xlfn._xlws.FILTER", vals, ExcelExpr(:eq, [test_lhs, test_rhs])]) => begin
            vals_str = func(vals)
            "$vals_str[$(func(test_lhs)) == $(func(test_rhs))]"
        end
        ExcelExpr(:call, ["_xlfn._xlws.FILTER", vals, ExcelExpr(:eq, [test_lhs, test_rhs]), if_empty_val]) => begin
            vals_str = func(vals)
            "$vals_str[($(func(test_lhs)) == $(func(test_rhs))).values.squeeze()]"
        end
        ExcelExpr(:call, ["_xlfn._xlws.FILTER", vals, ExcelExpr(:call, ["IFERROR", error_range, 0])]) => begin
            vals_str = func(vals)
            "$vals_str[~pd.isna($vals_str) & ($vals_str != 0.0)]"
        end
        ExcelExpr(:call, ["_xlfn._xlws.FILTER", vals, test_vals]) => begin
            # FILTER(range, range) is used as range[~pd.isna(range)]
            vals_str = func(vals)
            test_vals_str = func(test_vals)
            "$vals_str[~pd.isna($test_vals_str) & ($test_vals_str != 0.0)]"

        end
        ExcelExpr(:call, ["_xlfn.MAXIFS", vals, args...]) => begin
            vals_str = func(vals)
            mask = make_ifs_mask(args)

            "xl.max($vals_str[$mask])"
        end
        ExcelExpr(:call, ["_xlfn.MINIFS", vals, args...]) => begin
            vals_str = func(vals)
            mask = make_ifs_mask(args)

            "xl.min($vals_str[$mask])"
        end
        ExcelExpr(:call, ["SUMIFS", vals, args...]) => begin

            vals_str = @match vals begin
                ExcelExpr(:table_ref, [table, row_ref, col_ref, _, _]) where size(table)[1] == 1 => begin
                    if length(col_ref) == size(table)[2]
                        "$(getname(table)).iloc[0, :]"
                    else
                        col_name = [string(column_name(table, c)) for c in col_idx]
                        col_idx = "$(repr(col_name[begin])):$(repr(col_name[end]))"
                        first_row = row_name(table, 1)
                        "$(getname(table)).loc[$first_row, $cold_idx]"
                    end
                end
                _ => func(vals)
            end
            # vals_str = func(vals)
            mask = make_ifs_mask(args)

            "xl.xlsum($vals_str[$mask])"
        end
        ExcelExpr(:call, ["SUMIF", vals, "<>#N/A"]) => begin
            vals_str = func(vals)
            "xl.xlsum($vals_str[~pd.isna($vals_str)])"
        end
        ExcelExpr(:call, ["COUNTIFS", args...]) => begin
            mask = make_ifs_mask(args)
            "xl.xlsum($mask)"
        end
        ExcelExpr(:call, ["COUNTIF", args...]) => begin
            mask = make_ifs_mask(args)
            "xl.xlsum($mask)"
        end
        ExcelExpr(:call, ["AVERAGEIFS", range, args...]) => begin
            mask = make_ifs_mask(args)
            "xl.average($(func(range))[$mask])"
        end
        ExcelExpr(:/, [num::Number, denom]) => begin

            left_affinity, right_affinity = XLConvert.op_binding_affinity(:/)
            wrap_parens = left_affinity < XLConvert.bindingaffinity(ctx)
            rhs_str = XLConvert.convert(exporter, denom, XLConvert.withbindingaffinity(ctx, right_affinity))

            binop_str = "np.float64($num) / $rhs_str"
            if wrap_parens
                "(" * binop_str * ")"
            else
                binop_str
            end
        end
        ExcelExpr(:call, ["_xlfn.XLOOKUP",
            ExcelExpr(:&, [concat_lhs, "1"]),
            ExcelExpr(:&, [row_lhs, row_rhs]),
            ExcelExpr(:table_ref, [res_table, res_row_idx, res_col_idx, _, _]),
            0]
        ) => begin
            # Handles a very specific case to transform some code that's tough to make python match
            @assert length(res_col_idx) == 1
            if length(res_row_idx) != size(res_table)[1]
                @match_fail
            end

            col_name = XLConvert.column_name(res_table, first(res_col_idx))
            mask_str = "$(func(row_rhs)) == 1"
            ref_str = string(func(row_lhs), ".loc[$mask_str]")
            value_str = "$(getname(res_table)).loc[$mask_str, $(repr(col_name))]"
            "xl.xlookup($(func(concat_lhs)), $ref_str, $value_str, 0)"
        end
        ExcelExpr(:call, ["_xlfn.XLOOKUP",
            ExcelExpr(:&, [concat_lhs, concat_rhs]),
            ExcelExpr(:&, [row_lhs, row_rhs]),
            ExcelExpr(:table_ref, [res_table, res_row_idx, res_col_idx, _, _]),
            missing,
            1.0]
        ) => begin
            # Handles a very specific case to transform some code that's tough to make python match
            # if length(res_row_idx) != size(res_table)[1]
            #     @match_fail
            # end

            @assert length(res_col_idx) == 1

            # @show res_col_idx
            println("Applying xlookup -> pandas indexing rule")

            col_name = XLConvert.column_name(res_table, first(res_col_idx))
            row_subset = string(XLConvert.row_name(res_table, first(res_row_idx)), ":", XLConvert.row_name(res_table, last(res_row_idx)))
            "$(getname(res_table)).loc[$row_subset].loc[($(func(concat_rhs)) == $(func(row_rhs))) & ($(func(concat_lhs)) <= ($(func(row_lhs)))), $(repr(col_name))]"
        end
        ExcelExpr(:-, [ExcelExpr(:-, [ExcelExpr(:eq, [lhs, rhs])])]) => begin
            # equations like:
            # --(A1:A10 = B1)
            # this is used in a sum product to act like a mask, but xl_eq doesn't like ranges, so rewrite as basic ==
            "($(func(lhs)) == $(func(rhs)))"
        end
        ExcelExpr(
            :range,
            [
                ExcelExpr(:table_ref, range_left_args),
                ExcelExpr(:call, ["INDIRECT", ExcelExpr(:call, ["ADDRESS",
                    ExcelExpr(:call, ["ROW", ExcelExpr(:table_ref, row_args)]),
                    ExcelExpr(:-, [ExcelExpr(:+, [ExcelExpr(:call, ["COLUMN", ExcelExpr(:table_ref, col_args)]), add_rhs]), sub_rhs]),
                    # indirect_column,
                ])]),
            ],
        ) => begin

            # @show range_left_args row_args col_args add_rhs sub_rhs
            table = range_left_args[1]
            # table, row_idx, col_idx, _, _ = range_left_args
            row_idx = startrow(row_args[1]) + row_args[2] - startrow(table) - 1
            col_idx = startcol(col_args[1]) + col_args[3] - startcol(table) - 1
            tbl_name = getname(table)
            "$tbl_name.iloc[$row_idx, $col_idx:($(col_idx + 1) + int($(func(add_rhs)) - $(func(sub_rhs))))]"

        end
        ExcelExpr(
            :range,
            [
                ExcelExpr(:table_ref, range_left_args),
                ExcelExpr(:call, ["INDIRECT", ExcelExpr(:call, ["ADDRESS",
                    ExcelExpr(:call, ["ROW", ExcelExpr(:table_ref, row_args)]),
                    ExcelExpr(:-, [ExcelExpr(:+, [ExcelExpr(:call, ["COLUMN", ExcelExpr(:cell_ref, col_args)]), add_rhs]), sub_rhs]),
                ])]),
            ],
        ) => begin

            # @show range_left_args row_args col_args add_rhs sub_rhs
            table = range_left_args[1]
            # table, row_idx, col_idx, _, _ = range_left_args
            row_idx = startrow(row_args[1]) + row_args[2] - startrow(table) - 1
            col_idx = XLConvert.colnum(CellDependency(col_args[2], col_args[1])) - startcol(table)
            tbl_name = getname(table)
            "$tbl_name.iloc[$row_idx, $col_idx:($(col_idx + 1) + int($(func(add_rhs)) - $(func(sub_rhs))))]"

        end
        ExcelExpr(:call, ["IFERROR", ExcelExpr(:/, [1, denom]), 0]) => begin
            "xl.safe_recip($(func(denom)))"
        end
        _ => missing
    end
end


function get_cell_range_size(expr::FlatExpr, part_idx)
    @match expr.parts[part_idx] begin
        ExcelExpr(:sheet_ref, [sheet, child]) => get_cell_range_size(expr, child.i)
        ExcelExpr(:range, [FlatIdx(lhs_i), FlatIdx(rhs_i)]) => begin

            lhs_expr = expr.parts[lhs_i]
            rhs_expr = expr.parts[rhs_i]
            if !((lhs_expr.head == :cell_ref) && (rhs_expr.head == :cell_ref))
                return nothing
            end
            sheet = lhs_expr.args[2]
            if (sheet != rhs_expr.args[2])
                return nothing
            end

            lhs = lhs_expr.args[1]
            rhs = rhs_expr.args[1]

            start_col, start_row = XLConvert.parse_cell(lhs)
            end_col, end_row = XLConvert.parse_cell(rhs)

            @assert end_row >= start_row
            @assert end_col >= start_col

            return (end_row - start_row + 1, end_col - start_col + 1)
        end
        _ => nothing
    end
end

function get_spill_size(expr::FlatExpr)
    @match expr.parts[1] begin
        ExcelExpr(:call, ["TRANSPOSE", arg]) => begin
            cell_range_size = get_cell_range_size(expr, arg.i)
            isnothing(cell_range_size) && return nothing

            return reverse(cell_range_size)
        end
        _ => return nothing
    end
end

function handle_spill(wb::XLConvert.ExcelWorkbook2)
    spill_funcs = [
        "TRANSPOSE", "SEQUENCE", "SORT", "SORTBY", "UNIQUE",
    ]

    spilling_cells = Vector{CellDependency}()

    function can_spill(expr)
        if @ismatch expr.parts[1] ExcelExpr(:call, args)
            if args[1] in spill_funcs
                return true
            end
        end

        false
    end

    for (cell, data) in wb.cell_dict
        expr = get_expr(data)
        expr isa FlatExpr || continue

        if can_spill(expr)
            println("$cell is a formula that can spill")
            spill_size = get_spill_size(expr)
            println("Spill size = $spill_size")

            show(stdout, "text/plain", expr)


            colnum, rownum = XLConvert.get_coords(cell)
            # rownum += 1
            # colnum += 1
            spill_rows = rownum:(rownum+spill_size[1]-1)
            spill_cols = colnum:(colnum+spill_size[2]-1)
            # @show spill_rows spill_cols

            for r in spill_rows, c in spill_cols
                (r, c) == (rownum, colnum) && continue

                cell_dep = CellDependency(cell.sheet_name, XLConvert.index_to_cellname(c, r))
                # println("Setting $cell_dep as a spill cell")
                new_expr = XLConvert.insert_expr_front(expr, ExcelExpr(:spill_ref, Any[r-rownum+1, c-colnum+1, XLConvert.FlatIdx(1)]))
                wb.cell_dict[cell_dep] = XLConvert.SpillCell(cell_dep, new_expr)
                wb.cell_dependencies[cell_dep] = wb.cell_dependencies[cell]
            end

            push!(spilling_cells, cell)
        end
    end

    println("Found $(length(spilling_cells)) cells that can spill")
end

struct FunctionExpr
    func::XLConvert.FlatExpr
    params::Vector{ExcelExpr}
end


# struct CompressedWorkbook
#     xf::XLSX.XLSXFile
#     cell_numbering::ObjectNumbering{CellDependency}
#     cell_dict::Dict{CellDependency, Any}
#     cell_graph::Graphs.SimpleDiGraph{Int64}
#     key_values::Dict{String, Any}
# end

function make_compressed_workbook(wb::XLConvert.ExcelWorkbook2)
    exprs = [get_expr(cell) for cell in values(wb.cell_dict) if cell isa XLConvert.FormulaCell]

    func_exprs = Vector{FunctionExpr}()
    funcs = Set{FlatExpr}()

    for e in exprs
        func, params = XLConvert.functionalize(e)
        if func in funcs
            func = pop!(funcs, func)
        end
        push!(funcs, func)
    end


end

function recalc_deps!(wb::XLConvert.ExcelWorkbook2, cell::CellDependency)
    node = get_num(wb, cell)
    starting_deps = outneighbors(wb.cell_graph, node)
    keep_deps = Set{Int64}()
    start_deps_set = Set(starting_deps)

    # for dep in outneighbors(wb.cell_graph, node)
    #     rem_edge!(wb.cell_graph, node, dep)
    # end

    expr = get_expr(wb.cell_dict[cell])
    dep_cells = try
        XLConvert.get_expr_dependency_ranges(expr, wb.key_values)
    catch e
        @show cell
        @display expr
        throw(e)

    end
    unique!(dep_cells)

    function handle_dep(dep)
        if dep in start_deps_set
            push!(keep_deps, dep)
        else
            add_edge!(wb.cell_graph, node, dep)
        end
    end


    for dep in dep_cells
        if dep isa CellDependency
            num = XLConvert.get_num!(wb.cell_numbering, dep)
            handle_dep(num)
            # add_edge!(wb.cell_graph, node, num)
        else
            sheet = dep.first.sheet_name
            (start_col, start_row) = XLConvert.start_coord(dep)
            (end_col, end_row) = XLConvert.end_coord(dep)
            for r ∈ start_row:end_row, c ∈ start_col:end_col
                num = XLConvert.get_num!(wb.cell_numbering, CellDependency(sheet, c, r))
                handle_dep(num)
                # add_edge!(wb.cell_graph, node, num)
            end
        end
    end

    deps_to_remove = setdiff(start_deps_set, keep_deps)
    for d in deps_to_remove
        rem_edge!(wb.cell_graph, node, d)
    end

end

function apply_formula_override!(wb, cell, new_formula)
    # Need to change the cell_dict entry, and recalculate outneighbor dependencies
    cell_value = wb.cell_dict[cell] 
    @assert cell_value isa XLConvert.FormulaCell

    println("$cell has formula $(cell_value.cell.formula.formula), is being replaced with $new_formula")

    expr = XLConvert.toexpr(new_formula)
    expr = XLConvert.lower_sheet_names!(expr, cell.sheet_name)
    flat_expr = XLConvert.convert_to_flat_expr(expr)
    wb.cell_dict[cell] = XLConvert.FormulaCell(cell_value.cell, flat_expr)

    recalc_deps!(wb, cell)
end


function get_transpose_dep_forwards(wb::XLConvert.ExcelWorkbook2)
    dep_forwards = Dict{CellDependency, CellDependency}()

    for (cell, data) in wb.cell_dict
        expr = get_expr(data)
        expr isa FlatExpr || continue

        if @ismatch expr.parts[1] ExcelExpr(:call, ["TRANSPOSE", range])
            cell_row = rownum(cell)
            cell_col = colnum(cell)

            # println("$cell is a transpose")
            # (num_rows, num_cols)
            spill_size = get_spill_size(expr)
            transposed_range = part_to_workbook_range(expr, range)
            if isnothing(transposed_range)
                println("Found a tranpose that couldn't be cleanly forwarded, $cell")
                continue
            end
            # @display expr
            # Because I'm bad at being consistent, this range cells is indexed [col, row]
            range_cells = XLConvert.cells(transposed_range)
            first = cell
            out_region = WorkbookRegion(cell, CellDependency(cell.sheet_name, cell_col + spill_size[2] - 1, cell_row + spill_size[1] - 1))
            println("Forwarding to $out_region to $transposed_range")

            for r in 1:spill_size[1], c in 1:spill_size[2]
                spilled_cell = CellDependency(cell.sheet_name, cell_col + c - 1, cell_row + r - 1)
                # println("Creaeting dep forward $spilled_cell => $(range_cells[r, c])")
                dep_forwards[spilled_cell] = range_cells[r, c]
            end



        end
    end

    dep_forwards
end

function apply_dep_forwards!(wb::XLConvert.ExcelWorkbook2, forwards::Dict{CellDependency, CellDependency})

    handled = Set{Int64}()

    # modified_cells = Set{CellDependency}()

    for (original_cell, new_cell) in forwards
        start_node = get_num(wb, original_cell)
        users = inneighbors(wb.cell_graph, start_node)

        new_val = ExcelExpr(:cell_ref, Any[new_cell.cell, new_cell.sheet_name])

        for user in users
            user_cell = get_cell(wb, user)
            cell_value = wb.cell_dict[user_cell]
            expr = get_expr(cell_value)
            # @display expr

            empty!(handled)

            for (i, part) in enumerate(expr.parts)
                i in handled && continue

                @flat_match expr part begin
                    ExcelExpr(:cell_ref, [cell, sheet]) => begin 
                        ref_cell = CellDependency(sheet, cell)
                        if CellDependency(sheet, cell) == original_cell
                            expr.parts[i] = new_val
                        end

                        if ref_cell in keys(forwards)
                            forwarded = forwards[ref_cell]
                            expr.parts[i] = ExcelExpr(:cell_ref, Any[forwarded.cell, forwarded.sheet_name])
                        end

                    end
                    ExcelExpr(:range, [ExcelExpr(:cell_ref, [lhs, sheet]), ExcelExpr(:cell_ref, [rhs, sheet])]) => begin
                        lhs_cell = CellDependency(sheet, lhs)
                        rhs_cell = CellDependency(sheet, rhs)
                        if lhs_cell == CellDependency("Site Inputs", "C68")
                            if (lhs_cell in keys(forwards)) || (rhs_cell in keys(forwards))
                                @show original_cell
                                @show lhs_cell rhs_cell
                                @show (lhs_cell in keys(forwards))
                                @show (rhs_cell in keys(forwards))
                            end
                        end

                        if (lhs_cell in keys(forwards)) && !(rhs_cell in keys(forwards))
                            println("Found a range that wasn't both forwarded?")
                            @show lhs rhs
                            push!(handled, part.args[1].i)
                            push!(handled, part.args[2].i)
                        end
                        if (rhs_cell in keys(forwards)) && !(lhs_cell in keys(forwards))
                            println("Found a range that wasn't both forwarded?")
                            @show lhs rhs
                            push!(handled, part.args[1].i)
                            push!(handled, part.args[2].i)
                        end
                    end
                    _ => nothing
                end
            end
            # push!(modified_cells, user_cell)

            recalc_deps!(wb, user_cell)
        end
    end

    # for cell in modified_cells
    #     recalc_deps!(wb, cell)
    # end
end

function read_wb()
    # file = "scope=3.0_aspect=2.0_TEA.xlsm"
    file = "Modular TEA - master - v1.25.xlsm"
    wb = parse_workbook(file)

    # println("\n", "="^10, "Formula Replacements", "="^10)
    # These both seem like start one row too low
    # apply_formula_override!(wb, CellDependency("Equip&Mat Calcs", "W33"), "SUM(W4:W32)")
    # apply_formula_override!(wb, CellDependency("Equip&Mat Calcs", "W65"), "SUM(W36:W64)")
    # Make this one column wider, because other things will references this
    # range expecting it to be one longer than it is
    # apply_formula_override!(wb, CellDependency("Site Inputs", "C51"), "TRANSPOSE(vessels!F4:Y4)")

    
    wb
end

function number_of_paths(graph, source, destination)
    topo_sorted = topological_sort(graph)

    dp = zeros(Int64, nv(graph))
    dp[destination] = 1

    for idx in reverse(eachindex(topo_sorted))
        out = outneighbors(graph, topo_sorted[idx])
        for n in out
            dp[topo_sorted[idx]] += dp[n]
        end
    end

    dp[source]
end

function param_to_cell(wb::XLConvert.ExcelWorkbook2, param)
    nothing
end

function param_to_cell(wb::XLConvert.ExcelWorkbook2, param::ExcelExpr)
    @match param begin
        ExcelExpr(:cell_ref, [cell, sheet]) => CellDependency(sheet, cell)
        _ => nothing
    end

end

function find_repeated_functions(wb_in::XLConvert.ExcelWorkbook2)


    functions = XLConvert.ObjectNumbering(XLConvert.FlatExpr[])

    func_usages = Dict{Int64, Vector{CellDependency}}()

    cell_funcs = Dict{CellDependency, Tuple{Int64, Matrix{ExcelExpr}}}()

    for (cell, value) in wb_in.cell_dict
        value isa XLConvert.FormulaCell || continue
        expr = XLConvert.get_expr(value)
        expr isa XLConvert.FlatExpr || continue

        (func, params) = XLConvert.functionalize(expr)

        func_id = XLConvert.get_num!(functions, func)
        cell_funcs[cell] = (func_id, params)

        push!(get!(func_usages, func_id, CellDependency[]), cell)
    end

    func_ids = 1:length(functions)

    interesting_func_ids = filter(id -> length(XLConvert.get_obj(functions, id).parts) > 5, func_ids)

    funcs_by_usages = sort(interesting_func_ids, by = i -> length(func_usages[i]), rev = true)

    function get_cell_func_id(cell)
        isnothing(cell) && return nothing

        get(cell_funcs, cell, (nothing, ExcelExpr[]))[1]
    end
    @display XLConvert.get_obj(functions, 1)

    println("Top 5 most used functions:")
    for id in funcs_by_usages[1:10]
        usages = func_usages[id]
        println("Func id = $id has $(length(usages)) usages")
        @display XLConvert.get_obj(functions, id)
        params = reduce(vcat, [cell_funcs[c][2] for c in usages])
        param_cells = map(e -> param_to_cell(wb_in, e), params)
        param_cell_func_ids = get_cell_func_id.(param_cells)
        for col in 1:size(param_cell_func_ids, 2)
            ids = @view param_cell_func_ids[:, col]
            if !any(isnothing.(ids)) && allequal(ids)
                println("Found a set of parameters that are all defined similarly")
                @show col ids
            end
        end
        # @show allequal.(eachcol(param_cell_func_ids))
        # @display param_cells
        # @display get_cell_func_id.(param_cells)

        # for c in usages
        #     println("\t$c - $(cell_funcs[c][2])")
        # end


    end
end

function rename_cell_ref!(expr::XLConvert.FlatExpr, child::CellDependency, root::CellDependency, wb)
    expected_expr = ExcelExpr(:cell_ref, Any[child.cell, child.sheet_name])
    did_replace = false
    for (i, part) in enumerate(expr.parts)
        if part_to_cell_dependency(wb, part) == child
            expr.parts[i] = ExcelExpr(:cell_ref, root.cell, root.sheet_name)
            did_replace = true
        end
        # if @ismatch part ExcelExpr(:cell_ref, [expr_cell, expr_sheet])
        #     if CellDependency(expr_sheet, expr_cell) == child
        #         # println("Found rename!")
        #         expr.parts[i] = ExcelExpr(:cell_ref, root.cell, root.sheet_name)
        #         did_replace = true
        #     end
        # end
    end

    if !did_replace
        println("Couldn't find a replacement?")
        @display expr
    end

    did_replace
end

function merge_cell!(wb::XLConvert.ExcelWorkbook2, child::CellDependency, root::CellDependency)

    child_num = get_num(wb, child)
    root_num = get_num(wb, root)
    # @show child root
    # @show child_num root_num
    # @show inneighbors(wb.cell_graph, child_num)
    edges_to_remove = Vector{Edge}()
    edges_to_add = Vector{Edge}()
    for child_dep in inneighbors(wb.cell_graph, child_num)

        push!(edges_to_remove, Edge(child_dep, child_num))
        push!(edges_to_add, Edge(child_dep, root_num))

        # println("Removing ", Edge(child_dep, child_num))
        # @assert rem_edge!(wb.cell_graph, Edge(child_dep, child_num))
        # add_edge!(wb.cell_graph, Edge(child_dep, root_num))
        cell = get_cell(wb, child_dep)

        dict_value = wb.cell_dict[cell]

        rename_cell_ref!(get_expr(dict_value), child, root, wb)
    end

    for edge in edges_to_remove
        @assert rem_edge!(wb.cell_graph, edge)
    end
    for edge in edges_to_add
        add_edge!(wb.cell_graph, edge)
    end

    # @show inneighbors(wb.cell_graph, child_num)
    @assert length(inneighbors(wb.cell_graph, child_num)) == 0

    # rem_vertex!(wb.cell_graph, child_num)
    # XLConvert.rem_obj!(wb.cell_numbering, child_num)
end

function normalize_cell_refs(expr::XLConvert.FlatExpr)
    new_expr = deepcopy(expr)
    for part in new_expr.parts
        if @ismatch part ExcelExpr(:cell_ref, [cell, sheet])
            part.args[1] = replace(part.args[1], '$' => "")
        end
    end

    new_expr
end

function part_to_cell_dependency(wb::XLConvert.ExcelWorkbook2, part)
    if @ismatch part ExcelExpr(:cell_ref, [expr_cell, expr_sheet])
        # println("Found equivalent variables!")
        return CellDependency(expr_sheet, expr_cell)
    elseif @ismatch part ExcelExpr(:named_range, [name])
        sub_expr = wb.key_values[name]
        if length(sub_expr.parts) == 1
            return get_parent_cell(sub_expr)
        elseif length(sub_expr.parts) == 2
            if sub_expr.parts[1].head == :sheet_ref
                if @ismatch sub_expr.parts[2] ExcelExpr(:cell_ref, [expr_cell, expr_sheet])
                    return CellDependency(expr_sheet, expr_cell)
                end
            end
        end
    end

    nothing
end

function equivalence_analysis(wb::XLConvert.ExcelWorkbook2)

    equivalent_cells = Vector{Tuple{CellDependency, CellDependency}}()
    cells_to_remove = Vector{CellDependency}()

    function get_parent_cell(expr)
        part_to_cell_dependency(wb, expr.parts[1])
    end

    unique_formulas = Dict{XLConvert.FlatExpr, CellDependency}()

    for (cell, value) in wb.cell_dict
        value isa XLConvert.FormulaCell || continue
        expr = XLConvert.get_expr(value)
        expr isa XLConvert.FlatExpr || continue

        if length(expr.parts) == 1
            parent_cell = get_parent_cell(expr)
            if !isnothing(parent_cell)
                # if @ismatch expr.parts[1] ExcelExpr(:cell_ref, [expr_cell, expr_sheet])
                # println("Found equivalent variables!")
                # parent_cell = CellDependency(expr_sheet, expr_cell)
                row_dist = abs(XLConvert.rownum(cell) - XLConvert.rownum(parent_cell))
                col_dist = abs(XLConvert.colnum(cell) - XLConvert.colnum(parent_cell))
                if max(row_dist, col_dist) < 4
                    # println("$cell = $parent_cell")
                    push!(equivalent_cells, (cell, parent_cell))
                end
                # @show cell parent_cell
                # @display expr
            end
        else
            norm_expr = normalize_cell_refs(expr)
            if norm_expr in keys(unique_formulas)
                parent_cell = unique_formulas[norm_expr]
                dist = abs.(XLConvert.get_coords(cell) .- XLConvert.get_coords(parent_cell))
                if maximum(dist) < 4
                    println("Found cells with equivalent formula $(cell) = $(parent_cell)")
                    push!(equivalent_cells, (cell, parent_cell))
                end
            else
                unique_formulas[norm_expr] = cell
            end
        end
    end


    println("\nFound $(length(equivalent_cells)) equivalent cells")
    @display equivalent_cells
    for (child, root) in equivalent_cells
        push!(cells_to_remove, child)
        merge_cell!(wb, child, root)
    end

    for i in 1:3
        roots = unique!(map(v -> v[2], equivalent_cells))

        empty!(equivalent_cells)

        for root in roots
            root_num = get_num(wb, root)
            children = inneighbors(wb.cell_graph, root_num)
            child_exprs = [normalize_cell_refs(get_expr(wb.cell_dict[get_cell(wb, child)])) for child in children]
            unique_children = Dict{XLConvert.FlatExpr, Int64}()
            for (i, child_expr) in enumerate(child_exprs)
                if child_expr in keys(unique_children)
                    push!(equivalent_cells, get_cell.(Ref(wb), (children[i], children[unique_children[child_expr]])))
                else
                    unique_children[child_expr] = i
                end
            end
        end


        println("\nOn round $(i + 1): found $(length(equivalent_cells)) equivalent cells")
        @display equivalent_cells
        for (child, root) in equivalent_cells
            push!(cells_to_remove, child)
            merge_cell!(wb, child, root)
        end
    end


    # mask = ones(Bool, nv(wb.cell_graph))
    # for cell in cells_to_remove
    #     mask[get_num(wb, cell)] = false
    # end
    # (subgraph, vmap) = induced_subgraph(wb.cell_graph, findall(mask))

    # cell_numbering = XLConvert.ObjectNumbering(wb.cell_numbering.objs[vmap])

    # XLConvert.ExcelWorkbook(wb.xf, cell_numbering, wb.cell_dict, subgraph, wb.key_values)

    wb
end

function test_lookup_const_propagate(wb::XLConvert.ExcelWorkbook2)
    const_regions = Any[
        WorkbookRegion("Structure Calcs", "CM6", "CM47"),
        WorkbookRegion("Structure Calcs", "CM51", "CM92"),
        WorkbookRegion("Structure Calcs", "CM100", "CM141"),
        WorkbookRegion("Structure Calcs", "CM149", "CM190"),
        WorkbookRegion("Structure Calcs", "CM197", "CM238"),
        WorkbookRegion("Structure Calcs", "CN4", "FR4"),
        WorkbookRegion("Structure Calcs", "C6", "C47"),
        WorkbookRegion("Structure Calcs", "P4", "Y4"),
        WorkbookRegion("Structure Assumptions", "E3", "E66"),
        WorkbookRegion("Structure Assumptions", "AC2", "AZ2"),
        WorkbookRegion("Structure Assumptions", "Z3", "Z66"),
        WorkbookRegion("Structure Assumptions", "AA3", "AA66"),
        WorkbookRegion("Structure Assumptions", "AB3", "AB66"),
        WorkbookRegion("Anchor Sizing", "L2", "N2"),
        WorkbookRegion("Structure Calcs", "B6", "B47"),
        WorkbookRegion("Operations", "D3", "Q3"),
        # WorkbookRegion("Operations", "B45", "B48"),
        # WorkbookRegion("Operations", "C119", "C119"),
        # WorkbookRegion("Operations", "C138", "C138"),
        WorkbookRegion("Equip&Mat Assumptions", "E4", "E47"),
        WorkbookRegion("Equip&Mat Assumptions", "N51", "N59"),
        # WorkbookRegion("Equip&Mat Assumptions", "C48", "C66"),
        WorkbookRegion("Equip&Mat Assumptions", "C51", "C69"),
        WorkbookRegion("Equip&Mat Calcs", "W2", "BN2"),
        WorkbookRegion("Equip&Mat Calcs", "D30", "G30"),
        WorkbookRegion("Equip&Mat Calcs", "B31", "B42"),
        WorkbookRegion("Equip&Mat Calcs", "C4", "C34"),
        WorkbookRegion("Equip&Mat Calcs", "D119", "G119"),
        WorkbookRegion("vessels", "F5", "V5"),
        WorkbookRegion("vessels", "F14", "V14"),
    ]

    # test_cell = CellDependency("Structure Calcs", "EO85")
    # test_cell = CellDependency("Structure Calcs", "DJ74")
    # test_cell = CellDependency("Operations", "F136")
    # test_cell = CellDependency("Structure Calcs", "ED79")
    test_cell = CellDependency("Operations", "I46")

    # expr = get_expr(wb.cell_dict[test_cell])
    # println("Before")
    # @display expr

    # for i in 1:3
    #     expr = lookup_const_propagate!(wb, expr, const_regions; debug=true)
    #     println("After $i")
    #     @display expr
    # end

    # return

    cell_dict = Dict{CellDependency, XLConvert.CellTypes}()

    for (cell, value) in wb.cell_dict
        # if !(cell.sheet_name in ["Structure Calcs", "Equip&Mat Calcs", "Anchor Sizing", "Equip&Mat Assumptions", "Operations"])
        #     cell_dict[cell] = value
        #     continue
        # end
        # if !(cell.sheet_name in ["Operations"])
        #     cell_dict[cell] = value
        #     continue
        # end

        if !(value isa XLConvert.FormulaCell)
            cell_dict[cell] = value
            continue
        end

        expr = get_expr(value)
        if !(expr isa XLConvert.FlatExpr)
            cell_dict[cell] = value
            continue
        end

        num_parts = length(expr.parts)

        for i in 1:4
            expr = lookup_const_propagate!(wb, expr, const_regions)
        end

        expr = fix_indirect!(wb, expr)

        removed_parts = num_parts - length(expr.parts)
        if cell == test_cell
            println("For cell $cell removed $(removed_parts) parts from the expr")
        end
        # if removed_parts > 0
        #     println("For cell $cell removed $(removed_parts) parts from the expr")
        # end

        cell_dict[cell] = XLConvert.FormulaCell(value.cell, expr)

    end

    cell_list = collect(keys(cell_dict))
    cell_numbering = XLConvert.ObjectNumbering(cell_list)

    edge_list = Vector{Edge{Int64}}()
    # A relatively random (and hopefully conservative) guess that the average degree is 2
    sizehint!(edge_list, 2 * length(cell_numbering))

    @time "getting expr dependencies" for (i, cell) in enumerate(cell_numbering.objs)
        content = get(cell_dict, cell, MissingCell())
        if content isa XLConvert.FormulaCell
            # empty!(handled_deps)
            dep_cells = []
            try
                dep_cells = XLConvert.get_expr_dependency_ranges(content.expr, wb.key_values)
            catch e
                println("Error getting cell dependencies for cell $cell")
                # @show cell
                # @show content
                @show content.cell.formula
                # println("Expr:")
                # show(stdout, "text/plain", content.expr)
                @show e
                # throw(e)
                continue
            end

            unique!(dep_cells)

            for dep in dep_cells
                if dep isa CellDependency
                    num = XLConvert.get_num!(cell_numbering, dep)
                    push!(edge_list, Edge(i, num))
                else
                    sheet = dep.first.sheet_name
                    (start_col, start_row) = XLConvert.start_coord(dep)
                    (end_col, end_row) = XLConvert.end_coord(dep)
                    for r ∈ start_row:end_row, c ∈ start_col:end_col
                        num = XLConvert.get_num!(cell_numbering, CellDependency(sheet, c, r))
                        push!(edge_list, Edge(i, num))
                    end
                end
            end
        end

    end

    @time "graph construction" graph = Graphs.SimpleDiGraph(edge_list)

    XLConvert.ExcelWorkbook(wb.xf, cell_numbering, cell_dict, graph, wb.key_values)
end

function table_statements_have_same_equation(a::XLConvert.TableStatement, func_a, params_a, b::XLConvert.TableStatement, func_b, params_b)
    if ismissing(a.rhs_expr) || ismissing(b.rhs_expr)
        return a.rhs_expr === b.rhs_expr
    end
    if ismissing(func_a) || ismissing(func_b)
        return func_a === func_b
    end

    # @show func_a func_b

    # func_a, params_a = XLConvert.functionalize(a.rhs_expr)
    # func_b, params_b = XLConvert.functionalize(b.rhs_expr)
    are_equal = func_a == func_b
    if ismissing(are_equal)
        # @show func_a func_b
        return false
    end

    if func_a != func_b
        return false
    end

    if length(params_a) != length(params_b)
        return false
    end

    row_a, col_a = @view a.lhs_expr.args[2:3]
    row_b, col_b = @view b.lhs_expr.args[2:3]

    delta_y = row_a - row_b
    delta_x = col_a - col_b

    for (pa, pb) in zip(params_a, params_b)
        if pa == pb
            continue
        end

        if !XLConvert.equal_with_offset(pa, pb, delta_y, delta_x)
            if !!XLConvert.equal_with_offset(pa, pb, delta_x, delta_y)
                return false
            end
        end
    end

    return true


    # base_expr = a.rhs_expr
    # return equal_with_offset(base_expr, b.rhs_expr, delta_y, delta_x)
    # offset_expr = offset(b.rhs_expr, delta_y, delta_x)

    # base_expr === offset_expr
end

function flexible_table_broadcast_transform_2d!(statements::Vector{AbstractStatement})
    stmt_graph = make_statement_graph(statements)
    # @time closure = transitiveclosure(stmt_graph)
    topo_sorted = topological_sort(reverse(stmt_graph))
    stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    # stmt_to_level = Dict((stmt => stmt_topo_levels[i]) for (i, stmt) in enumerate(statements))
    # table_statements::Vector{TableStatement} = filter(s -> isa(s, TableStatement), statements)
    table_statement_idxs::Vector{Int64} = filter(i -> (isa(statements[i], XLConvert.TableStatement) && length(XLConvert.get_set_cells(statements[i])) == 1), eachindex(statements))
    # @show table_statement_idxs statements[table_statement_idxs]
    # node_nums = used_subset.node_nums


    function get_stmt_table(stmt::XLConvert.TableStatement)
        stmt.lhs_expr.args[1]
    end

    function num_parts(::Any)
        1
    end
    function num_parts(e::XLConvert.FlatExpr)
        length(e.parts)
    end

    function get_stmt_num_expr_parts(stmt::XLConvert.TableStatement)
        num_parts(stmt.rhs_expr)
    end

    path_seen = zeros(Bool, nv(stmt_graph))
    function is_independent(node_a::Int64, node_b::Int64)
        lvl_a = stmt_topo_levels[node_a]
        lvl_b = stmt_topo_levels[node_b]
        if lvl_a == lvl_b
            return true
        end
        return false

        # has_path = if lvl_b > lvl_a
        #     has_path_within(stmt_graph, node_b, node_a, lvl_b - lvl_a, path_seen)
        # else
        #     has_path_within(stmt_graph, node_a, node_b, lvl_a - lvl_b, path_seen)
        # end

        # if has_path
        #     println("Nodes have the same equation but are not independent")
        #     println(statements[node_a])
        #     println(statements[node_b])

        # end

        # !has_path
    end

    # get_num_parts = i -> length(statements[i].rhs_expr.parts)
    get_num_parts = i -> get_stmt_num_expr_parts(statements[i])

    # inner_group_by_func = (i, j) -> (get_num_parts(i) == get_num_parts(j)) && table_statements_have_same_equation(statements[i], statements[j]) && is_independent(i, j)
    # inner_group_by_func = (i, j) -> table_statements_have_same_equation(statements[i], statements[j]) && is_independent(i, j)

    # same_table_and_level = group_to_dict(table_statement_idxs, i -> (get_stmt_table(statements[i]), stmt_topo_levels[i]))
    same_table_and_level = XLConvert.group_to_dict(table_statement_idxs, i -> get_stmt_table(statements[i]))
    @show length(same_table_and_level)
    groups = Vector{Vector{Int64}}()
    @time "grouping" for group in values(same_table_and_level)
        # @show length(group)
        # @show group
        # @time closure = transitiveclosure(stmt_graph)
        same_num_parts = XLConvert.group_to_dict(group, get_num_parts)
        # @show same_num_parts
        # @show length(same_num_parts)
        for sub_group in values(same_num_parts)
            # functionalized = XLConvert.funtcionalize.(sub_group)
            # @show length(sub_group)
            # functionalized = (i -> XLConvert.functionalize(statements[i])).(sub_group)
            functionalized = Dict(i => XLConvert.functionalize(statements[i].rhs_expr) for i in sub_group)
            same_equations = XLConvert.group_by(sub_group, (i, j) -> is_independent(i, j) && table_statements_have_same_equation(statements[i], functionalized[i][1], functionalized[i][2], statements[j], functionalized[j][1], functionalized[j][2]))
            # @show same_equations
            append!(groups, same_equations)
        end
        # same_equations = group_by(group, (i, j) -> table_statements_have_same_equation(statements[i], statements[j]) && not_interdependent(closure, i, j))
        # same_equations = group_by(group, inner_group_by_func)
        # append!(groups, same_equations)
    end
    # @show groups

    new_statements = copy(statements)

    grouped_by_level = groups

    num_table_stmts = 0

    for group in grouped_by_level
        if length(group) <= 1
            continue
        end
        table = statements[group[1]].lhs_expr.args[1]
        sheet = table.sheet_name

        # get_row_num = s -> rownum(s.assigned_vars[1])
        # get_col_num = s -> colnum(s.assigned_vars[1])
        get_row_num = i -> rownum(statements[i].assigned_vars[1])
        get_col_num = i -> colnum(statements[i].assigned_vars[1])

        row_nums = get_row_num.(group)
        col_nums = get_col_num.(group)
        coords = zip(col_nums, row_nums) |> collect
        coord_to_statement = Dict(c => s for (c, s) in zip(coords, group))

        sort!(coords)
        regions = XLConvert.get_2d_regions(coords)
        # println("Group size = $(length(group))")
        for region in regions
            cols, rows = region
            region_area = (length(cols) * length(rows))
            if region_area < 5
                continue
            end
            # if length(cols) > 1
            #     println("\t$region")
            #     println("\tSize: $(region_area)")
            # end
            # println("\t$region")
            # println("\tSize: $(region_area)")


            # # region_coords = vec([(c, r) for c in cols, r in rows])
            # # run_statements = [coord_to_statement[coord] for coord in region_coords]
            # # region_coords = vec([(c, r) for c in cols, r in rows])
            run_statements = vec([coord_to_statement[(c, r)] for c in cols, r in rows])

            # first_expr = statements[run_statements[1]].rhs_expr

            # broadcasted = try
            #     convert_to_broadcasted(first_expr, length(rows) - 1, length(cols) - 1)
            # catch ex
            #     println("Failed to broadcast, would have broadcasted a run of $(region_area)")
            #     @show ex
            #     @show sheet rows cols
            #     show(stdout, "text/plain", first_expr)

            #     statement_group = statements[run_statements]
            #     statement_set = Set(statement_group)
            #     length_before = length(new_statements)
            #     # filter!(s -> !(s in statement_group), new_statements)
            #     filter!(!in(statement_set), new_statements)
            #     length_after = length(new_statements)
            #     println("Remove $(length_before - length_after) statements to put them in a group")
            #     push!(new_statements, GroupedStatement(statement_group))
            #     continue
            # end

            length_before = length(new_statements)
            statement_group = statements[run_statements]
            statement_set = Set(statement_group)
            filter!(!in(statement_set), new_statements)
            # # filter!(s -> !(s in statements[run_statements]), new_statements)
            length_after = length(new_statements)
            # println("Remove $(length_before - length_after) statements")
            push!(new_statements, XLConvert.GroupedStatement(sort(statement_group, by = s -> XLConvert.get_set_cells(s)[1])))


            # # lhs_expr = ExcelExpr(:table_ref, table, (run_idx[begin]:run_idx[end]) .- startrow(table) .+ 1, run_statements[1].lhs_expr.args[3], (true, true), (true, true))
            # table_rows = rows .- startrow(table) .+ 1
            # table_cols = cols .- startcol(table) .+ 1
            # # @show table
            # # @show table_rows, table_cols
            # lhs_expr = ExcelExpr(:table_ref, table, table_rows, table_cols, (true, true), (true, true))
            # # println("Broadcasting $(region_area) statements together. $(lhs_expr)")
            # # @show statement_group
            # # @show table_statements_have_same_equation(statement_group[1], statement_group[2])
            # # @show group statements[group]
            # # @show inner_group_by_func(group[1], group[2])
            # # @assert table_statements_have_same_equation(statement_group[1], statement_group[2])


            # # lhs_vars = reduce(vcat, get_set_cells.(statements[run_statements]))
            # lhs_vars = reduce(vcat, get_set_cells.(statement_group))
            # # @show lhs_vars
            # # new_statement = TableStatement(lhs_expr, lhs_vars, broadcasted, reduce(vcat, get_cell_deps.(statements[run_statements])) |> unique |> collect, true)
            # new_statement = TableStatement(lhs_expr, lhs_vars, broadcasted, reduce(vcat, get_cell_deps.(statement_group)) |> unique |> collect, true)

            # push!(new_statements, new_statement)
            num_table_stmts += 1

        end
    end

    length_before = length(statements)
    length_after = length(new_statements)
    println("flexible_table_broadcast_transform_2d!: Removed $(length_before - length_after) statements")
    println("flexible_table_broadcast_transform_2d!: Created $num_table_stmts group statements")

    new_statements
end


function StandardTable(xf, sheet, name, top_left, bottom_right, column_name_row, row_name_col)
    top_left_cell = CellDependency(sheet, top_left)
    bottom_right_cell = CellDependency(sheet, bottom_right)

    column_start = CellDependency(sheet, colnum(top_left_cell), column_name_row)
    column_end = CellDependency(sheet, colnum(bottom_right_cell), column_name_row)

    row_col_num = XLSX.decode_column_number(row_name_col)

    row_start = CellDependency(sheet, row_col_num, rownum(top_left_cell))
    row_end = CellDependency(sheet, row_col_num, rownum(bottom_right_cell))

    DefTable(xf, sheet, name, top_left, bottom_right, "$(column_start.cell):$(column_end.cell)", "$(row_start.cell):$(row_end.cell)")
end
function StandardTable(xf, sheet, name, top_left, bottom_right, column_name_row)
    top_left_cell = CellDependency(sheet, top_left)
    bottom_right_cell = CellDependency(sheet, bottom_right)

    column_start = CellDependency(sheet, colnum(top_left_cell), column_name_row)
    column_end = CellDependency(sheet, colnum(bottom_right_cell), column_name_row)

    DefTable(xf, sheet, name, top_left, bottom_right, "$(column_start.cell):$(column_end.cell)", "")
end

function split_range(range_ref::AbstractString)
    start_ref, end_ref = string.(split(range_ref, ":", limit = 2))
    start_ref, end_ref
end

function NameRangeTable(xf, sheet, name, range_ref::AbstractString)
    top_left, bottom_right = split_range(range_ref)
    DefTable(xf, sheet, name, top_left, bottom_right, "", "")
end

function ColumnNamesTable(xf, source::ExcelTable, name)
    isempty(source.column_names_range) && error("Table $(getname(source)) has no column names range")
    top_left, bottom_right = split_range(source.column_names_range)
    DefTable(xf, source.sheet_name, name, top_left, bottom_right, string(top_left, ":", bottom_right), "")
end

function RowNamesTable(xf, source::ExcelTable, name)
    if ismissing(source.row_names_range) || isempty(source.row_names_range)
        error("Table $(getname(source)) has no row names range")
    end
    top_left, bottom_right = split_range(source.row_names_range)
    DefTable(xf, source.sheet_name, name, top_left, bottom_right, string(top_left, ":", top_left), string(top_left, ":", bottom_right))
end

function table_by_name(tables, table_name::AbstractString)
    idx = findfirst(t -> t.table_name == table_name, tables)
    isnothing(idx) && error("Table \"$table_name\" was not found")
    tables[idx]
end

function add_table!(tables, xf, sheet, spec::Tuple{AbstractString, AbstractString, AbstractString, Integer})
    name, top_left, bottom_right, column_name_row = spec
    push!(tables, StandardTable(xf, sheet, name, top_left, bottom_right, column_name_row))
end

function add_table!(tables, xf, sheet, spec::Tuple{AbstractString, AbstractString, AbstractString, Integer, AbstractString})
    name, top_left, bottom_right, column_name_row, row_name_col = spec
    push!(tables, StandardTable(xf, sheet, name, top_left, bottom_right, column_name_row, row_name_col))
end

function add_table!(tables, xf, sheet, spec::Tuple{AbstractString, AbstractString, AbstractString, AbstractString, AbstractString})
    name, top_left, bottom_right, column_names_range, row_names_range = spec
    push!(tables, DefTable(xf, sheet, name, top_left, bottom_right, column_names_range, row_names_range))
end

function add_tables!(tables, xf, sheet, specs)
    for spec in specs
        add_table!(tables, xf, sheet, spec)
    end
end

function add_row_names_table!(tables, xf, source_table_name::AbstractString, new_table_name::AbstractString)
    push!(tables, RowNamesTable(xf, table_by_name(tables, source_table_name), new_table_name))
end

function add_column_names_table!(tables, xf, source_table_name::AbstractString, new_table_name::AbstractString)
    push!(tables, ColumnNamesTable(xf, table_by_name(tables, source_table_name), new_table_name))
end

function make_vessels_tables(xf)
    tables = ExcelTable[]

    add_tables!(tables, xf, "vessels", [
        ("vsl_choices", "F14", "T43", 4, "B"),
    ])
    add_tables!(tables, xf, "vessels", [
        ("type_info", "F5", "T9", 4, "B"),
        ("from_library", "F103", "T122", 6, "B"),
        ("intermediate_res", "F145", "T151", 6, "B"),
        ("recorded_results", "F61", "T99", 4, "B"),
    ])

    tables
end

function make_structure_calcs_tables(xf)
    tables = ExcelTable[]

    add_tables!(tables, xf, "Structure Calcs", [
        ("tasks", "CN4", "FH4", 4),
        ("component_calc", "B6", "CG47", 4, "B"),
        ("component_vessel_task_times", "CN6", "FH47", 4, "CM"),
        ("component_deck_hand_time", "CN51", "FH92", 4, "CM"),
        ("component_task_quantity", "CN100", "FH141", 4, "CM"),
        ("component_task_annual_vsl_time", "CN149", "FH190", 4, "CM"),
        ("component_task_annual_deck_hand_time", "CN197", "FH238", 4, "CM"),
        ("monthly_machine_use", "CN251", "FH262", 4, "CM"),
        ("task_location", "CN242", "FO242", 4),
        ("task_operation", "CN243", "FO243", 4),
        ("task_vessel", "CN244", "FO244", 4),
    ])
    add_tables!(tables, xf, "Structure Calcs", [
        ("vssl_A_hours", "C55", "F68", 54, "B"),
        ("vssl_B_hours", "C75", "F88", 74, "B"),
        ("vssl_C_hours", "C95", "F108", 94, "B"),
        ("vssl_A_deck_hand_hours", "H55", "K68", 54, "B"),
        ("vssl_B_deck_hand_hours", "H75", "K88", 74, "B"),
        ("vssl_C_deck_hand_hours", "H95", "K108", 94, "B"),
        ("machine_hours", "C114", "F127", 113, "B"),
        ("machine_deck_hand_hours", "H114", "K127", 113, "B"),
    ])
    add_row_names_table!(tables, xf, "vssl_A_hours", "vssl_A_ops")
    add_row_names_table!(tables, xf, "vssl_B_hours", "vssl_B_ops")
    add_row_names_table!(tables, xf, "vssl_C_hours", "vssl_C_ops")
    add_row_names_table!(tables, xf, "machine_hours", "machine_hours_rows")
    add_column_names_table!(tables, xf, "vssl_A_hours", "vssl_A_hours_loc")
    add_column_names_table!(tables, xf, "vssl_B_hours", "vssl_B_hours_loc")
    add_column_names_table!(tables, xf, "vssl_C_hours", "vssl_C_hours_loc")
    add_column_names_table!(tables, xf, "machine_hours", "machine_hours_loc")
    add_column_names_table!(tables, xf, "vssl_A_deck_hand_hours", "vssl_A_deck_hand_loc")
    add_column_names_table!(tables, xf, "vssl_B_deck_hand_hours", "vssl_B_deck_hand_loc")
    add_column_names_table!(tables, xf, "vssl_C_deck_hand_hours", "vssl_C_deck_hand_loc")
    add_column_names_table!(tables, xf, "machine_deck_hand_hours", "machine_deck_hand_hours_loc")

    tables
end

function make_operations_tables(xf)
    tables = ExcelTable[]

    add_tables!(tables, xf, "Operations", [
        ("vessel_inputs", "D4", "Q18", 3, "B"),
        ("vsl_A_annual_time", "D45", "Q48", 3, "B"),
        ("vsl_B_annual_time", "D51", "Q54", 3, "B"),
        ("vsl_C_annual_time", "D57", "Q60", 3, "B"),
        ("vsl_A_deck_hand_annual", "D64", "Q67", 3, "B"),
        ("vsl_B_deck_hand_annual", "D70", "Q73", 3, "B"),
        ("vsl_C_deck_hand_annual", "D76", "Q79", 3, "B"),
        ("non_vessel_deck_hand_annual_requirements", "D82", "Q85", 3, "B"),
        ("dates", "D87", "Q88", 3, "B"),
        ("work_hours_per_day", "D90", "Q90", 3, "B"),
        ("capacity_per_month", "D93", "Q104", 3, "B"),
        ("month_used_even_years", "D108", "Q119", 3, "B"),
        ("month_used_odd_years", "D122", "Q133", 3, "B"),
        ("crew_reqs", "D135", "Q136", 3, "B"),
        # ("vessel_A", "D139", "Q183", 138, "B"),
        # ("vessel_B", "D187", "Q231", 186, "B"),
        # ("vessel_C", "D234", "Q278", 233, "B"),
        ("vessel_A", "D138", "Q183", 3, "B"),
        ("vessel_B", "D186", "Q231", 3, "B"),
        ("vessel_C", "D233", "Q278", 3, "B"),
        ("non_vessel_ops", "D281", "Q283", 3, "B"),
        ("num_employees_required", "D286", "Q294", 3, "B"),
        ("empoloyment_totals", "D296", "Q301", 3, "B"),
        ("financing_mult", "D304", "Z304", 3, "B"),
        ("vessel_months_needed", "D343", "E354", 336, "B"),
        # ("min_vessels_required", "D357", "E372", 336, "B"),
        ("min_vessels_required", "D357", "E368", 336, "B"),
        ("vessels_required", "D372", "E372", 336, "B"),
        ("weather_day_portion", "D394", "E405", 336, "B"),
        ("maintenance_dates", "D407", "E408", 336, "B"),
        ("vessel_days_used", "D411", "E422", 336, "B"),
        ("vessel_days_rented", "D425", "E436", 336, "B"),
        ("annual_costs", "D307", "Q315", 3, "B"),
        ("total_costs", "R307", "R315", 306, "B"),
    ])
    add_tables!(tables, xf, "Operations", [
        ("month_info", "T109", "X132", 108),
    ])
    add_row_names_table!(tables, xf, "month_used_even_years", "even_years_months")
    add_row_names_table!(tables, xf, "month_used_odd_years", "odd_years_months")
    add_row_names_table!(tables, xf, "vessel_days_rented", "vessel_days_rented_rows")
    add_row_names_table!(tables, xf, "vsl_A_annual_time", "vsl_A_annual_time_rows")
    add_row_names_table!(tables, xf, "vsl_B_annual_time", "vsl_B_annual_time_rows")
    add_row_names_table!(tables, xf, "vsl_C_annual_time", "vsl_C_annual_time_rows")
    add_row_names_table!(tables, xf, "vsl_A_deck_hand_annual", "vsl_A_deck_hand_annual_rows")
    add_row_names_table!(tables, xf, "vsl_B_deck_hand_annual", "vsl_B_deck_hand_annual_rows")
    add_row_names_table!(tables, xf, "vsl_C_deck_hand_annual", "vsl_C_deck_hand_annual_rows")
    add_row_names_table!(tables, xf, "vessel_days_used", "vessel_days_used_rows")
    add_column_names_table!(tables, xf, "vessel_A", "vessel_A_cols")
    add_column_names_table!(tables, xf, "vessel_B", "vessel_B_col")
    add_column_names_table!(tables, xf, "vessel_C", "vessel_C_col")

    tables
end

function make_structure_assumptions_tables(xf)
    tables = ExcelTable[]

    add_tables!(tables, xf, "Structure Assumptions", [
        ("task_settings", "B3", "BA73", 2, "F"),
    ])
    add_tables!(tables, xf, "Structure Assumptions", [
        ("assembly", "D77", "AH100", 76, "C"),
    ])

    tables
end

function make_growth_and_site_tables(xf)
    tables = ExcelTable[]

    add_tables!(tables, xf, "growth model", [
        ("growth_curve", "B21", "G274", 18),
        ("harvest_period", "H21", "Y117", 18),
    ])
    add_tables!(tables, xf, "Site Inputs", [
        ("weather_days", "C68", "AD88", 67),
        ("sig_wave_height", "B21", "Q59", 20),
        ("oyster_site", "C103", "P2677", 101),
    ])
    add_tables!(tables, xf, "Site Inputs", [
        ("vessel_weather_days", "AA21", "AF32", 20, "AA"),
        ("oyster_site_params", "L95", "P98", 101, "K"),
    ])

    tables
end

function make_machines_tables(xf)
    tables = ExcelTable[]
    table_names = [
        ("info", "D3", "AB27"),
        ("estimated_cost_breakdown", "D30", "AB36"),
        ("calcs", "D38", "AB56"),
        ("power_support", "D59", "AB64"),
        ("combined", "D67", "AB71"),
        ("used_on_vessel", "D74", "AB87"),
        ("quantity_required", "D90", "AB103"),
        ("num_required_in_year", "D104", "AB104"),
        ("rented_machinery", "D112", "AB114"),
        ("time_needed", "D119", "AB130"),
        ("min_needed", "D133", "AB144"),
    ]
    add_tables!(tables, xf, "Machines", [(name, top_left, bottom_right, 2, "B") for (name, top_left, bottom_right) in table_names])

    for name in first.(table_names)
        add_row_names_table!(tables, xf, name, "$(name)_rows")
    end

    tables
end

function make_equip_mat_tables(xf)
    tables = ExcelTable[]

    add_tables!(tables, xf, "Equip&Mat Assumptions", [
        ("tasks", "B4", "BG47", 3, "H"),
        ("equipment", "C51", "AG69", 47, "C"),
        ("materials", "C73", "N90", 69, "C"),
    ])

    add_tables!(tables, xf, "Equip&Mat Calcs", [
        ("tasks", "W2", "BN2", 2),
    ])
    add_tables!(tables, xf, "Equip&Mat Calcs", [
        ("vssl_or_machine_time", "W4", "BN34", 2, "B"),
        ("deck_hand_time", "W38", "BN66", 2, "U"),
        ("monthly_machine_use", "W77", "BN88", 2, "V"),
        ("task_location", "W69", "BU69", 2),
        ("task_operation", "W70", "BU70", 2),
        ("task_vessel", "W71", "BU71", 2),
        ("task_machine", "W73", "BU73", 2),
        ("equipment", "B4", "N34", 2, "B"),
    ])
    add_tables!(tables, xf, "Equip&Mat Calcs", [
        ("per_item_names", "K2", "N2", 2),
    ])
    add_tables!(tables, xf, "Equip&Mat Calcs", [
        ("vssl_A_hours", "D120", "G133", 119, "B"),
        ("vssl_B_hours", "D139", "G152", 138, "B"),
        ("vssl_C_hours", "D157", "G170", 156, "B"),
    ])
    add_tables!(tables, xf, "Equip&Mat Calcs", [
        ("non_vssl_hours", "D175", "G188", 174),
        ("non_vssl_hours_loc", "D174", "G174", 174),
    ])
    add_row_names_table!(tables, xf, "vssl_A_hours", "vssl_A_ops")
    add_row_names_table!(tables, xf, "vssl_B_hours", "vssl_B_ops")
    add_row_names_table!(tables, xf, "vssl_C_hours", "vssl_C_ops")
    add_column_names_table!(tables, xf, "vssl_A_hours", "vssl_A_hours_loc")
    add_column_names_table!(tables, xf, "vssl_B_hours", "vssl_B_hours_loc")
    add_column_names_table!(tables, xf, "vssl_C_hours", "vssl_C_hours_loc")

    add_tables!(tables, xf, "Equip&Mat Calcs", [
        ("deck_hand_vssl_A_hours", "K120", "N133", 119, "B"),
    ])
    add_tables!(tables, xf, "Equip&Mat Calcs", [
        ("deck_hand_vssl_A_hours_loc", "K119", "N119", 119),
    ])
    add_tables!(tables, xf, "Equip&Mat Calcs", [
        ("deck_hand_vssl_B_hours", "K139", "N152", 138, "B"),
    ])
    add_tables!(tables, xf, "Equip&Mat Calcs", [
        ("deck_hand_vssl_B_hours_loc", "K138", "N138", 138),
    ])
    add_tables!(tables, xf, "Equip&Mat Calcs", [
        ("deck_hand_vssl_C_hours", "K157", "N170", 156, "B"),
    ])
    add_tables!(tables, xf, "Equip&Mat Calcs", [
        ("deck_hand_vssl_C_hours_loc", "K156", "N156", 156),
        ("deck_hand_non_vssl_hours", "K175", "N188", 174),
        ("deck_hand_non_vssl_hours_loc", "K174", "N174", 174),
    ])

    tables
end

function make_labor_tables(xf)
    tables = ExcelTable[]

    add_tables!(tables, xf, "Labor", [
        ("inputs", "D6", "L10", 5, "B"),
        ("required_per_month", "D13", "L24", 5, "B"),
        ("monthly_overtime", "D27", "L38", 5, "B"),
        ("totals", "D40", "L46", 5, "B"),
    ])


    tables
end

function make_misc_tables(xf)
    tables = ExcelTable[]

    add_tables!(tables, xf, "3. Seed", [
        ("annual_energy", "AE4", "AI10", 3, "AD"),
        ("capex", "AO4", "BI30", 3, "AM"),
    ])
    add_tables!(tables, xf, "oyster gear", [
        ("gear", "B5", "AB28", 3),
    ])
    add_tables!(tables, xf, "KAM Results", [
        ("structural", "B9", "G47", 8, "B"),
    ])
    add_tables!(tables, xf, "Vessel_Library", [
        ("library", "D6", "Q39", 3, "B"),
        ("workable_wave_height", "E42", "Q42", 41, "B"),
        ("weather_day_portion", "E43", "Q54", 41, "B"),
    ])
    add_tables!(tables, xf, "standard tasks", [
        ("structure_related", "B3", "AX66", 2, "F"),
        ("equip_related", "B70", "Z143", 2, "F"),
        ("equip_flags", "AA70", "BC143", 69),
    ])
    add_tables!(tables, xf, "material properties", [
        ("props", "C3", "W37", 2, "B"),
    ])
    add_tables!(tables, xf, "Anchor Sizing", [
        ("anchors", "L3", "N44", 2, "J"),
        ("deployment", "C58", "D60", 57, "C"),
        ("sediment_type", "C63", "F66", 62, "C"),
    ])
    add_tables!(tables, xf, "Anchor Sizing", [
        ("k_soils", "J82", "M82", 82),
        ("a_soils", "D82", "H82", 82),
    ])
    add_tables!(tables, xf, "growth- cohort group 1", [
        ("growth", "B10", "BX157", 9),
    ])

    add_tables!(tables, xf, "Results", [
        ("annual_cost_summary", "G3", "I16", 9, "F"),
    ])

    tables
end

function fix_indirect!(wb::XLConvert.ExcelWorkbook2, expr_in::XLConvert.FlatExpr; debug::Bool = false)
    expr = copy(expr_in)
    region_a = WorkbookRegion("Structure Assumptions", "I77", "I100")
    switch_names_ = ["none", "transverse_line_quant", "mid_mooring_floats_per_mooring_leg", "GL_sets", "seg_per_trans_line", "GL_per_set"]
    switch_args_a = Any[]
    for n in switch_names_
        push!(switch_args_a, n)
        push!(switch_args_a, ExcelExpr(:named_range, n))
    end

    region_b = WorkbookRegion("Structure Assumptions", "K77", "K100")
    switch_names_ = ["none", "transverse_line_quant", "load_anchor_line", "load_cult_line", "load_transverse_line", "GL_per_set"]
    switch_args_b = Any[]
    for n in switch_names_
        push!(switch_args_b, n)
        push!(switch_args_b, ExcelExpr(:named_range, n))
    end

    region_c = WorkbookRegion("Structure Assumptions", "F77", "F100")
    switch_names_ = ["mooring_assemb_onshore_tog", "at_base", "HDPE_pipe_deploy", "on_farm"]
    switch_args_c = Any[]
    for n in switch_names_
        push!(switch_args_c, n)
        push!(switch_args_c, ExcelExpr(:named_range, n))
    end

    region_args = [(region_a, switch_args_a), (region_b, switch_args_b), (region_c, switch_args_c)]

    did_work = false
    for (i, part) in enumerate(expr.parts)
        @flat_match expr part begin
            ExcelExpr(:call, ["INDIRECT", ExcelExpr(:cell_ref, [cell, sheet])]) => begin
                c = CellDependency(sheet, cell)
                for (region, switch_args) in region_args
                    if c in region
                        expr.parts[i] = ExcelExpr(:call, Any["SWITCH", part.args[2], switch_args...])
                        did_work = true
                        break
                    end
                end
                # if CellDependency(sheet, cell) in region_a
                #     expr.parts[i] = ExcelExpr(:call, Any["SWITCH", part.args[2], switch_args_a...])
                #     did_work = true
                # end
            end
            ExcelExpr(:call, ["INDIRECT", ExcelExpr(:call, ["_xlfn.XLOOKUP", val, ref, ExcelExpr(:range, [ExcelExpr(:cell_ref, [start, sheet]), ExcelExpr(:cell_ref, [stop, sheet])])])]) => begin
                r = WorkbookRegion(sheet, start, stop)
                for (region, switch_args) in region_args
                    if r in region
                        expr.parts[i] = ExcelExpr(:call, Any["SWITCH", part.args[2], switch_args...])
                        did_work = true
                        break
                    end
                end
                # if WorkbookRegion(sheet, start, stop) in region_a
                #     # @display expr_in
                #     # XLConvert.toexpr("IF($cell = \"none\", none, IF($cell = ))")
                #     expr.parts[i] = ExcelExpr(:call, Any["SWITCH", part.args[2], switch_args_a...])
                #     did_work = true
                # end
            end
            ExcelExpr(:call, ["INDIRECT", rest...]) => begin
                # println("Found XLOOKUP that wasn't handled")

            end
            _ => continue
        end

    end

    if did_work
        expr = XLConvert.convert_to_flat_expr(XLConvert.convert_to_expr(expr))
        # @display expr
    end

    expr
end

function indirect_fixes(wb::XLConvert.ExcelWorkbook2)

    cell_dict = Dict{CellDependency, XLConvert.CellTypes}()

    for (cell, value) in wb.cell_dict
        # if !(cell.sheet_name in ["Structure Calcs", "Equip&Mat Calcs", "Anchor Sizing", "Equip&Mat Assumptions", "Operations"])
        #     cell_dict[cell] = value
        #     continue
        # end

        if !(value isa XLConvert.FormulaCell)
            cell_dict[cell] = value
            continue
        end

        expr = get_expr(value)
        if !(expr isa XLConvert.FlatExpr)
            cell_dict[cell] = value
            continue
        end

        num_parts = length(expr.parts)
        # @show cell


        expr = fix_indirect!(wb, expr)
        # removed_parts = num_parts - length(expr.parts)
        # if removed_parts > 0
        #     println("For cell $cell removed $(removed_parts) parts from the expr")
        # end

        cell_dict[cell] = XLConvert.FormulaCell(value.cell, expr)

    end

    cell_list = collect(keys(cell_dict))
    cell_numbering = XLConvert.ObjectNumbering(cell_list)

    edge_list = Vector{Edge{Int64}}()
    # A relatively random (and hopefully conservative) guess that the average degree is 2
    sizehint!(edge_list, 2 * length(cell_numbering))

    @time "getting expr dependencies" for (i, cell) in enumerate(cell_numbering.objs)
        content = get(cell_dict, cell, MissingCell())
        if content isa XLConvert.FormulaCell
            # empty!(handled_deps)
            dep_cells = []
            try
                dep_cells = XLConvert.get_expr_dependency_ranges(content.expr, wb.key_values)
            catch e
                println("Error getting cell dependencies for cell $cell")
                # @show cell
                # @show content
                @show content.cell.formula
                # println("Expr:")
                # show(stdout, "text/plain", content.expr)
                @show e
                # throw(e)
                continue
            end

            unique!(dep_cells)

            for dep in dep_cells
                if dep isa CellDependency
                    num = XLConvert.get_num!(cell_numbering, dep)
                    push!(edge_list, Edge(i, num))
                else
                    sheet = dep.first.sheet_name
                    (start_col, start_row) = XLConvert.start_coord(dep)
                    (end_col, end_row) = XLConvert.end_coord(dep)
                    for r ∈ start_row:end_row, c ∈ start_col:end_col
                        num = XLConvert.get_num!(cell_numbering, CellDependency(sheet, c, r))
                        push!(edge_list, Edge(i, num))
                    end
                end
            end
        end

    end

    @time "graph construction" graph = Graphs.SimpleDiGraph(edge_list)

    XLConvert.ExcelWorkbook(wb.xf, cell_numbering, cell_dict, graph, wb.key_values)

end

function get_subset(wb_in::XLConvert.ExcelWorkbook2)
    wb = wb_in

    target_output = CellDependency("Results", "C6")

    # Structural total annual cost
    # target_output = CellDependency("Structure Calcs", "BV48")
    # Structural total cost
    # target_output = CellDependency("Structure Calcs", "BO48")
    # target_output = CellDependency("Equip&Mat Calcs", "W65")
    # target_output = CellDependency("Operations", "R313")

    # target_output = CellDependency("vessels", "AA30")
    # all_target_outputs = [target_output, CellDependency("Structure Calcs", "CS6")]
    # all_target_outputs = vec(XLConvert.cells(WorkbookRegion("Site Inputs", "L103", "P2677")))
    all_target_outputs = [target_output]
    # @show all_target_outputs
    # inputs = [CellDependency("vessels", "AA13"), CellDependency("vessels", "AA16"), CellDependency("vessels", "AA17"), CellDependency("vessels", "AA18")]
    @show nv(wb.cell_graph) ne(wb.cell_graph)

    # test_cell = CellDependency("Operations", "I46")

    @time "lookup_const_propagate" wb = test_lookup_const_propagate(wb)
    @show nv(wb.cell_graph) ne(wb.cell_graph)


    # dep_forwards = get_transpose_dep_forwards(wb)
    # apply_dep_forwards!(wb, dep_forwards)


    inputs = Vector{CellDependency}()
    append!(inputs, XLConvert.cells(WorkbookRegion("aggregated oyster growth", "F3", "G157")))
    append!(inputs, XLConvert.cells(WorkbookRegion("aggregated oyster growth", "J11", "K157")))
    append!(inputs, XLConvert.cells(WorkbookRegion("Structure Calcs", "FG5", "FH262")))
    append!(inputs, XLConvert.cells(WorkbookRegion("Operations", "F337", "H448")))

    @time used_subset = XLConvert.get_workbook_subset(wb, all_target_outputs)

    @time used_subset = XLConvert.force_cells_to_be_value(used_subset, inputs)

    if check_cycles(used_subset)
        throw("Subset has cycles!")
    end


    @show nv(used_subset.cell_graph) ne(used_subset.cell_graph)

    used_subset
end

function check_cycles(used_subset::XLConvert.ExcelWorkbook2)
    cycle = find_cycle(used_subset.cell_graph)
    if !isnothing(cycle)
        println("Found a cycle!")
        for (i, node) in enumerate(cycle)
            cell = XLConvert.get_cell(used_subset, node)
            # println("\t$i: $cell, $(get_cell_value(used_subset, cell))")
            println("\t$i: $cell")
            @display get_expr(used_subset.cell_dict[cell])
        end
        true
    else
        println("Didn't find any cycles")
        false
    end
end

function make_tables(used_subset::XLConvert.ExcelWorkbook2)
    xf = used_subset.xf
    tables = ExcelTable[]

    append!(tables, make_vessels_tables(xf))
    append!(tables, make_structure_calcs_tables(xf))
    append!(tables, make_operations_tables(xf))
    append!(tables, make_structure_assumptions_tables(xf))
    append!(tables, make_growth_and_site_tables(xf))
    append!(tables, make_machines_tables(xf))
    append!(tables, make_equip_mat_tables(xf))
    append!(tables, make_labor_tables(xf))
    append!(tables, make_misc_tables(xf))

    extra_ranges = [
        ("oyster Husbandry model", "A13", "JI25")
        ("oyster Husbandry model", "A31", "JI116")
        ("oyster Husbandry model", "I29", "AF30")
        ("oyster Husbandry model", "I8", "AF9")
        ("oyster Husbandry model", "CK8", "DH11")
        ("oyster Husbandry model", "DM6", "DW11")
        ("oyster Husbandry model", "EE5", "EO11")
        ("oyster Husbandry model", "HC6", "HM11")
        ("oyster Husbandry model", "IH7", "IR11")
        ("oyster Husbandry model", "FI6", "FS11")
        ("oyster Husbandry model", "GK6", "GU11")
        ("oyster Husbandry model", "GK28", "GU28")
        ("oyster Husbandry model", "AI30", "BF30")
        ("oyster Husbandry model", "CK30", "DZ30")
        # ("growth- cohort group 1", "B10", "BX157")
        ("growth- cohort group 2", "B11", "BX82")
        ("growth- cohort group 3", "B11", "BX82")
        ("aggregated oyster growth", "E4", "G4")
        ("aggregated oyster growth", "I4", "K4")
        ("Lists", "B15", "B18")
        ("Equip&Mat Assumptions", "C95", "C98")
    ]

    for (sheet, top_left, bottom_right) in extra_ranges
        left, top = XLConvert.parse_cell(top_left)
        right, bottom = XLConvert.parse_cell(bottom_right)

        table_name = "$(top_left)_$(bottom_right)"
        col_names = XLSX.encode_column_number.(left:right)
        # old_name = getname(table_to_grow)
        table = ExcelTable(sheet, table_name, top_left, bottom_right, "", "", col_names, missing)
        push!(tables, table)
    end
    # tables = [propulsion_table]
    # tables = XLConvert.ExcelTable[]
    @time find_tables!(tables, used_subset)

    tables
end

function group_and_func_statements(statements)
    println("-"^40)
    println("Group statements")
    println("-"^40)
    @time grouped_statements = group_statements(statements)
    println("-"^40)
    println("Group statements (again)")
    println("-"^40)
    @time grouped_statements = group_statements(grouped_statements)
    println("-"^40)
    println("Add functions")
    println("-"^40)
    @time statements_with_funcs = add_functions(grouped_statements, min_intermediates = 2)

    statements_with_funcs
end

function make_raw_statements(used_subset::XLConvert.ExcelWorkbook2, tables)
    @time "Make statements" statements = make_statements(used_subset)
    @time "if_multiple_transform" if_multiple_transform!(statements)
    @time "if_toggle_transform" if_toggle_transform!(statements)
    @time "round_if_transform" round_if_transform!(statements)
    # @time "is_blank_transform" is_blank_transform!(statements)
    @time "table_ref_transform" table_ref_transform!(statements, tables)
    table_stmts = filter(s -> s isa XLConvert.TableStatement, statements)
    @show length(table_stmts)

    statements
end

function generate_statements(used_subset::XLConvert.ExcelWorkbook2, tables)
    statements = begin
        @time "Make statements" statements = make_statements(used_subset)
        @time "if_multiple_transform" if_multiple_transform!(statements)
        @time "if_toggle_transform" if_toggle_transform!(statements)
        @time "round_if_transform" round_if_transform!(statements)
        @time "if_one_zero_transform" if_one_zero_transform!(statements)
        @time "average_two_cell_range_transform" average_two_cell_range_transform!(statements)
        # @time "is_blank_transform" is_blank_transform!(statements)
        @time "table_ref_transform" table_ref_transform!(statements, tables)
        @time "indirect_transform!" indirect_transform!(statements)
        table_stmts = filter(s -> s isa XLConvert.TableStatement, statements)
        @show length(table_stmts)
        # return statements
        @time "table_col_row_name_transform" table_col_row_name_transform!(statements, used_subset.xf)

        println("-"^40)
        println("new broadcast")
        println("-"^40)
        @time statements = new_broadcast(statements, debug = false)

        println("-"^40)
        println("new broadcast again")
        println("-"^40)
        @time statements = new_broadcast(statements, debug = false)


        # println("-"^40)
        # println("Table Broadcast Transform")
        # println("-"^40)
        # @time statements = table_broadcast_transform_2d!(statements)
        # println("-"^40)
        # println("Flexible Table Broadcast Transform")
        # println("-"^40)
        # @time statements = flexible_table_broadcast_transform_2d!(statements)
        println("-"^40)
        println("Group statements")
        println("-"^40)
        @time grouped_statements = group_statements(statements)
        println("-"^40)
        println("Add functions")
        println("-"^40)
        @time statements_with_funcs = add_functions(grouped_statements, min_intermediates = 2)
        # println("-"^40)
        # println("Group statements (again)")
        # println("-"^40)
        # @time statements_with_funcs = group_statements(statements_with_funcs)
        # statements_with_funcs = add_functions(statements, min_intermediates=2)

        statements_with_funcs
    end

    statements

end

function analyze_sheet_dependencies(statements::Vector{AbstractStatement}, stmt_graph::DiGraph{Int64}; test_sheet = nothing)
    function get_statement_sheets(stmt::AbstractStatement)
        set_cells = get_set_cells(stmt)
        unique((c -> c.sheet_name).(set_cells))
    end

    function summarize_sheet_list(sheets::Vector{String}; max_items::Int = 3)
        isempty(sheets) && return "<none>"
        sorted_sheets = sort(sheets)
        shown = sorted_sheets[1:min(max_items, length(sorted_sheets))]
        out = join(shown, ", ")
        if length(sorted_sheets) > max_items
            out *= ", ..."
        end
        out
    end

    function get_statement_label(i::Int)
        set_cells = sort(get_set_cells(statements[i]))
        sheets = get_statement_sheets(statements[i])
        sheet_suffix = length(sheets) <= 1 ? "" : " [sheets: $(summarize_sheet_list(sheets))]"

        if isempty(set_cells)
            return "stmt[$i]$sheet_suffix"
        elseif length(set_cells) == 1
            c = set_cells[1]
            return "$(c.sheet_name)!$(c.cell)$sheet_suffix"
        else
            c = set_cells[1]
            return "$(c.sheet_name)!$(c.cell) (+$(length(set_cells) - 1) cells)$sheet_suffix"
        end
    end

    function summarize_cells(cells::Set{CellDependency}; max_items::Int = 6)
        if isempty(cells)
            return "none"
        end

        sorted_cells = sort!(collect(cells))
        shown = sorted_cells[1:min(max_items, length(sorted_cells))]
        text = join(["$(c.sheet_name)!$(c.cell)" for c in shown], ", ")
        if length(sorted_cells) > max_items
            text *= ", ..."
        end

        text
    end

    cell_deps_sets = Vector{Union{Nothing, Set{CellDependency}}}(nothing, length(statements))
    set_cells_sets = Vector{Union{Nothing, Set{CellDependency}}}(nothing, length(statements))
    # cell_deps_sets = [Set(get_cell_deps(s)) for s in statements]
    # set_cells_sets = [Set(get_set_cells(s)) for s in statements]
    function get_cell_deps_sets(stmt_i::Int)
        cache = cell_deps_sets[stmt_i]
        if isnothing(cache)
            cache = Set(get_cell_deps(statements[stmt_i]))
            cell_deps_sets[stmt_i] = cache
            cache
        else
            cache
        end
    end

    function get_set_cells_sets(stmt_i::Int)
        cache = set_cells_sets[stmt_i]
        if isnothing(cache)
            cache = Set(get_set_cells(statements[stmt_i]))
            set_cells_sets[stmt_i] = cache
            cache
        else
            cache
        end
    end

    function get_edge_cells(src_i::Int, dst_i::Int)
        # src_deps = Set(get_cell_deps(statements[src_i]))
        # dst_sets = Set(get_set_cells(statements[dst_i]))
        src_deps = get_cell_deps_sets(src_i)
        dst_sets = get_set_cells_sets(dst_i)
        intersect(src_deps, dst_sets)
    end

    statements_by_sheet = group_to_dict(1:length(statements), i -> get_statement_sheets(statements[i]))
    grouped_entries = collect(statements_by_sheet)
    sort!(grouped_entries, by = kv -> join(first(kv), "|"))

    sheet_reports = Dict{String, Any}()

    for (sheets, sheet_statements) in grouped_entries
        # Only analyze statements that set cells on exactly one sheet.
        length(sheets) == 1 || continue
        sheet = sheets[1]

        if !isnothing(test_sheet) && sheet != test_sheet
            continue
        end

        for i in sheet_statements

        end
        # cell_deps_sets = [Set(get_cell_deps(s)) for s in statements]
        # set_cells_sets = [Set(get_set_cells(s)) for s in statements]

        # sheet_statements is a list of statement indices, which correspond to nodes in stmt_graph.
        sheet_node_set = Set(sheet_statements)

        # Inputs to this sheet that originate from statements on other sheets.
        external_inputs_by_sheet = Dict{String, Set{CellDependency}}()
        for stmt_i in sheet_statements
            for dep in get_cell_deps(statements[stmt_i])
                dep.sheet_name == sheet && continue
                push!(get!(()->Set{CellDependency}(), external_inputs_by_sheet, dep.sheet_name), dep)
            end
        end

        # Outputs from this sheet that are consumed by statements on other sheets.
        external_outputs_by_sheet = Dict{String, Set{CellDependency}}()
        for producer_i in sheet_statements
            # produced_cells = Set(get_set_cells(statements[producer_i]))
            produced_cells = get_set_cells_sets(producer_i)
            isempty(produced_cells) && continue

            for consumer_i in inneighbors(stmt_graph, producer_i)
                consumer_i in sheet_node_set && continue

                # used_cells = intersect(produced_cells, Set(get_cell_deps(statements[consumer_i])))
                used_cells = intersect(produced_cells, get_cell_deps_sets(consumer_i))
                isempty(used_cells) && continue

                consumer_sheets = get_statement_sheets(statements[consumer_i])
                isempty(consumer_sheets) && (consumer_sheets = ["<unknown>"])

                for consumer_sheet in consumer_sheets
                    union!(get!(()->Set{CellDependency}(), external_outputs_by_sheet, consumer_sheet), used_cells)
                end
            end
        end

        # Detect chains that start on this sheet, leave it, and eventually return to it.
        bridge_paths = Dict{Tuple{Int64, Int64}, Vector{Int64}}()
        for start_node in sheet_statements
            queue = Int64[]
            parents = Dict{Int64, Int64}()
            seen_external = Set{Int64}()

            for n in outneighbors(stmt_graph, start_node)
                n in sheet_node_set && continue
                push!(queue, n)
                push!(seen_external, n)
                parents[n] = start_node
            end

            q_i = 1
            while q_i <= length(queue)
                node = queue[q_i]
                q_i += 1

                for next_node in outneighbors(stmt_graph, node)
                    if next_node in sheet_node_set
                        if next_node != start_node
                            pair = (Int64(start_node), Int64(next_node))
                            if !haskey(bridge_paths, pair)
                                path = Int64[next_node]
                                cur = node
                                while true
                                    push!(path, cur)
                                    if cur == start_node
                                        break
                                    end
                                    cur = parents[cur]
                                end
                                reverse!(path)
                                bridge_paths[pair] = path
                            end
                        end
                        continue
                    end

                    if !(next_node in seen_external)
                        push!(seen_external, next_node)
                        parents[next_node] = node
                        push!(queue, next_node)
                    end
                end
            end
        end

        bridge_boundary_inputs_by_sheet = Dict{String, Set{CellDependency}}()
        bridge_boundary_outputs_by_sheet = Dict{String, Set{CellDependency}}()
        boundary_edge_counts = Dict{Tuple{Int64, Int64}, Int64}()

        for path in values(bridge_paths)
            length(path) >= 2 || continue

            first_edge = (path[1], path[2])
            last_edge = (path[end-1], path[end])

            boundary_edge_counts[first_edge] = get(boundary_edge_counts, first_edge, 0) + 1
            boundary_edge_counts[last_edge] = get(boundary_edge_counts, last_edge, 0) + 1

            for c in get_edge_cells(first_edge[1], first_edge[2])
                push!(get!(()->Set{CellDependency}(), bridge_boundary_inputs_by_sheet, c.sheet_name), c)
            end
            for c in get_edge_cells(last_edge[1], last_edge[2])
                push!(get!(()->Set{CellDependency}(), bridge_boundary_outputs_by_sheet, c.sheet_name), c)
            end
        end

        # num_external_inputs = sum(length(v) for v in values(external_inputs_by_sheet))
        # num_external_outputs = sum(length(v) for v in values(external_outputs_by_sheet))
        num_external_inputs = sum(length, values(external_inputs_by_sheet), init = 0)
        num_external_outputs = sum(length, values(external_outputs_by_sheet), init = 0)

        println("\n" * "-"^50)
        println("Sheet \"$sheet\" diagnostics")
        println("- statements: $(length(sheet_statements))")
        println("- external input cells: $num_external_inputs")
        println("- external output cells: $num_external_outputs")
        println("- cross-sheet pass-through chains: $(length(bridge_paths))")

        if !isempty(external_inputs_by_sheet)
            println("  Inputs by sheet:")
            for dep_sheet in sort(collect(keys(external_inputs_by_sheet)))
                cells = external_inputs_by_sheet[dep_sheet]
                println("    <- $dep_sheet: $(length(cells)) ($(summarize_cells(cells)))")
            end
        end

        if !isempty(external_outputs_by_sheet)
            println("  Outputs used by sheet:")
            for out_sheet in sort(collect(keys(external_outputs_by_sheet)))
                cells = external_outputs_by_sheet[out_sheet]
                println("    -> $out_sheet: $(length(cells)) ($(summarize_cells(cells)))")
            end
        end

        if !isempty(bridge_paths)
            println("  Pass-through boundary cells (where cross-sheet coupling happens):")
            if !isempty(bridge_boundary_inputs_by_sheet)
                println("    External -> this sheet edge cells:")
                for dep_sheet in sort(collect(keys(bridge_boundary_inputs_by_sheet)))
                    cells = bridge_boundary_inputs_by_sheet[dep_sheet]
                    println("      <- $dep_sheet: $(length(cells)) ($(summarize_cells(cells)))")
                end
            end
            if !isempty(bridge_boundary_outputs_by_sheet)
                println("    This sheet -> external edge cells:")
                for out_sheet in sort(collect(keys(bridge_boundary_outputs_by_sheet)))
                    cells = bridge_boundary_outputs_by_sheet[out_sheet]
                    println("      -> $out_sheet: $(length(cells)) ($(summarize_cells(cells)))")
                end
            end

            println("  Most common pass-through boundary edges:")
            sorted_edges = collect(boundary_edge_counts)
            sort!(sorted_edges, by = kv -> kv[2], rev = true)
            for (edge, count) in sorted_edges[1:min(6, length(sorted_edges))]
                src_i, dst_i = edge
                edge_cells = get_edge_cells(src_i, dst_i)
                src_sheets = summarize_sheet_list(get_statement_sheets(statements[src_i]))
                dst_sheets = summarize_sheet_list(get_statement_sheets(statements[dst_i]))

                println("    * $(get_statement_label(src_i)) -> $(get_statement_label(dst_i))")
                println("      chains: $count | src sheets: [$src_sheets] | dst sheets: [$dst_sheets]")
                println("      shared edge cells: $(summarize_cells(edge_cells))")
            end

            println("  Example pass-through chains:")
            bridge_pairs = collect(keys(bridge_paths))
            sort!(bridge_pairs, by = p -> (get_statement_label(p[1]), get_statement_label(p[2])))

            for pair in bridge_pairs[1:min(6, length(bridge_pairs))]
                path = bridge_paths[pair]
                path_labels = join(get_statement_label.(path), " -> ")
                external_sheets = Set{String}()

                if length(path) > 2
                    for node in path[2:(end-1)]
                        for node_sheet in get_statement_sheets(statements[node])
                            node_sheet == sheet && continue
                            push!(external_sheets, node_sheet)
                        end
                    end
                end

                first_edge_cells = length(path) >= 2 ? get_edge_cells(path[1], path[2]) : Set{CellDependency}()
                last_edge_cells = length(path) >= 2 ? get_edge_cells(path[end-1], path[end]) : Set{CellDependency}()

                external_str = isempty(external_sheets) ? "<unknown>" : join(sort(collect(external_sheets)), ", ")
                println("    * $(get_statement_label(pair[1])) reaches $(get_statement_label(pair[2])) via [$external_str]")
                println("      path: $path_labels")
                println("      boundary A edge cells: $(summarize_cells(first_edge_cells))")
                println("      boundary B edge cells: $(summarize_cells(last_edge_cells))")
            end
        end

        sheet_reports[sheet] = Dict(
            :sheet_statements => sheet_statements,
            :external_inputs_by_sheet => external_inputs_by_sheet,
            :external_outputs_by_sheet => external_outputs_by_sheet,
            :bridge_paths => bridge_paths,
            :bridge_boundary_inputs_by_sheet => bridge_boundary_inputs_by_sheet,
            :bridge_boundary_outputs_by_sheet => bridge_boundary_outputs_by_sheet,
            :boundary_edge_counts => boundary_edge_counts,
        )
    end

    sheet_reports
end

function export_statements(used_subset::XLConvert.ExcelWorkbook2, statements, tables)
    xf = used_subset.xf
    @time "get_all_referenced_cells" all_ref_cells = get_all_referenced_cells(used_subset)
    # var_names_map = make_var_names_map(all_ref_cells, wb.xf)

    println("Making var names map")
    @time var_names_map = make_var_names_map(all_ref_cells, used_subset)

    for (cell, name) in var_names_map
        if name == "yield"
            @show cell name
            var_names_map[cell] = "yield_val"
        end
    end

    used_cell_set = Set(get_all_referenced_cells(used_subset))
    used_names = Set{String}()
    vessels_sheet = xf["vessels"]
    columns = ("AA", "AB", "AC")
    column_names = ("loaded", "unloaded", "max_cruise")
    for row in 43:332
        cells = [CellDependency("vessels", "$col$row") for col in columns]
        cells_used = in.(cells, Ref(used_cell_set))

        # If none of the row cells are used, move on
        any(cells_used) || continue

        base_name = vessels_sheet["Y$row"]
        unit = vessels_sheet["Z$row"]

        # Skip is base_name is missing
        ismissing(base_name) && continue

        base_name = XLConvert.normalize_var_name(strip(base_name))

        name = if base_name in used_names
            if ismissing(unit)
                base_name * "_$row"
            else
                base_name * "_" * unit
            end
        else
            base_name
        end
        name = XLConvert.normalize_var_name(name)

        push!(used_names, name)

        if false && sum(cells_used) == 1
            println("Single name! $name")
            var_names_map[cells[cells_used][1]] = name
        else
            for i in eachindex(cells)
                if cells_used[i]
                    var_names_map[cells[i]] = name * "_" * column_names[i]
                end
            end
        end
    end
    # vessel_solving_tbls = DefTable(xf, "vessels", "vsl_dsn", "AA43", "AC332", "AA6:AC6", "Y43:Y332")
    # for key in keys(var_names_map)
    #     if key in vessel_solving_tbls
    #         r = rownum(key) - startrow(vessel_solving_tbls) + 1
    #         c = colnum(key) - startcol(vessel_solving_tbls) + 1
    #         r_name = XLConvert.row_name(vessel_solving_tbls, r)
    #         c_name =XLConvert.column_name(vessel_solving_tbls, c)
    #         var_names_map[key] = XLConvert.normalize_var_name("$(r_name)_$(c_name)")
    #     end
    # end



    println("Setting names from tables!")
    all_ref_cell_set = Set(all_ref_cells)
    @time for t in tables
        # set_names_from_table!(var_names_map, all_ref_cells, t)
        set_names_from_table!(var_names_map, all_ref_cell_set, t)
    end

    handlers = Vector{AbstractHandler}()
    handlers = [EdgeCaseHandler(), IndirectXlookupRangeHandler(), BasicOpHandler(), TableRefHandler(), EverythingElseHandler()]

    # key_values_dict = Dict((p[1] => FormulaParser.toexpr(repr(p[2]))) for p in XLSX.get_workbook(wb.xf).workbook_names)
    key_values_dict = copy(used_subset.key_values)
    for (k, value) in key_values_dict
        if value isa XLConvert.FlatExpr
            key_values_dict[k] = XLConvert.insert_table_refs(value, tables)
        end
    end

    println("Infer types")
    cell_types = Dict{CellDependency, Any}()
    for cell_dep in all_ref_cells
        # cell_types[cell] = Any
        cell_data = get_cell_value(used_subset, cell_dep)
        try
            ws = xf[string(cell_dep.sheet_name)]
            type = if (cell_data isa MissingCell)
                Missing
            else
                getdatatype(ws, cell_data.cell)
            end
            cell_types[cell_dep] = type
        catch
            cell_types[cell_dep] = Any
        end

        # if type === Any
        #     println("Cell $(XLConvert.to_string(cell_dep)) has Any type!")
        # end
    end
    # @time cell_types = infer_types(used_subset)

    new_names_map = var_names_map
    # exporter = JuliaExporter(used_subset, new_names_map, tables, key_values_dict, handlers, cell_types)
    exporter = PythonExporter(used_subset, new_names_map, tables, key_values_dict, handlers, cell_types)


    println("Write file")
    @time write_file(exporter, "tea/modular_tea.py", used_subset, statements)

    statements, exporter
end

function get_statements_and_tables(wb_in::XLConvert.ExcelWorkbook2)
    wb = wb_in
    xf = wb.xf

    used_subset = wb

    cell_graph = used_subset.cell_graph
    @show nv(cell_graph) ne(cell_graph)
    node_nums = 1:nv(cell_graph)
    by_degree = sort(node_nums, by = v -> length(outneighbors(cell_graph, v)), rev = true)
    for n in by_degree[begin:2]
        cell = get_cell(used_subset, n)
        println("Cell $cell has $(length(outneighbors(cell_graph, n))) outneighbors")
        # @show length(outneighbors(cell_graph, n))
        # @display get_expr(wb.cell_dict[cell])
    end
    check_cycles(used_subset)


    println("-"^40)
    println("Finding Tables")
    println("-"^40)
    tables = make_tables(used_subset)
    # statements = get_statements(used_subset, tables)
    statements = generate_statements(used_subset, tables)

    cell_to_statement = XLConvert.make_cell_to_statement_dict(statements)

    const_cells = [
        CellDependency("Anchor Sizing", "AE26"),
        CellDependency("oyster Husbandry model", "DM6"),
        CellDependency("Equip&Mat Assumptions", "AC3"),
        CellDependency("Site Inputs", "C68"),
        CellDependency("Equip&Mat Calcs", "K2"),
        CellDependency("Equip&Mat Calcs", "W2"),
        CellDependency("Equip&Mat Calcs", "B120"),
        CellDependency("Equip&Mat Calcs", "D119"),
        CellDependency("Equip&Mat Calcs", "B139"),
        CellDependency("Equip&Mat Calcs", "B157"),
        CellDependency("Equip&Mat Calcs", "B175"),
        CellDependency("Equip&Mat Calcs", "D138"),
        CellDependency("Equip&Mat Calcs", "D156"),
        CellDependency("Equip&Mat Calcs", "K119"),
        CellDependency("Equip&Mat Calcs", "K138"),
        CellDependency("Equip&Mat Calcs", "K156"),
        CellDependency("Equip&Mat Calcs", "K174"),
        CellDependency("Structure Calcs", "P4"),
        CellDependency("Structure Calcs", "B55"),
        CellDependency("Structure Calcs", "C54"),
        CellDependency("Structure Calcs", "B75"),
        CellDependency("Structure Calcs", "B95"),
        CellDependency("Structure Calcs", "C74"),
        CellDependency("Structure Calcs", "C94"),
        CellDependency("Structure Calcs", "CN4"),
        CellDependency("Structure Calcs", "B114"),
        CellDependency("Structure Calcs", "H113"),
        CellDependency("Structure Calcs", "H54"),
        CellDependency("Structure Calcs", "H74"),
        CellDependency("Structure Calcs", "H94"),
        CellDependency("Structure Assumptions", "AD2"),
        CellDependency("Structure Assumptions", "X76"),
        CellDependency("Machines", "B70"),
        CellDependency("Machines", "B86"),
        CellDependency("Machines", "B74"),
        CellDependency("Machines", "B90"),
        CellDependency("Labor", "D5"),
    ]
    for cell in const_cells
        stmt = get(cell_to_statement, cell, nothing)
        if isnothing(stmt)
            println("Const cell $cell wasn't actually set anywhere")
            continue
        end

        set_cells = get_set_cells(stmt)
        if length(set_cells) == 1 && set_cells[1] in const_cells
            lhs = set_cells[1]
            if stmt isa XLConvert.StandardStatement
                stmt.rhs_expr = xf[lhs.sheet_name][lhs.cell]
            elseif stmt isa XLConvert.TableStatement
                stmt.rhs_expr = xf[lhs.sheet_name][lhs.cell]
            end
        end
    end
    # for stmt in statements
    #     set_cells = get_set_cells(stmt)
    #     if length(set_cells) == 1 && set_cells[1] in const_cells
    #         lhs = set_cells[1]
    #         if stmt isa XLConvert.StandardStatement
    #             stmt.rhs_expr = xf[lhs.sheet_name][lhs.cell]
    #         elseif stmt isa XLConvert.TableStatement
    #             stmt.rhs_expr = xf[lhs.sheet_name][lhs.cell]
    #         end
    #     end
    # end

    override_values = [
        (CellDependency("oyster Husbandry model", "H9"), 0.0)
        (CellDependency("KAM Results", "J10"), 0.0)
        (CellDependency("KAM Results", "J11"), 0.0)
        (CellDependency("KAM Results", "J13"), 0.0)
        (CellDependency("KAM Results", "J14"), 0.0)
        (CellDependency("KAM Results", "J31"), 0.0)
        (CellDependency("KAM Results", "J32"), 0.0)
        (CellDependency("KAM Results", "J34"), 0.0)
        (CellDependency("KAM Results", "J35"), 0.0)
        (CellDependency("KAM Results", "J40"), 0.0)
        (CellDependency("KAM Results", "J41"), 0.0)
    ]
    for (cell, new_value) in override_values
        stmt = get(cell_to_statement, cell, nothing)
        if isnothing(stmt)
            println("overridden value at $cell wasn't associated with a statement")
            continue
        end

        if cell in get_set_cells(stmt)
            # @show stmt
            if stmt isa XLConvert.StandardStatement
                stmt.rhs_expr = new_value
            else
                throw("Don't know how to override statement $stmt")
            end
        end
    end

    missing_to_zero_regions = [
        WorkbookRegion("Machines", "D13", "AB13")
    ]

    for region in missing_to_zero_regions
        for cell in XLConvert.cells(region)
            stmt = get(cell_to_statement, cell, nothing)
            if isnothing(stmt)
                println("$cell wasn't associated with a statement, whend oing missing to zero")
                continue
            end
            if stmt isa XLConvert.TableStatement
                if ismissing(stmt.rhs_expr)
                    stmt.rhs_expr = 0
                end
                # @show stmt.rhs_expr
            elseif stmt isa XLConvert.BroadcastedStatement
                # @show stmt.func_expr
            end
            # @show stmt
        end
    end

    statements, tables
end

function run(wb_in::XLConvert.ExcelWorkbook2)
    wb = wb_in
    xf = wb.xf

    used_subset = wb
    statements, tables = get_statements_and_tables(wb_in)

    export_statements(used_subset, statements, tables)

    statements, tables
end

function debug_types_for_statement(wb, statements, cell_dep)
    target_output = CellDependency("Results", "C26")
    all_target_outputs = [target_output]

    @time used_subset = get_workbook_subset(wb, all_target_outputs)

    key_values_dict = copy(wb.key_values)
    # for (k, value) in key_values_dict
    #     if value isa XLConvert.FlatExpr
    #         key_values_dict[k] = XLConvert.insert_table_refs(value, tables)
    #     end
    # end

    println("Infer types")
    @time cell_types = infer_types(used_subset)
    node = findfirst(s -> cell_dep in XLConvert.get_set_cells(s), statements)
    @show node
    statement = statements[node]
    @show statement
    if statement isa XLConvert.StandardStatement || statement isa XLConvert.TableStatement
        expr = statement.rhs_expr
        XLConvert.print_type_debug(expr, cell_dep.sheet_name, cell_types, key_values_dict)
    elseif statement isa XLConvert.FunctionStatement
        sub_node = findfirst(s -> cell_dep in XLConvert.get_set_cells(s), statement.intermediates)
        @show sub_node
        statement = statement.intermediates[sub_node]
        if statement isa XLConvert.StandardStatement || statement isa XLConvert.TableStatement
            expr = statement.rhs_expr
            XLConvert.print_type_debug(expr, cell_dep.sheet_name, cell_types, key_values_dict)
        end
    end

end

function get_long_exprs(wb::XLConvert.ExcelWorkbook)
    formulas = [p for p in pairs(wb.cell_dict) if p[2] isa XLConvert.FormulaCell]

    exprs = [Pair(p[1], p[2].expr) for p in formulas if p[2].expr isa XLConvert.FlatExpr]

    function expr_length(p)
        e = p[2]
        length(e.parts)
    end

    sort!(exprs, by = expr_length, rev = true)

    for (cell, expr) in exprs[begin:10]

        println(cell)

        show(stdout, "text/plain", expr)
        println("\n")
    end


end

