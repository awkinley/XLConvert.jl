
function infer_types(used_subset::XLConvert.WorkbookSubset)
    graph = used_subset.graph
    cycles = Graphs.simplecycles(graph)
    if !isempty(cycles)
        println("Removing $(length(cycles)) cycles from the statement graph, this is almost certainly incorrect.")
        for cycle in cycles
            rem_edge!(graph, cycle[end], cycle[begin])
        end
    end
    # rev_cycles = Graphs.simplecycles(reverse(graph))
    # @show rev_cycles
    topo_levels = get_topo_levels_bottom_up(graph)
    max_level = maximum(values(topo_levels))
    grouped_by_level = Dict((l => [kv.first for kv in topo_levels if kv.second == l]) for l in 0:max_level)

    wb = used_subset.wb
    all_ref_cells = get_all_referenced_cells(wb)

    cell_types = Dict{CellDependency, Any}()

    input_cells = all_ref_cells[grouped_by_level[0]]
    for cell_dep in input_cells
        cell_data = get_cell_value(wb, cell_dep)
        ws = wb.xf[string(cell_dep.sheet_name)]
        type = if (cell_data isa MissingCell)
            Missing
        else
            getdatatype(ws, cell_data.cell)
        end
        if type === Any
            println("Cell $(XLConvert.to_string(cell_dep)) has Any type!")
        end
        cell_types[cell_dep] = type
    end


    # We do this because there's a cycle, and otherwise we'll get a type inference error
    # for cell_dep in deps_to_remove
    #     cell_data = get_cell_value(wb, cell_dep)
    #     ws = wb.xf[string(cell_dep.sheet_name)]
    #     type = if cell_data isa MissingCell
    #         Missing
    #     else
    #         getdatatype(ws, cell_data.cell)
    #     end
    #     if type == Any
    #         println("Cell $(to_string(cell_dep)) has Any type!")
    #     end
    #     cell_types[cell_dep] = type
    # end

    key_values_dict = wb.key_values
    begin
        for level in 1:max_level
            # @show level
            level_cells = all_ref_cells[grouped_by_level[level]]
            for cell_dep in level_cells
                cell_data = get_cell_value(wb, cell_dep)
                type = try
                    get_type(get_expr(cell_data), cell_dep.sheet_name, cell_types, key_values_dict)
                catch e
                    @show cell_dep
                    @show cell_data
                    @show get_expr(cell_data)
                    throw(e)
                end

                if type == Any
                    # @show cell_dep
                    # @show cell_data
                    # pprintln(cell_data.expr)
                    # println("Cell $(XLConvert.to_string(cell_dep)) has Any type!")
                    # failed = true
                    # break
                end
                cell_types[cell_dep] = type
            end
        end
    end

    cell_types

end

struct FastHashedFlatExpr
    expr::XLConvert.FlatExpr
end

function Base.hash(x::FastHashedFlatExpr, h::UInt)
    h = hash(:FlatExpr, h)
    for e in x.expr.parts
        h = hash(e.head, h)
    end

    h
end
Base.:(==)(a::FastHashedFlatExpr, b::FastHashedFlatExpr) = a.expr == b.expr
Base.isequal(a::FastHashedFlatExpr, b::FastHashedFlatExpr) = isequal(a.expr, b.expr)

function merge_tables(a::XLConvert.ExcelTable, b::XLConvert.ExcelTable)
    a.sheet_name != b.sheet_name && return nothing

    if XLConvert.startrow(a) == XLConvert.startrow(b) && XLConvert.endrow(a) == XLConvert.endrow(b)
        if XLConvert.endcol(a) + 1 == XLConvert.startcol(b)
            # Merge columns
            top_left = a.top_left
            bottom_right = b.bottom_right
            table_name = "$(top_left)_$(bottom_right)"
            col_names = XLSX.encode_column_number.(XLConvert.startcol(a):XLConvert.endcol(b))
            new_table = ExcelTable(a.sheet_name, table_name, top_left, bottom_right, "", "", col_names, missing)
            return new_table
        elseif XLConvert.endcol(b) + 1 == XLConvert.startcol(a)
            # Merge columns
            top_left = b.top_left
            bottom_right = a.bottom_right
            table_name = "$(top_left)_$(bottom_right)"
            col_names = XLSX.encode_column_number.(XLConvert.startcol(b):XLConvert.endcol(a))
            new_table = ExcelTable(a.sheet_name, table_name, top_left, bottom_right, "", "", col_names, missing)
            return new_table
        else
            return nothing
        end
    elseif XLConvert.startcol(a) == XLConvert.startcol(b) && XLConvert.endcol(a) == XLConvert.endcol(b)
        if XLConvert.endrow(a) + 1 == XLConvert.startrow(b)
            # Merge rows
            top_left = a.top_left
            bottom_right = b.bottom_right
            table_name = "$(top_left)_$(bottom_right)"
            col_names = XLSX.encode_column_number.(XLConvert.startcol(a):XLConvert.endcol(b))
            new_table = ExcelTable(a.sheet_name, table_name, top_left, bottom_right, "", "", col_names, missing)
            return new_table
        elseif XLConvert.endrow(b) + 1 == XLConvert.startrow(a)
            # Merge rows
            top_left = b.top_left
            bottom_right = a.bottom_right
            table_name = "$(top_left)_$(bottom_right)"
            col_names = XLSX.encode_column_number.(XLConvert.startcol(b):XLConvert.endcol(a))
            new_table = ExcelTable(a.sheet_name, table_name, top_left, bottom_right, "", "", col_names, missing)
            return new_table
        else
            return nothing
        end
    end

    nothing
end

function find_tables_in_sheet(sheet_name, cells)

    formula_cells = filter(c -> c.expr isa XLConvert.FlatExpr, cells)
    # formula_cells = cells
    # @show length(formula_cells)

    functionalized = [XLConvert.FormulaCell(cell.cell, XLConvert.functionalize(cell.expr)[1]) for cell in formula_cells]
    @show typeof(functionalized)

    # uses_same_function = XLConvert.group_to_dict(functionalized, c -> string(c.expr))
    uses_same_function = XLConvert.group_to_dict(functionalized, c -> FastHashedFlatExpr(c.expr))

    tables = Vector{XLConvert.ExcelTable}()

    # @show length(keys(uses_same_function))
    for func in keys(uses_same_function)
        # @show func
        formula_cells = uses_same_function[func]
        cells = [CellDependency(sheet_name, string(f.cell.ref)) for f in formula_cells]
        # row_nums = rownum.(cells)
        # col_nums = colnum.(cells)
        coords = XLConvert.get_coords.(cells)
        # coords = zip(col_nums, row_nums) |> collect
        coord_to_statement = Dict(c => s for (c, s) in zip(coords, cells))
        sort!(coords)


        regions = XLConvert.get_2d_regions(coords)
        # regions_old = XLConvert.get_2d_regions_old(coords)
        # if regions != regions_old
        #     @show coords
        #     @show regions
        #     @show regions_old
        # end
        # println("Group size = $(length(group))")
        for region in regions
            cols, rows = region
            region_area = (length(cols) * length(rows))
            if region_area < 2
                continue
            end
            # if length(cols) > 1
            #     println("\t$region")
            #     println("\tSize: $(region_area)")
            # end

            top_left = coord_to_cell_name(rows[1], cols[1])
            bottom_right = coord_to_cell_name(rows[end], cols[end])

            table_name = "$(top_left)_$(bottom_right)"
            col_names = XLSX.encode_column_number.(cols)
            table = ExcelTable(sheet_name, table_name, top_left, bottom_right, "", "", col_names, missing)
            push!(tables, table)

        end

        # @show length(uses_same_function[func])
    end


    println("Made $(length(tables)) tables on sheet $sheet_name")
    sort!(tables, by = t -> (XLConvert.startrow(t), XLConvert.startcol(t)))

    # new_tables = Vector{XLConvert.ExcelTable}()
    not_done = true
    total_merges = 0
    while not_done
        new_tables = Vector{XLConvert.ExcelTable}()

        merge_count = 0
        for table in tables
            did_merge = false
            # @show table.table_name
            for i in eachindex(new_tables)
                b = new_tables[i]
                # @show b.table_name
                new_table = merge_tables(table, b)
                # @show isnothing(new_table)

                if !isnothing(new_table)
                    new_tables[i] = new_table
                    merge_count += 1
                    did_merge = true
                    break
                end
            end

            if !did_merge
                push!(new_tables, table)
            end
        end

        total_merges += merge_count
        not_done = merge_count != 0
        tables = new_tables
    end


    println("Did $total_merges table merges to end up with $(length(tables)) final tables")

    tables
end

function get_ranges(expr::FlatExpr)
    ranges = Vector{Tuple{String, String, String}}()
    # handled = Set{Int}()
    for (i, part) in enumerate(expr.parts)
        # i in handled && continue

        @match part begin
            # ExcelExpr(:cell_ref, [cell, sheet]) => push!(deps, CellDependency(sheet, cell))
            # ExcelExpr(:sheet_ref, (sheet_name, ref)) => get_expr_dependencies(ref, key_values)
            # ExcelExpr(:named_range, [name]) => append!(deps, get_expr_dependencies(key_values[name], key_values))
            ExcelExpr(:range, [FlatIdx(lhs_i), FlatIdx(rhs_i)]) => begin
                lhs_expr = expr.parts[lhs_i]
                rhs_expr = expr.parts[rhs_i]
                if !((lhs_expr.head == :cell_ref) && (rhs_expr.head == :cell_ref))
                    # throw("Don't know how to get dependencies for $(expr)")
                    continue
                end
                sheet = lhs_expr.args[2]
                if (sheet != rhs_expr.args[2])
                    throw("Don't know how to get dependencies for $(expr)")
                end

                lhs = lhs_expr.args[1]
                rhs = rhs_expr.args[1]

                push!(ranges, (lhs_expr.args[2], lhs, rhs))
            end
            _ => continue
        end
    end

    ranges
end

function coord_to_cell_name(row, col)
    string(XLSX.encode_column_number(col), row)
end

function find_untabled_ranges(used_subset::XLConvert.WorkbookSubset, tables::Vector{XLConvert.ExcelTable})
    used_cells = XLConvert.get_used_cells(used_subset)
    wb = used_subset.wb

    function handle_expr(expr::XLConvert.FlatExpr)
        ranges = get_ranges(expr)
        for (sheet, lhs, rhs) in ranges
            start_col, start_row = XLConvert.parse_cell(lhs)
            end_col, end_row = XLConvert.parse_cell(rhs)

            range_cells = [CellDependency(sheet, XLConvert.index_to_cellname(col, r)) for r in start_row:end_row, col in start_col:end_col]
            cell_table_idx = zeros(Int, size(range_cells))
            for (table_idx, table) in enumerate(tables)
                @. cell_table_idx[range_cells ∈ (table,)] .= table_idx
            end
            unique_tables = unique(cell_table_idx)
            if length(unique_tables) != 1 || cell_table_idx[1] == 0
                # println("Found untabled range! $(lhs):$(rhs)")
                # println("Cell Table Idx:")
                # show(stdout, "text/plain", cell_table_idx)
                # println("")
            end
            # Range overlaps with a single table
            if length(unique_tables) == 2 && 0 in unique_tables
                idx_to_grow = first(setdiff(unique_tables, Set([0])))
                table_to_grow = tables[idx_to_grow]

                left = min(start_col, startcol(table_to_grow))
                right = max(end_col, endcol(table_to_grow))
                top = min(start_row, startrow(table_to_grow))
                bottom = max(end_row, endrow(table_to_grow))

                top_left = coord_to_cell_name(top, left)
                bottom_right = coord_to_cell_name(bottom, right)

                table_name = "$(top_left)_$(bottom_right)"
                col_names = XLSX.encode_column_number.(left:right)
                old_name = getname(table_to_grow)
                tables[idx_to_grow] = ExcelTable(sheet, table_name, top_left, bottom_right, "", "", col_names, missing)
                println("Growing $(old_name) to $(getname(tables[idx_to_grow]))")
            end

            # Range overlaps with no tables
            if all(unique_tables .== 0) && length(range_cells) > 4
                top_left = coord_to_cell_name(start_row, start_col)
                bottom_right = coord_to_cell_name(end_row, end_col)

                table_name = "$(top_left)_$(bottom_right)"
                col_names = XLSX.encode_column_number.(start_col:end_col)
                push!(tables, ExcelTable(sheet, table_name, top_left, bottom_right, "", "", col_names, missing))
                println("Creating new table $(table_name)")

            end
        end
    end

    for cell in used_cells
        if !(cell in keys(wb.cell_dict))
            continue
        end

        expr = XLConvert.get_expr(wb.cell_dict[cell])
        if !(expr isa XLConvert.FlatExpr)
            continue
        end

        handle_expr(expr)

    end

    for expr in values(wb.key_values)
        if !(expr isa XLConvert.FlatExpr)
            continue
        end

        handle_expr(expr)
    end

end


function find_tables(used_subset::XLConvert.WorkbookSubset)
    all_tables = Vector{XLConvert.ExcelTable}()

    find_tables!(all_tables, used_subset)
end

function find_tables!(starting_tables::Vector{ExcelTable}, used_subset::XLConvert.WorkbookSubset)
    wb = used_subset.wb
    cells_by_sheet = XLConvert.group_to_dict(filter(c -> wb.cell_dict[c] isa XLConvert.FormulaCell, keys(wb.cell_dict)), c -> c.sheet_name)

    cells_by_sheet = XLConvert.group_to_dict(filter(c -> wb.cell_dict[c] isa XLConvert.FormulaCell, keys(wb.cell_dict)), c -> c.sheet_name)
    for sheet_name in keys(cells_by_sheet)
        sheet_tables = find_tables_in_sheet(sheet_name, [wb.cell_dict[c] for c in cells_by_sheet[sheet_name]])
        append!(starting_tables, sheet_tables)
    end
    find_untabled_ranges(used_subset, starting_tables)

    starting_tables
end

function get_statements(used_subset::WorkbookSubset, tables::Vector{ExcelTable})
    @time "Make statements" statements = make_statements(used_subset)
    @time "if_multiple_transform" if_multiple_transform!(statements)
    @time "if_toggle_transform" if_toggle_transform!(statements)
    @time "round_if_transform" round_if_transform!(statements)
    # @time "is_blank_transform" is_blank_transform!(statements)
    @time "table_ref_transform" table_ref_transform!(statements, tables)
    table_stmts = filter(s -> s isa XLConvert.TableStatement, statements)
    @show length(table_stmts)

    println("-"^40)
    println("Table Broadcast Transform")
    println("-"^40)
    @time statements = table_broadcast_transform_2d!(statements)
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
    # statements_with_funcs = add_functions(statements, min_intermediates=2)

    statements_with_funcs
end

function export_julia(wb::ExcelWorkbook2, output_targets::Vector{CellDependency}, file_name::String, tables::Vector{ExcelTable})

    @time used_subset = get_workbook_subset(wb, output_targets)

    println("-"^40)
    println("Finding Tables")
    println("-"^40)
    @time find_tables!(tables, used_subset)

    cycles = Graphs.simplecycles(used_subset.graph)
    all_ref_cells = get_all_referenced_cells(wb)
    println("-"^40)
    println("Cycles")
    println("-"^40)
    @time "showing cycles" for cycle in cycles
        @show all_ref_cells[cycle]
    end

    statements = get_statements(used_subset, tables)

    @time "get_all_referenced_cells" all_ref_cells = get_all_referenced_cells(wb)

    println("Making var names map")
    @time var_names_map = make_var_names_map(all_ref_cells, wb)
    println("Setting names from tables!")
    @time for t in tables
        set_names_from_table!(var_names_map, all_ref_cells, t)
    end

    handlers = Vector{AbstractHandler}()
    handlers = [BasicOpHandler(), TableRefHandler(), EverythingElseHandler()]

    key_values_dict = copy(wb.key_values)
    for (k, value) in key_values_dict
        if value isa XLConvert.FlatExpr
            key_values_dict[k] = XLConvert.insert_table_refs(value, tables)
        end
    end

    println("Infer types")
    @time cell_types = infer_types(used_subset)

    exporter = JuliaExporter(wb, var_names_map, tables, key_values_dict, handlers, cell_types)


    println("Write file")
    @time write_file(exporter, file_name, wb, statements)

end