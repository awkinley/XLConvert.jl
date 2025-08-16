if false
    include("./src/XLConvert.jl")
end
using XLConvert
using XLSX
using Graphs

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

function coord_to_cell_name(row, col)
    string(XLSX.encode_column_number(col), row)
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

function find_tables_in_sheet(sheet_name, cells)

    # formula_cells = filter(c -> c isa XLConvert.FormulaCell, cells)
    formula_cells = cells
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

function find_tables(used_subset::XLConvert.WorkbookSubset)
    wb = used_subset.wb
    cells_by_sheet = XLConvert.group_to_dict(filter(c -> wb.cell_dict[c] isa XLConvert.FormulaCell, keys(wb.cell_dict)), c -> c.sheet_name)

    all_tables = Vector{XLConvert.ExcelTable}()

    for sheet_name in keys(cells_by_sheet)
        # sheet_tables = find_tables_in_sheet(sheet_name, [wb.cell_dict[c] for c in cells_by_sheet[sheet_name] if wb.cell_dict[c] isa XLConvert.FormulaCell])
        sheet_tables = find_tables_in_sheet(sheet_name, [wb.cell_dict[c] for c in cells_by_sheet[sheet_name]])
        append!(all_tables, sheet_tables)
    end

    # sheet_name = "Energy Feed & Utility Prices"
    # find_tables_in_sheet(sheet_name, [wb.cell_dict[c] for c in cells_by_sheet[sheet_name]])

    # for sheet_name in keys(cells_by_sheet)
    # end
    all_tables
end

function read_wb()
    file = "current-central-biomass-gasification-version-oct20.xlsm"
    wb = parse_workbook(file)
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

function get_statements(wb::XLConvert.ExcelWorkbook2)
    xf = wb.xf

    # return

    target_output = CellDependency("Results", "C26")
    all_target_outputs = [target_output]

    @time "get_workbook_subset" used_subset = get_workbook_subset(wb, all_target_outputs)

    println("-"^40)
    println("Finding Tables")
    println("-"^40)
    @time "find_tables" tables = find_tables(used_subset)
    push!(tables, DefTable(xf, "Depreciation", "MACRS", "C15", "H35", "C14:H14", ""))

    # return

    # cycles = Graphs.simplecycles(used_subset.graph)
    # all_ref_cells = get_all_referenced_cells(wb)
    # println("-"^40)
    # println("Cycles")
    # println("-"^40)
    # for cycle in cycles
    #     @show all_ref_cells[cycle]
    # end

    # # tables = begin
    # #     values_tab = DefTable(xf, "Sheet1", "value_series", "C6", "C17", "C5:C5", "")

    # #     []
    # # end

    @time "make_statements" statements = make_statements(used_subset)
    if_multiple_transform!(statements)
    if_toggle_transform!(statements)
    round_if_transform!(statements)
    table_ref_transform!(statements, tables)

    println("-"^40)
    println("Table Broadcast Transform")
    println("-"^40)
    @time statements = table_broadcast_transform_2d!(statements)

    statements
end

function find_tables(wb)
    target_output = CellDependency("Results", "C26")
    all_target_outputs = [target_output]

    @time used_subset = get_workbook_subset(wb, all_target_outputs)

    println("-"^40)
    println("Finding Tables")
    println("-"^40)
    @time tables = find_tables(used_subset)
    push!(tables, DefTable(wb.xf, "Depreciation", "MACRS", "C15", "H35", "C14:H14", ""))

end

function run(wb::XLConvert.ExcelWorkbook2)
    # file = "current-central-biomass-gasification-version-oct20.xlsm"
    # wb = parse_workbook(file)
    xf = wb.xf

    # return

    target_output = CellDependency("Results", "C26")
    all_target_outputs = [target_output]

    @time used_subset = get_workbook_subset(wb, all_target_outputs)

    println("-"^40)
    println("Finding Tables")
    println("-"^40)
    @time tables = find_tables(used_subset)
    push!(tables, DefTable(xf, "Depreciation", "MACRS", "C15", "H35", "C14:H14", ""))

    # return

    cycles = Graphs.simplecycles(used_subset.graph)
    all_ref_cells = get_all_referenced_cells(wb)
    println("-"^40)
    println("Cycles")
    println("-"^40)
    @time "showing cycles" for cycle in cycles
        @show all_ref_cells[cycle]
    end

    # tables = begin
    #     values_tab = DefTable(xf, "Sheet1", "value_series", "C6", "C17", "C5:C5", "")

    #     []
    # end

    statements = begin
        @time "Make statements" statements = make_statements(used_subset)
        @time "if_multiple_transform" if_multiple_transform!(statements)
        @time "if_toggle_transform" if_toggle_transform!(statements)
        @time "round_if_transform" round_if_transform!(statements)
        @time "table_ref_transform" table_ref_transform!(statements, tables)
        table_stmts = filter(s -> s isa XLConvert.TableStatement, statements)
        @show length(table_stmts)

        # statements

        # statements = table_broadcast_transform_2d!(statements, used_subset)

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
        # println("-"^40)
        # println("Group statements (again x2)")
        # println("-"^40)
        # grouped_statements = group_statements(grouped_statements)
        # println("-"^40)
        # println("Add functions")
        # println("-"^40)
        # statements_with_funcs = add_functions(grouped_statements, min_intermediates=2)
        # statements_with_funcs = add_functions(statements, min_intermediates=2)

        # statements_with_funcs
        grouped_statements
        # statements
    end

    # Graph has edges parent -> child
    # so an arrow represents an dependency (not a usage)
    # stmt_graph = make_statement_graph(statements)

    # function is_of_interest(c)
    #     c.sheet_name == "Debt Financing Calculations" && colnum(c) == 2 && rownum(c) == 56
    # end
    # function get_set_cell_coords(i)
    #     c = XLConvert.get_set_cells(statements[i])[1]
    #     (rownum(c), colnum(c))
    # end

    # statements_of_interest = [i for i in eachindex(statements) if isa(statements[i], XLConvert.TableStatement) && is_of_interest(XLConvert.get_set_cells(statements[i])[1])]


    # stmt_graph = make_statement_graph(statements)

    # stmt_topo_levels = get_topo_levels_bottom_up(stmt_graph)
    # # stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    # for stmt_i in statements_of_interest
    #     XLConvert.debug_group_statements(statements, stmt_graph, stmt_topo_levels, stmt_i)
    # end

    # return

    # topo_sorted = topological_sort(reverse(stmt_graph))
    # # @show statements[topo_sorted[end]]
    # stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    # stmt_to_level = Dict((stmt => stmt_topo_levels[i]) for (i, stmt) in enumerate(statements))
    # # table_statements::Vector{TableStatement} = filter(s -> isa(s, TableStatement), statements)
    # table_statement_idxs::Vector{Int64} = filter(i -> isa(statements[i], XLConvert.TableStatement), eachindex(statements))

    # function is_of_interest(c)
    #     c.sheet_name == "Debt Financing Calculations" && colnum(c) == 2
    # end
    # function get_set_cell_coords(i)
    #     c = XLConvert.get_set_cells(statements[i])[1]
    #     (rownum(c), colnum(c))
    # end

    # statements_of_interest = [i for i in table_statement_idxs if is_of_interest(XLConvert.get_set_cells(statements[i])[1])]
    # sort!(statements_of_interest, by=i -> stmt_topo_levels[i])
    # rev_g = reverse(stmt_graph)
    # for i in eachindex(statements_of_interest)
    #     for stmt_j in statements_of_interest[i + 1:end]
    #         num_paths = number_of_paths(rev_g, statements_of_interest[i], stmt_j)
    #         if num_paths > 0
    #             stmt_a = statements[statements_of_interest[i]]
    #             stmt_b = statements[stmt_j]
    #             @show stmt_a stmt_b
    #             @show num_paths
    #         end
    #     end
    # end
    # topo_sorted_idx = [findfirst(isequal(i), topo_sorted) for i in statements_of_interest]
    # min_idx, max_idx = extrema(topo_sorted_idx)
    # @show min_idx max_idx

    # has_interdep = false
    # # for idx in (min_idx + 1):(max_idx - 1)
    # for idx in min_idx:max_idx
    #     # idx in topo_sorted_idx && continue

    #     node = topo_sorted[idx]
    #     parents = inneighbors(stmt_graph, node)
    #     for p in parents
    #         if p in statements_of_interest
    #             @show idx node p
    #             @show statements[node]
    #             @show statements[p]
    #             has_interdep = true
    #             break
    #         end
    #     end

    #     if has_interdep
    #         break
    #     end
    # end

    # @show has_interdep

    # for stmt_idx in sort(statements_of_interest, by=get_set_cell_coords)
    # for stmt_idx in sort(statements_of_interest)
    #     @show statements[stmt_idx]
    #     level = stmt_topo_levels[stmt_idx]
    #     println("Level = $level") 
    # end


    # return


    @time "get_all_referenced_cells" all_ref_cells = get_all_referenced_cells(wb)
    # var_names_map = make_var_names_map(all_ref_cells, wb.xf)

    println("Making var names map")
    @time var_names_map = make_var_names_map(all_ref_cells, wb)
    println("Setting names from tables!")
    @time for t in tables
        set_names_from_table!(var_names_map, all_ref_cells, t)
    end

    handlers = Vector{AbstractHandler}()
    handlers = [BasicOpHandler(), TableRefHandler(), EverythingElseHandler()]

    # key_values_dict = Dict((p[1] => FormulaParser.toexpr(repr(p[2]))) for p in XLSX.get_workbook(wb.xf).workbook_names)
    key_values_dict = wb.key_values

    println("Infer types")
    @time cell_types = infer_types(used_subset)

    new_names_map = var_names_map
    exporter = JuliaExporter(wb, new_names_map, tables, key_values_dict, handlers, cell_types)


    println("Write file")
    @time write_file(exporter, "current_biomass_gasification.jl", wb, statements)
end
