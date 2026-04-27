
function get_all_cells(sheet)
    cells = Vector{XLSX.Cell}()
    for row in XLSX.eachrow(sheet)
        append!(cells, values(row.rowcells))
    end
    # XLSX.getcellrange(sheet, XLSX.get_dimension(sheet))
    cells
end


function build_referenced_formula_dict(formulas)
    out = Dict([f.formula.id => f for f in filter(f -> f.formula isa XLSX.ReferencedFormula, formulas)])
    out
end

can_parse_formula(cell::XLSX.Cell) = can_parse_formula(cell.formula)
can_parse_formula(formula::XLSX.AbstractFormula) = false
can_parse_formula(formula::XLSX.FormulaReference) = true # Don't yet actualy know how to do the offsettings
function can_parse_formula(formula::XLSX.ReferencedFormula)
    str = formula.formula
    try
        toexpr(parse_formula(str)) !== nothing
        # tokens = RBNF.runlexer(ExcelFormula, str)
        # ast, ctx = RBNF.runparser(Start, tokens)
        # ast !== nothing
    catch _
        false
    end
end
function can_parse_formula(formula::XLSX.Formula)
    str = formula.formula
    try
        toexpr(parse_formula(str)) !== nothing
        # tokens = RBNF.runlexer(ExcelFormula, str)
        # ast, ctx = RBNF.runparser(Start, tokens)
        # ast !== nothing
    catch _
        false
    end
end

get_all_formulas(excel_file, sheet_name) = filter(has_formula, filter(!isempty, get_all_cells(excel_file[sheet_name])))

function get_cell_value(ctx::XLSX.Worksheet, cell)
    v = ctx[string(replace(cell, '$' => ""))]
    v

end

function execs_correctly(sheet, formula_cell, xl_ctx)
    try
        true_value = sheet[formula_cell.ref]
        calc_value = eval(toexpr(parse_formula(formula_cell.formula.formula)), xl_ctx)
        xl_compare(true_value, calc_value)
    catch _
        false
    end
end

function test_sheet_eval(xl, sheet_idx)


    sheets = XLSX.sheetnames(xl)
    sheet_name = sheets[sheet_idx]
    xl_ctx = ExcelContext(sheet_name,
        Dict([(name => xl[name]) for name in XLSX.sheetnames(xl)]),
        Dict((p[1] => toexpr(parse_formula(repr(p[2])))) for p in XLSX.get_workbook(xl).workbook_names),
    )

    println("Sheet: $(sheet_name)")
    sheet = xl[sheet_name]
    sheet_cells = get_all_cells(sheet)
    # filter(!isempty, sheet_cells)
    sheet_formulas = filter(has_formula, sheet_cells)
    # length(sheet_formulas)
    # sum(can_parse_formula.(sheet_formulas))
    # sheet_formulas[findall((f -> !can_parse_formula(f)).(sheet_formulas))]
    @assert all(can_parse_formula.(sheet_formulas)) "Couldn't parse every sheet formula"

    parsable = filter(can_parse_formula, sheet_formulas)
    reference_formula_dict = build_referenced_formula_dict(sheet_formulas)

    total_formulas = length(parsable)
    not_parse = []
    can_parse = []
    formula_refs = []
    errors = []
    random = []

    for (i, cell) in enumerate(parsable)
        # try
        if cell.datatype == "e"
            push!(errors, cell)
        elseif cell.formula isa XLSX.FormulaReference
            formula = cell.formula
            base_formula_cell = reference_formula_dict[formula.id]

            if occursin("RAND()", base_formula_cell.formula.formula)
                push!(random, cell)
                continue
            end
            start_row = base_formula_cell.ref.row_number
            start_col = base_formula_cell.ref.column_number
            end_row = cell.ref.row_number
            end_col = cell.ref.column_number
            delta_x = end_col - start_col
            delta_y = end_row - start_row

            try
                true_value = sheet[cell.ref]
                expression = offset(toexpr(parse_formula(base_formula_cell.formula.formula)), delta_y, delta_x)
                calc_value = eval(expression, xl_ctx)
                if xl_compare(true_value, calc_value)
                    push!(can_parse, cell)
                else
                    push!(not_parse, cell)
                end
            catch _
                push!(not_parse, cell)
            end
        elseif occursin("RAND()", cell.formula.formula)
            push!(random, cell)
        elseif execs_correctly(sheet, cell, xl_ctx)
            push!(can_parse, cell)
        else
            push!(not_parse, cell)
        end
        # catch _
        #     @warn "test failed?" i cell
        # end
    end

    num_correct = length(can_parse)
    num_not_correct = length(not_parse)
    num_formula_references = length(formula_refs)
    num_error_output = length(errors)
    num_rand = length(random)

    println("Total formulas = $(total_formulas)")
    println("Num correct = $(num_correct)")
    println("Num not correct = $(num_not_correct)")
    println("Num formula references = $(num_formula_references)")
    println("Num error output = $(num_error_output)")
    println("Num rand = $(num_rand)")

    sheet, xl_ctx, not_parse
end


function get_sheet_exprs(xl, sheet_name)
    sheet = xl[sheet_name]
    sheet_cells = get_all_cells(sheet)
    sheet_formulas = filter(has_formula, sheet_cells)

    reference_formula_dict = build_referenced_formula_dict(sheet_formulas)

    cell_exprs = Dict()
    errors = []

    for (i, cell) in enumerate(sheet_formulas)
        if cell.datatype == "e"
            push!(errors, cell)

        end
        if cell.formula isa XLSX.FormulaReference
            formula = cell.formula
            base_formula_cell = reference_formula_dict[formula.id]

            start_row = base_formula_cell.ref.row_number
            start_col = base_formula_cell.ref.column_number
            end_row = cell.ref.row_number
            end_col = cell.ref.column_number
            delta_x = end_col - start_col
            delta_y = end_row - start_row

            expression = offset(toexpr(parse_formula(base_formula_cell.formula.formula)), delta_y, delta_x)
            cell_exprs[cell.ref.name] = expression
        else

            try
                expression = toexpr(parse_formula(cell.formula.formula))
                cell_exprs[cell.ref.name] = expression
            catch e
                @show sheet_name
                @show cell
                @show cell.formula.formula
            end
        end
    end
    cell_exprs
end


get_all_exprs(xl) = Dict((sheet_name => get_sheet_exprs(xl, sheet_name)) for sheet_name in XLSX.sheetnames(xl))

function make_ctx(current_sheet, xl)
    ExcelContext(current_sheet,
        Dict([(name => xl[name]) for name in XLSX.sheetnames(xl)]),
        Dict((p[1] => toexpr(parse_formula(repr(p[2])))) for p in XLSX.get_workbook(xl).workbook_names),
    )
end

function named_range_to_cell(workbook::ExcelWorkbook, name::AbstractString)
    named_range_dict = XLSX.get_workbook(workbook.xf).workbook_names
    range = named_range_dict[name]
    sheet_part, cell_part = split(string(range), "!")
    sheet = strip(sheet_part, '\'')
    CellDependency(sheet, cell_part)
end



function get_topo_levels_bottom_up(wb::WorkbookSubset)
    topo_sorted = topological_sort(reverse(wb.graph))

    topo_levels = Dict{Int64, Int64}()
    for node in filter(v -> v ∈ wb.used_nodes, topo_sorted)
        dependencies = outneighbors(wb.graph, node)
        topo_levels[node] = maximum(k -> topo_levels[k] + 1, dependencies; init = 0)
    end

    topo_levels
end



function get_topo_levels_top_down(wb::WorkbookSubset)
    topo_sorted = topological_sort(wb.graph)

    topo_levels = Dict{Int64, Int64}()
    for node in filter(v -> v ∈ wb.used_nodes, topo_sorted)
        # We have to filter in this case, and not in the bottom up case
        # because it's not possible for a cell to depend on a value not in used_nodes
        # but it is possible for a cell not in used_nodes to depend on one that is
        dependents = filter(in(wb.used_nodes), inneighbors(wb.graph, node))
        topo_levels[node] = minimum(k -> topo_levels[k] - 1, dependents; init = 0)
    end

    # Will be a negative number
    min_level = minimum(values(topo_levels))

    # Change the range of levels from -n:0 to 0:n
    for k in keys(topo_levels)
        topo_levels[k] -= min_level
    end


    topo_levels
end

function getcell(xf, cell::CellDependency)
    sheet = xf[string(cell.sheet_name)]
    XLSX.getcell(sheet, cell.cell)
end


function to_rhs(exporter::JuliaExporter, wb::ExcelWorkbook, cell_ref::CellDependency)
    if !(cell_ref in keys(wb.cell_dict))
        return "missing"
    end

    cell_value = wb.cell_dict[cell_ref]

    expr = get_expr(cell_value)
    expr = insert_table_refs(expr, tables) |> convert_if_multiple
    convert(exporter, expr, cell_ref.sheet_name)
end


function set_names_from_table!(name_map, cell_dependencies, table::ExcelTable)
    r_start = startrow(table)
    c_start = startcol(table)
    r_end = endrow(table)
    c_end = endcol(table)
    for c in c_start:c_end, r in r_start:r_end
        cell = CellDependency(table.sheet_name, c, r)
        if cell in cell_dependencies
            row_idx = r - startrow(table) + 1
            col_name = column_name(table, c - startcol(table) + 1)
            r_name = row_name(table, r - startrow(table) + 1)

            # var_name = "$(getname(table))[$(repr(row_idx)), $(repr(col_name))]"
            var_name = "$(getname(table)).loc[$(repr(r_name)), $(repr(col_name))]"
            name_map[cell] = var_name

        end
    end
end

# contexts = Dict((sheet_name => make_ctx(sheet_name, xf)) for sheet_name in XLSX.sheetnames(xf))



function DefTable(xf::XLSX.XLSXFile, sheet_name, table_name, top_left, bottom_right, column_names_range, row_names_range)
    column_names = missing

    startcol, startrow = parse_cell(top_left)
    endcol = parse_cell(bottom_right)[1]
    if !isempty(column_names_range)
        column_names = string.(xf[sheet_name][column_names_range])
        for i in eachindex(column_names)
            if ismissing(column_names[i])
                column_names[i] = "missing_$(i)"
            elseif column_names[i] in column_names[begin:(i-1)]
                column_names[i] *= "_" * XLSX.encode_column_number(startcol + i - 1)
            end
        end
    else
        startcol = parse_cell(top_left)[1]
        endcol = parse_cell(bottom_right)[1]
        num_cols = 1 + endcol - startcol
        column_names = 1:num_cols
    end

    row_names = missing
    if row_names_range != ""
        row_names = xf[sheet_name][row_names_range]
        for i in eachindex(row_names)
            if ismissing(row_names[i])
                row_names[i] = "missing_$(i)"
            elseif row_names[i] in row_names[begin:(i - 1)]
                row_names[i] *= string("_", startrow + i - 1)
            end
        end

    end
    ExcelTable(
        sheet_name,
        table_name,
        top_left,
        bottom_right,
        column_names_range,
        row_names_range,
        replace.(string.(column_names), '$' => ""),
        ismissing(row_names) ? missing : replace.(string.(row_names), '$' => ""),
    )
end

function group_to_dict(values, get_key)
    res = Dict()
    for value in values
        k = get_key(value)
        push!(get!(Vector, res, k), value)
        # if haskey(res, k)
        #     push!(res[k], value)
        # else
        #     res[k] = [value]
        # end
    end
    res
end

function group_to_dict(values::AbstractArray{T}, get_key) where {T}
    res = Dict{Any, Vector{T}}()
    for value in values
        k = get_key(value)
        push!(get!(Vector{T}, res, k), value)
        # if haskey(res, k)
        #     push!(res[k], value)
        # else
        #     res[k] = [value]
        # end
    end
    res
end

function compute_delayed_grouping(graph, used_nodes, max_level)
    topo_sorted = topological_sort(graph)

    topo_levels = Dict{Int64, Int64}()
    for node in filter(v -> v ∈ used_nodes, topo_sorted)
        dependents = filter(in(used_nodes), inneighbors(graph, node))
        if isempty(dependents)
            topo_levels[node] = max_level
        else
            topo_levels[node] = minimum(map(k -> topo_levels[k], dependents)) - 1
        end
    end
    grouped_by_level_delayed = Dict((l => [kv.first for kv in topo_levels if kv.second == l]) for l in 0:max_level)
    grouped_by_level_delayed, topo_levels
end

function find_level_compressions(graph, used_nodes, grouped_by_level, topo_levels)
    new_grouped_by_level = deepcopy(grouped_by_level)

    max_level = maximum(keys(grouped_by_level))
    function can_be_smushed(node)
        dependencies = filter(in(used_nodes), outneighbors(graph, node))
        node_level = topo_levels[node]
        if isempty(dependencies)
            # Technically, if the node was at a non-zero level, it could be
            # freely moved down, but I'm not sure why that would be desirable
            return false, missing, node_level
        end

        dep_levels = [topo_levels[n] for n in dependencies]
        level_limiters = findall(l -> l == node_level - 1, dep_levels)


        if length(level_limiters) == 1
            max_compress_level = maximum(filter(l -> l != node_level - 1, dep_levels), init = 0)
            true, dependencies[level_limiters[1]], max_compress_level
        else
            false, missing, node_level
        end
    end

    function remove_from_current_level(node)
        current_level = topo_levels[node]
        filter!(!=(node), new_grouped_by_level[current_level])
    end

    visited = Set{Int64}()
    squashed = Set{Int64}()

    for level in reverse(0:max_level)
        level_nodes = grouped_by_level[level]

        for node in level_nodes

            smushable, level_limiter, max_compress_level = can_be_smushed(node)
            compress_group = [node]
            while smushable && !(level_limiter in visited) && (level_limiter > max_compress_level)
                push!(compress_group, level_limiter)
                smushable, level_limiter, next_max_compress_level = can_be_smushed(level_limiter)
                max_compress_level = max(max_compress_level, next_max_compress_level)
            end

            union!(visited, compress_group)

            if length(compress_group) > 20
                println("Found a group of $(length(compress_group)) nodes that can be smushed")
                insert_level = max_level
                for n in compress_group
                    dependants = filter(neighbor -> (neighbor in used_nodes) && !(neighbor in compress_group), inneighbors(graph, n))
                    limits = [topo_levels[n] - 1 for n in dependants]
                    insert_level = min(insert_level, minimum(limits; init = max_level))
                    # println("\t$(var_names_map[all_referenced_nodes[n]])")
                end
                @show insert_level

                # insert_level = topo_levels[node]
                # insert_level = max_compress_level


                remove_from_current_level.(compress_group)

                append!(new_grouped_by_level[insert_level], reverse(compress_group))
                union!(squashed, compress_group)

            end


        end

    end



    function node_max_level(node, topo_levels)
        if node in squashed
            return topo_levels[node]
        end
        dependants = filter(in(used_nodes), inneighbors(graph, node))
        if isempty(dependants)
            return max_level
        end

        limits = [topo_levels[n] - 1 for n in dependants]
        return minimum(limits)
    end

    new_topo_levels = Dict{Int64, Int64}()
    for (level, level_nodes) in new_grouped_by_level
        for n in level_nodes
            new_topo_levels[n] = level
        end
    end


    did_move = true

    while did_move
        @show did_move
        moved_any = false
        for level in reverse(0:max_level)
            level_nodes = new_grouped_by_level[level]

            for node in level_nodes
                highest_level = node_max_level(node, new_topo_levels)
                if highest_level != level
                    # println("moving $(var_names_map[all_referenced_nodes[node]]) from level $level to $highest_level")
                    filter!(!=(node), new_grouped_by_level[level])
                    push!(new_grouped_by_level[highest_level], node)
                    new_topo_levels[node] = highest_level
                    moved_any = true

                end
            end
        end
        did_move = moved_any
    end

    new_grouped_by_level
end

function group_by(itr::Vector{T}, by) where {T}

    groups = Vector{Vector{T}}()
    # find_func = (item, g) -> by(item, first(g))

    for item in itr
        found = false
        for g in groups
            if by(item, first(g))
                push!(g, item)
                found = true
                break
            end
        end
        if !found
            push!(groups, [item])
        end
        # # ind = findfirst(g -> by(item, first(g)), groups)
        # ind = findfirst(Base.Fix1(find_func, item), groups)
        # if ind === nothing
        #     push!(groups, [item])
        # else
        #     push!(groups[ind], item)
        # end
    end

    groups
end

# function group_by(itr::AbstractArray{T}, by) where {T}

#     groups = Vector{Vector{T}}()
#     find_func = (item, g) -> by(item, first(g))

#     for item in itr
#         # ind = findfirst(g -> by(item, first(g)), groups)
#         ind = findfirst(Base.Fix1(find_func, item), groups)
#         if ind === nothing
#             push!(groups, [item])
#         else
#             push!(groups[ind], item)
#         end
#     end

#     groups
# end

# This function is currently unused
# function find_column_ops(table::ExcelTable, all_cell_dict::Dict{CellDependency,CellTypes}, topo_levels::Dict{Int64,Int64}, node_nums, var_names)
function find_column_ops(table::ExcelTable, used_subset::WorkbookSubset, topo_levels, exporter)
    sheet = table.sheet_name
    all_cell_dict = used_subset.wb.cell_dict
    node_nums = used_subset.node_nums

    ctx = make_ctx(sheet, used_subset.wb.xf)

    function cell_have_same_equation(a::CellDependency, b::CellDependency)
        levels = unique(map(c -> (c in keys(node_nums) && node_nums[c] in keys(topo_levels)) ? topo_levels[node_nums[c]] : -1, (a, b)))
        if length(levels) != 1
            return false
        end
        if a in keys(all_cell_dict) && b in keys(all_cell_dict)
            a_data = all_cell_dict[a]
            b_data = all_cell_dict[b]

            if a_data isa FormulaCell && b_data isa FormulaCell
                base_expr = a_data.expr
                delta_y = rownum(a) - rownum(b)

                base_expr == offset(b_data.expr, delta_y, 0)
            elseif a_data isa ValueCell && b_data isa ValueCell1
                result = a_data.value == b_data.value
                if ismissing(result)
                    ismissing(a_data.value) && ismissing(b_data.value)
                else
                    result
                end
            else
                false
            end
        else
            !(a in keys(all_cell_dict) || b in keys(all_cell_dict))
        end
    end

    resulting_lines = Dict{Int64, Vector{Any}}()

    for col in startcol(table):endcol(table)
        # row = startrow(table)
        row_cells = [CellDependency(sheet, index_to_cellname(col, row)) for row in startrow(table):endrow(table)]
        equal_cell_groups = group_by(row_cells, cell_have_same_equation)
        col_name = column_name(table, 1 + col - startcol(table))
        # @show col_name
        # if col_name != "total annualized financing cost"
        #     continue
        # end

        for group in equal_cell_groups
            runs = get_contiguous_runs(sort(rownum.(group)))

            for run in runs
                if length(run) <= 4
                    continue
                end
                @show run

                run_cells = [CellDependency(sheet, index_to_cellname(col, r)) for r in run]
                levels = unique(map(c -> (c in keys(node_nums) && node_nums[c] in keys(topo_levels)) ? topo_levels[node_nums[c]] : -1, run_cells))
                @show levels
                if !(-1 in levels) && length(levels) == 1
                    range_string = "$(to_string(run_cells[begin])) to $(to_string(run_cells[end]))"
                    println("$range_string is a single formula")
                    println("topo_levels = $levels")
                    # @show all_cell_dict[run_cells[begin]].cell
                    first_expr = get_expr(all_cell_dict[run_cells[begin]])

                    broadcasted = try
                        convert_to_broadcasted(first_expr, run[end] - run[begin], 0)
                    catch ex
                        @show ex
                        println("Failed to broadcast, would have broadcased a run of $(length(run))")
                        @show sheet run[1]
                        continue
                    end


                    col_name = repr(column_name(table, col - startcol(table) + 1))

                    row_idx = if run[1] == startrow(table) && run[end] == endrow(table)
                        "!"
                    else
                        row_start_idx = run[1] - startrow(table) + 1
                        row_stop_idx = run[end] - startrow(table) + 1
                        "$(row_start_idx):$(row_stop_idx)"
                    end

                    # @show first_expr
                    # @show broadcasted
                    expr_table_refs = insert_table_refs(broadcasted, tables) |> convert_if_multiple
                    lhs = "$(getname(table))[$(row_idx), $col_name]"
                    # @show expr_table_refs

                    line = if contains_if(expr_table_refs)
                        function_expr, params = functionalize(expr_table_refs, [])
                        function make_function_string(function_name, function_expr, num_params)
                            params_str = join(["param_$i" for i in 1:num_params], ", ")
                            expr_str = convert(exporter, function_expr, sheet)
                            # expr_str = xl_expr_to_julia(function_expr, ctx, var_names, tables)
                            "function $function_name($params_str)\n\t$expr_str\nend\n"
                        end
                        function_name = "func_$(normalize_var_name(sheet))_$(run_cells[1].cell)_$(run_cells[end].cell)"
                        func_str = make_function_string(function_name, function_expr, length(params))
                        convert(exporter, function_expr, sheet)
                        # params_strings = [xl_expr_to_julia(param_expr, ctx, var_names, tables) for param_expr in params]
                        params_strings = [convert(exporter, param_expr, sheet) for param_expr in params]
                        func_params = join(params_strings, ", ")
                        rhs = "$function_name($(func_params))"
                        line = "$func_str@. $lhs = $rhs\n"

                    else
                        rhs = convert(exporter, expr_table_refs, sheet)
                        # rhs = xl_expr_to_julia(expr_table_refs, ctx, var_names, tables)
                        "# $range_string\n@. $lhs = $rhs\n"
                    end
                    println(table.table_name)
                    print(line)
                    # @show line
                    if !(levels[1] in keys(resulting_lines))
                        resulting_lines[levels[1]] = []
                    end

                    push!(resulting_lines[levels[1]], (run_cells, line))
                    # @show "$lhs .= $rhs\n"
                    # push!(result_lines, "$lhs .= $rhs\n")
                end
            end

        end
    end

    resulting_lines
end


function make_input_struct(all_dependencies, cell_dict, dependency_dict, used_nodes, var_names, tables)
    input_cells = []
    for node in used_nodes
        cell = all_dependencies[node]
        if cell in keys(cell_dict)
            cell_data = cell_dict[cell]
            if cell_data isa ValueCell
                push!(input_cells, cell)
            end

        else
            println("cell $(to_string(cell)) no in cell_dict?")
            @show var_names[cell]
        end
    end

    lines = Vector{String}()
    table_assigment_lines = Vector{String}()
    input_name_map = Dict{CellDependency, String}()

    push!(lines, "@kwdef struct Inputs")

    for cell in sort(input_cells)
        struct_name = var_names[cell]

        is_table_cell = false
        for table in tables
            if cell in table
                c = colnum(cell)
                r = rownum(cell)
                r_name = string(row_name(table, r - startrow(table) + 1))
                col_name = string(column_name(table, c - startcol(table) + 1))
                struct_name = "$(getname(table))_$(normalize_var_name(r_name))_$(normalize_var_name(col_name))"
                push!(table_assigment_lines, "$(var_names[cell]) = input.$struct_name\n")
                is_table_cell = true
                break
            end
        end
        if !is_table_cell
            input_name_map[cell] = string("input.", struct_name)
        end

        value = cell_dict[cell].value
        if value isa Dates.Time
            value = value.instant
        end
        @show struct_name value repr(value)
        push!(lines, "\t$struct_name = $(repr(value))")
    end

    push!(lines, "end\n")

    join(lines, "\n"), input_name_map, table_assigment_lines, input_cells
end

function renumber_levels!(grouped_by_level)
    min_level, max_level = extrema(keys(grouped_by_level))
    grouped_copy = copy(grouped_by_level)
    for k in keys(grouped_by_level)
        delete!(grouped_by_level, k)
    end

    current_level = 0
    for l in min_level:max_level
        if l in keys(grouped_copy) && !isempty(grouped_copy[l])
            grouped_by_level[current_level] = grouped_copy[l]
            current_level += 1
        end
    end
end

function find_table_containing_cell(cell::CellDependency, tables)
    findfirst(t -> cell in t, tables)
end


function get_level_dependencies(grouped_by_level, workbook_subset::WorkbookSubset, tables)
    dependencies = workbook_subset.wb.cell_dependencies
    all_cells = get_all_referenced_cells(workbook_subset.wb)

    for level in sort(collect(keys(grouped_by_level)))
        cells = map(n -> all_cells[n], grouped_by_level[level])
        all_cell_deps = [dependencies[c] for c in cells if c in keys(dependencies)]
        level_deps = reduce(union, all_cell_deps, init = [])

        cell_deps = []
        table_deps = []
        for cell in level_deps
            table = find_table_containing_cell(cell, tables)
            if isnothing(table)
                push!(cell_deps, cell)
            else
                push!(table_deps, table)
            end
        end


        println("Level $level has $(length(level_deps)) dependencies")
        println("$(length(cell_deps)) cell deps, $(length(unique(table_deps))) table deps")
    end

end

function get_used_cells(subset::WorkbookSubset)
    all_cells = get_all_referenced_cells(subset.wb)

    all_cells[subset.used_nodes]
end

function get_cell_value(workbook::ExcelWorkbook, cell::CellDependency)
    value = get(workbook.cell_dict, cell, MissingCell())
    if ismissing(value)
        MissingCell()
    else
        value
    end
end

function make_statements(subset::WorkbookSubset)
    workbook::ExcelWorkbook = subset.wb


    lhs_cells = get_used_cells(subset)

    statements = Vector{AbstractStatement}(undef, length(lhs_cells) + 1)

    for i in eachindex(lhs_cells)
        cell = lhs_cells[i]
        cell_value = get_cell_value(workbook, cell)
        rhs_expr = get_expr(cell_value)

        # stmt = StandardStatement(cell, rhs_expr, get(workbook.cell_dependencies, cell, []))
        stmt = StandardStatement(cell, rhs_expr, get_dependent_cells(workbook, cell))

        statements[i] = stmt
    end

    output_stmt = OutputStatement(subset.output_cells)
    statements[end] = output_stmt

    statements
end

function make_statements(workbook::ExcelWorkbook2)
    # workbook::ExcelWorkbook = subset.wb


    lhs_cells = get_all_referenced_cells(workbook)

    # statements = Vector{AbstractStatement}(undef, length(lhs_cells) + 1)
    statements = Vector{AbstractStatement}(undef, length(lhs_cells))

    for i in eachindex(lhs_cells)
        cell = lhs_cells[i]
        cell_value = get_cell_value(workbook, cell)
        rhs_expr = get_expr(cell_value)

        # stmt = StandardStatement(cell, rhs_expr, get(workbook.cell_dependencies, cell, []))
        stmt = StandardStatement(cell, rhs_expr, get_dependent_cells(workbook, cell))

        statements[i] = stmt
    end

    # output_stmt = OutputStatement(subset.output_cells)
    # statements[end] = output_stmt

    statements
end


function make_cell_to_statement_dict(statements::Vector{AbstractStatement})
    cell_to_statement = Dict{CellDependency, AbstractStatement}()
    for statement in statements
        set_cells = get_set_cells(statement)
        for cell in set_cells
            if cell in keys(cell_to_statement)
                @show statement
                @show cell
                @show cell_to_statement[cell]
                println("Statement", statement, "sets cell", cell)
                println("but statement", cell_to_statement[cell], "already set that cell")
                @assert !(cell in keys(cell_to_statement))
            end

            cell_to_statement[cell] = statement
        end
    end

    cell_to_statement
end

function make_statement_graph(statements::Vector{AbstractStatement})
    # cell_to_statement = make_cell_to_statement_dict(statements)
    # cell_to_statement = Dict{CellDependency,AbstractStatement}()
    # for statement in statements
    #     set_cells = get_set_cells(statement)
    #     for cell in set_cells
    #         if cell in keys(cell_to_statement)
    #             @show statement
    #             @show cell
    #             @show cell_to_statement[cell]
    #         end
    #         @assert !(cell in keys(cell_to_statement))

    #         cell_to_statement[cell] = statement
    #     end
    # end

    cell_to_statement = Dict{CellDependency, Int64}()
    for (i, statement) in enumerate(statements)
        set_cells = get_set_cells(statement)
        for cell in set_cells
            if cell in keys(cell_to_statement)
                @show statement
                @show cell
                @show cell_to_statement[cell]
                println("Statement", statement, "sets cell", cell)
                println("but statement ", cell_to_statement[cell], " already set that cell")
                @assert !(cell in keys(cell_to_statement))
            end

            cell_to_statement[cell] = i
        end
    end

    cell_to_statement

    # statement_nums = Dict{AbstractStatement, Int64}(n => i for (i, n) in enumerate(statements))
    # adj_matrix = zeros(Bool, (length(statements), length(statements)))
    # stmt_end_nodes = Vector{Int64}()

    edge_list = Vector{Edge{Int64}}()
    sizehint!(edge_list, length(statements))
    for (start_node, statement) in enumerate(statements)
        # start_node = statement_nums[statement]
        # start_node = statement_nums[statement]
        cell_deps::Vector{CellDependency} = get_cell_deps(statement)
        # empty!(stmt_end_nodes)

        for cell_dep in cell_deps
            # cell_dep::CellDependency
            # end_statement = get(cell_to_statement, cell_dep, nothing)
            end_node = get(cell_to_statement, cell_dep, nothing)
            # if !(cell_dep in keys(cell_to_statement))
            # if isnothing(end_statement)
            if isnothing(end_node)
                @show statement
                @show get_set_cells(statement)
                @show cell_dep
                continue
            end
            # end_statement = cell_to_statement[cell_dep]
            # end_node = statement_nums[end_statement]

            # In cases like grouped statements, it's possible for nodes to depend on themselves
            # we should just be able to ignore that
            if start_node != end_node
                # if !(end_node in stmt_end_nodes)
                #     push!(stmt_end_nodes, end_node)
                # end
                push!(edge_list, Edge(start_node, end_node))
            end
        end

        # append!(edge_list, Edge.(start_node, stmt_end_nodes))
    end


    # nested_edge_list = [[(statement_nums{statement}, statement_nums[cell_to_statement[cell_dep]]) for cell_dep in get_cell_deps(statement)] for statement in statements]
    # edge_list = reduce(vcat, nested_edge_list)

    # graph = Graphs.SimpleDiGraphFromIterator(Edge.(edge_list))
    # graph = Graphs.SimpleDiGraphFromIterator(edge_list)
    graph = Graphs.SimpleDiGraph(edge_list)
    # graph = Graphs.SimpleDiGraph(adj_matrix)
    cycle = find_cycle(graph)
    if !isnothing(cycle)
        println("Found a cycle!")
        for (i, node) in enumerate(cycle)
            stmt = statements[node]
            # println("\tStatement setting $(get_set_cells(stmt))")
            # println("\t$i: $cell, $(get_cell_value(used_subset, cell))")
            println("\t$i: $stmt")
            # @display get_expr(used_subset.cell_dict[cell])
        end
        throw("statement graph had a cycle!")

    end

    cycles = Graphs.simplecycles(graph)
    if !isempty(cycles)
        println("Removing $(length(cycles)) cycles from the statement graph, this is almost certainly incorrect.")
        for cycle in cycles
            for stmt in statements[cycle]
                println("\tStatement setting $(get_set_cells(stmt))")
            end
            rem_edge!(graph, cycle[end], cycle[begin])
            println("-"^20)
        end
    end

    graph
end

# export_statements(io::IO, exporter, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement}) = export_statements_levels_with_moving(io, exporter, wb, statements)
# export_statements(io::IO, exporter, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement}) = export_statements_levels(io, exporter, wb, statements)
export_statements(io::IO, exporter, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement}) = export_statements_global_affinity(io, exporter, wb, statements)

function export_statements_optimized(io::IO, exporter, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement})
    stmt_graph = make_statement_graph(statements)

    non_input_statements = Vector{AbstractStatement}()

    for i in eachindex(statements)
        if length(outneighbors(stmt_graph, i)) > 0
            push!(non_input_statements, statements[i])
        end
    end

    stmt_graph = make_statement_graph_relaxed(non_input_statements)

    @time optimized_order, costs = optimized_topological_sa_big_jump(reverse(stmt_graph), ; distance_decay = 0.5, iterations = 10000000, initial_temp = 20, cooling_rate = 0.9999995)

    # @time optimized_order, costs = optimized_topological_sa_big_jump(reverse(stmt_graph); starting_order=nothing, distance_decay=0.5, iterations=1000000, initial_temp=15, cooling_rate=0.99999)
    # sorted_order = topological_sort(reverse(stmt_graph))
    # stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    # stmt_topo_levels = get_topo_levels_bottom_up(stmt_graph)
    # bottom_up_levels = get_topo_levels_bottom_up(stmt_graph)
    # input_statements = Set([kv.first for kv in bottom_up_levels if kv.second == 0])

    # max_level = maximum(values(stmt_topo_levels))

    for statement in non_input_statements[optimized_order]
        # for statement in non_input_statements[sorted_order]
        write(io, export_statement(exporter, wb, statement))
    end
    # grouped_by_level = Dict((l => [kv.first for kv in topo_levels if kv.second == l]) for l in 0:max_level)
end

function export_statements_levels_with_moving(io::IO, exporter, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement})
    stmt_graph = make_statement_graph(statements)
    cell_to_statement = make_cell_to_statement_dict(statements)
    stmt_to_node = Dict([n => i for (i, n) in enumerate(statements)])

    stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    # stmt_topo_levels = get_topo_levels_bottom_up(stmt_graph)
    bottom_up_levels = get_topo_levels_bottom_up(stmt_graph)
    input_statements = Set([kv.first for kv in bottom_up_levels if kv.second == 0])

    max_level = maximum(values(stmt_topo_levels))
    # statement_strings = [export_statement(exporter, wb, s) for s in statements]

    level_order_indices = []

    for level in 0:max_level
        level_statement_idx = [kv.first for kv in stmt_topo_levels if (kv.second == level) && !(kv.first in input_statements)]

        sort!(level_statement_idx, by = i -> get_set_cells(statements[i])[1])

        push!(level_order_indices, level_statement_idx)
    end

    for _ in 1:1
        any_moved = false

        for level in (max_level-2):-1:2
            stmt_idxs = level_order_indices[level]
            new_level_order_indices = []

            for n in stmt_idxs
                dep_stmts = outneighbors(stmt_graph, n)

                stmts_to_move = []

                for dep in dep_stmts
                    dep_usages = inneighbors(stmt_graph, dep)
                    if length(dep_usages) == 1 && !(dep in input_statements)
                        push!(stmts_to_move, dep)
                    end
                end

                deleteat!(level_order_indices[level-1], findall(level_order_indices[level-1] .∈ (stmts_to_move,)))
                if length(stmts_to_move) > 0
                    println("Moving $(length(stmts_to_move)) statements up")
                    # @show statements[n]
                    # @show statements[stmts_to_move]
                    any_moved = true
                end

                append!(new_level_order_indices, stmts_to_move)
                push!(new_level_order_indices, n)
            end

            level_order_indices[level] = new_level_order_indices
        end

        if !any_moved
            break
        end
    end

    for level in 0:max_level
        # level_statement_idx = [kv.first for kv in stmt_topo_levels if (kv.second == level) && !(kv.first in input_statements)]
        # level_statements = statements[level_statement_idx]

        # level_statements = sort(level_statements, by=s -> get_set_cells(s)[1])

        println("Level: $(level)")
        write(io, "# Level $(level)\n")

        level_statements = statements[level_order_indices[level+1]]

        for s in level_statements
            node = stmt_to_node[s]
            dependents = statements[inneighbors(stmt_graph, node)]
            if !isempty(dependents)
                usages = "[" * join(to_string.((exporter,), dependents), ", ") * "]"
                # write(io, "# Used in $(length(dependents)) places: $usages\n")
            end

            write(io, export_statement(exporter, wb, s))
        end

        write(io, "\n\n")
    end
    # grouped_by_level = Dict((l => [kv.first for kv in topo_levels if kv.second == l]) for l in 0:max_level)
end

function export_statements_levels(io::IO, exporter, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement})
    stmt_graph = make_statement_graph(statements)
    cell_to_statement = make_cell_to_statement_dict(statements)
    stmt_to_node = Dict([n => i for (i, n) in enumerate(statements)])

    stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    # stmt_topo_levels = get_topo_levels_bottom_up(stmt_graph)
    bottom_up_levels = get_topo_levels_bottom_up(stmt_graph)
    input_statements = Set([kv.first for kv in bottom_up_levels if kv.second == 0])

    max_level = maximum(values(stmt_topo_levels))

    for level in 0:max_level
        level_statement_idx = [kv.first for kv in stmt_topo_levels if (kv.second == level) && !(kv.first in input_statements)]
        level_statements = statements[level_statement_idx]

        level_statements = sort(level_statements, by = s -> get_set_cells(s)[1])

        # println("Level: $(level)")
        write(io, "# Level $(level)\n")

        for s in level_statements
            node = stmt_to_node[s]
            dependents = statements[inneighbors(stmt_graph, node)]
            if !isempty(dependents)
                # usages = if length(dependents) > 5
                #     "[" * join(to_string.((exporter,), dependents[1:4]), ", ") * ", ..., " * to_string(exporter, dependents[end]) * "]"
                # else
                #     "[" * join(to_string.((exporter,), dependents), ", ") * "]"
                # end
                # usages = "[" * join(to_string.((exporter,), dependents), ", ") * "]"
                # write(io, "# Used in $(length(dependents)) places: $usages\n")
            end

            write(io, export_statement(exporter, wb, s))
        end

        write(io, "\n\n")
    end
    # grouped_by_level = Dict((l => [kv.first for kv in topo_levels if kv.second == l]) for l in 0:max_level)
end

function export_statements_levels_table_grouped(io::IO, exporter, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement})
    stmt_graph = make_statement_graph(statements)

    stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    bottom_up_levels = get_topo_levels_bottom_up(stmt_graph)
    input_statements = Set([kv.first for kv in bottom_up_levels if kv.second == 0])

    max_level = maximum(values(stmt_topo_levels))

    level_order_indices = [Int64[] for _ in 0:max_level]
    statement_levels = Dict{Int64, Int64}()
    for level in 0:max_level
        level_statement_idx = [kv.first for kv in stmt_topo_levels if (kv.second == level) && !(kv.first in input_statements)]
        sort!(level_statement_idx, by = i -> get_set_cells(statements[i])[1])
        level_order_indices[level+1] = level_statement_idx
        for idx in level_statement_idx
            statement_levels[idx] = level
        end
    end

    table_name(table) = try
        string(getname(table))
    catch
        "<unknown>"
    end

    move_diag = Dict{Int64, Vector{Tuple{Int64, Any}}}((l => Tuple{Int64, Any}[]) for l in 0:max_level)

    # Move same-table TableStatements upward across levels, promoting directly
    # to the highest feasible destination level when valid.
    any_moved = true
    while any_moved
        any_moved = false

        for source_level in 0:(max_level - 1)
            source_idx = level_order_indices[source_level+1]

            i = 1
            while i <= length(source_idx)
                node = source_idx[i]
                stmt = statements[node]
                if !(stmt isa TableStatement)
                    i += 1
                    continue
                end

                table = get_set_table(stmt)
                node_deps = [d for d in outneighbors(stmt_graph, node) if d in keys(statement_levels)]
                node_users = [u for u in inneighbors(stmt_graph, node) if u in keys(statement_levels)]

                if isempty(node_users)
                    i += 1
                    continue
                end

                # Keep this pass focused on table grouping.
                has_same_table_user_above = any(u -> (statement_levels[u] > source_level) &&
                                                  (statements[u] isa TableStatement) &&
                                                  (get_set_table(statements[u]) == table), node_users)
                if !has_same_table_user_above
                    i += 1
                    continue
                end

                # Highest legal level is bounded by earliest user.
                max_feasible_level = minimum(statement_levels[u] for u in node_users)
                if max_feasible_level <= source_level
                    i += 1
                    continue
                end

                moved = false
                for dest_level in max_feasible_level:-1:(source_level + 1)
                    # dependencies must be at or below destination
                    if any(statement_levels[d] > dest_level for d in node_deps)
                        continue
                    end
                    # users must be at or above destination
                    if any(statement_levels[u] < dest_level for u in node_users)
                        continue
                    end

                    target_idx = level_order_indices[dest_level+1]
                    target_pos = Dict((n => p) for (p, n) in enumerate(target_idx))

                    lower = maximum((get(target_pos, d, 0) + 1 for d in node_deps); init = 1)
                    upper = minimum((get(target_pos, u, length(target_idx) + 1) for u in node_users); init = length(target_idx) + 1)
                    if lower > upper
                        continue
                    end

                    # Prefer insertion near same-table users in destination.
                    same_table_users_in_dest = [u for u in node_users if statement_levels[u] == dest_level &&
                                                                   (statements[u] isa TableStatement) &&
                                                                   (get_set_table(statements[u]) == table)]
                    preferred = if isempty(same_table_users_in_dest)
                        upper
                    else
                        minimum(get(target_pos, u, upper) for u in same_table_users_in_dest)
                    end
                    insert_pos = min(max(preferred, lower), upper)

                    deleteat!(source_idx, i)
                    insert!(target_idx, insert_pos, node)
                    statement_levels[node] = dest_level

                    push!(move_diag[dest_level], (source_level, table))
                    any_moved = true
                    moved = true
                    break
                end

                if !moved
                    i += 1
                end
            end
        end
    end

    function order_level_table_aware(level_idx::Vector{Int64}, level::Int64)
        if length(level_idx) <= 1
            return copy(level_idx), String[]
        end

        in_level = Set(level_idx)
        base_pos = Dict((idx => pos) for (pos, idx) in enumerate(level_idx))
        in_level_prereq_count = Dict((idx => 0) for idx in level_idx)
        in_level_dependents = Dict((idx => Int64[]) for idx in level_idx)

        for idx in level_idx
            for dep in outneighbors(stmt_graph, idx)
                if dep in in_level
                    in_level_prereq_count[idx] += 1
                    push!(in_level_dependents[dep], idx)
                end
            end
        end

        available = [idx for idx in level_idx if in_level_prereq_count[idx] == 0]
        sort!(available, by = idx -> base_pos[idx])

        ordered = Int64[]
        last_table = missing
        local_reorder_counts = Dict{Any, Int64}()

        while !isempty(available)
            default_choice = available[1]
            chosen = default_choice

            if !ismissing(last_table)
                same_table = filter(idx -> (statements[idx] isa TableStatement) && get_set_table(statements[idx]) == last_table, available)
                if !isempty(same_table)
                    chosen = same_table[1]
                    if chosen != default_choice
                        local_reorder_counts[last_table] = get(local_reorder_counts, last_table, 0) + 1
                    end
                end
            end

            deleteat!(available, findfirst(==(chosen), available))
            push!(ordered, chosen)

            if statements[chosen] isa TableStatement
                last_table = get_set_table(statements[chosen])
            else
                last_table = missing
            end

            for user in in_level_dependents[chosen]
                in_level_prereq_count[user] -= 1
                if in_level_prereq_count[user] == 0
                    push!(available, user)
                end
            end
            sort!(available, by = idx -> base_pos[idx])
        end

        comments = String[]
        if length(ordered) != length(level_idx)
            # Fallback: keep remaining statements in base order if constraints became cyclic.
            remaining = filter(idx -> !(idx in ordered), level_idx)
            append!(ordered, remaining)
            push!(comments, "# [diag] warning: unresolved same-level ordering cycle at level $level; used base order fallback for $(length(remaining)) statement(s)")
        end

        for (table, count) in local_reorder_counts
            push!(comments, "# [diag] grouped TableStatements within level $level for table \"$(table_name(table))\" (reorder step(s): $count)")
        end

        ordered, comments
    end

    for level in 0:max_level
        write(io, "# Level $(level)\n")

        if !isempty(move_diag[level])
            grouped_moves = Dict{Any, Tuple{Int64, Set{Int64}}}()
            for (source_level, table) in move_diag[level]
                count, levels = get(grouped_moves, table, (0, Set{Int64}()))
                push!(levels, source_level)
                grouped_moves[table] = (count + 1, levels)
            end

            for (table, (count, levels)) in grouped_moves
                from_levels = sort(collect(levels))
                from_levels_str = join(from_levels, ", ")
                write(io, "# [diag] cross-level grouping moved $count statement(s) from level(s) [$from_levels_str] into level $level for table \"$(table_name(table))\"\n")
            end
        end

        ordered_idx, level_diag = order_level_table_aware(level_order_indices[level+1], level)
        for comment in level_diag
            write(io, "$comment\n")
        end

        for idx in ordered_idx

            dependents = statements[inneighbors(stmt_graph, idx)]
            if !isempty(dependents)
                usages = if length(dependents) > 5
                    "[" * join(to_string.((exporter,), dependents[1:4]), ", ") * ", ..., " * to_string(exporter, dependents[end]) * "]"
                else
                    "[" * join(to_string.((exporter,), dependents), ", ") * "]"
                end
                usages = "[" * join(to_string.((exporter,), dependents), ", ") * "]"
                write(io, "# Used in $(length(dependents)) places: $usages\n")
            end
            write(io, export_statement(exporter, wb, statements[idx]))
        end

        write(io, "\n\n")
    end
end

statement_table_key(stmt::AbstractStatement) = missing
statement_table_key(stmt::TableStatement) = get_set_table(stmt)
statement_table_key(stmt::BroadcastedStatement) = get_set_table(stmt)

function get_global_affinity_order(wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement})
    stmt_graph = make_statement_graph(statements)
    stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    bottom_up_levels = get_topo_levels_bottom_up(stmt_graph)
    input_statements = Set([kv.first for kv in bottom_up_levels if kv.second == 0])

    active_indices = filter(i -> !(i in input_statements), collect(eachindex(statements)))
    if isempty(active_indices)
        # write(io, "# Global affinity schedule: no non-input statements to export.\n")
        return Int64[]
    end

    function statement_sheet_key(stmt::AbstractStatement)
        cells = get_set_cells(stmt)
        isempty(cells) && return missing
        base_sheet = string(cells[1].sheet_name)
        if any(c -> string(c.sheet_name) != base_sheet, cells)
            return "<mixed>"
        end
        base_sheet
    end


    max_level = maximum(values(stmt_topo_levels))
    base_order = Int64[]
    for level in 0:max_level
        level_idxs = filter(i -> stmt_topo_levels[i] == level && (i in active_indices), eachindex(statements))
        sort!(level_idxs, by = i -> isempty(get_set_cells(statements[i])) ? "" : string(get_set_cells(statements[i])[1]))
        append!(base_order, level_idxs)
    end
    base_rank = Dict((idx => pos) for (pos, idx) in enumerate(base_order))

    active_set = Set(active_indices)
    pending_deps = Dict{Int64, Int64}()
    for idx in active_indices
        pending_deps[idx] = count(dep -> dep in active_set, outneighbors(stmt_graph, idx))
    end

    ready = [idx for idx in active_indices if pending_deps[idx] == 0]
    sort!(ready, by = idx -> base_rank[idx])

    table_weight = 100
    sheet_weight = 15
    max_event_diags = 300

    last_table = missing
    last_sheet = missing
    ordered = Int64[]
    event_diags = String[]

    table_affinity_picks = 0
    sheet_affinity_picks = 0
    dual_affinity_picks = 0
    base_picks = 0
    ready_size_sum = 0
    max_ready_size = 0

    while !isempty(ready)
        ready_size_sum += length(ready)
        max_ready_size = max(max_ready_size, length(ready))

        base_choice = ready[1]
        chosen = base_choice
        chosen_score = -1
        chosen_table_match = false
        chosen_sheet_match = false

        for candidate in ready
            cand_stmt = statements[candidate]
            cand_table = statement_table_key(cand_stmt)
            cand_sheet = statement_sheet_key(cand_stmt)

            table_match = !ismissing(last_table) && !ismissing(cand_table) && cand_table == last_table
            sheet_match = !ismissing(last_sheet) && !ismissing(cand_sheet) && cand_sheet == last_sheet
            score = table_match * table_weight + sheet_match * sheet_weight

            if score > chosen_score || (score == chosen_score && base_rank[candidate] < base_rank[chosen])
                chosen = candidate
                chosen_score = score
                chosen_table_match = table_match
                chosen_sheet_match = sheet_match
            end
        end

        if chosen_score == 0 || chosen == base_choice
            base_picks += 1
        else
            if chosen_table_match && chosen_sheet_match
                dual_affinity_picks += 1
            elseif chosen_table_match
                table_affinity_picks += 1
            elseif chosen_sheet_match
                sheet_affinity_picks += 1
            end

            if length(event_diags) < max_event_diags
                chosen_stmt = statements[chosen]
            end
        end

        deleteat!(ready, findfirst(==(chosen), ready))
        push!(ordered, chosen)

        chosen_stmt = statements[chosen]
        last_table = statement_table_key(chosen_stmt)
        last_sheet = statement_sheet_key(chosen_stmt)

        for user in inneighbors(stmt_graph, chosen)
            if !(user in active_set)
                continue
            end
            pending_deps[user] -= 1
            if pending_deps[user] == 0
                push!(ready, user)
            end
        end
        sort!(ready, by = idx -> base_rank[idx])
    end

    return ordered
end

function export_statements_global_affinity(io::IO, exporter, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement})
    stmt_graph = make_statement_graph(statements)
    stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    bottom_up_levels = get_topo_levels_bottom_up(stmt_graph)
    input_statements = Set([kv.first for kv in bottom_up_levels if kv.second == 0])

    active_indices = filter(i -> !(i in input_statements), collect(eachindex(statements)))
    if isempty(active_indices)
        write(io, "# Global affinity schedule: no non-input statements to export.\n")
        return
    end

    function statement_sheet_key(stmt::AbstractStatement)
        cells = get_set_cells(stmt)
        isempty(cells) && return missing
        base_sheet = string(cells[1].sheet_name)
        if any(c -> string(c.sheet_name) != base_sheet, cells)
            return "<mixed>"
        end
        base_sheet
    end


    table_name(table) = try
        string(getname(table))
    catch
        "<unknown>"
    end

    max_level = maximum(values(stmt_topo_levels))
    base_order = Int64[]
    for level in 0:max_level
        level_idxs = filter(i -> stmt_topo_levels[i] == level && (i in active_indices), eachindex(statements))
        sort!(level_idxs, by = i -> isempty(get_set_cells(statements[i])) ? "" : string(get_set_cells(statements[i])[1]))
        append!(base_order, level_idxs)
    end
    base_rank = Dict((idx => pos) for (pos, idx) in enumerate(base_order))

    active_set = Set(active_indices)
    pending_deps = Dict{Int64, Int64}()
    for idx in active_indices
        pending_deps[idx] = count(dep -> dep in active_set, outneighbors(stmt_graph, idx))
    end

    ready = [idx for idx in active_indices if pending_deps[idx] == 0]
    sort!(ready, by = idx -> base_rank[idx])

    table_weight = 100
    sheet_weight = 15
    max_event_diags = 300

    last_table = missing
    last_sheet = missing
    ordered = Int64[]
    event_diags = String[]

    table_affinity_picks = 0
    sheet_affinity_picks = 0
    dual_affinity_picks = 0
    base_picks = 0
    ready_size_sum = 0
    max_ready_size = 0

    while !isempty(ready)
        ready_size_sum += length(ready)
        max_ready_size = max(max_ready_size, length(ready))

        base_choice = ready[1]
        chosen = base_choice
        chosen_score = -1
        chosen_table_match = false
        chosen_sheet_match = false

        for candidate in ready
            cand_stmt = statements[candidate]
            cand_table = statement_table_key(cand_stmt)
            cand_sheet = statement_sheet_key(cand_stmt)

            table_match = !ismissing(last_table) && !ismissing(cand_table) && cand_table == last_table
            sheet_match = !ismissing(last_sheet) && !ismissing(cand_sheet) && cand_sheet == last_sheet
            score = table_match * table_weight + sheet_match * sheet_weight

            if score > chosen_score || (score == chosen_score && base_rank[candidate] < base_rank[chosen])
                chosen = candidate
                chosen_score = score
                chosen_table_match = table_match
                chosen_sheet_match = sheet_match
            end
        end

        if chosen_score == 0 || chosen == base_choice
            base_picks += 1
        else
            if chosen_table_match && chosen_sheet_match
                dual_affinity_picks += 1
            elseif chosen_table_match
                table_affinity_picks += 1
            elseif chosen_sheet_match
                sheet_affinity_picks += 1
            end

            if length(event_diags) < max_event_diags
                # reason = chosen_table_match && chosen_sheet_match ? "table+sheet" :
                #          chosen_table_match ? "table" : "sheet"
                chosen_stmt = statements[chosen]
                chosen_table = statement_table_key(chosen_stmt)
                chosen_sheet = statement_sheet_key(chosen_stmt)
                chosen_desc = to_string(exporter, chosen_stmt)
                base_desc = to_string(exporter, statements[base_choice])
                table_str = ismissing(chosen_table) ? "none" : table_name(chosen_table)
                sheet_str = ismissing(chosen_sheet) ? "none" : string(chosen_sheet)
                # push!(event_diags, "# [diag] affinity pick at step $(length(ordered)+1): reason=$reason table=\"$table_str\" sheet=\"$sheet_str\" over base=$base_desc chose=$chosen_desc ready=$(length(ready))")
            end
        end

        deleteat!(ready, findfirst(==(chosen), ready))
        push!(ordered, chosen)

        chosen_stmt = statements[chosen]
        last_table = statement_table_key(chosen_stmt)
        last_sheet = statement_sheet_key(chosen_stmt)

        for user in inneighbors(stmt_graph, chosen)
            if !(user in active_set)
                continue
            end
            pending_deps[user] -= 1
            if pending_deps[user] == 0
                push!(ready, user)
            end
        end
        sort!(ready, by = idx -> base_rank[idx])
    end

    if length(ordered) != length(active_indices)
        missing_nodes = filter(i -> !(i in ordered), active_indices)
        sort!(missing_nodes, by = i -> base_rank[i])
        append!(ordered, missing_nodes)
        push!(event_diags, "# [diag] warning: scheduler left $(length(missing_nodes)) statement(s) unscheduled; appended in base order")
    end

    avg_ready = ready_size_sum / max(length(ordered), 1)
    omitted_event_diags = max(0, (table_affinity_picks + sheet_affinity_picks + dual_affinity_picks) - length(event_diags))

    write(io, "# Global Topological Schedule (Affinity Heuristic)\n")
    write(io, "# [diag] table affinity weight=$table_weight, sheet affinity weight=$sheet_weight\n")
    write(io, "# [diag] scheduled $(length(ordered)) statement(s), excluded $(length(input_statements)) input/base statement(s)\n")
    write(io, "# [diag] decisions: base=$base_picks table=$table_affinity_picks sheet=$sheet_affinity_picks table+sheet=$dual_affinity_picks\n")
    write(io, "# [diag] ready-set stats: avg=$(round(avg_ready, digits=2)) max=$max_ready_size\n")
    if omitted_event_diags > 0
        write(io, "# [diag] omitted $omitted_event_diags detailed affinity event line(s) to keep output manageable\n")
    end
    for line in event_diags
        write(io, "$line\n")
    end
    write(io, "\n")

    previous_table = missing
    previous_sheet = missing
    for (step, idx) in enumerate(ordered)
        stmt = statements[idx]
        current_table = statement_table_key(stmt)
        current_sheet = statement_sheet_key(stmt)

        if current_table !== previous_table || current_sheet !== previous_sheet
            if !ismissing(current_table)
                write(io, "# [diag] context step $step: table=\"$(table_name(current_table))\" sheet=\"$(ismissing(current_sheet) ? "none" : string(current_sheet))\"\n")
            elseif !ismissing(current_sheet)
                write(io, "# [diag] context step $step: sheet=\"$(current_sheet)\"\n")
            else
                write(io, "# [diag] context step $step: statement without sheet/table context\n")
            end
        end

        dependents = statements[inneighbors(stmt_graph, idx)]
        if !isempty(dependents)
            usages = if length(dependents) > 3
                "[" * join(to_string.((exporter,), dependents[1:2]), ", ") * ", ..., " * to_string(exporter, dependents[end]) * "]"
            else
                "[" * join(to_string.((exporter,), dependents), ", ") * "]"
            end
            write(io, "# Used in $(length(dependents)) places: $usages\n")
        end

        write(io, export_statement(exporter, wb, stmt))
        previous_table = current_table
        previous_sheet = current_sheet
    end

    write(io, "\n")
end


function get_input_statements(statements::AbstractArray{AbstractStatement})
    stmt_graph = make_statement_graph(statements)

    stmt_topo_levels_bottom_up = get_topo_levels_bottom_up(stmt_graph)
    statements_bottom_up = group_to_dict(1:length(statements), s -> stmt_topo_levels_bottom_up[s])

    input_statements = statements[statements_bottom_up[0]]
    input_statements
end

function get_input_comment(exporter::JuliaExporter, statements::AbstractArray{AbstractStatement}, stmt_graph, statement_num)
    usages = inneighbors(stmt_graph, statement_num)

    num_children = length(usages)
    comment = "used in $num_children statements"
    # if num_children == 1
    for usage in usages

        child_stmt = statements[usage]
        output_cells = join([exporter.var_names[c] for c in get_set_cells(child_stmt)], ", ")
        if length(output_cells) > 80
            output_cells = output_cells[1:77] * "..."
        end
        comment *= ", [" * output_cells * "]"

    end
    # end

    # for stmt_i in usages
    #     stmt = statements[stmt_i]

    # end

    comment
end

function get_input_comment(exporter::PythonExporter, statements::AbstractArray{AbstractStatement}, stmt_graph, statement_num)
    usages = inneighbors(stmt_graph, statement_num)

    num_children = length(usages)
    comment = "used in $num_children statements"
    # if num_children == 1
    for usage in usages

        child_stmt = statements[usage]
        output_cells = join([exporter.var_names[c] for c in get_set_cells(child_stmt)], ", ")
        if length(output_cells) > 80
            output_cells = output_cells[1:77] * "..."
        end
        comment *= ", [" * output_cells * "]"

    end
    # end

    # for stmt_i in usages
    #     stmt = statements[stmt_i]

    # end

    comment
end

function make_input_struct(exporter, statements::AbstractArray{AbstractStatement})
    stmt_graph = make_statement_graph(statements)

    stmt_topo_levels_bottom_up = get_topo_levels_bottom_up(stmt_graph)
    statements_bottom_up = group_to_dict(1:length(statements), s -> stmt_topo_levels_bottom_up[s])

    input_statements = statements[statements_bottom_up[0]]
    # input_statements = get_input_statements(statements)

    input_standard_stmts = filter(s -> s isa StandardStatement, input_statements)
    input_assigned_vars = map(s -> s.assigned_var, input_standard_stmts)
    sort_order = sortperm(input_assigned_vars)

    var_names = [exporter.var_names[c] for c in input_assigned_vars[sort_order]]
    var_values = [convert(exporter, stmt.rhs_expr, stmt.assigned_var.sheet_name) for stmt in input_standard_stmts[sort_order]]
    var_types = [exporter.cell_types[c] for c in input_assigned_vars[sort_order]]

    input_statement_nums = filter(s -> statements[s] isa StandardStatement, statements_bottom_up[0])[sort_order]

    var_comments = [get_input_comment(exporter, statements, stmt_graph, i) for i in input_statement_nums]

    struct_str = make_struct(exporter, "Inputs", var_names; var_types = var_types, ismutable = true, default_values = var_values, var_comments = var_comments)

    struct_str, var_names
end

function make_dataframe_declaration(exporter::JuliaExporter, wb::ExcelWorkbook, table::ExcelTable)
    lhs = getname(table)
    num_rows, num_cols = size(table)

    sheet = table.sheet_name
    # println(getname(table))

    col_defs = Vector{String}()
    sizehint!(col_defs, num_cols)
    if is_transposed(table)
        start_r = startrow(table)
        for r in startrow(table):endrow(table)
            col_cells = [CellDependency(sheet, c, r) for c in startcol(table):endcol(table)]

            types = reduce(union_types, [get(exporter.cell_types, c, Missing) for c in col_cells])
            col_values = if types == Missing
                "Vector{Missing}(missing, $num_rows)"
            elseif types == Float64
                "zeros($num_rows)"
            elseif types == Any
                "Vector{Any}(missing, $num_rows)"
            elseif types isa DataType
                "Vector{Union{$types, Missing}}(missing, $num_rows)"
            else
                push!(types, Missing)

                "Vector{Union{$(join(string.(types),","))}}(missing, $num_rows)"
            end

            col_name = column_name(table, r - start_r + 1)
            # col_def = "$(repr(col_name)) => fill!(Vector{$type_str}(undef, $num_rows), $initial_value)"
            col_def = "$(repr(col_name)) => $col_values"
            push!(col_defs, col_def)
            # println("Col: $col_name type: $(types)")
        end
    else
        start_c = startcol(table)
        for c in startcol(table):endcol(table)
            col_cells = [CellDependency(sheet, c, r) for r in startrow(table):endrow(table)]

            types = reduce(union_types, [get(exporter.cell_types, c, Missing) for c in col_cells])
            col_values = if types == Missing
                "Vector{Missing}(missing, $num_rows)"
            elseif types == Float64
                "zeros($num_rows)"
            elseif types == Any
                "Vector{Any}(missing, $num_rows)"
            elseif types isa DataType
                "Vector{Union{$types, Missing}}(missing, $num_rows)"
            else
                push!(types, Missing)

                "Vector{Union{$(join(string.(types),","))}}(missing, $num_rows)"
            end

            col_name = column_name(table, c - start_c + 1)
            # col_def = "$(repr(col_name)) => fill!(Vector{$type_str}(undef, $num_rows), $initial_value)"
            col_def = "$(repr(col_name)) => $col_values"
            push!(col_defs, col_def)
            # println("Col: $col_name type: $(types)")
        end
    end
    # col_names = [column_name(table, c) for c in 1:num_cols]
    # line_str = "\t$lhs = DataFrame(Base.convert(Matrix{Any}, zeros($num_rows, $num_cols)), [$(join(repr.(col_names), ", "))])"
    line_str = "\t$lhs = DataFrame($(join(col_defs, ", ")))"
    # line_str = "\t$lhs = Base.convert(Matrix{Any}, zeros($num_rows, $num_cols))"

    line_str
end

function make_dataframe_declaration(exporter::PythonExporter, wb::ExcelWorkbook, table::ExcelTable)
    lhs = getname(table)
    num_rows, num_cols = size(table)

    sheet = table.sheet_name
    # println(getname(table))

    col_defs = Vector{String}()
    sizehint!(col_defs, num_cols)
    if is_transposed(table)
        start_r = startrow(table)
        for r in startrow(table):endrow(table)
            col_cells = [CellDependency(sheet, c, r) for c in startcol(table):endcol(table)]

            # types = reduce(union_types, [get(exporter.cell_types, c, Missing) for c in col_cells])
            # col_values = if types == Missing
            #     "Vector{Missing}(missing, $num_rows)"
            # elseif types == Float64
            #     "zeros($num_rows)"
            # elseif types == Any
            #     "Vector{Any}(missing, $num_rows)"
            # elseif types isa DataType
            #     "Vector{Union{$types, Missing}}(missing, $num_rows)"
            # else
            #     push!(types, Missing)

            #     "Vector{Union{$(join(string.(types),","))}}(missing, $num_rows)"
            # end

            col_name = column_name(table, r - start_r + 1)
            # col_def = "$(repr(col_name)) => fill!(Vector{$type_str}(undef, $num_rows), $initial_value)"
            col_def = "$(repr(col_name)): $col_values"
            push!(col_defs, col_def)
            # println("Col: $col_name type: $(types)")
        end
    else
        start_c = startcol(table)
        for c in startcol(table):endcol(table)
            col_cells = [CellDependency(sheet, c, r) for r in startrow(table):endrow(table)]

            # types = reduce(union_types, [get(exporter.cell_types, c, Missing) for c in col_cells])
            # col_values = if types == Missing
            #     "Vector{Missing}(missing, $num_rows)"
            # elseif types == Float64
            #     "zeros($num_rows)"
            # elseif types == Any
            #     "Vector{Any}(missing, $num_rows)"
            # elseif types isa DataType
            #     "Vector{Union{$types, Missing}}(missing, $num_rows)"
            # else
            #     push!(types, Missing)

            #     "Vector{Union{$(join(string.(types),","))}}(missing, $num_rows)"
            # end

            col_name = column_name(table, c - start_c + 1)
            # col_def = "$(repr(col_name)) => fill!(Vector{$type_str}(undef, $num_rows), $initial_value)"
            # col_def = "$(repr(col_name)): $col_values"
            # col_def = "$(repr(col_name)): np.zeros($num_rows)"
            col_def = "$(repr(col_name))"
            push!(col_defs, col_def)
            # println("Col: $col_name type: $(types)")
        end
    end
    # col_names = [column_name(table, c) for c in 1:num_cols]
    # line_str = "\t$lhs = DataFrame(Base.convert(Matrix{Any}, zeros($num_rows, $num_cols)), [$(join(repr.(col_names), ", "))])"
    # line_str = "\t$lhs = pd.DataFrame({$(join(col_defs, ",\n\t"))})"
    cols_str = join(col_defs, ", ")
    as_str(val) = val isa AbstractString ? repr(val) : string(val)
    row_names = row_name.(Ref(table), 1:num_rows)

    index_str = if row_names == 1:length(row_names)
        "range(1, $(length(row_names) + 1))"
    else
        "[" * join(as_str.(row_names), ", ") * "]"
    end
    line_str = """\t$lhs = pd.DataFrame(
    \t\tnp.zeros(($num_rows, $num_cols), dtype=object),
    \t\tcolumns=[$cols_str],
    \t\tindex=$index_str,
    \t)"""
    # line_str = "\t$lhs = Base.convert(Matrix{Any}, zeros($num_rows, $num_cols))"

    line_str
end

function make_input_table_struct(exporter::JuliaExporter, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement})
    tables = exporter.tables

    lines = Vector{String}()

    struct_names = getname.(tables)
    # var_types = repeat(["Matrix"], length(tables))
    var_types = repeat(["DataFrame"], length(tables))
    struct_def = make_struct(exporter, "Tables", struct_names; var_types = var_types)
    push!(lines, struct_def)

    input_statements = get_input_statements(statements)
    input_table_stmts = filter(s -> s isa TableStatement, input_statements)
    grouped_by_set_table = group_to_dict(input_table_stmts, get_set_table)
    push!(lines, "function make_input_tables()")

    for table in tables

        top_left = CellDependency(table.sheet_name, startcol(table), startrow(table))
        bottom_right = CellDependency(table.sheet_name, endcol(table), endrow(table))
        push!(lines, "\t# Table $top_left:$bottom_right")
        push!(lines, make_dataframe_declaration(exporter, wb, table))

        if !(table in keys(grouped_by_set_table))
            println("Table $table is not set by anything")
            continue
        end

        group = grouped_by_set_table[table]

        sort!(group, by = s -> get_set_cells(s)[1])

        get_row_num = s -> rownum(s.assigned_vars[1])
        get_col_num = s -> colnum(s.assigned_vars[1])

        row_nums = get_row_num.(group)
        col_nums = get_col_num.(group)
        coords = zip(col_nums, row_nums) |> collect
        coord_to_statement = Dict(c => s for (c, s) in zip(coords, group))
        regions = get_2d_regions(coords)
        for region in regions
            cols, rows = region
            region_coords = vec([(c, r) for c in cols, r in rows])

            if length(region_coords) < 3
                for c in region_coords
                    s = coord_to_statement[c]
                    string = export_statement(exporter, wb, s)
                    push!(lines, indent(rstrip(string), 1))
                end
            else
                region_statements = map(c -> coord_to_statement[c], region_coords)

                first_statement = region_statements[1]

                lhs = convert_to_broadcasted(first_statement.lhs_expr, length(rows) - 1, length(cols) - 1)
                # stmts = Matrix{AbstractStatement}(undef, length(rows), length(cols))
                stmts = [coord_to_statement[(c, r)] for r in rows, c in cols]
                convert_stmt = s -> convert(exporter, s.rhs_expr, table.sheet_name)
                stmt_strs = convert_stmt.(stmts)
                # rhs_strings = map(s -> convert(exporter, s.rhs_expr, table.sheet_name), region_statements)
                joined = if length(cols) > 1
                    join(map(v -> join(v, " "), eachrow(stmt_strs)), ";")
                else
                    join(vec(stmt_strs), ", ")
                end
                # joined = join(map(v -> join(v, " "), eachrow(stmt_strs)), ";")
                # joined = join(rhs_strings, ", ")
                rhs_str = "[$joined]"
                lhs_str = convert(exporter, lhs, table.sheet_name)
                str = "$lhs_str .= $rhs_str"
                push!(lines, indent(str, 1))
            end
        end
        # get_set_cells.(group)


        # for s in group
        #     string = export_statement(exporter, wb, s)
        #     push!(lines, indent(rstrip(string), 1))
        # end
    end
    push!(lines, "")

    # for table in tables

    # end
    # push!(lines, "")

    push!(lines, "\tTables(")
    for table in tables
        push!(lines, "\t\t" * getname(table) * ",")
    end

    push!(lines, "\t)")
    push!(lines, "end")

    join(lines, "\n")
end

function make_input_table_struct(exporter::PythonExporter, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement})
    tables = exporter.tables

    lines = Vector{String}()

    struct_names = getname.(tables)
    # var_types = repeat(["Matrix"], length(tables))
    var_types = repeat(["pd.DataFrame"], length(tables))
    struct_def = make_struct(exporter, "Tables", struct_names; var_types = var_types)
    push!(lines, struct_def)

    input_statements = get_input_statements(statements)
    # for s in input_statements
    #     if !(s isa TableStatement || s isa BroadcastedStatement)
    #         println("Input statement of unknown type: $s")
    #     end
    # end
    input_table_stmts = filter(s -> s isa TableStatement || s isa BroadcastedStatement, input_statements)
    grouped_by_set_table = group_to_dict(input_table_stmts, get_set_table)
    push!(lines, "def make_input_tables():")
    cell_to_stmt = make_cell_to_statement_dict(input_table_stmts)

    for table in tables
        top_left = CellDependency(table.sheet_name, startcol(table), startrow(table))
        bottom_right = CellDependency(table.sheet_name, endcol(table), endrow(table))
        push!(lines, "\t# Table $(table.sheet_name)!$(top_left.cell):$(bottom_right.cell)")

        push!(lines, make_dataframe_declaration(exporter, wb, table))

        if !(table in keys(grouped_by_set_table))
            continue
        end
        group = grouped_by_set_table[table]

        sort!(group, by = s -> get_set_cells(s)[1])

        num_set_cells = s -> length(s.assigned_vars)
        get_row_num = s -> rownum(s.assigned_vars[1])
        get_col_num = s -> colnum(s.assigned_vars[1])

        debug = getname(table) == "tab_oyster_Husbandry_model_HC6_HM11"
        if debug
            println("Found debug table")
            @display group
        end

        # sets_single_cell_mask = num_set_cells.(group) .== 1

        # for s in findall(.!sets_single_cell_mask)
        #     if debug
        #         @show group[s]
        #     end
        #     string = export_statement(exporter, wb, group[s])
        #     push!(lines, indent(rstrip(string), 1))
        # end
        set_cells = reduce(vcat, get_set_cells.(group))
        row_nums = rownum.(set_cells)
        col_nums = colnum.(set_cells)

        # row_nums = get_row_num.(group)[sets_single_cell_mask]
        # col_nums = get_col_num.(group)[sets_single_cell_mask]


        coords = zip(col_nums, row_nums) |> collect
        coord_to_cell = Dict(c => cell for (c, cell) in zip(coords, set_cells))
        # coord_to_statement = Dict(c => s for (c, s) in zip(coords, group[sets_single_cell_mask]))
        # if debug
        #     for (c, r) in coords
        #         if row_name(table, r - startrow(table) + 1) == "Harvest: change out filled container"
        #             statement = coord_to_statement[(c, r)]
        #             # @show r c statement
        #         end
        #     end
        # end
        regions = get_2d_regions(coords)
        for region in regions
            cols, rows = region
            region_coords = vec([(c, r) for c in cols, r in rows])

            if length(region_coords) < 1
                for c in region_coords
                    s = coord_to_statement[c]
                    string = export_statement(exporter, wb, s)
                    push!(lines, indent(rstrip(string), 1))
                end
            else
                # region_statements = map(c -> coord_to_statement[c], region_coords)

                # first_statement = region_statements[1]

                cells = [coord_to_cell[(c, r)] for r in rows, c in cols]
                first_statement = cell_to_stmt[cells[1, 1]]
                lhs = convert_to_broadcasted(first_statement.lhs_expr, length(rows) - 1, length(cols) - 1)
                if lhs.head == :broadcast_protect
                    lhs = lhs.args[1]
                end
                # stmts = Matrix{AbstractStatement}(undef, length(rows), length(cols))
                # stmts = [coord_to_statement[(c, r)] for r in rows, c in cols]

                # convert_stmt(s::TableStatement) = convert(exporter, s.rhs_expr, table.sheet_name)
                # stmt_strs = convert_stmt.(stmts)
                # joined = if length(cols) > 1 && length(rows) > 1
                #     join(map(v -> string('[', join(v, ", "), ']'), eachrow(stmt_strs)), ", ")
                # else
                #     join(vec(stmt_strs), ", ")
                # end
                # joined = join(map(v -> join(v, " "), eachrow(stmt_strs)), ";")
                # joined = join(rhs_strings, ", ")
                # rhs_str = "[$joined]"
                lhs_str = convert(exporter, lhs, table.sheet_name)
                # str = "$lhs_str = $rhs_str"
                top_left = cells[1, 1]
                bottom_right = cells[end, end]
                tbl_rows = rows .- startrow(table) .+ 1
                tbl_cols = cols .- startcol(table) .+ 1
                # row_idx = if length(rows) == 1
                #     repr(row_name(table, first(tbl_rows)))
                # else
                #     repr(row_name(table, first(tbl_rows))) * ":" * repr(row_name(table, last(tbl_rows)))
                # end
                row_idx = repr(row_name(table, first(tbl_rows))) * ":" * repr(row_name(table, last(tbl_rows)))
                # col_idx = if length(cols) == 1
                #     repr(column_name(table, first(tbl_cols)))
                # else
                #     repr(column_name(table, first(tbl_cols))) * ":" * repr(column_name(table, last(tbl_cols)))
                # end
                col_idx = repr(column_name(table, first(tbl_cols))) * ":" * repr(column_name(table, last(tbl_cols)))
                lhs_str = "$(getname(table)).loc[$row_idx, $col_idx]"

                str = "$lhs_str = xl.load_range(\"$(top_left.sheet_name)\", \"$(top_left.cell)\", \"$(bottom_right.cell)\")"
                push!(lines, indent(str, 1))
            end
        end
        # get_set_cells.(group)


        # for s in group
        #     string = export_statement(exporter, wb, s)
        #     push!(lines, indent(rstrip(string), 1))
        # end
        push!(lines, "")

    end
    push!(lines, "")

    push!(lines, "\treturn Tables(")
    for table in tables
        push!(lines, "\t\t" * getname(table) * ",")
    end

    push!(lines, "\t)\n")
    # push!(lines, "end")

    join(lines, "\n")
end

function write_file(exporter::PythonExporter, file_name::AbstractString, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement})
    output_file = open(file_name, "w")

    write(output_file, "from dataclasses import dataclass\n")
    write(output_file, "import datetime\n")
    write(output_file, "from typing import Any\n")
    write(output_file, "import numpy as np\n")
    write(output_file, "import pandas as pd\n")
    write(output_file, "import python_funcs as xl\n\n")
    # write(output_file, "using DataFrames\n\n")
    # write(output_file, "using Dates\n\n")
    # write(
    #     output_file,
    #     """
    #     function if_multiple(dividend, divisor, value)
    #     xl_compare(xl_mod(dividend, divisor), 0) ? value : 0.0
    #     end

    #     """,
    # )
    for func_stmt in filter(s -> s isa FunctionStatement, statements)
        write(output_file, get_function_string(exporter, wb, func_stmt), "\n")
    end
    for func_stmt in filter(s -> s isa GenericFunctionStatement, statements)
        write(output_file, get_function_string(exporter, wb, func_stmt), "\n")
    end
    for group_stmt in filter(s -> s isa GroupedStatement, statements)
        func_str = get_function_string(exporter, wb, group_stmt)
        if !isnothing(func_str)
            write(output_file, func_str, "\n")
        end
    end

    for out_stmt in filter(s -> s isa OutputStatement, statements)
        write(output_file, make_outupt_struct(exporter, wb, out_stmt))
    end

    input_struct_str, input_struct_vars = make_input_struct(exporter, statements)
    write(output_file, input_struct_str)

    write(output_file, make_input_table_struct(exporter, wb, statements), "\n")

    # write(output_file, "def calculate(inputs:Inputs, tables:Tables)\n")
    write(output_file, "inputs = Inputs()\n")
    write(output_file, "tables = make_input_tables()\n\n")

    for table in exporter.tables
        lhs = getname(table)
        write(output_file, "$lhs = tables.$lhs\n")
    end

    starting_var_names = copy(exporter.var_names)

    reverse_var_names = Dict(values(exporter.var_names) .=> keys(exporter.var_names))

    for var in input_struct_vars
        cell_ref = reverse_var_names[var]

        # @info "making var have input before it" cell_ref var

        exporter.var_names[cell_ref] = "inputs.$var"
        # println("Var name for $cell_ref is $(exporter.var_names[cell_ref])")
        # @show cell_ref exporter.var_names[cell_ref]
    end


    export_statements(output_file, exporter, wb, statements)

    filter!(p -> false, exporter.var_names)
    for (k, v) in starting_var_names
        exporter.var_names[k] = v
    end

    # write(output_file, "end\n")
    # write(output_file, "\ncalculate()")

    # run_str = """function run_crest_solar()
    #     inputs = Inputs()
    #     tables = make_input_tables()
    #     calculate(inputs, tables)
    # end"""
    # write(output_file, "\n", run_str, "\n")


    close(output_file)
end

function write_file(exporter::JuliaExporter, file_name::AbstractString, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement})
    output_file = open(file_name, "w")

    write(output_file, "using XLConvert\n")
    write(output_file, "using DataFrames\n\n")
    write(output_file, "using Dates\n\n")
    write(
        output_file,
        """
        function if_multiple(dividend, divisor, value)
        xl_compare(xl_mod(dividend, divisor), 0) ? value : 0.0
        end

        """,
    )
    for func_stmt in filter(s -> s isa FunctionStatement, statements)
        write(output_file, get_function_string(exporter, wb, func_stmt), "\n")
    end
    for group_stmt in filter(s -> s isa GroupedStatement, statements)
        func_str = get_function_string(exporter, wb, group_stmt)
        if !isnothing(func_str)
            write(output_file, func_str, "\n")
        end
    end

    for out_stmt in filter(s -> s isa OutputStatement, statements)
        write(output_file, make_outupt_struct(exporter, wb, out_stmt))
    end

    input_struct_str, input_struct_vars = make_input_struct(exporter, statements)
    write(output_file, input_struct_str)

    write(output_file, make_input_table_struct(exporter, wb, statements), "\n")

    write(output_file, "function calculate(inputs::Inputs, tables::Tables)\n")

    for table in exporter.tables
        lhs = getname(table)
        write(output_file, "$lhs = tables.$lhs\n")
    end

    starting_var_names = copy(exporter.var_names)

    reverse_var_names = Dict(values(exporter.var_names) .=> keys(exporter.var_names))

    for var in input_struct_vars
        cell_ref = reverse_var_names[var]

        # @info "making var have input before it" cell_ref var

        exporter.var_names[cell_ref] = "inputs.$var"
        # println("Var name for $cell_ref is $(exporter.var_names[cell_ref])")
        # @show cell_ref exporter.var_names[cell_ref]
    end


    export_statements(output_file, exporter, wb, statements)

    filter!(p -> false, exporter.var_names)
    for (k, v) in starting_var_names
        exporter.var_names[k] = v
    end

    write(output_file, "end\n")
    # write(output_file, "\ncalculate()")

    run_str = """function run_crest_solar()
        inputs = Inputs()
        tables = make_input_tables()
        calculate(inputs, tables)
    end"""
    write(output_file, "\n", run_str, "\n")


    close(output_file)
end

function get_graphviz_stmt_name(exporter::JuliaExporter, wb::ExcelWorkbook, statement::StandardStatement)
    exporter.var_names[statement.assigned_var]
end
function get_graphviz_stmt_name(exporter::JuliaExporter, wb::ExcelWorkbook, statement::TableStatement)
    cell_ref = statement.assigned_vars[1]
    sheet = cell_ref.sheet_name

    convert(exporter, statement.lhs_expr, sheet)
end
function get_graphviz_stmt_name(exporter::JuliaExporter, wb::ExcelWorkbook, statement::GroupedStatement)
    set_cells = get_set_cells(statement)
    names = exporter.var_names
    "$(names[set_cells[begin]])...$(names[set_cells[end]])"
end
function get_graphviz_stmt_name(exporter::JuliaExporter, wb::ExcelWorkbook, statement::FunctionStatement)
    "calculate_$(exporter.var_names[statement.assigned_var])"
end

function write_graphviz(exporter::JuliaExporter, file_name::AbstractString, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement})

    colors = [
        "dimgray",
        "maroon",
        "green",
        "navy",
        "goldenrod",
        "mediumaquamarine",
        "red",
        "yellow",
        "lime",
        "mediumorchid",
        "mediumspringgreen",
        "blue",
        "coral",
        "fuchsia",
        "dodgerblue",
        "plum",
        "deeppink",
        "lightskyblue",
        "bisque",
    ]
    color_map = Dict(sheet => colors[i] for (i, sheet) in enumerate(XLSX.sheetnames(wb.xf)))
    stmt_graph = make_statement_graph(statements)

    begin
        lines = ["digraph {rankdir=LR;graph [size=\"60,60!\"];node[style=filled];"]
        for edge ∈ edges(stmt_graph)
            src_node = Graphs.src(edge)
            dst_node = Graphs.dst(edge)
            push!(lines, "$dst_node -> $src_node")
        end
        push!(lines, "")

        # @show statements

        for node in vertices(stmt_graph)
            stmt = statements[node]
            if isempty(get_set_cells(stmt))
                continue
            end
            cell_ref = get_set_cells(stmt)[begin]
            # node_str = to_string(all_referenced_nodes[node])
            node_str = replace(get_graphviz_stmt_name(exporter, wb, stmt), "\"" => "\\\"")
            push!(lines, "$node [label=\"$node_str\", fillcolor=$(color_map[cell_ref.sheet_name])]")
        end
        push!(lines, "}")
        write(file_name, join(lines, "\n"))


    end
end

is_data_copy_statement(stmt) = false
function is_data_copy_statement(stmt::StandardStatement)
    expr = stmt.rhs_expr
    return expr.head == :cell_ref
end
function is_data_copy_statement(stmt::TableStatement)
    length(stmt.assigned_vars) != 1 && return false

    @match stmt.rhs_expr begin
        ExcelExpr(:cell_ref, _) => true
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => length(row_idx) == 1 && length(col_idx) == 1
        _ => false
    end
end

function get_statement_stats(statements::AbstractArray{AbstractStatement})
    statement_types = typeof.(statements)

    for statement_type in unique(statement_types)
        num_occurances = sum(statement_types .== statement_type)
        println("$num_occurances occurances of $statement_type")
    end

    stmt_graph = make_statement_graph(statements)
    stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    stmt_topo_levels_bottom_up = get_topo_levels_bottom_up(stmt_graph)
    statements_bottom_up = group_to_dict(1:length(statements), s -> stmt_topo_levels_bottom_up[s])

    num_levels = maximum(values(stmt_topo_levels))
    println("Number of Levels: $num_levels")

    println("Number of value inputs: $(length(statements_bottom_up[0]))")

    is_base_input = s -> stmt_topo_levels_bottom_up[s] == 0
    non_input_statements = statements[filter(!is_base_input, 1:length(statements))]

    copy_statements = filter(is_data_copy_statement, non_input_statements)
    num_copy_statements = length(copy_statements)

    println("There are $num_copy_statements statements that just copy a value")
    # for s in copy_statements
    #     println(s)
    # end

end

function get_statement_parents_idx(stmt_graph, stmt_idx::Int)
    findall(bfs_parents(stmt_graph, stmt_idx, dir = :out) .> 0)
end

function get_statement_child_idx(stmt_graph, stmt_idx::Int)
    findall(bfs_parents(stmt_graph, stmt_idx, dir = :in) .> 0)
end

function statement_set_string(stmt_idx)
    stmt = grouped_statements[stmt_idx]
    set_cells = get_set_cells(stmt)
    set_vars = map(c -> new_names_map[c], set_cells)
    if length(set_vars) == 1
        "$(set_vars[1])"
    else
        "$(set_vars[1]) ... $(set_vars[end])"
    end
end


function indent(str::AbstractString, level::Int)
    lines = split(str, "\n")
    join(["\t"^level * l for l in lines], "\n")
end

function look_for_function(statements, output_stmt_idx, stmt_graph)
    visited = Set{Int64}([output_stmt_idx])

    inputs = []
    intermediate = []

    queue = Vector{Int64}([output_stmt_idx])

    while !isempty(queue)
        n = popfirst!(queue)

        dep_stmts = outneighbors(stmt_graph, n)

        for dep in dep_stmts
            if dep in visited
                continue
            end

            push!(visited, dep)
            dep_usages = inneighbors(stmt_graph, dep)
            is_only_used_in_function = all(((u in intermediate) || u == output_stmt_idx) for u in dep_usages)
            is_not_base_input = length(outneighbors(stmt_graph, dep)) > 0
            is_not_function_stmt = !(statements[dep] isa FunctionStatement)
            if is_only_used_in_function && is_not_base_input && is_not_function_stmt
                push!(intermediate, dep)
                push!(queue, dep)
            else
                if dep != output_stmt_idx
                    push!(inputs, dep)
                end
            end

        end
    end

    inputs, intermediate
end

function get_func_inputs(input_statements)
    required_resources = Set()
    for stmt in input_statements
        if stmt isa TableStatement
            push!(required_resources, get_set_table(stmt))
        else
            for c in get_set_cells(stmt)
                push!(required_resources, c)
            end
        end
    end

    required_resources
end

function look_for_functions(statements; min_intermediates = 3, max_inputs = 30)
    new_statements = copy(statements)

    stmt_graph = make_statement_graph(statements)
    # stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    stmt_topo_levels_bottom_up = get_topo_levels_bottom_up(stmt_graph)

    grouped_by_level = group_to_dict(1:length(statements), s -> stmt_topo_levels_bottom_up[s])
    max_level = maximum(keys(grouped_by_level))


    captured_stmts = Set{Int64}()

    # for level in 1:max_level - 1
    for level in reverse(1:(max_level-1))
        level_stmts = grouped_by_level[level]
        # println("Level: $level")

        for s in level_stmts
            if !(statements[s] isa StandardStatement)
                continue
            end
            if !(statements[s] in new_statements)
                continue
            end

            inputs, intermediate = look_for_function(statements, s, stmt_graph)
            input_statements = statements[inputs]
            # if length(intermediate) >= min_intermediates && length(inputs) < max_inputs
            if length(intermediate) >= min_intermediates && length(get_func_inputs(input_statements)) < max_inputs
                if any((i in captured_stmts) for i in intermediate)
                    # println("Intermediate value is the result of a function!")
                    continue
                end

                # println("Output $s: $(statement_set_string(s))")
                # println("\tNum inputs = $(length(inputs))")
                # # for input in inputs
                # #     println("\t$(statement_set_string(input))")
                # #     # println("\t$input: $(get_set_cells(statements[input]))")
                # # end
                # println("\tNum intermediates = $(length(intermediate))")
                # for s in intermediate
                #     println("\t\t$(statement_set_string(s))")
                #     # println("\t$s: $(get_set_cells(statements[s]))")

                #     # dep_usages = inneighbors(stmt_graph, s)
                #     # println("\tusages: $dep_usages")
                # end


                push!(captured_stmts, s)

                statement = statements[s]
                # function_statements = [statements[reverse(intermediate)]..., statement]
                function_statements = Vector{AbstractStatement}(undef, length(intermediate) + 1)
                function_statements[begin:(end-1)] .= statements[reverse(intermediate)]
                function_statements[end] = statement
                # push!(function_statements, statement)
                func_statement = FunctionStatement(statement.assigned_var, statements[inputs], function_statements)

                # statement_group = [statement, statements[intermediate]...]
                filter!(s -> !(s in function_statements), new_statements)
                push!(new_statements, func_statement)
            end
        end

    end
    println("Made $(length(captured_stmts)) functions!")
    new_statements
end

function add_functions(statements; min_intermediates = 3, max_inputs = 30)
    start_len = length(statements)

    new_statements = look_for_functions(statements; min_intermediates = min_intermediates, max_inputs = max_inputs)
    while (start_len != length(new_statements))
        start_len = length(new_statements)
        new_statements = look_for_functions(new_statements; min_intermediates = min_intermediates, max_inputs = max_inputs)
    end
    new_statements
end
