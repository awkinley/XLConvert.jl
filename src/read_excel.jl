# include("excel_expr.jl")
# include("formula_parser.jl")
# include("excel_formula.jl")
# include("type_infer.jl")
# include("cell_dependency.jl")
# include("excel_table.jl")
# include("excel_workbook.jl")
# include("export_julia.jl")
# include("statement.jl")
# include("workbook_subset.jl")
# include("variable_naming.jl")

# include("transforms/if_multiple.jl")
# include("transforms/table_broadcast.jl")
# include("transforms/group_statements.jl")


# function XLSX.getdata(xl::XLSX.XLSXFile, s::AbstractString)

#     if XLSX.is_valid_sheet_cellname(s)
#         return XLSX.getdata(xl, XLSX.SheetCellRef(s))
#     elseif XLSX.is_valid_sheet_cellrange(s)
#         return XLSX.getdata(xl, XLSX.SheetCellRange(s))
#     elseif XLSX.is_valid_sheet_column_range(s)
#         return XLSX.getdata(xl, XLSX.SheetColumnRange(s))
#     elseif XLSX.is_workbook_defined_name(xl, s)

#         v = XLSX.get_defined_name_value(xl.workbook, s)
#         if XLSX.is_defined_name_value_a_constant(v)
#             return v
#         elseif XLSX.is_defined_name_value_a_reference(v)
#             if typeof(v) == XLSX.SheetCellRef
#                 v2 = XLSX.SheetCellRef(replace(v.sheet, "'" => ""), v.cellref)
#             else
#                 v2 = XLSX.SheetCellRange(replace(v.sheet, "'" => ""), v.rng)
#             end
#             try
#                 return XLSX.getdata(xl, v)
#             catch
#             # println("Removing single quotes and retrying")
#             finally
#                 return XLSX.getdata(xl, v2)
#             end
#         else
#             error("Unexpected defined name value: $v.")
#         end
#     end

#     error("$s is not a valid sheetname or cell/range reference.")
# end

# file = "excel_to_code/ZM - Kelp for kelp TEA_LCA v0.56.xlsm.xlsx";
# file = "excel_to_code/UMaine Monhegan Farm Verification v0.09.xlsm";

# xf = XLSX.readxlsx(file)

# assumptions = xf["1. Assumptions"]


# kelp_cost_cell = XLSX.getcell(assumptions, "C4")
# assumptions[1:5, 2:5]

function get_all_cells(sheet)
    cells = Vector{XLSX.Cell}()
    for row in XLSX.eachrow(sheet)
        append!(cells, values(row.rowcells))
    end
    # XLSX.getcellrange(sheet, XLSX.get_dimension(sheet))
    cells
end

# all_cells = get_all_cells(assumptions)

# has_formula(cell) = false
# has_formula(cell::XLSX.Cell) = !isempty(cell.formula)

# filter(!isempty, all_cells)
# all_formulas = filter(has_formula, all_cells)


function build_referenced_formula_dict(formulas)
    out = Dict([f.formula.id => f for f in filter(f -> f.formula isa XLSX.ReferencedFormula, formulas)])
    out
end

# excel_functions = readlines("./excel_to_code/ExcelBuiltinFunctionList.txt")

# function parse_to_json(formula_string)
#     @time command = `xlparser $(formula_string)`
#     @time output_json = read(command, String)
#     @time JSON.parse(output_json)
# end

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

# every_formula = reduce(vcat, [get_all_formulas(xf, sheet_name) for sheet_name in XLSX.sheetnames(xf)])

# length(every_formula)
# @time sum(can_parse_formula.(every_formula))
# @profview sum(can_parse_formula.(every_formula))
# @profile sum(can_parse_formula.(every_formula))
# failing_formulas = every_formula[findall((f -> !can_parse_formula(f)).(every_formula))]
# failing_formulas[1].formula.formula

function get_cell_value(ctx::XLSX.Worksheet, cell)
    v = ctx[string(replace(cell, '$' => ""))]
    # ismissing(v) ? 0 : v
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

# sheet, xl_ctx, not_parse = test_sheet_eval(xf, 7);
# sheet_cells = get_all_cells(sheet)
# sheet_formulas = filter(has_formula, sheet_cells)
# reference_formula_dict = build_referenced_formula_dict(sheet_formulas)

# cell = not_parse[1]
# base_formula_cell = reference_formula_dict[cell.formula.id]
# start_row = base_formula_cell.ref.row_number
# start_col = base_formula_cell.ref.column_number
# end_row = cell.ref.row_number
# end_col = cell.ref.column_number
# delta_x = end_col - start_col
# delta_y = end_row - start_row

# true_value = sheet[cell.ref]
# xl_expr = toexpr(parse_formula(base_formula_cell.formula.formula))
# expression = offset(toexpr(parse_formula(base_formula_cell.formula.formula)), delta_y, delta_x)
# calc_value = eval(expression, xl_ctx)
# xl_compare(true_value, calc_value)

# test = not_parse[1]
# sheet[test.ref]
# toexpr(parse_formula(test.formula.formula))
# parse_formula(test.formula.formula) |> toexpr
# # toexpr(parse_formula("1 - (2 * 5)*6"))
# eval(toexpr(parse_formula(test.formula.formula)), xl_ctx)
# sheet[test.ref]
# xl_compare(sheet[test.ref], exec(toexpr(parse_formula(test.formula.formula)), xl_ctx))


function get_sheet_exprs(xl, sheet_name)
    # xl_ctx = ExcelContext(sheet_name,
    #     Dict([(name => xl[name]) for name in XLSX.sheetnames(xl)]),
    #     Dict((p[1] => toexpr(parse_formula(repr(p[2])))) for p in XLSX.get_workbook(xl).workbook_names)
    # )

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

function get_expr_dependencies(expr, key_values::Dict)
    return []
end

function get_expr_dependencies(expr::FlatExpr, key_values::Dict)
    deps = Vector{CellDependency}()
    handled = Set{Int}()
    for (i, part) in enumerate(expr.parts)
        i in handled && continue

        @match part begin
            ExcelExpr(:cell_ref, [cell, sheet]) => push!(deps, CellDependency(sheet, cell))
            # ExcelExpr(:sheet_ref, (sheet_name, ref)) => get_expr_dependencies(ref, key_values)
            ExcelExpr(:named_range, [name]) => get_expr_dependencies(key_values[name], key_values)
            ExcelExpr(:range, [FlatIdx(lhs_i), FlatIdx(rhs_i)]) => begin
                lhs_expr = expr.parts[lhs_i]
                rhs_expr = expr.parts[rhs_i]
                if !((lhs_expr.head == :cell_ref) && (rhs_expr.head == :cell_ref))
                    throw("Don't know how to get dependencies for $(expr)")
                end
                sheet = lhs_expr.args[2]
                if (sheet != rhs_expr.args[2])
                    throw("Don't know how to get dependencies for $(expr)")
                end

                lhs = lhs_expr.args[1]
                rhs = rhs_expr.args[1]

                start_col, start_row = parse_cell(lhs)
                end_col, end_row = parse_cell(rhs)

                @assert end_row >= start_row
                @assert end_col >= start_col

                for col in start_col:end_col, r in start_row:end_row
                    push!(deps, CellDependency(sheet, index_to_cellname(col, r)))
                end

                push!(handled, lhs_i)
                push!(handled, rhs_i)
            end
            _ => continue
        end
    end

    deps
end


function get_expr_dependencies(expr::ExcelExpr, key_values::Dict)::Vector{CellDependency}
    @match expr begin
        ExcelExpr(:cell_ref, [cell, sheet]) => [CellDependency(sheet, cell)]
        # ExcelExpr(:sheet_ref, (sheet_name, ref)) => get_expr_dependencies(ref, key_values)
        ExcelExpr(:named_range, [name]) => begin
            value = key_values[name]
            if value isa ExcelExpr
                get_expr_dependencies(value, key_values)
            else
                Vector{CellDependency}()
            end
        end
        ExcelExpr(:range, [ExcelExpr(:cell_ref, [lhs, sheet]), ExcelExpr(:cell_ref, [rhs, sheet])]) => begin
            start_col, start_row = parse_cell(lhs)
            end_col, end_row = parse_cell(rhs)

            @assert end_row >= start_row
            @assert end_col >= start_col
            [CellDependency(sheet, index_to_cellname(c, r)) for c ∈ start_col:end_col for r ∈ start_row:end_row]
        end
        # TODO: Handle?
        ExcelExpr(:range, [lhs, rhs]) => throw("Don't know how to get dependencies for $(expr)")
        ExcelExpr(op, args) => begin
            res = Vector{CellDependency}()
            for a in args
                if a isa ExcelExpr
                    append!(res, get_expr_dependencies(a, key_values))
                end
            end
            res
        end
    end
end

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

function get_topo_levels_bottom_up(graph::SimpleDiGraph)
    topo_sorted = topological_sort(reverse(graph))

    topo_levels = Dict{Int64, Int64}()
    for node in topo_sorted
        dependencies = outneighbors(graph, node)
        topo_levels[node] = maximum(k -> topo_levels[k] + 1, dependencies; init = 0)
    end

    topo_levels
end

function get_topo_levels_top_down(graph::SimpleDiGraph)
    topo_sorted = topological_sort(graph)

    topo_levels = Dict{Int64, Int64}()
    for node in topo_sorted
        # We have to filter in this case, and not in the bottom up case
        # because it's not possible for a cell to depend on a value not in used_nodes
        # but it is possible for a cell not in used_nodes to depend on one that is
        dependents = inneighbors(graph, node)
        # topo_levels[node] = minimum(map(k -> topo_levels[k] - 1, dependents); init=0)
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


insert_table_refs(expr, tables) = expr

function cell_fixed_row(cell::String)
    cell[1] == '$'
end

function cell_fixed_col(cell::String)
    cell[1] == '$'

    first_num = findfirst(isdigit, cell)

    cell[first_num-1] == '$'
end

function insert_table_refs(expr::FlatExpr, tables)
    # @info "In primary insert_table_refs" expr

    new_expr = copy(expr)

    deps = Vector{CellDependency}()
    handled = Set{Int}()
    for (i, part) in enumerate(new_expr.parts)
        i in handled && continue
        @match part begin
            ExcelExpr(:cell_ref, [cell, sheet::String]) => begin
                c, r = parse_cell(cell)
                for table in tables
                    if (sheet == table.sheet_name
                        &&
                        (r, c) in table)
                        row_idx = r - startrow(table) + 1
                        col_idx = c - startcol(table) + 1
                        fixed_row = cell_fixed_row(cell)
                        fixed_col = cell_fixed_col(cell)
                        # fixed_row = offset(expr, 1, 0) == expr
                        # fixed_col = offset(expr, 0, 1) == expr

                        new_expr.parts[i] = ExcelExpr(:table_ref, Any[table, row_idx, col_idx, (fixed_row, fixed_row), (fixed_col, fixed_col)])
                    end
                end
            end
            # ExcelExpr(:sheet_ref, (sheet_name, ref)) => ExcelExpr(:sheet_ref, sheet_name, insert_table_refs(ref, tables))
            # ExcelExpr(:named_range, (name,)) => convert_to_broadcasted(get_key_value(ctx, name), ctx, row_offset, col_offset)

            ExcelExpr(:range, [FlatIdx(lhs_i), FlatIdx(rhs_i)]) => begin
                lhs_expr = new_expr.parts[lhs_i]
                rhs_expr = new_expr.parts[rhs_i]
                if !((lhs_expr.head == :cell_ref) && (rhs_expr.head == :cell_ref))
                    throw("Don't know how to get dependencies for $(expr)")
                end
                sheet = lhs_expr.args[2]
                if (sheet != rhs_expr.args[2])
                    throw("Don't know how to get dependencies for $(expr)")
                end

                lhs = lhs_expr.args[1]
                rhs = rhs_expr.args[1]

                start_c, start_r = parse_cell(lhs)
                stop_c, stop_r = parse_cell(rhs)

                for table in tables
                    if (sheet == table.sheet_name
                        && (start_r, start_c) in table
                        && (stop_r, stop_c) in table)
                        col_idx = (start_c:stop_c) .- startcol(table) .+ 1
                        row_start_idx = start_r - startrow(table) + 1
                        row_stop_idx = stop_r - startrow(table) + 1

                        # is_fixed_row(expr) = offset(expr, 1, 0) == expr
                        # is_fixed_col(expr) = offset(expr, 0, 1) == expr
                        fixed_row = (cell_fixed_row(lhs), cell_fixed_row(rhs))
                        fixed_col = (cell_fixed_col(lhs), cell_fixed_col(rhs))
                        # fixed_row = tuple(is_fixed_row.(part.args)...)
                        # fixed_col = tuple(is_fixed_col.(part.args)...)

                        new_expr.parts[i] = ExcelExpr(:table_ref, Any[table, row_start_idx:row_stop_idx, col_idx, fixed_row, fixed_col])

                        push!(handled, lhs_i)
                        push!(handled, rhs_i)

                        continue
                    end
                end
            end
            _ => continue
        end
    end
    new_expr
end
function insert_table_refs(expr::ExcelExpr, tables)
    # @info "In primary insert_table_refs" expr
    @match expr begin
        ExcelExpr(:cell_ref, [cell, sheet]) => begin
            c, r = parse_cell(cell)
            for table in tables
                if (sheet == table.sheet_name
                    && (r, c) in table)
                    # && c >= startcol(table)
                    # && c <= endcol(table)
                    # && r >= startrow(table)
                    # && r <= endrow(table))
                    row_idx = r - startrow(table) + 1
                    col_idx = c - startcol(table) + 1
                    fixed_row = offset(expr, 1, 0) == expr
                    fixed_col = offset(expr, 0, 1) == expr

                    return ExcelExpr(:table_ref, table, row_idx, col_idx, (fixed_row, fixed_row), (fixed_col, fixed_col))
                end
            end

            expr
            # range_start = CellDependency(ctx.current_sheet, cell)
            # range_stop = offset(range_start, row_offset, col_offset)
            # # offset 
            # # [CellDependency(ctx.current_sheet, cell)]
            # ExcelExpr(:range, (ExcelExpr(:cell_ref, (range_start.cell,)), ExcelExpr(:cell_ref, (range_stop.cell,))))
        end
        # ExcelExpr(:sheet_ref, (sheet_name, ref)) => ExcelExpr(:sheet_ref, sheet_name, insert_table_refs(ref, tables))
        # ExcelExpr(:named_range, (name,)) => convert_to_broadcasted(get_key_value(ctx, name), ctx, row_offset, col_offset)
        ExcelExpr(:range, [ExcelExpr(:cell_ref, [lhs, sheet]), ExcelExpr(:cell_ref, [rhs, sheet])]) => begin
            start_c, start_r = parse_cell(lhs)
            stop_c, stop_r = parse_cell(rhs)
            # @info "found range" start_c start_r stop_c stop_r

            for table in tables
                # @info "Checking table" ctx.current_sheet == table.sheet_name start_c >= startcol(table) stop_c <= endcol(table) start_r >= startrow(table) stop_r <= endrow(table)
                if (sheet == table.sheet_name
                    && (start_r, start_c) in table
                    && (stop_r, stop_c) in table)
                    # && start_c >= startcol(table)
                    # && stop_c <= endcol(table)
                    # && start_r >= startrow(table)
                    # && stop_r <= endrow(table))
                    # @info "Found a matching table"
                    # col_names = [column_name(table, c - startcol(table) + 1) for c in start_c:stop_c]
                    # col_idx = [c - startcol(table) + 1 for c in start_c:stop_c]
                    col_idx = (start_c:stop_c) .- startcol(table) .+ 1
                    row_start_idx = start_r - startrow(table) + 1
                    row_stop_idx = stop_r - startrow(table) + 1

                    is_fixed_row(expr) = offset(expr, 1, 0) == expr
                    is_fixed_col(expr) = offset(expr, 0, 1) == expr
                    fixed_row = tuple(is_fixed_row.(expr.args)...)
                    fixed_col = tuple(is_fixed_col.(expr.args)...)
                    # fixed_row = offset(expr.args[1], 1, 0) == expr.args[1]
                    # fixed_col = offset(expr, 0, 1) == expr

                    return ExcelExpr(:table_ref, table, row_start_idx:row_stop_idx, col_idx, fixed_row, fixed_col)
                end
            end
            expr
        end
        ExcelExpr(:range, [lhs, rhs]) => throw("Don't know how to get dependencies for $(expr)")
        ExcelExpr(op, args) => begin
            new_args = similar(args)
            map!(a -> insert_table_refs(a, tables), new_args, args)
            # mapped = map(a -> insert_table_refs(a, tables), args)
            # ExcelExpr(op, mapped...)
            ExcelExpr(op, new_args)
            # reduce(vcat, mapped)
        end
    end
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
    c_start, r_start = parse_cell(table.top_left)
    c_end, r_end = parse_cell(table.bottom_right)
    for c in c_start:c_end, r in r_start:r_end
        cell = CellDependency(table.sheet_name, index_to_cellname(c, r))
        if cell in keys(cell_dependencies)
            row_idx = r - startrow(table) + 1
            col_name = column_name(table, c - startcol(table) + 1)

            var_name = "$(getname(table))[$(repr(row_idx)), $(repr(col_name))]"
            name_map[cell] = var_name

        end
    end
    # for cell in cell_dependencies
    #     c, r = parse_cell(cell.cell)
    #     if (cell.sheet_name == table.sheet_name && (r, c) in table)
    #         row_idx = r - startrow(table) + 1
    #         # col_idx = c - startcol(table) + 1
    #         # var_name = "$(getname(table))[$(repr(row_idx)), $(repr(col_idx))]"
    #         col_name = column_name(table, c - startcol(table) + 1)

    #         var_name = "$(getname(table))[$(repr(row_idx)), $(repr(col_name))]"
    #         # println("cell = $(repr(cell)) $var_name")
    #         name_map[cell] = var_name
    #     end
    # end
end

# contexts = Dict((sheet_name => make_ctx(sheet_name, xf)) for sheet_name in XLSX.sheetnames(xf))



function DefTable(xf::XLSX.XLSXFile, sheet_name, table_name, top_left, bottom_right, column_names_range, row_names_range)
    column_names = missing
    if !isempty(column_names_range)
        column_names = xf[sheet_name][column_names_range]
        for i in eachindex(column_names)
            if ismissing(column_names[i])
                column_names[i] = "missing_$(i)"
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


iscontiguous(values::Vector{T}) where {T <: Integer} = values == minimum(values):maximum(values)
function get_contiguous_runs(values::Vector{T}) where {T <: Integer}
    if isempty(values)
        return values
    end
    if length(values) == 1
        return [values]
    end

    result = Vector{Vector{T}}()
    current = Vector{T}()
    last_value = values[1]
    push!(current, last_value)

    for next_value in values[2:end]
        if next_value == last_value + 1
            push!(current, next_value)
        else
            push!(result, current)
            current = [next_value]
        end
        last_value = next_value
    end
    push!(result, current)

    result
end

contains_if(expr) = false
function contains_if(expr::ExcelExpr)
    @match expr begin
        ExcelExpr(:call, ["IF", args...]) => begin
            true
        end
        ExcelExpr(op, args) => begin
            any(contains_if(a) for a in args)
        end
    end
end

convert_to_broadcasted(expr, row_offset, col_offset) = expr

function convert_to_broadcasted(expr::ExcelExpr, row_offset, col_offset)
    @match expr begin
        ExcelExpr(:cell_ref, [cell, sheet]) => begin
            range_start = CellDependency(sheet, cell)
            stop_cell = offset(expr, row_offset, col_offset).args[1]
            range_stop = CellDependency(sheet, stop_cell)
            # range_stop = offset(range_start, row_offset, col_offset)
            # offset 
            # [CellDependency(ctx.current_sheet, cell)]
            if range_start != range_stop
                ExcelExpr(:range, ExcelExpr(:cell_ref, range_start.cell, sheet), ExcelExpr(:cell_ref, range_stop.cell, sheet))
            else
                expr
            end
        end
        # ExcelExpr(:sheet_ref, (sheet_name, ref)) => ExcelExpr(:sheet_ref, sheet_name, convert_to_broadcasted(ref, sheet_name, row_offset, col_offset))
        # ExcelExpr(:named_range, (name,)) => convert_to_broadcasted(get_key_value(ctx, name), ctx, row_offset, col_offset)
        ExcelExpr(:range, [ExcelExpr(:cell_ref, [lhs, sheet]), ExcelExpr(:cell_ref, [rhs, sheet])]) => begin
            function is_fixed(cell)
                offset_cell_parse_rgx = r"([$]?[A-Z]+)([$]?[0-9]+)"
                cell_match = match(offset_cell_parse_rgx, cell)
                @assert cell_match.match == cell "Cell didn't parse properly"
                cell_match[1][1] == '$' && cell_match[2][1] == '$'
            end
            if is_fixed(lhs) && is_fixed(rhs)
                # I think this isn't strictly correct, but it might be good enough
                mapped = map(a -> convert_to_broadcasted(a, row_offset, col_offset), expr.args)
                sub_expr = ExcelExpr(expr.head, mapped...)

                ExcelExpr(:broadcast_protect, sub_expr)
            else
                throw("Converting ranged to broadcasted is complicated $(expr)")
            end
        end

        ExcelExpr(:table_ref, [table, row_idx, col_idx, fixed_row, fixed_col]) => begin
            if fixed_row == (true, true) && fixed_col == (true, true)
                return ExcelExpr(:broadcast_protect, expr)
            end

            row_idx = @match fixed_row begin
                (true, true) => row_idx
                (false, false) => begin
                    if length(row_idx) == 1
                        row_idx[1]:(row_idx[1]+row_offset)
                    else
                        throw("Broadcasting table row range ref is complicated")
                    end
                end
                (true, false) => throw("Broadcasting partially fixed table refs is complicated")
                (false, true) => throw("Broadcasting partially fixed table refs is complicated")
            end
            col_idx = @match fixed_col begin
                (true, true) => col_idx
                (false, false) => begin
                    if length(col_idx) == 1
                        col_idx[1]:(col_idx[1]+col_offset)
                    else
                        throw("Broadcasting table col range ref is complicated")
                    end
                end
                (true, false) => throw("Broadcasting partially fixed table refs is complicated")
                (false, true) => throw("Broadcasting partially fixed table refs is complicated")
            end

            ExcelExpr(:table_ref, table, row_idx, col_idx, fixed_row, fixed_col)
        end

        # ExcelExpr(:call, ("IF", args...)) => begin
        #     throw("Broadcasting if doesn't work!")
        # end
        ExcelExpr(:range, [lhs, rhs]) => throw("Don't know how to get dependencies for $(expr)")
        ExcelExpr(op, args) => begin
            mapped = map(a -> convert_to_broadcasted(a, row_offset, col_offset), args)
            ExcelExpr(op, mapped...)
            # reduce(vcat, mapped)
        end
    end
end

function functionalize!(expr, previous_params::Vector{ExcelExpr})
    expr
end

function functionalize_at!(exprs::Vector{Any}, index::Int, previous_params::Vector{ExcelExpr})
    expr = exprs[index]
    if expr isa ExcelExpr
        if expr.head in (:cell_ref, :sheet_ref, :named_range, :range, :table_ref)
            push!(previous_params, expr)
            exprs[index] = ExcelExpr(:func_param, length(previous_params))
        else
            for i in eachindex(expr.args)
                functionalize_at!(expr.args, i, previous_params)
            end

        end
    end
end

function functionalize!(expr::ExcelExpr, previous_params::Vector{ExcelExpr})
    if expr.head in (:cell_ref, :sheet_ref, :named_range, :range, :table_ref)
        push!(previous_params, expr)
        return ExcelExpr(:func_param, length(previous_params))
    end

    for i in eachindex(expr.args)
        functionalize_at!(expr.args, i, previous_params)
    end

    expr

    # @match expr begin
    #     ExcelExpr(:cell_ref, [cell, sheet]) => begin
    #         push!(previous_params, expr)
    #         ExcelExpr(:func_param, length(previous_params))
    #     end
    #     ExcelExpr(:sheet_ref, [sheet_name, ref]) => begin
    #         push!(previous_params, expr)
    #         ExcelExpr(:func_param, length(previous_params))
    #     end
    #     ExcelExpr(:named_range, [name,]) => begin 
    #         push!(previous_params, expr)
    #         ExcelExpr(:func_param, length(previous_params))
    #     end
    #     ExcelExpr(:range, args) => begin
    #         push!(previous_params, expr)
    #         ExcelExpr(:func_param, length(previous_params))
    #     end
    #     ExcelExpr(:table_ref, args) => begin
    #         push!(previous_params, expr)
    #         ExcelExpr(:func_param, length(previous_params))
    #     end

    #     ExcelExpr(op, args) => begin
    #         mapped_exprs = similar(args)
    #         # mapped_exprs = []
    #         # for a in args
    #         for i in eachindex(args)
    #             if args[i] isa ExcelExpr
    #                 e, previous_params = functionalize(args[i], previous_params)
    #                 mapped_exprs[i] = e
    #             else
    #                 mapped_exprs[i] = args[i]
    #             end
    #             # push!(mapped_exprs, e)
    #         end
    #         # mapped = map(a -> functionalize(a, ctx, previous_params), args)
    #         # (ExcelExpr(op, mapped_exprs...), previous_params)
    #         (ExcelExpr(op, mapped_exprs), previous_params)
    #         # reduce(vcat, mapped)
    #     end
    # end
end

function functionalize(expr, previous_params::Vector{ExcelExpr})
    expr
end
function functionalize(expr::ExcelExpr, previous_params::Vector{ExcelExpr})
    if expr.head in (:cell_ref, :sheet_ref, :named_range, :range, :table_ref)
        # (ExcelExpr(:func_param, Any[length(previous_params) + 1]), [previous_params... expr])
        push!(previous_params, expr)
        ExcelExpr(:func_param, length(previous_params))
    else
        op = expr.head
        args = expr.args
        mapped_exprs = similar(args)
        for i in eachindex(args)
            if typeof(args[i]) == ExcelExpr
                mapped_exprs[i] = functionalize(args[i], previous_params)
            else
                mapped_exprs[i] = args[i]
            end
        end
        ExcelExpr(op, mapped_exprs)
    end
end

function functionalize(expr, previous_params::Matrix{ExcelExpr})
    (expr, previous_params)
end

function functionalize(expr::FlatExpr, previous_params::Vector{ExcelExpr})
    to_remove = Vector{Int}()
    for i in eachindex(expr.parts)
        part = expr.parts[i]
        if part.head in (:cell_ref, :sheet_ref, :named_range, :range, :table_ref)
            for arg in part.args
                if arg isa FlatIdx
                    push!(to_remove, arg.i)
                end
            end
        end
    end
    new_parts = Vector{ExcelExpr}(undef, length(expr.parts) - length(to_remove))
    add_i = 1

    # new_expr = copy(expr)
    # to_remove = Vector{Int}()
    for i in eachindex(expr.parts)
        i in to_remove && continue

        part = expr.parts[i]
        if part.head in (:cell_ref, :sheet_ref, :named_range, :range, :table_ref)
            push!(previous_params, convert_to_expr(part, expr))
            new_parts[add_i] = ExcelExpr(:func_param, Any[length(previous_params)])
        else
            # new_args = similar(part.args)
            if all(v -> !(v isa FlatIdx), part.args)
                new_parts[add_i] = part
                # push!(new_parts, part)
            else
                new_args = copy(part.args)
                # for (i, arg) in enumerate(part.args)
                for (i, arg) in enumerate(new_args)
                    if arg isa FlatIdx
                        offset = 0
                        for t in to_remove
                            if t < arg.i
                                offset += 1
                            end
                        end

                        new_i = arg.i - offset
                        new_args[i] = FlatIdx(new_i)
                    end
                end
                new_parts[add_i] = ExcelExpr(part.head, new_args)
                # push!(new_parts, ExcelExpr(part.head, new_args))
            end
        end
        add_i += 1
    end

    # new_parts = Vector{ExcelExpr}()
    # for (i, part) in enumerate(new_expr.parts)
    #     i in to_remove && continue

    #     new_args = similar(part.args)

    #     for (i, arg) in enumerate(part.args)
    #         if arg isa FlatIdx
    #             new_i = arg.i - sum(to_remove .< arg.i)
    #             new_args[i] = FlatIdx(new_i)
    #         else
    #             new_args[i] = arg
    #         end
    #     end
    #     push!(new_parts, ExcelExpr(part.head, new_args))
    # end
    # FlatExpr(filter(e -> e.head != :missing, new_expr.parts))
    FlatExpr(new_parts)
    # new_expr
end

function functionalize(expr)
    params = Vector{ExcelExpr}()
    new_expr = functionalize(expr, params)
    (new_expr, permutedims(params))

    # params = Vector{ExcelExpr}()

    # new_expr = functionalize!(deepcopy(expr), params)
    # (new_expr, permutedims(params))
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
        for g in groups
            if by(item, first(g))
                push!(g, item)
                break
            end
        end
        push!(groups, [item])
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
    function in_table(cell, table)
        c = colnum(cell)
        r = rownum(cell)
        # cell.sheet_name == table.sheet_name && c >= startcol(table) && c <= endcol(table) && r >= startrow(table) && r <= endrow(table)
        cell.sheet_name == table.sheet_name && (r, c) in table
    end

    for cell in sort(input_cells)
        struct_name = var_names[cell]

        is_table_cell = false
        for table in tables
            if in_table(cell, table)
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

        push!(lines, "\t$struct_name = $(repr(cell_dict[cell].value))")
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
    r = rownum(cell)
    c = colnum(cell)
    findfirst(t -> cell.sheet_name == t.sheet_name && ((r, c) in t), tables)
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

        stmt = StandardStatement(cell, rhs_expr, get(workbook.cell_dependencies, cell, []))

        statements[i] = stmt
    end

    output_stmt = OutputStatement(subset.output_cells)
    statements[end] = output_stmt

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
            end
            @assert !(cell in keys(cell_to_statement))

            cell_to_statement[cell] = statement
        end
    end

    cell_to_statement
end

function make_statement_graph(statements::Vector{AbstractStatement})
    cell_to_statement = make_cell_to_statement_dict(statements)
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

    statement_nums = Dict{AbstractStatement, Int64}([n => i for (i, n) in enumerate(statements)])
    # adj_matrix = zeros(Bool, (length(statements), length(statements)))

    edge_list = Vector{Edge{Int64}}()
    for statement in statements
        start_node = statement_nums[statement]
        cell_deps::Vector{CellDependency} = get_cell_deps(statement)
        for cell_dep in cell_deps
            # cell_dep::CellDependency
            if !(cell_dep in keys(cell_to_statement))
                @show get_set_cells(statement)
                @show cell_dep
            end
            end_statement = cell_to_statement[cell_dep]
            end_node = statement_nums[end_statement]

            # In cases like grouped statements, it's possible for nodes to depend on themselves
            # we should just be able to ignore that
            if start_node != end_node
                # adj_matrix[start_node, end_node] = true
                push!(edge_list, Edge(start_node, end_node))
            end
        end
    end


    # nested_edge_list = [[(statement_nums[statement], statement_nums[cell_to_statement[cell_dep]]) for cell_dep in get_cell_deps(statement)] for statement in statements]
    # edge_list = reduce(vcat, nested_edge_list)

    # graph = Graphs.SimpleDiGraphFromIterator(Edge.(edge_list))
    # graph = Graphs.SimpleDiGraphFromIterator(edge_list)
    graph = Graphs.SimpleDiGraph(edge_list)
    # graph = Graphs.SimpleDiGraph(adj_matrix)

    cycles = Graphs.simplecycles(graph)
    if !isempty(cycles)
        println("Removing $(length(cycles)) cycles from the statement graph, this is almost certainly incorrect.")
        for cycle in cycles
            rem_edge!(graph, cycle[end], cycle[begin])
        end
    end

    graph
end

# export_statements(io::IO, exporter, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement}) = export_statements_levels_with_moving(io, exporter, wb, statements)
export_statements(io::IO, exporter, wb::ExcelWorkbook, statements::AbstractArray{AbstractStatement}) = export_statements_levels(io, exporter, wb, statements)

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
                usages = "[" * join(to_string.((exporter,), dependents), ", ") * "]"
                write(io, "# Used in $(length(dependents)) places: $usages\n")
            end

            write(io, export_statement(exporter, wb, s))
        end

        write(io, "\n\n")
    end
    # grouped_by_level = Dict((l => [kv.first for kv in topo_levels if kv.second == l]) for l in 0:max_level)
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

function make_input_struct(exporter::JuliaExporter, statements::AbstractArray{AbstractStatement})
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
    start_c = startcol(table)

    col_defs = Vector{String}()
    sizehint!(col_defs, num_cols)

    for c in startcol(table):endcol(table)
        col_cells = [CellDependency(sheet, index_to_cellname(c, r)) for r in startrow(table):endrow(table)]

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
    # col_names = [column_name(table, c) for c in 1:num_cols]
    # line_str = "\t$lhs = DataFrame(Base.convert(Matrix{Any}, zeros($num_rows, $num_cols)), [$(join(repr.(col_names), ", "))])"
    line_str = "\t$lhs = DataFrame($(join(col_defs, ", ")))"
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
        push!(lines, make_dataframe_declaration(exporter, wb, table))
        # lhs = getname(table)
        # num_rows, num_cols = size(table)
        # col_names = [column_name(table, c) for c in 1:num_cols]
        # # line_str = "\t$lhs = DataFrame(Base.convert(Matrix{Any}, zeros($num_rows, $num_cols)), [$(join(repr.(col_names), ", "))])"
        # line_str = "\t$lhs = Base.convert(Matrix{Any}, zeros($num_rows, $num_cols))"
        # push!(lines, line_str)
        # write(output_file, line_str * "\n")
    end
    push!(lines, "")

    for table in tables
        if !(table in keys(grouped_by_set_table))
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

    push!(lines, "\tTables(")
    for table in tables
        push!(lines, "\t\t" * getname(table) * ",")
    end

    push!(lines, "\t)")
    push!(lines, "end")

    join(lines, "\n")
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

        exporter.var_names[cell_ref] = "inputs.$var"
    end


    export_statements(output_file, exporter, wb, statements)

    filter!(p -> false, exporter.var_names)
    for (k, v) in starting_var_names
        exporter.var_names[k] = v
    end

    write(output_file, "end\n")
    # write(output_file, "\ncalculate()")

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
                func_statement = FunctionStatement(statement.assigned_var, statements[inputs], [statements[reverse(intermediate)]..., statement])

                statement_group = [statement, statements[intermediate]...]
                filter!(s -> !(s in statement_group), new_statements)
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
