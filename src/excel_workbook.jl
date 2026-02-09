
struct MissingCell
end

struct ValueCell1
    cell::XLSX.Cell
    value::Any
end

ValueCell = ValueCell1

struct FormulaCell
    cell::XLSX.Cell
    expr::Union{FlatExpr, ExcelExpr, Float64, Int64, String, Missing}
end

struct SpillCell
    cell::CellDependency
    expr::Union{FlatExpr, ExcelExpr, Float64, Int64, String, Missing}
end

# FormulaCell = FormulaCell2

CellTypes = Union{ValueCell, FormulaCell, SpillCell}

function get_expr(cell::FormulaCell)
    cell.expr
end

function get_expr(cell::SpillCell)
    cell.expr
end

function get_expr(cell::ValueCell)
    cell.value
end

function get_expr(::MissingCell)
    missing
end



struct ExcelWorkbook2
    # xf::XLSX.XLSXFile
    # cell_dict::Dict{CellDependency, CellTypes}
    # cell_dependencies::Dict{CellDependency, Vector{CellDependency}}
    # key_values::Dict{String, Any}
    xf::XLSX.XLSXFile
    cell_numbering::ObjectNumbering{CellDependency}
    cell_dict::Dict{CellDependency, Any}
    cell_graph::Graphs.SimpleDiGraph{Int64}
    key_values::Dict{String, Any}
end

function get_cell(wb::ExcelWorkbook2, num::Int64)
    get_obj(wb.cell_numbering, num)
end 
function get_num(wb::ExcelWorkbook2, cell::CellDependency)
    get_num(wb.cell_numbering, cell)
end 

function get_dependent_cells(wb::ExcelWorkbook2, cell::CellDependency)
    num = get_num(wb, cell)
    map(n -> get_cell(wb, n), outneighbors(wb.cell_graph, num))
end

ExcelWorkbook = ExcelWorkbook2

has_formula(cell) = false
has_formula(cell::XLSX.Cell) = !isempty(cell.formula)

is_ref_formula(c) = c.formula isa XLSX.ReferencedFormula

function lower_sheet_names(expr, current_sheet::AbstractString)
    return expr
end

function lower_sheet_names(expr::ExcelExpr, current_sheet::AbstractString)
    @match expr begin
        ExcelExpr(:cell_ref, [cell]) => ExcelExpr(:cell_ref, cell, current_sheet)
        ExcelExpr(:cols, [columns]) => begin
            println("Got column expr")
            @show expr
            ExcelExpr(:cols, current_sheet, columns)
        end
        ExcelExpr(:sheet_ref, [sheet_name, ref]) => lower_sheet_names(ref, sheet_name)
        ExcelExpr(op, args) => begin
            ExcelExpr(op, lower_sheet_names.(args, current_sheet))
        end
    end
end

function lower_sheet_names!(expr, current_sheet::AbstractString)
end

function lower_sheet_names!(expr::ExcelExpr, current_sheet::AbstractString)
    @match expr begin
        ExcelExpr(:cell_ref, [cell]) => begin
            push!(expr.args, current_sheet)
        end
        ExcelExpr(:cols, [columns]) => begin
            println("Got column expr")
            @show expr
            push!(expr.args, current_sheet)
            # ExcelExpr(:cols, current_sheet, columns)
        end
        ExcelExpr(:sheet_ref, [sheet_name, ref]) => lower_sheet_names!(ref, sheet_name)
        ExcelExpr(op, args) => begin
            for arg in args
                lower_sheet_names!(arg, current_sheet)
            end
            # ExcelExpr(op, lower_sheet_names.(args, current_sheet))
        end
    end
end

function convert_cell(sheet, sheet_name, cell::XLSX.Cell)
    if has_formula(cell)
        try
            @assert !(cell.formula isa XLSX.FormulaReference)
            expr = toexpr(cell.formula.formula)
            lower_sheet_names!(expr, sheet_name)
            FormulaCell(cell, convert_to_flat_expr(expr))
            # FormulaCell(cell, expr)
        catch e
            println("Failed to parse cell formula")
            println(cell.formula.formula)
            # throw(e)
            ValueCell(cell, XLSX.getdata(sheet, cell))
        end
    else
        ValueCell(cell, XLSX.getdata(sheet, cell))
    end
end

function offset_formula_cell(new_cell::XLSX.Cell, formula_cell::FormulaCell)
    cell_ref = formula_cell.cell.ref
    start_row = cell_ref.row_number
    start_col = cell_ref.column_number

    end_col, end_row = parse_cell(new_cell.ref.name)
    delta_x = end_col - start_col
    delta_y = end_row - start_row

    expression = offset(formula_cell.expr, delta_y, delta_x)
    FormulaCell(new_cell, expression)
end

function get_cell_dict(xl)
    cell_dict = Dict{CellDependency, CellTypes}()

    for sheet_name in XLSX.sheetnames(xl)
        println("Parsing worksheet $(sheet_name)")

        sheet = xl[sheet_name]
        all_cells = filter(!isempty, get_all_cells(sheet))

        ref_cells_dict = Dict{Int64, CellDependency}()
        formula_refs_to_handle = Vector{XLSX.Cell}()

        for row in XLSX.eachrow(sheet)
            for cell in values(row.rowcells)
                cell_dep = CellDependency(sheet_name, cell.ref.name)

                if cell.formula isa XLSX.FormulaReference
                    formula = cell.formula
                    if formula.id in keys(ref_cells_dict)
                        formula_cell = cell_dict[ref_cells_dict[formula.id]]
                        cell_dict[cell_dep] = offset_formula_cell(cell, formula_cell)
                    else
                        # This would happen if a formula is referenced before
                        # it's defined, I don't know if this actually happens in
                        # practice. I haven't seen it happen, but haven't looked too hard.
                        push!(formula_refs_to_handle, cell)
                    end
                else
                    cell_dict[cell_dep] = convert_cell(sheet, sheet_name, cell)

                    if cell.formula isa XLSX.ReferencedFormula
                        ref_cells_dict[cell.formula.id] = cell_dep
                    end
                end

            end
        end

        # @show length(formula_refs_to_handle)
        for cell in formula_refs_to_handle
            cell_dep = CellDependency(sheet_name, cell.ref.name)
            formula_cell = cell_dict[ref_cells_dict[cell.formula.id]]
            cell_dict[cell_dep] = offset_formula_cell(cell, formula_cell)
        end
    end

    cell_dict
end

function get_all_dependencies(cell_dict::Dict{CellDependency, CellTypes}, key_values)
    output = Dict{CellDependency, Vector{CellDependency}}()

    for (cell, content) in cell_dict
        if content isa FormulaCell
            try
                output[cell] = get_expr_dependencies(content.expr, key_values)
            catch e
                println("Error getting cell dependencies for cell $cell")
                # @show cell
                # @show content
                @show content.cell.formula
                println("Expr:")
                show(stdout, "text/plain", content.expr)
                @show e
                # throw(e)
            end
        end
    end

    output
end

function get_extra_named_values!(xf::XLSX.XLSXFile)
    xroot = XLSX.xmlroot(xf, "xl/workbook.xml")
    @assert EzXML.nodename(xroot) == "workbook" "Malformed xl/workbook.xml. Root node name should be 'workbook'. Got '$(EzXML.nodename(xroot))'."

    # workbook to be parsed
    workbook = XLSX.get_workbook(xf)

    existing_keys = keys(workbook.workbook_names)
    # named ranges
    for node in EzXML.eachelement(xroot)
        if EzXML.nodename(node) == "definedNames"
            for defined_name_node in EzXML.eachelement(node)
                @assert EzXML.nodename(defined_name_node) == "definedName"
                defined_value_string = EzXML.nodecontent(defined_name_node)
                name = defined_name_node["name"]

                local defined_value::XLSX.DefinedNameValueTypes
                if !(name in existing_keys)
                    defined_value = defined_value_string
                else
                    continue
                end


                if haskey(defined_name_node, "localSheetId")
                    # is a Worksheet level name

                    # localSheetId is the 0-based index of the Worksheet in the order
                    # that it is displayed on screen.
                    # Which is the order of the elements under <sheets> element in workbook.xml .
                    localSheetId = parse(Int, defined_name_node["localSheetId"]) + 1
                    sheetId = workbook.sheets[localSheetId].sheetId
                    # println("worksheet_names ($sheetId, $name) = $defined_value")
                    workbook.worksheet_names[(sheetId, name)] = defined_value
                else
                    # is a Workbook level name
                    # println("worksheet_names $name = $defined_value")
                    workbook.workbook_names[name] = defined_value
                end
            end

            break
        end
    end

    nothing
end

function parse_workbook(filepath::AbstractString)
    println("Opening excel file...")
    @time xf = XLSX.readxlsx(filepath)
    get_extra_named_values!(xf)
    println("Getting cell dict...")
    @time cell_dict = get_cell_dict(xf)
    println("Getting cell dependencies...")

    # parsed_key_values = Dict((p[1] => lower_sheet_names(FormulaParser.toexpr(repr(p[2])), "")) for p in XLSX.get_workbook(xf).workbook_names)
    # parsed_key_values = Dict((p[1] => lower_sheet_names(toexpr(repr(p[2])), "")) for p in XLSX.get_workbook(xf).workbook_names)
    parsed_key_values = Dict{String, Any}()
    for p in XLSX.get_workbook(xf).workbook_names
        try
            expr = toexpr(string(p[2]))
            lower_sheet_names!(expr, "")
            parsed_key_values[p[1]] = convert_to_flat_expr(expr)
            # parsed_key_values[p[1]] = expr
        catch e
            println("Failed to parse named range formula")
            println(p[1] * ": " * string(p[2]))
            @show e
        end
    end


    for (key, value) in XLSX.get_workbook(xf).worksheet_names
        name = key[2]
        try
            expr = toexpr(string(value))
            lower_sheet_names!(expr, "")
            parsed_key_values[name] = convert_to_flat_expr(expr)
            # parsed_key_values[p[1]] = expr
        catch e
            println("Failed to parse named range formula")
            println(p[1] * ": " * string(p[2]))
            @show e
        end
    end



    # parsed_key_values = Dict((p[1] => lower_sheet_names(toexpr(string(p[2])), "")) for p in XLSX.get_workbook(xf).workbook_names)
    # @show keys(XLSX.get_workbook(xf).workbook_names)

    cell_list = collect(keys(cell_dict))
    cell_numbering = ObjectNumbering(cell_list)

    edge_list = Vector{Edge{Int64}}()
    # A relatively random (and hopefully conservative) guess that the average degree is 2
    sizehint!(edge_list, 2 * length(cell_numbering))
    
    @time "getting expr dependencies" for (i, cell) in enumerate(cell_numbering.objs)
        content = get(cell_dict, cell, MissingCell())
        if content isa XLConvert.FormulaCell
            # empty!(handled_deps)
            dep_cells = []
            try
                dep_cells = get_expr_dependency_ranges(content.expr, parsed_key_values)
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
                    num = get_num!(cell_numbering, dep)
                    push!(edge_list, Edge(i, num))
                else
                    sheet = dep.first.sheet_name
                    (start_col, start_row) = start_coord(dep)
                    (end_col, end_row) = end_coord(dep)
                    for r ∈ start_row:end_row, c ∈ start_col:end_col
                        num = get_num!(cell_numbering, CellDependency(sheet, c, r))
                        push!(edge_list, Edge(i, num))
                    end
                end
            end
        end

    end

    @time "graph construction" graph = Graphs.SimpleDiGraph(edge_list)

    ExcelWorkbook(xf, cell_numbering, cell_dict, graph, parsed_key_values)

    # @time cell_dependencies = get_all_dependencies(cell_dict, parsed_key_values)

    # ExcelWorkbook(xf, cell_dict, cell_dependencies, parsed_key_values)
end

function get_all_referenced_cells(workbook::ExcelWorkbook)
    workbook.cell_numbering.objs
    # unioned = Set(keys(workbook.cell_dependencies))
    # for cells in values(workbook.cell_dependencies)
    #     union!(unioned, cells)
    # end

    # # collect(union(keys(workbook.cell_dependencies), values(workbook.cell_dependencies)...))
    # collect(unioned)
end

function get_workbook_subset(workbook::XLConvert.ExcelWorkbook2, output_cells::Vector{CellDependency})
    all_referenced_nodes = get_all_referenced_cells(workbook)
    for cell in output_cells
        if cell ∉ all_referenced_nodes
            println("Adding cell $cell to all_referenced_nodes")
            push!(all_referenced_nodes, cell)
        end
    end

    graph = workbook.cell_graph

    # @show length(all_referenced_nodes) nv(graph)
    # @assert length(all_referenced_nodes) == nv(graph)
    target_used_nodes = zeros(Bool, nv(graph))

    for target in output_cells
        target_node_num = get_num(workbook, target)

        parents = bfs_parents(graph, target_node_num, dir = :out)
        @. target_used_nodes |= parents > 0
    end

    used_nodes_list = findall(target_used_nodes)

    cell_dict = workbook.cell_dict
    # for n in used_nodes_list
    #     cell = get_cell(workbook, n)
    #     cell_val = workbook.cell_dict[cell]
    #     if subgraph_mask[n] || cell_val isa XLConvert.ValueCell || cell_val isa XLConvert.MissingCell
    #         cell_dict[cell] = cell_val
    #     else
    #         # If all the the inputs to this cell (which is an additional input, and would normally just be a value cell)
    #         # are present in the inputs, then it can actually be a formula cell
    #         # This shows up in cases like
    #         # input_a = 10.0
    #         # input_b = input_a
    #         # expr_a = input_a + input_b
    #         # The induced_subgraph will record this interdependency, so it's nicer to keep it
    #         # Note, this doesn't handle cases where only some of the inputs are
    #         # present, in which case we  modify the subgraph to remove the edges between inputs. (see below)
    #         if all([value_inputs_mask[c] for c in outneighbors(graph, n)])
    #             cell_dict[cell] = cell_val
    #         else
    #             cell_dict[cell] = XLConvert.ValueCell(cell_val.cell, workbook.xf[cell.sheet_name][cell.cell])
    #         end
    #     end
    # end

    # value_input_nodes = Set(findall(value_inputs_mask))

    (subgraph, vmap) = induced_subgraph(graph, used_nodes_list)

    # # Modifying the subgraph to remove edges to value input nodes
    # for (new_node, old_node) in enumerate(vmap)
    #     if old_node in value_input_nodes && cell_dict[get_cell(workbook, old_node)] isa XLConvert.ValueCell

    #         edges_to_remove = [Edge(new_node, dep) for dep in outneighbors(subgraph, new_node)]
    #         for e in edges_to_remove
    #             rem_edge!(subgraph, e)
    #         end
    #     end
    # end

    cell_numbering = XLConvert.ObjectNumbering(workbook.cell_numbering.objs[vmap])

    XLConvert.ExcelWorkbook(workbook.xf, cell_numbering, cell_dict, subgraph, workbook.key_values)
end

function get_workbook_subset(workbook::XLConvert.ExcelWorkbook2, output_cells::Vector{CellDependency}, input_cells::Vector{CellDependency})
    all_referenced_nodes = get_all_referenced_cells(workbook)
    for cell in output_cells
        if cell ∉ all_referenced_nodes
            println("Adding cell $cell to all_referenced_nodes")
            push!(all_referenced_nodes, cell)
        end
    end

    graph = workbook.cell_graph

    # @show length(all_referenced_nodes) nv(graph)
    # @assert length(all_referenced_nodes) == nv(graph)
    target_used_nodes = zeros(Bool, nv(graph))

    for target in output_cells
        target_node_num = get_num(workbook, target)

        parents = bfs_parents(graph, target_node_num, dir = :out)
        @. target_used_nodes |= parents > 0
    end

    input_children = zeros(Bool, nv(graph))
    for input in input_cells
        node_num = get_num(workbook, input)
        children = bfs_parents(graph, node_num, dir  = :in)
        @. input_children |=  children > 0
    end
    if !isempty(input_cells)
        subgraph_mask = target_used_nodes .& input_children
    else
        subgraph_mask = target_used_nodes
    end

    value_inputs_mask = zeros(Bool, nv(graph))

    for n in findall(subgraph_mask)
        for dependent in outneighbors(graph, n)
            if !subgraph_mask[dependent]
                value_inputs_mask[dependent] = true
            end
        end
    end

    used_nodes_list = findall(subgraph_mask .| value_inputs_mask)

    cell_dict = Dict{CellDependency, Any}()
    for n in used_nodes_list
        cell = get_cell(workbook, n)
        cell_val = workbook.cell_dict[cell]
        if subgraph_mask[n] || cell_val isa XLConvert.ValueCell || cell_val isa XLConvert.MissingCell
            cell_dict[cell] = cell_val
        else
            # If all the the inputs to this cell (which is an additional input, and would normally just be a value cell)
            # are present in the inputs, then it can actually be a formula cell
            # This shows up in cases like
            # input_a = 10.0
            # input_b = input_a
            # expr_a = input_a + input_b
            # The induced_subgraph will record this interdependency, so it's nicer to keep it
            # Note, this doesn't handle cases where only some of the inputs are
            # present, in which case we  modify the subgraph to remove the edges between inputs. (see below)
            if all([value_inputs_mask[c] for c in outneighbors(graph, n)])
                cell_dict[cell] = cell_val
            else
                cell_dict[cell] = XLConvert.ValueCell(cell_val.cell, workbook.xf[cell.sheet_name][cell.cell])
            end
        end
    end

    value_input_nodes = Set(findall(value_inputs_mask))

    (subgraph, vmap) = induced_subgraph(graph, used_nodes_list)

    # Modifying the subgraph to remove edges to value input nodes
    for (new_node, old_node) in enumerate(vmap)
        if old_node in value_input_nodes && cell_dict[get_cell(workbook, old_node)] isa XLConvert.ValueCell

            edges_to_remove = [Edge(new_node, dep) for dep in outneighbors(subgraph, new_node)]
            for e in edges_to_remove
                rem_edge!(subgraph, e)
            end
        end
    end

    cell_numbering = XLConvert.ObjectNumbering(workbook.cell_numbering.objs[vmap])

    XLConvert.ExcelWorkbook(workbook.xf, cell_numbering, cell_dict, subgraph, workbook.key_values)
end