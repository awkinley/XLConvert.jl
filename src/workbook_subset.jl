struct WorkbookSubset
    wb::ExcelWorkbook
    output_cells::Vector{CellDependency}
    node_nums::Dict{CellDependency,Int}
    used_nodes::Vector{Int64}
    graph::Graphs.SimpleDiGraph{Int64}
end

# WorkbookSubset = WorkbookSubset2

function get_workbook_subset(workbook::ExcelWorkbook, output_cells::Vector{CellDependency})
    all_referenced_nodes = get_all_referenced_cells(workbook)

    node_nums = Dict([n => i for (i, n) in enumerate(all_referenced_nodes)])
    nested_edge_list = [[(node_nums[n], node_nums[parent]) for parent in workbook.cell_dependencies[n]] for n in keys(workbook.cell_dependencies)]
    edge_list = reduce(vcat, nested_edge_list)

    graph = Graphs.SimpleDiGraphFromIterator(Edge.(edge_list))

    target_used_nodes = Set{Int64}()
    for target in output_cells
        target_node_num = node_nums[target]

        used_verts = findall(bfs_parents(graph, target_node_num, dir=:out) .> 0)

        union!(target_used_nodes, used_verts)
    end
    used_nodes_list = sort(collect(target_used_nodes))

    WorkbookSubset(workbook, output_cells, node_nums, used_nodes_list, graph)
end