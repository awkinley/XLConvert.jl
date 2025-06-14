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

    # node_nums = Dict([n => i for (i, n) in enumerate(all_referenced_nodes)])
    node_nums = Dict{CellDependency, Int}(n => i for (i, n) in enumerate(all_referenced_nodes))
    # graph = Graphs.SimpleDiGraph()
    # add_vertices!(graph, length(all_referenced_nodes))
    num_edges = sum(length, values(workbook.cell_dependencies))
    @show length(all_referenced_nodes)
    # @show num_edges
    edge_list = Vector{Edge{Int64}}(undef, num_edges)

    # nested_edge_list = [[(node_nums[n], node_nums[parent]) for parent in workbook.cell_dependencies[n]] for n in keys(workbook.cell_dependencies)]
    # edge_list = reduce(vcat, nested_edge_list)
    # adj_matrix = zeros(Bool, (length(all_referenced_nodes), length(all_referenced_nodes)))
    idx = 1
    for n in keys(workbook.cell_dependencies)
        i = node_nums[n]

        for parent in workbook.cell_dependencies[n]
            edge_list[idx] = Edge(i, node_nums[parent])
            idx += 1
            # add_edge!(graph, i, node_nums[parent])
            # adj_matrix[i, node_nums[parent]] = true
        end
    end

    graph = Graphs.SimpleDiGraph(edge_list)
    # graph = Graphs.SimpleDiGraph(adj_matrix)
    target_used_nodes = zeros(Bool, nv(graph))

    # target_used_nodes = Set{Int64}()
    for target in output_cells
        target_node_num = node_nums[target]

        parents = bfs_parents(graph, target_node_num, dir=:out)
        # @show parents
        @. target_used_nodes |= parents > 0

        # used_verts = findall(bfs_parents(graph, target_node_num, dir=:out) .> 0)

        # union!(target_used_nodes, used_verts)
    end
    used_nodes_list = findall(target_used_nodes)
    # used_nodes_list = collect(target_used_nodes)
    # sort!(used_nodes_list)
    # used_nodes_list = sort(collect(target_used_nodes))

    WorkbookSubset(workbook, output_cells, node_nums, used_nodes_list, graph)
end