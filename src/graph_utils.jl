
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


function bfs_multi_parents(g::AbstractGraph, sources; dir=:out)
    if dir == :out
        Graphs._bfs_parents(g, sources, outneighbors)
    else
        Graphs._bfs_parents(g, sources, inneighbors)
    end
end

"""
    find_cycle(g::SimpleDiGraph)

Search for a directed cycle in `g`.

Returns:
- A vector of vertex indices representing a cycle (with the first vertex
  repeated at the end), or
- `nothing` if the graph is acyclic.
"""
function find_cycle(g::SimpleDiGraph)
    n = nv(g)
    visited = falses(n)
    on_stack = falses(n)
    parent = fill(0, n)

    cycle = nothing

    function dfs(u)
        visited[u] = true
        on_stack[u] = true

        for v in outneighbors(g, u)
            if cycle !== nothing
                return
            elseif !visited[v]
                parent[v] = u
                dfs(v)
            elseif on_stack[v]
                # Found a back edge u -> v, reconstruct cycle
                path = [v]
                cur = u
                while cur != v
                    push!(path, cur)
                    cur = parent[cur]
                end
                push!(path, v)   # close the cycle
                reverse!(path)
                cycle = path
                return
            end
        end

        on_stack[u] = false
    end

    for v in 1:n
        if !visited[v]
            dfs(v)
            if cycle !== nothing
                return cycle
            end
        end
    end

    return nothing
end