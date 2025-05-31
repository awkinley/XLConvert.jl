function debug_group_statements(statements::Vector{AbstractStatement}, graph, topo_levels, node)

    max_level = maximum(values(topo_levels))
    grouped_by_level = Dict((l => [kv.first for kv in topo_levels if kv.second == l]) for l in 0:max_level)

    function can_be_smushed(node)
        dependencies = outneighbors(graph, node)
        node_level = topo_levels[node]

        statement = statements[node]
        # Don't smush output statement, since it doens't play well in groups
        # (since it can't be part of a for loop)
        if statement isa OutputStatement
            println("Can't smush because it's an output statement")
            return false, missing, node_level
        end

        if isempty(dependencies)
            # Technically, if the node was at a non-zero level, it could be
            # freely moved down, but I'm not sure why that would be desirable
            println("Can't smush because it has no dependencies")
            return false, missing, node_level
        end

        for d in dependencies
            println(statements[d])
        end
        dep_levels = [topo_levels[n] for n in dependencies]
        @show node_level dep_levels
        level_limiters = findall(l -> l == node_level - 1, dep_levels)

        @show length(level_limiters)

        if length(level_limiters) == 1
            max_compress_level = maximum(filter(l -> l != node_level - 1, dep_levels), init=0)
            # max_compress_level = maximum(dep_levels[level_limiters], init=0)
            true, dependencies[level_limiters[1]], max_compress_level
        else
            false, missing, node_level
        end
    end
    visited = Set{Int64}()

    new_statements = copy(statements)


    smushable, level_limiter, max_compress_level = can_be_smushed(node)
    @show smushable level_limiter max_compress_level
    compress_group = [node]
    if smushable
        @show topo_levels[level_limiter] > max_compress_level
        @show length(outneighbors(graph, level_limiter))
    end

    while smushable && !(level_limiter in visited) && (topo_levels[level_limiter] > max_compress_level) && (length(outneighbors(graph, level_limiter)) > 0)
        push!(compress_group, level_limiter)
        smushable, level_limiter, next_max_compress_level = can_be_smushed(level_limiter)
        @show smushable level_limiter max_compress_level
        max_compress_level = max(max_compress_level, next_max_compress_level)
    end

    union!(visited, compress_group)

    @show compress_group
    if length(compress_group) > 10
        println("Found a group of $(length(compress_group)) nodes that can be smushed")
        for node in compress_group
            println(get_set_cells(statements[node]))
        end
    end

end

function group_statements(statements::Vector{AbstractStatement}, graph, topo_levels)

    max_level = maximum(values(topo_levels))
    grouped_by_level = Dict((l => [kv.first for kv in topo_levels if kv.second == l]) for l in 0:max_level)

    function can_be_smushed(node)
        dependencies = outneighbors(graph, node)
        node_level = topo_levels[node]

        statement = statements[node]
        # Don't smush output statement, since it doens't play well in groups
        # (since it can't be part of a for loop)
        if statement isa OutputStatement
            return false, missing, node_level
        end

        if isempty(dependencies)
            # Technically, if the node was at a non-zero level, it could be
            # freely moved down, but I'm not sure why that would be desirable
            return false, missing, node_level
        end

        dep_levels = [topo_levels[n] for n in dependencies]
        level_limiters = findall(l -> l == node_level - 1, dep_levels)


        if length(level_limiters) == 1
            max_compress_level = maximum(filter(l -> l != node_level - 1, dep_levels), init=0)
            # max_compress_level = maximum(dep_levels[level_limiters], init=0)
            true, dependencies[level_limiters[1]], max_compress_level
        else
            false, missing, node_level
        end
    end
    visited = Set{Int64}()

    new_statements = copy(statements)

    for level in reverse(0:max_level)
        level_nodes = grouped_by_level[level]

        for node in level_nodes
            if node in visited
                continue
            end

            smushable, level_limiter, max_compress_level = can_be_smushed(node)
            compress_group = [node]
            while smushable && !(level_limiter in visited) && (topo_levels[level_limiter] > max_compress_level) && (length(outneighbors(graph, level_limiter)) > 0)
                push!(compress_group, level_limiter)
                smushable, level_limiter, next_max_compress_level = can_be_smushed(level_limiter)
                max_compress_level = max(max_compress_level, next_max_compress_level)
            end

            union!(visited, compress_group)

            if length(compress_group) > 10
                println("Found a group of $(length(compress_group)) nodes that can be smushed")
                for node in compress_group[1:5]
                    println(get_set_cells(statements[node]))
                end
                statement_group = [statements[n] for n in compress_group]
                filter!(s -> !(s in statement_group), new_statements)

                push!(new_statements, GroupedStatement(reverse(statement_group)))
            end
        end
    end

    new_statements
end

function group_statements(statements::Vector{AbstractStatement})
    stmt_graph = make_statement_graph(statements)
    stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    grouped = group_statements(statements, stmt_graph, stmt_topo_levels)
    stmt_graph = make_statement_graph(grouped)
    stmt_topo_levels = get_topo_levels_bottom_up(stmt_graph)
    group_statements(grouped, stmt_graph, stmt_topo_levels)

end