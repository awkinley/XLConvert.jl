
function try_smush_node(statements::Vector{AbstractStatement}, graph, topo_levels, node; visited = missing, debug = false)
    max_level = maximum(values(topo_levels))
    # grouped_by_level = Dict((l => [kv.first for kv in topo_levels if kv.second == l]) for l in 0:max_level)

    start_level = topo_levels[node]
    if debug
        @show start_level
        @show statements[node]
    end
    if start_level == 0
        return Int64[]
    end

    deps_at_level = Dict{Int64, Set{Int64}}()
    function add_node(new_node)
        level = topo_levels[new_node]
        if level in keys(deps_at_level)
            push!(deps_at_level[level], new_node)
        else
            deps_at_level[level] = Set{Int64}([new_node])
        end
    end

    add_node(node)


    current_level = start_level

    while true
        level_nodes = get(deps_at_level, current_level, Vector{Int64}())
        if debug
            @show current_level
            @show level_nodes
            for n in level_nodes
                println(statements[n])
            end
        end

        if length(level_nodes) > 2
            debug && println("Breaking because there's more than two level nodes")
            break
        end

        was_visited = false
        for n in level_nodes
            dependencies = outneighbors(graph, n)
            for dep in dependencies
                add_node(dep)
                if !ismissing(visited)
                    was_visited |= dep in visited
                end
            end
        end


        current_level -= 1

        if was_visited
            debug && println("Breaking because a dependency was visited")
            break
        end

        if current_level == 0
            debug && println("Breaking because current level = 0")
            break
        end
    end

    current_level += 1

    grouped_nodes = Vector{Int64}()

    function not_input(node)
        length(outneighbors(graph, node)) > 0
    end

    for lvl in reverse(current_level:start_level)
        append!(grouped_nodes, filter(not_input, get(deps_at_level, lvl, Vector{Int64}())))
    end
    @assert length(unique(grouped_nodes)) == length(grouped_nodes)
    if debug
        @show grouped_nodes
    end

    return grouped_nodes
end


function debug_group_statements(statements::Vector{AbstractStatement}, graph, topo_levels, node)

    max_level = maximum(values(topo_levels))
    # grouped_by_level = Dict((l => [kv.first for kv in topo_levels if kv.second == l]) for l in 0:max_level)

    function can_be_smushed(node)
        @show node
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

        println("Dependencies")
        for d in dependencies
            println("\t", statements[d])
        end
        dep_levels = [topo_levels[n] for n in dependencies]
        @show node_level dep_levels
        level_limiters = findall(l -> l == node_level - 1, dep_levels)

        @show length(level_limiters)

        if length(level_limiters) <= 3
            max_compress_level = maximum(filter(l -> l != node_level - 1, dep_levels), init = 0)
            # max_compress_level = maximum(dep_levels[level_limiters], init=0)
            true, dependencies[level_limiters], max_compress_level
        else
            println("Not smushable because more than three level limiters")
            false, missing, node_level
        end
    end
    visited = Set{Int64}()

    new_statements = copy(statements)

    smushable, level_limiter, max_compress_level = can_be_smushed(node)

    min_level = max_compress_level

    @show smushable level_limiter max_compress_level
    compress_group = [node]
    if smushable
        @show [topo_levels[n] for n in level_limiter] .> max_compress_level
        @show length(reduce(vcat, map(Base.Fix1(outneighbors, graph), level_limiter)))
        if !ismissing(level_limiter)
            @show statements[level_limiter]
        end
    end

    # while smushable && !(level_limiter in visited) && (topo_levels[level_limiter] > max_compress_level) && (length(outneighbors(graph, level_limiter)) > 0)
    while smushable && !any([v in visited for v in level_limiter]) && all([topo_levels[limiter] for limiter in level_limiter] .> min_level) && all(length.(outneighbors.((graph,), level_limiter)) .> 0)
        append!(compress_group, level_limiter)
        new_level_limiters = []
        for n in level_limiter
            is_smushable, local_level_limiter, next_max_compress_level = can_be_smushed(n)
            @show next_max_compress_level
            smushable = smushable && is_smushable
            # append!(new_level_limiters, filter(!in(new_level_limiters), local_level_limiter))
            if !ismissing(local_level_limiter)
                append!(new_level_limiters, filter(!in(new_level_limiters), local_level_limiter))
            end
            println("min_level = $min_level, next_max_compress_level = $next_max_compress_level")
            min_level = max(min_level, next_max_compress_level)
            println("Afterwards:")
            @show min_level
        end
        println("Outside loop, min_level = $min_level")

        level_limiter = new_level_limiters
        smushable &= 0 < length(level_limiter) <= 3
        # smushable, level_limiter, next_max_compress_level = can_be_smushed(level_limiter)
        @show smushable level_limiter max_compress_level
        # if !ismissing(level_limiter)
        #     @show statements[level_limiter]
        # end
        # max_compress_level = max(max_compress_level, next_max_compress_level)
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


        # if length(level_limiters) == 1
        if length(level_limiters) <= 3
            max_compress_level = maximum(filter(l -> l != node_level - 1, dep_levels), init = 0)
            # max_compress_level = maximum(dep_levels[level_limiters], init=0)
            true, dependencies[level_limiters], max_compress_level
        else
            false, missing, node_level
        end
    end
    visited = Set{Int64}()

    new_statements = copy(statements)

    for level in reverse(0:max_level)
        level_nodes = grouped_by_level[level]

        for node in level_nodes
            stmt = statements[node]
            do_debug = CellDependency("Cash Flow Analysis", "C108") in get_set_cells(stmt)
            if do_debug
                @show node stmt
                @show node in visited
            end
            if node in visited
                continue
            end
            # @show node statements[node]

            smushable, level_limiter, max_compress_level = can_be_smushed(node)
            compress_group = [node]
            # while smushable && !(level_limiter in visited) && (topo_levels[level_limiter] > max_compress_level) && (length(outneighbors(graph, level_limiter)) > 0)
            #     push!(compress_group, level_limiter)
            #     smushable, level_limiter, next_max_compress_level = can_be_smushed(level_limiter)
            #     max_compress_level = max(max_compress_level, next_max_compress_level)
            # end
            if do_debug
                @show [topo_levels[n] for n in level_limiter]
                @show max_compress_level
            end

            # @show level_limiter

            while smushable && !any([v in visited for v in level_limiter]) && all([topo_levels[limiter] for limiter in level_limiter] .> max_compress_level) && all(length.(outneighbors.((graph,), level_limiter)) .> 0)
                append!(compress_group, level_limiter)
                new_level_limiters = []
                for n in level_limiter
                    is_smushable, local_level_limiter, next_max_compress_level = can_be_smushed(n)
                    if do_debug
                        @show n is_smushable local_level_limiter
                    end
                    smushable = smushable && is_smushable
                    if !ismissing(local_level_limiter)
                        append!(new_level_limiters, filter(!in(new_level_limiters), local_level_limiter))
                    end
                    max_compress_level = max(max_compress_level, next_max_compress_level)

                    smushable || break
                end

                level_limiter = new_level_limiters
                # @show level_limiter
                smushable &= 0 < length(level_limiter) <= 3

                if do_debug
                    @show [topo_levels[n] for n in level_limiter]
                    @show max_compress_level
                end

                if do_debug
                    @show smushable
                    @show !any([v in visited for v in level_limiter])
                    @show all([topo_levels[limiter] for limiter in level_limiter] .> max_compress_level)
                    @show all(length.(outneighbors.((graph,), level_limiter)) .> 0)
                end
            end


            if do_debug
                @show compress_group
            end
            if length(compress_group) > 10
                union!(visited, compress_group)
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

function new_group_statements(statements::Vector{AbstractStatement}, graph, topo_levels)

    max_level = maximum(values(topo_levels))
    grouped_by_level = Dict((l => [kv.first for kv in topo_levels if kv.second == l]) for l in 0:max_level)

    visited = Set{Int64}()

    new_statements = copy(statements)

    for level in reverse(0:max_level)
        level_nodes = grouped_by_level[level]

        for node in level_nodes
            stmt = statements[node]
            if stmt isa OutputStatement
                continue
            end

            # @show stmt
            do_debug = CellDependency("Cash Flow Analysis", "C108") in get_set_cells(stmt)
            if do_debug
                @show node stmt
                @show node in visited
            end
            if node in visited
                continue
            end

            compress_group = try_smush_node(statements, graph, topo_levels, node, visited = visited)
            # @show node statements[node]
            if do_debug
                @show compress_group
            end
            # Technically, this could be alright if the value is then only used internally
            # But because a group will generally introduce a scope, it's important that any values set inside that are available outside
            # This is true for table statements, since the table persists, but for other kinds of statements, the value won't persist
            # It also causes functions to have more parameters than needed
            has_non_table_statement = any(n -> !isa(statements[n], TableStatement), @view compress_group[begin:end-1])
            if length(compress_group) > 10 && !has_non_table_statement
                union!(visited, compress_group)
                println("Found a group of $(length(compress_group)) nodes that can be smushed")
                for node in compress_group[1:5]
                    # println(get_set_cells(statements[node]))
                    println(statements[node])
                end
                statement_group = [statements[n] for n in compress_group]
                for n in compress_group
                    if !(statements[n] in new_statements)
                        @show stmt
                        @show statements[n]
                        @show n in visited
                        throw("Tried to group a statement that wasn't available")
                    end
                end
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
    # grouped = group_statements(statements, stmt_graph, stmt_topo_levels)
    grouped = new_group_statements(statements, stmt_graph, stmt_topo_levels)
    stmt_graph = make_statement_graph(grouped)
    stmt_topo_levels = get_topo_levels_bottom_up(stmt_graph)
    # group_statements(grouped, stmt_graph, stmt_topo_levels)
    new_group_statements(grouped, stmt_graph, stmt_topo_levels)

end
function debug_group_statements(statements::Vector{AbstractStatement}, set_cell::CellDependency)
    stmt_graph = make_statement_graph(statements)
    stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    # grouped = group_statements(statements, stmt_graph, stmt_topo_levels)
    # stmt_graph = make_statement_graph(grouped)
    # stmt_topo_levels = get_topo_levels_bottom_up(stmt_graph)
    node = findfirst(s -> set_cell in get_set_cells(s), statements)
    try_smush_node(statements, stmt_graph, stmt_topo_levels, node; debug = true)
    # debug_group_statements(statements, stmt_graph, stmt_topo_levels, node)

end
