function fill_pos!(pos, order::Vector{Int64})
    for (i, o) in enumerate(order)
        pos[o] = i
    end
end

function make_pos(order::Vector{Int64})
    pos = Vector{Int64}(undef, length(order))
    for (i, o) in enumerate(order)
        pos[o] = i
    end
    pos
end

function cost(edge_list::Vector{Tuple{Int64,Int64}}, order::Vector{Int64}, pos::Vector{Int64})
    # pos = make_pos(order)
    fill_pos!(pos, order)
    total_cost = Float64(0)
    # for edge in Graphs.edges(graph)
    for edge in edge_list
        # p1 = pos[src(edge)]
        # p2 = pos[dst(edge)]
        p1 = pos[edge[1]]
        p2 = pos[edge[2]]
        dist = abs(Float64(p1 - p2))
        total_cost += sqrt(dist)
    end
    total_cost
    # sum(abs(pos[src(edge)] - pos[dst(edge)])^distance_decay for edge in Graphs.edges(g))
    # sum(sqrt(abs(pos[src(edge)] - pos[dst(edge)])) for edge in Graphs.edges(graph))
end

function optimized_topological_local_heuristic(g::SimpleDiGraph; iterations=1000)

    n = nv(g)
    order = topological_sort(g)
    edge_list = [(src(e), dst(e)) for e in Graphs.edges(g)]

    is_valid_topological(g, order) || error("Invalid initial topological order")

    # Cost function
    pos = make_pos(order)

    current_order = copy(order)
    current_cost = cost(edge_list, current_order, pos)
    @show "Starting cost" current_cost

    costs = Vector{Float64}()
    push!(costs, current_cost)

    # Neighbor generation
    function neighbor!(order, node_idx)
        node = order[node_idx]

        fill_pos!(pos, order)

        # pos = make_pos(order)
        # pos = Dict(v => i for (i, v) in enumerate(order))
        get_pos = n -> pos[n]

        succs = outneighbors(g, node)
        latest = minimum(get_pos, succs, init=n + 1) - 1
        if latest > node_idx
            deleteat!(order, node_idx)
            insert!(order, latest, node)
        end

        order
    end

    best_order = copy(current_order)
    best_cost = current_cost

    # Annealing loop
    for _ in 1:iterations
        for node_idx in n:-1:1
            neighbor!(current_order, node_idx)
            current_cost = cost(edge_list, current_order, pos)
            push!(costs, current_cost)
            if current_cost < best_cost
                copyto!(best_order, current_order)
                best_cost = current_cost
            end
        end
    end

    @show "Best cost" best_cost

    return best_order, costs
end

function optimized_topological_sa_big_jump(g::SimpleDiGraph; starting_order=nothing,
    distance_decay=2.0, initial_temp=1000.0, cooling_rate=0.95,
    iterations=1000, rng=Random.default_rng())

    # Create graph and mappings
    # all_nodes = unique(vcat(original_order, first.(edges), last.(edges)))
    # label_to_id = Dict(label => i for (i, label) in enumerate(all_nodes))
    # id_to_label = Dict(i => label for (label, i) in label_to_id)
    # n = length(all_nodes)

    # g = SimpleDiGraph(n)
    n = nv(g)
    order = isnothing(starting_order) ? topological_sort(g) : starting_order

    edge_list = [(src(e), dst(e)) for e in Graphs.edges(g)]

    # Initial topological sort using original order
    is_valid_topological(g, order) || error("Invalid initial topological order")

    # Cost function
    pos = make_pos(order)

    current_order = copy(order)
    current_cost = cost(edge_list, current_order, pos)
    @show "Starting cost" current_cost

    costs = Vector{Float64}()
    push!(costs, current_cost)

    best_order = copy(current_order)
    best_cost = current_cost
    temp = initial_temp


    # Neighbor generation
    function neighbor!(order)
        node_idx = rand(rng, 1:n)
        node = order[node_idx]

        fill_pos!(pos, order)

        # pos = make_pos(order)
        # pos = Dict(v => i for (i, v) in enumerate(order))
        get_pos = n -> pos[n]

        # Calculate valid move range
        preds = inneighbors(g, node)
        earliest = maximum(get_pos, preds, init=0) + 1

        succs = outneighbors(g, node)
        latest = minimum(get_pos, succs, init=n + 1) - 1
        # latest = isempty(succs) ? n : minimum(findfirst(isequal(s), order) for s in succs) - 1

        earliest > latest && return nothing
        new_pos = rand(rng, earliest:latest)
        new_pos == node_idx && return nothing

        deleteat!(order, node_idx)
        insert!(order, new_pos, node)

        # if !is_valid_topological(g, order)
        #     println("Made invalid order")
        #     @show node earliest latest new_pos order
        # end
        order
    end

    new_order = zeros(Int64, size(current_order))

    # Annealing loop
    for k in 1:iterations
        # new_order = copy(current_order)
        copyto!(new_order, current_order)
        # new_order .= current_order
        delta_e = neighbor!(new_order)
        delta_e === nothing && continue

        new_cost = cost(edge_list, new_order, pos)
        δ = new_cost - current_cost

        if δ < 0 || rand(rng) < exp(-δ / temp)
            current_order .= new_order
            current_cost = new_cost
            if current_cost < best_cost
                best_order = copy(new_order)
                best_cost = new_cost
            end
            # current_cost < best_cost && (best_order, best_cost) = (copy(new_order), new_cost)
        end

        push!(costs, current_cost)

        # temp *= cooling_rate
        temp = initial_temp * (1 - k / iterations)
    end

    @show "Best cost" best_cost
    # return [id_to_label[v] for v in best_order]
    return best_order, costs
end

function optimized_topological_sa(g::SimpleDiGraph; starting_order=nothing,
    distance_decay=2.0, initial_temp=1000.0, cooling_rate=0.95,
    iterations=1000, rng=Random.default_rng())

    # Create graph and mappings
    # all_nodes = unique(vcat(original_order, first.(edges), last.(edges)))
    # label_to_id = Dict(label => i for (i, label) in enumerate(all_nodes))
    # id_to_label = Dict(i => label for (label, i) in label_to_id)
    # n = length(all_nodes)

    # g = SimpleDiGraph(n)
    n = nv(g)
    order = isnothing(starting_order) ? topological_sort(g) : starting_order

    # order = topological_sort(g)
    edge_list = [(src(e), dst(e)) for e in Graphs.edges(g)]
    # @show "original order" order
    # edge_ids = Tuple{Int,Int}[]
    # for (u, v) in edges
    #     add_edge!(g, label_to_id[u], label_to_id[v])
    #     push!(edge_ids, (label_to_id[u], label_to_id[v]))
    # end

    # Initial topological sort using original order
    # order = [label_to_id[n] for n in original_order]
    is_valid_topological(g, order) || error("Invalid initial topological order")

    # Cost function
    pos = make_pos(order)

    current_order = copy(order)
    current_cost = cost(edge_list, current_order, pos)
    @show "Starting cost" current_cost

    costs = Vector{Float64}()
    push!(costs, current_cost)

    best_order = copy(current_order)
    best_cost = current_cost
    temp = initial_temp

    function calc_node_energy(pos, node)
        preds = inneighbors(g, node)
        succs = outneighbors(g, node)

        p1 = pos[node]

        total_cost = 0.0
        for n in preds
            p2 = pos[n]
            dist = abs(Float64(p1 - p2))
            total_cost += sqrt(dist)
        end

        for n in succs
            p2 = pos[n]
            dist = abs(Float64(p1 - p2))
            total_cost += sqrt(dist)
        end

        total_cost
    end

    # Neighbor generation
    function neighbor!(order)
        node_idx = rand(rng, 1:n)
        node = order[node_idx]

        fill_pos!(pos, order)

        # pos = make_pos(order)
        # pos = Dict(v => i for (i, v) in enumerate(order))
        get_pos = n -> pos[n]

        # Calculate valid move range
        preds = inneighbors(g, node)
        earliest = maximum(get_pos, preds, init=0) + 1

        succs = outneighbors(g, node)
        latest = minimum(get_pos, succs, init=n + 1) - 1
        # latest = isempty(succs) ? n : minimum(findfirst(isequal(s), order) for s in succs) - 1

        earliest > latest && return nothing
        # new_pos = rand(rng, earliest:latest)
        new_pos = node_idx + (rand(rng, Bool) ? 1 : -1)
        # new_pos == node_idx && return nothing
        ((new_pos > latest) || (new_pos < earliest)) && return nothing

        swap_node = order[new_pos]

        initial_energy = calc_node_energy(pos, swap_node) + calc_node_energy(pos, node)

        tmp = order[new_pos]
        order[new_pos] = order[node_idx]
        order[node_idx] = tmp

        pos[swap_node] = node_idx
        pos[node] = new_pos

        new_energy = calc_node_energy(pos, swap_node) + calc_node_energy(pos, node)

        delta_energy = new_energy - initial_energy
        # deleteat!(order, node_idx)
        # insert!(order, new_pos, node)

        # if !is_valid_topological(g, order)
        #     println("Made invalid order")
        #     @show node earliest latest new_pos order
        # end
        return delta_energy
    end

    new_order = zeros(Int64, size(current_order))

    # Annealing loop
    for _ in 1:iterations
        # new_order = copy(current_order)
        copyto!(new_order, current_order)
        # new_order .= current_order
        delta_e = neighbor!(new_order)
        delta_e === nothing && continue

        # if !is_valid_topological(g, new_order)
        #     current_order = new_order  # Accept invalid moves temporarily
        #     revert = neighbor!(new_order)  # Try to fix
        #     ((revert === nothing) || !is_valid_topological(g, revert)) && continue
        # end

        # new_cost = cost(g, new_order, pos)
        # new_cost = cost(edge_list, new_order, pos)
        # full_delta = new_cost - current_cost
        δ = delta_e
        # @show δ full_delta
        # @assert abs(δ - delta_e) < 1e-8
        new_cost = current_cost + δ

        if δ < 0 || rand(rng) < exp(-δ / temp)
            # if δ < 0
            # copyto!(current_order, new_order)
            current_order .= new_order
            current_cost = new_cost
            if current_cost < best_cost
                best_order = copy(new_order)
                best_cost = new_cost
            end
            # current_cost < best_cost && (best_order, best_cost) = (copy(new_order), new_cost)
        end

        push!(costs, current_cost)

        temp *= cooling_rate
    end

    @show "Best cost" best_cost
    # return [id_to_label[v] for v in best_order]
    return best_order, costs
end

function is_valid_topological(g::SimpleDiGraph, order::Vector{Int})
    visited = Set{Int}()
    for v in order
        for p in inneighbors(g, v)
            p ∉ visited && return false
        end
        push!(visited, v)
    end
    length(visited) == nv(g)
end


# edges = [(2, 3), (3, 5), (1, 4), (4, 5), (5, 6), (6, 7), (7, 8), (1, 7), (3, 8)]

# graph = Graphs.SimpleDiGraphFromIterator(Edge.(edges))
# optimized_order = optimized_topological_sa(graph; distance_decay=0.5)
# println(['A' + (n - 1) for n in optimized_order])