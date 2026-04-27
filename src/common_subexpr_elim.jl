
function are_same_expr(a::ExcelExpr, b::ExcelExpr, part_equivalences::Dict{Int, Set{Int}})
    a.head == b.head || return false
    length(a.args) == length(b.args) || return false

    # at this point the heads are the same, and the number of arguments are the same
    # the only thing left to check is whether the arguments are the same, or equivalent

    for (a_arg, b_arg) in zip(a.args, b.args)
        ismissing(a_arg) && return false
        ismissing(b_arg) && return false

        if (a_arg isa FlatIdx) && (b_arg isa FlatIdx)
            if !(b_arg.i in get(part_equivalences, a_arg.i, Set()))
                return false
            end
        elseif a_arg != b_arg
            return false
        end
    end


    true
end

function flat_expr_graph(expr::XLConvert.FlatExpr)

    edge_list = Vector{Edge{Int64}}()

    for i in eachindex(expr.parts)
        part = expr.parts[i]

        for arg in part.args
            if arg isa FlatIdx
                push!(edge_list, Edge(i, Int64(arg.i)))
            end
        end
    end

    Graphs.SimpleDiGraph(edge_list)
end

function common_subexpression_elimination(expr::XLConvert.FlatExpr)
    n = length(expr.parts)

    # part_equivalences = zeros(Bool, n, n)
    expr_graph = flat_expr_graph(expr)

    duplicate_parts = Vector{Int}()
    seen_parts = Set{XLConvert.ExcelExpr}()

    # lowest level leaf equivalence groups
    # currently computed in O(N^2), which isn't optimal
    equiv_groups = []

    for (i, part) in enumerate(expr.parts)
        if any([i in g for g in equiv_groups])
            continue
        end

        rest_idx = i .+ findall(isequal(part), expr.parts[(i+1):end])
        # part_equivalences[i, rest_idx] .= true
        # part_equivalences[rest_idx, i] .= true

        if !isempty(rest_idx)
            push!(equiv_groups, [i, rest_idx...])
        end
    end

    # dictionary of "equivalent" expression parts
    # equivalent mean that their sub-trees match
    # contains the full list for every tree node
    part_equivalences = Dict{Int, Set{Int}}()

    # fill part_equivalences for leaf nodes based on equiv_groups
    for g in equiv_groups
        # println("Group of size $(length(g))")
        s = Set(g)
        for idx in g
            # println("\t$idx $(expr.parts[idx])")
            part_equivalences[idx] = s
        end
    end


    # go backwards through the expression parts
    # are_same_expr determines sub-tree equivalence
    # going backwards means this can be O(N)
    # filling in part_equivalences
    for (i, part) in Iterators.reverse(enumerate(expr.parts))
        i in keys(part_equivalences) && continue
        i == 1 && continue

        equal_idxs = findall(s -> are_same_expr(part, s, part_equivalences), expr.parts[1:(i-1)])
        if !isempty(equal_idxs)

            group = [i, equal_idxs...]
            push!(equiv_groups, group)
            for idx in group
                @assert !(idx in keys(part_equivalences))
                part_equivalences[idx] = Set(group)
            end
        end
    end

    # not sure if this part is necessary
    # it's basically a recomputation of part_equivalences to ensure full mapping
    # but it seems like this should be maintained?
    for g in unique(values(part_equivalences))
        # println("Group of size $(length(g))")
        s = Set(g)
        for idx in g
            # println("\t$idx $(expr.parts[idx])")
            part_equivalences[idx] = s
        end
    end

    # @display part_equivalences
    # @display equiv_groups

    new_parts = Vector{XLConvert.ExcelExpr}()

    equivalent_subexpr_groups = vec(unique(values(part_equivalences)))
    

    # For every group of equivalent subexpressions:
    #
    # 1. Determine if the group needs to be its own variable
    # This is true if any of the group members have parents outside the removed parts 
    # 
    # 2. Determine which groups it depends on, for the purposes of ordering


    # part_remappings = Dict(keys(part_equivalences) .=> minimum.(values(part_equivalences)))
    # @show part_remappings
    # kept_parts = unique(values(part_remappings))
    # @show equivalent_subexpr_groups

    removed_parts = Set(keys(part_equivalences))

    required_groups = Int[]


    for (i, group) in enumerate(equivalent_subexpr_groups)
        required_as_var = false
        for n in group
            for parent in inneighbors(expr_graph, n)
                if !(parent in removed_parts)
                    push!(required_groups, i)
                    required_as_var = true
                    break
                end
            end
            if required_as_var
                break
            end
        end
    end

    # @display required_groups


    sub_exprs = Vector{XLConvert.FlatExpr}()
    subexpr_mapping = Dict()

    for i in eachindex(required_groups)
        group = equivalent_subexpr_groups[required_groups[i]]


        # @show group
        node = minimum(group)
        # @show node
        sub_tree = findall(>(0), Graphs.bfs_parents(expr_graph, node))
        # @show sub_tree
        sub_parts = deepcopy.(expr.parts[sub_tree])

        for part in sub_parts
            for i in eachindex(part.args)
                arg = part.args[i]
                if arg isa FlatIdx
                    new_i = findfirst(==(arg.i), sub_tree)
                    part.args[i] = XLConvert.FlatIdx(new_i)
                end
            end
        end

        if length(sub_parts) >= 1
            for g in group
                subexpr_mapping[g] = i
            end
            push!(sub_exprs, XLConvert.FlatExpr(sub_parts))
        else
            for g in group
                pop!(removed_parts, g)
            end

        end

        # show(stdout, "text/plain", XLConvert.FlatExpr(sub_parts))
    end


    for (i, part) in enumerate(expr.parts)
        i in removed_parts && continue


        if all(v -> !(v isa FlatIdx), part.args)
            push!(new_parts, part)
            continue
        end

        new_args = copy(part.args)
        for (i, arg) in enumerate(new_args)
            if arg isa FlatIdx
                if arg.i in keys(subexpr_mapping)
                    # new_args[i] = ExcelExpr(:func_param, Any[subexpr_mapping[arg.i]])
                    new_args[i] = ExcelExpr(:var_ref, Any[subexpr_mapping[arg.i]])
                else

                    offset = 0
                    for t in removed_parts
                        if t < arg.i
                            offset += 1
                        end
                    end

                    new_i = arg.i - offset
                    new_args[i] = FlatIdx(new_i)
                end
            end
        end
        push!(new_parts, ExcelExpr(part.head, new_args))

    end

    # for (i, e) in enumerate(sub_exprs)
    #     print("cse_$i =")
    #     show(stdout, "text/plain", e)
    # end

    # show(stdout, "text/plain", FlatExpr(new_parts))


    sub_exprs, FlatExpr(new_parts)


    # part_equivalences

end