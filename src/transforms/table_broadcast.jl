
function get_2d_regions(coords::Vector{Tuple{Int,Int}})
    isempty(coords) && return []

    if length(coords) == 1
        c, r = coords[1]
        return [(c:c, r:r)]
    end


    vertical_runs = Vector{Tuple{Int, UnitRange{Int}}}()

    cols = Vector{Int}()

    i = 1
    while i <= length(coords)
        c, start_r = coords[i]
        i += 1
        end_r = start_r
        while (i <= length(coords)) && (coords[i] == (c, end_r + 1))
            end_r += 1
            i += 1
        end

        push!(vertical_runs, (c, start_r:end_r))
    end


    regions = Vector{Tuple{UnitRange{Int},UnitRange{Int}}}()

    i = 1

    while !isempty(vertical_runs)
    # while i <= length(vertical_runs)
        # column, rows = vertical_runs[i]
        # @info "Outer while loop" vertical_runs
        column, rows = popfirst!(vertical_runs)

        end_c = column

        do_continue = true
        while do_continue
            do_continue = false
            j = 1
            while j <= length(vertical_runs)
                # @info "Inner while loop" j vertical_runs
                # @show j vertical_runs
            # for (other_c, other_rows) in vertical_runs
                other_c, other_rows = vertical_runs[j]
                if other_c != end_c + 1 
                    j += 1
                    continue
                end

                if (first(other_rows) <= first(rows)) && (last(other_rows) >= last(rows))
                    if first(other_rows) < first(rows)
                        # println("Would push $((other_c, first(other_rows):(first(rows) - 1)))")
                        push!(vertical_runs, (other_c, first(other_rows):(first(rows) - 1)))
                    end
                    if last(other_rows) > last(rows)
                        # println("Would push $((other_c, (last(rows) + 1):last(other_rows)))")
                        push!(vertical_runs, (other_c, (last(rows) + 1):last(other_rows)))
                    end
                    # do_continue = true

                    end_c += 1
                    deleteat!(vertical_runs, j)
                else
                    j += 1
                end
            end
        end

        push!(regions, (column:end_c, rows))

        i += 1
    end


    sort!(regions; by=a -> first.(a))
end

function get_2d_regions_old(coords::Vector{Tuple{Int,Int}})
    # Given a list of sorted (column, row) points, we find 2d contiguous rectangles.


    isempty(coords) && return []

    if length(coords) == 1
        c, r = coords[1]
        return [(c:c, r:r)]
    end

    regions = Vector{Tuple{UnitRange{Int},UnitRange{Int}}}()

    # Make a copy so that we can be destructive to it to keep track of what we've handled thus far
    points = copy(coords)

    # A basic heuristic is used for finding rectangles, which is to greedily take any runs in a column,
    # and then try and expand that single column to the right. 
    # Only expand to the right if all the same rows are represented in that column.

    while !isempty(points)
        start_c, start_r = popfirst!(points)
        end_r = start_r
        while !isempty(points) && points[1] == (start_c, end_r + 1)
            _, end_r = popfirst!(points)
        end

        end_c = start_c + 1

        next_col_points = [(end_c, r) for r in start_r:end_r]
        while all(p -> p in points, next_col_points)
            end_c += 1
            # Remove the points that are now claimed by this region
            filter!(p -> !(p in next_col_points), points)
            next_col_points = [(end_c, r) for r in start_r:end_r]
        end
        end_c -= 1

        push!(regions, (start_c:end_c, start_r:end_r))

    end

    regions
end


function exprs_equal_with_offset(a::ExcelExpr, b::ExcelExpr, rows::Int, cols::Int)
    if a.head != b.head || length(a.args) != length(b.args)
        return false
    end

    if a.head == :cell_ref
        return a == offset(b, rows, cols)
    elseif a.head == :table_ref
        # @show a
        # @show offset(b, rows, cols)
        # @show isequal(a, offset(b, rows, cols))
        return isequal(a, offset(b, rows, cols))
        # return a == offset(b, rows, cols)
    else
        for i in 1:length(a.args)
            if !equal_with_offset(a.args[i], b.args[i], rows, cols)
                return false
            end
        end
        # for (a_arg, b_arg) in zip(a.args, b.args)
        #     if !equal_with_offset(a_arg, b_arg, rows, cols)
        #         return false
        #     end
        # end
        return true
    end
end

function equal_with_offset(a, b, rows::Int, cols::Int)
    if (a isa ExcelExpr) && (b isa ExcelExpr)
        return exprs_equal_with_offset(a, b, rows, cols)
    end

    if (a isa ExcelExpr) || (b isa ExcelExpr)
        return false
    end

    a === b
end

function table_statements_have_same_equation(a::TableStatement, b::TableStatement)
    if ismissing(a.rhs_expr) || ismissing(b.rhs_expr)
        return a.rhs_expr === b.rhs_expr
    end

    base_expr = a.rhs_expr
    row_a, col_a = a.lhs_expr.args[2:3]
    row_b, col_b = b.lhs_expr.args[2:3]

    delta_y = row_a - row_b
    delta_x = col_a - col_b
    return equal_with_offset(base_expr, b.rhs_expr, delta_y, delta_x)
    # offset_expr = offset(b.rhs_expr, delta_y, delta_x)

    # base_expr === offset_expr
end

function not_interdependent(closure, node_a, node_b)
    !(node_b in outneighbors(closure, node_a)) && !(node_a in outneighbors(closure, node_b))
end
function has_path_within(
    g::AbstractGraph{T},
    u::Integer,
    v::Integer,
    max_steps::Integer,
    seen::Vector{Bool},
) where {T}
    fill!(seen, false)
    # seen = zeros(Bool, nv(g))
    # (seen[u] || seen[v]) && return false
    u == v && return true # cannot be separated
    next = Vector{Tuple{T, Int32}}()
    push!(next, (u, 0))
    seen[u] = true
    while !isempty(next)
        src, level = popfirst!(next) # get new element from queue
        # if level > max_steps
        #     continue
        # end

        for vertex in outneighbors(g, src)
            vertex == v && return true
            if !seen[vertex] && level <= max_steps
                push!(next, (vertex, level + 1)) # push onto queue
                seen[vertex] = true
            end
        end
    end
    return false
end

function table_broadcast_transform_2d!(statements::Vector{AbstractStatement})
    stmt_graph = make_statement_graph(statements)
    # @time closure = transitiveclosure(stmt_graph)
    topo_sorted = topological_sort(reverse(stmt_graph))
    stmt_topo_levels = get_topo_levels_top_down(stmt_graph)
    stmt_to_level = Dict((stmt => stmt_topo_levels[i]) for (i, stmt) in enumerate(statements))
    # table_statements::Vector{TableStatement} = filter(s -> isa(s, TableStatement), statements)
    table_statement_idxs::Vector{Int64} = filter(i -> isa(statements[i], TableStatement), eachindex(statements))
    # node_nums = used_subset.node_nums


    function get_stmt_table(stmt::TableStatement)
        stmt.lhs_expr.args[1]
    end

    path_seen = zeros(Bool, nv(stmt_graph))
    function is_independent(node_a::Int64, node_b::Int64)
        lvl_a = stmt_topo_levels[node_a]
        lvl_b = stmt_topo_levels[node_b]
        if lvl_a == lvl_b
            return true
        end

        if lvl_b > lvl_a
            return !has_path_within(stmt_graph, node_b, node_a, lvl_b - lvl_a, path_seen)
        else
            return !has_path_within(stmt_graph, node_a, node_b, lvl_a - lvl_b, path_seen)
        end
    end

    # same_table_and_level = group_to_dict(table_statement_idxs, i -> (get_stmt_table(statements[i]), stmt_topo_levels[i]))
    same_table_and_level = group_to_dict(table_statement_idxs, i -> get_stmt_table(statements[i]))
    groups = Vector{Vector{Int64}}()
    for group in values(same_table_and_level)
        # @time closure = transitiveclosure(stmt_graph)
        # same_equations = group_by(group, (i, j) -> table_statements_have_same_equation(statements[i], statements[j]) && not_interdependent(closure, i, j))
        same_equations = group_by(group, (i, j) -> table_statements_have_same_equation(statements[i], statements[j]) && is_independent(i, j))
        append!(groups, same_equations)
    end

    new_statements = copy(statements)

    grouped_by_level = groups

    num_table_stmts = 0

    for group in grouped_by_level
        if length(group) <= 1
            continue
        end
        table = statements[group[1]].lhs_expr.args[1]
        sheet = table.sheet_name

        # get_row_num = s -> rownum(s.assigned_vars[1])
        # get_col_num = s -> colnum(s.assigned_vars[1])
        get_row_num = i -> rownum(statements[i].assigned_vars[1])
        get_col_num = i -> colnum(statements[i].assigned_vars[1])

        row_nums = get_row_num.(group)
        col_nums = get_col_num.(group)
        coords = zip(col_nums, row_nums) |> collect
        coord_to_statement = Dict(c => s for (c, s) in zip(coords, group))

        regions = get_2d_regions(sort(coords))
        # println("Group size = $(length(group))")
        for region in regions
            cols, rows = region
            region_area = (length(cols) * length(rows))
            if region_area < 2
                continue
            end
            # if length(cols) > 1
            #     println("\t$region")
            #     println("\tSize: $(region_area)")
            # end


            region_coords = vec([(c, r) for c in cols, r in rows])
            run_statements = [coord_to_statement[coord] for coord in region_coords]

            first_expr = statements[run_statements[1]].rhs_expr

            broadcasted = try
                convert_to_broadcasted(first_expr, length(rows) - 1, length(cols) - 1)
            catch ex
                @show ex
                @show sheet rows cols
                println("Failed to broadcast, would have broadcasted a run of $(length(region_coords))")

                statement_group = statements[run_statements]
                statement_set = Set(statement_group)
                length_before = length(new_statements)
                # filter!(s -> !(s in statement_group), new_statements)
                filter!(!in(statement_set), new_statements)
                length_after = length(new_statements)
                println("Remove $(length_before - length_after) statements to put them in a group")
                push!(new_statements, GroupedStatement(statement_group))
                continue
            end

            length_before = length(new_statements)
            statement_group = statements[run_statements]
            statement_set = Set(statement_group)
            filter!(!in(statement_set), new_statements)
            # filter!(s -> !(s in statements[run_statements]), new_statements)
            length_after = length(new_statements)
            # println("Remove $(length_before - length_after) statements")


            # lhs_expr = ExcelExpr(:table_ref, table, (run_idx[begin]:run_idx[end]) .- startrow(table) .+ 1, run_statements[1].lhs_expr.args[3], (true, true), (true, true))
            table_rows = rows .- startrow(table) .+ 1
            table_cols = cols .- startcol(table) .+ 1
            # println("Table start row = $(startrow(table)), start col = $(startcol(table))")
            # @show table
            # @show table_rows, table_cols
            lhs_expr = ExcelExpr(:table_ref, table, table_rows, table_cols, (true, true), (true, true))

            # lhs_vars = reduce(vcat, get_set_cells.(statements[run_statements]))
            lhs_vars = reduce(vcat, get_set_cells.(statement_group))
            # @show lhs_vars
            # new_statement = TableStatement(lhs_expr, lhs_vars, broadcasted, reduce(vcat, get_cell_deps.(statements[run_statements])) |> unique |> collect, true)
            new_statement = TableStatement(lhs_expr, lhs_vars, broadcasted, reduce(vcat, get_cell_deps.(statement_group)) |> unique |> collect, true)

            push!(new_statements, new_statement)
            num_table_stmts += 1

        end
    end

    length_before = length(statements)
    length_after = length(new_statements)
    println("table_broadcast_transform_2d!: Removed $(length_before - length_after) statements")
    println("table_broadcast_transform_2d!: Created $num_table_stmts table statements")

    new_statements
end

# function table_broadcast_transform_2d!(statements::Vector{AbstractStatement}, used_subset::WorkbookSubset, topo_levels)
#     table_statements::Vector{TableStatement} = filter(s -> isa(s, TableStatement), statements)
#     node_nums = used_subset.node_nums


#     function get_stmt_table(stmt::TableStatement)
#         stmt.lhs_expr.args[1]
#     end

#     level_cache = Dict{TableStatement,Int}()
#     function get_stmt_level(stmt::TableStatement)
#         if stmt in keys(level_cache)
#             return level_cache[stmt]
#         end

#         get_level = c -> (c in keys(node_nums) && node_nums[c] in keys(topo_levels)) ? topo_levels[node_nums[c]] : -1
#         get_stmt_level = stmt -> maximum(get_level, get_set_cells(stmt); init=0)
#         level = get_stmt_level(stmt)
#         level_cache[stmt] = level

#         level
#     end

#     same_table_and_level = group_to_dict(table_statements, stmt -> (get_stmt_table(stmt), get_stmt_level(stmt)))
#     groups = Vector{Vector{TableStatement}}()
#     for group in values(same_table_and_level)
#         same_equations = group_by(group, table_statements_have_same_equation)
#         append!(groups, same_equations)
#     end


#     new_statements = copy(statements)

#     grouped_by_level = groups
#     # grouped_by_level = group_by(table_statements, same_table_equation_and_level)

#     num_table_stmts = 0

#     for group in grouped_by_level
#         if length(group) <= 1
#             continue
#         end
#         table = group[1].lhs_expr.args[1]
#         sheet = table.sheet_name

#         get_row_num = s -> rownum(s.assigned_vars[1])
#         get_col_num = s -> colnum(s.assigned_vars[1])

#         row_nums = get_row_num.(group)
#         col_nums = get_col_num.(group)
#         coords = zip(col_nums, row_nums) |> collect
#         coord_to_statement = Dict(c => s for (c, s) in zip(coords, group))

#         regions = get_2d_regions(sort(coords))
#         # println("Group size = $(length(group))")
#         for region in regions
#             cols, rows = region
#             region_area = (length(cols) * length(rows))
#             if region_area < 2
#                 continue
#             end
#             # if length(cols) > 1
#             #     println("\t$region")
#             #     println("\tSize: $(region_area)")
#             # end


#             region_coords = vec([(c, r) for c in cols, r in rows])
#             run_statements = [coord_to_statement[coord] for coord in region_coords]

#             first_expr = run_statements[1].rhs_expr

#             broadcasted = try
#                 convert_to_broadcasted(first_expr, length(rows) - 1, length(cols) - 1)
#             catch ex
#                 @show ex
#                 println("Failed to broadcast, would have broadcased a run of $(length(region_coords))")
#                 # @show sheet run[1]
#                 continue
#             end

#             length_before = length(new_statements)
#             filter!(s -> !(s in run_statements), new_statements)
#             length_after = length(new_statements)
#             # println("Remove $(length_before - length_after) statements")


#             # lhs_expr = ExcelExpr(:table_ref, table, (run_idx[begin]:run_idx[end]) .- startrow(table) .+ 1, run_statements[1].lhs_expr.args[3], (true, true), (true, true))
#             table_rows = rows .- startrow(table) .+ 1
#             table_cols = cols .- startcol(table) .+ 1
#             # println("Table start row = $(startrow(table)), start col = $(startcol(table))")
#             # @show table
#             # @show table_rows, table_cols
#             lhs_expr = ExcelExpr(:table_ref, table, table_rows, table_cols, (true, true), (true, true))

#             lhs_vars = reduce(vcat, get_set_cells.(run_statements))
#             # @show lhs_vars
#             new_statement = TableStatement(lhs_expr, lhs_vars, broadcasted, reduce(vcat, get_cell_deps.(run_statements)) |> unique |> collect, true)

#             push!(new_statements, new_statement)
#             num_table_stmts += 1

#         end
#     end

#     length_before = length(statements)
#     length_after = length(new_statements)
#     println("table_broadcast_transform_2d!: Removed $(length_before - length_after) statements")
#     println("table_broadcast_transform_2d!: Created $num_table_stmts table statements")

#     new_statements
# end

function table_broadcast_transform_2d!(statements::Vector{AbstractStatement}, used_subset::WorkbookSubset)
    topo_levels = get_topo_levels_top_down(used_subset)
    table_broadcast_transform_2d!(statements, used_subset, topo_levels)
end

# This function is currently unused
function table_broadcast_transform!(statements::Vector{AbstractStatement}, used_subset::WorkbookSubset, topo_levels)
    table_statements::Vector{TableStatement} = filter(s -> isa(s, TableStatement), statements)

    node_nums = used_subset.node_nums

    statements_by_table = group_by(table_statements, (a, b) -> getname(a.lhs_expr.args[1]) == getname(b.lhs_expr.args[1]))
    for table_statements in statements_by_table
        grouped_by_col = group_by(table_statements, (a, b) -> colnum.(get_set_cells(a)) == colnum.(get_set_cells(b)))
        for col_group in grouped_by_col
            example_statement = col_group[1]
            sheet = example_statement.lhs_expr.args[1].sheet_name
            # if sheet != "5.Cap-ex"
            #     continue
            # end

            table = col_group[1].lhs_expr.args[1]

            equal_cell_groups = group_by(col_group, table_statements_have_same_equation)

            for group in equal_cell_groups
                if length(group) > 1
                    example_statement = group[1]

                    println("$(example_statement.assigned_vars), $(length(group))")
                end
                get_row_num = s -> rownum(s.assigned_vars[1])
                sorted_statements = sort(group, by=get_row_num)
                # @show get_row_num.(sorted_statements)



                row_nums = get_row_num.(sorted_statements)
                # row_indexes = rownum.(map(s -> s.assigned_vars[1], sorted_statements)) .- (starting_row - 1)
                starting_row = row_nums[begin]

                # @show row_nums

                runs = get_contiguous_runs(row_nums)

                for run_idx in runs
                    if length(run_idx) < 2
                        continue
                    end
                    # run_statements = sorted_statements[run_idx]
                    statement_indices = [findfirst(r -> r == idx, row_nums) for idx in run_idx]
                    # @show row_nums
                    # @show run_idx
                    # @show statement_indices
                    run_statements = sorted_statements[statement_indices]
                    lhs_vars = reduce(vcat, get_set_cells.(run_statements))

                    levels = unique(map(c -> (c in keys(node_nums) && node_nums[c] in keys(topo_levels)) ? topo_levels[node_nums[c]] : -1, lhs_vars))

                    if !(-1 in levels) && length(levels) == 1
                        first_expr = run_statements[1].rhs_expr

                        broadcasted = try
                            convert_to_broadcasted(first_expr, run_idx[end] - run_idx[begin], 0)
                        catch ex
                            @show ex
                            println("Failed to broadcast, would have broadcased a run of $(length(run))")
                            # @show sheet run[1]
                            continue
                        end

                        length_before = length(statements)
                        filter!(s -> !(s in run_statements), statements)
                        length_after = length(statements)
                        println("Remove $(length_before - length_after) statements")


                        lhs_expr = ExcelExpr(:table_ref, table, (run_idx[begin]:run_idx[end]) .- startrow(table) .+ 1, run_statements[1].lhs_expr.args[3], (true, true), (true, true))

                        # @show lhs_vars
                        new_statement = TableStatement(lhs_expr, lhs_vars, broadcasted, reduce(vcat, get_cell_deps.(run_statements)) |> unique |> collect, true)

                        push!(statements, new_statement)
                    end
                end
            end
        end
    end
end