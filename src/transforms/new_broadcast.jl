"""
As part of converting an excel workbook into code (just julia code at the moment), we want to try and improve the generated code as much as possible.

Part of this is converting repeated formulas (that might arrise from copy and
pasting a formula between rows of a table for example), into a format that is
more like a person would write.

This function aims to very generically identify these situations on each sheet.
It has to ensure all the required properties for the conversion to be valid at
met, this includes the statements not being inter-depdenent, having the same
formula, parameters acting accordingly to how broadcasting should work, etc.



"""

function new_broadcast(statements::Vector{AbstractStatement})
    stmt_graph = make_statement_graph(statements)
    new_broadcast(statements, stmt_graph)
end

function new_broadcast(statements::Vector{AbstractStatement}, stmt_graph::DiGraph; debug::Bool = true)

    new_statements = copy(statements)

    statements_by_sheet = group_to_dict(1:length(statements), s -> (c -> c.sheet_name).(get_set_cells(statements[s])))

    topo = topological_sort(stmt_graph)

    for (sheets, sheet_statements) in statements_by_sheet
        # @show sheet
        length(sheets) == 1 || continue
        # Some debugging lines
        # sheets[1] == "growth- cohort group 1"  || continue
        # sheets[1] == "growth- cohort group 2" || continue
        # startswith(sheets[1], "growth- cohort") || continue
        # sheets[1] == "Structure Calcs" || continue
        # sheets[1] == "Operations" || continue
        # sheets[1] == "oyster Husbandry model" || continue

        sheet = sheets[1]

        println("$sheet has $(length(sheet_statements)) statements")

        func_numbering = ObjectNumbering(Vector{Any}())
        func_params = Vector{Matrix{ExcelExpr}}(undef, length(sheet_statements))
        statement_funcs = Vector{Int64}(undef, length(sheet_statements))

        stmt_to_sheet_i = Dict(n => i for (i, n) in enumerate(sheet_statements))

        for (i, stmt_i) in enumerate(sheet_statements)
            stmt = statements[stmt_i]
            func, params = functionalize(stmt.rhs_expr)

            num = get_num!(func_numbering, func isa FlatExpr ? FastHashedFlatExpr(func) : func)
            statement_funcs[i] = num

            func_params[i] = params
        end

        uses_same_function = group_to_dict(sheet_statements, c -> statement_funcs[stmt_to_sheet_i[c]])

        @show length(uses_same_function)


        for (func_i, stmt_indices) in uses_same_function
            func = get_obj(func_numbering, func_i)
            func isa FastHashedFlatExpr || continue
            length(stmt_indices) <= 1 && continue

            if debug
                println("\nFunction has $(length(stmt_indices)) usages")
            end

            # @display func.expr

            stmt_lhs = [get_set_cells(statements[s])[1] for s in stmt_indices]
            sheet_stmt_i = [stmt_to_sheet_i[i] for i in stmt_indices]

            # test_cell = CellDependency("growth- cohort group 2", "AB64")
            # test_cell = CellDependency("growth- cohort group 2", "G12")
            # test_cell = CellDependency("Structure Calcs", "BR6")
            test_cell = CellDependency("oyster Husbandry model", "A32")

            # if test_cell in stmt_lhs
            #     println("Has $(test_cell)!!")
            # else
            #     continue
            # end

            sort_order = sortperm(stmt_lhs)

            stmt_lhs = stmt_lhs[sort_order]

            params = func_params[sheet_stmt_i, :]
            params = reduce(vcat, params[sort_order])

            stmt_nodes = stmt_indices[sort_order]

            stmt_antichains = partition_to_antichains(stmt_graph, topo, stmt_nodes)

            if debug
                println("Number of antichains = $(length(stmt_antichains))")
            end
            # @show length(stmt_antichains)

            if length(stmt_antichains) == length(stmt_indices)
                if debug
                    println("A group of $(length(stmt_indices)) were all inter-dependent, and so could not be broadcasted")
                    if length(stmt_lhs) > 5
                        @show stmt_lhs[begin:5]
                    else
                        @show stmt_lhs
                    end
                end

                continue
            end

            for chain in stmt_antichains

                if length(chain) <= 2
                    if debug
                        println("A chain of length 2 or less was skipped over for broadcasting")
                        @show statements[chain]
                    end

                    continue
                end

                if debug
                    println("Chain of length $(length(chain))")
                end

                chain_lhs = [get_set_cells(statements[s])[1] for s in chain]
                # if test_cell in chain_lhs
                #     println("chain has test_cell!!")
                # else
                #     continue
                # end
                sort_order = sortperm(chain_lhs)
                chain_sheet_i = [stmt_to_sheet_i[i] for i in chain]

                params = func_params[chain_sheet_i, :]
                params = reduce(vcat, params[sort_order])

                if length(chain_lhs) > 3
                    @show chain_lhs[begin:3]
                    # @display params[begin:5, :]
                else
                    @show chain_lhs
                    # @display params
                end

                length_before = length(new_statements)
                statement_set = Set(statements[chain])
                filter!(!in(statement_set), new_statements)

                broadcasted_stmts = do_broadcast_reduction(statements[chain], func, params, debug = debug)
                append!(new_statements, broadcasted_stmts)
                length_after = length(new_statements)

                if debug
                    println("Reduced the number of statements by $(length_before - length_after) by broadcasting, which gave $(length(broadcasted_stmts)) statements")
                end

            end



        end
    end

    println("Net removed $(length(statements) - length(new_statements)) total statements")

    new_statements

end

function get_broadcast_end(stmts::Vector{AbstractStatement}, run_statement_idxs, params; debug::Bool = true)
    if length(run_statement_idxs) == 1
        return 1
    end

    start_stmt_i = run_statement_idxs[1]
    params1 = params[start_stmt_i, :]
    stmt_i2 = run_statement_idxs[2]
    params2 = params[stmt_i2, :]


    equal_params = Vector{Int}()
    changing_params = Vector{Int}()
    for (i, (p1, p2)) in enumerate(zip(params1, params2))
        if p1 == p2
            push!(equal_params, i)
        else
            push!(changing_params, i)
        end
    end

    start_cell = get_set_cells(stmts[run_statement_idxs[1]])[1]

    for (i, s_i) in enumerate(run_statement_idxs[2:end])
        cell = get_set_cells(stmts[s_i])[1]

        row_offset = rownum(cell) - rownum(start_cell)
        col_offset = colnum(cell) - colnum(start_cell)
        for p_i in equal_params
            if params[s_i, p_i] != params1[p_i]
                if debug
                    # println("Changing params not equal")
                    # @show params[s_i, p_i]
                    # @show params1[p_i]
                end
                return i
            end
        end

        for p_i in changing_params
            if !exprs_equal_with_offset(params[s_i, p_i], params1[p_i], row_offset, col_offset)
                if debug
                    # println("Changing params not equal with offset")
                    # @show p_i findfirst(==(p_i), changing_params)
                    # @show params[s_i, p_i]
                    # @show params1[p_i]
                    # @show row_offset col_offset
                end
                return i
            end
        end
    end

    return length(run_statement_idxs)
end

"""
Expects a list of statements that all use functionalized expression func.
The statements should probably also be sorted by set_cell.
The statements should set the same sheet.
Statements should not be inter-dependent. (Can be achived with partition_to_antichains).

Attempts to both do the broadcasting, as well as provide meaningful diagnositics for why it couldn't.
"""
function do_broadcast_reduction(statements::Vector{AbstractStatement}, func, params; debug::Bool = true)
    new_statements = Vector{AbstractStatement}()

    set_a_single_cell = [length(get_set_cells(s)) == 1 for s in statements]
    stmts = statements[set_a_single_cell]
    append!(new_statements, statements[.!set_a_single_cell])
    masked_params = params[set_a_single_cell, :]

    set_cells = [get_set_cells(s)[1] for s in stmts]
    row_nums = rownum.(set_cells)
    col_nums = colnum.(set_cells)
    coords = zip(col_nums, row_nums) |> collect
    coord_to_statement_i = Dict(c => s for (c, s) in zip(coords, eachindex(stmts)))
    sort!(coords)
    regions = get_2d_regions(coords)

    if debug
        println("Num regions = $(length(regions))")
    end

    for region in regions
        cols, rows = region
        region_area = (length(cols) * length(rows))

        if debug
            firstcol = XLSX.encode_column_number(cols[1])
            lastcol = XLSX.encode_column_number(cols[end])
            println("rows = $rows, cols = $firstcol:$lastcol")
        end

        run_statement_idxs = vec([coord_to_statement_i[(c, r)] for r in rows, c in cols])
        # run_statement_idxs = vec([coord_to_statement_i[(c, r)] for c in cols, r in rows])

        if region_area == 1
            append!(new_statements, stmts[run_statement_idxs])
            continue
        end

        # @show get_set_cells.(stmts[run_statement_idxs[begin:2]])

        last_i = 1
        ranges = Vector{UnitRange{Int}}()

        while last_i <= length(run_statement_idxs)
            i = get_broadcast_end(stmts, @view(run_statement_idxs[last_i:end]), masked_params)
            # @show last_i i

            push!(ranges, last_i:(last_i+i-1))
            last_i += i
        end

        for range in ranges
            if length(range) == 1
                if debug
                    println("Range has length of 1")
                end
                append!(new_statements, stmts[run_statement_idxs[range]])
                continue
            end

            if debug
                println("Range has $(length(range)) statements")
                # @show stmts[run_statement_idxs[range]][begin:3]
            end

            range_statements = stmts[run_statement_idxs[range]]

            mask = isa.(range_statements, [TableStatement])

            if sum(mask) != length(range)
                if debug
                    println("A range seemed to broadcast but $(100 * sum(mask) / length(range))% of the statements were table statements")
                    not_tbl_stmts = range_statements[.!mask]
                    if length(not_tbl_stmts) >= 3
                        @show get_set_cells.(not_tbl_stmts[begin:3])
                    else
                        @show get_set_cells.(not_tbl_stmts)
                    end
                end

                append!(new_statements, range_statements[.!mask])
                range_statements = range_statements[mask]

            end

            if !isempty(range_statements)
                push!(new_statements, GroupedStatement(range_statements))
            end
        end

    end

    return new_statements
end


"""
Partition `nodes` into antichains under full-graph reachability.

Arguments:
- g     : global DAG
- topo  : topological_sort(g), precomputed once
- nodes : subset to partition

Returns:
Vector of antichains, each a Vector{Int}.
"""
function partition_to_antichains(
    g::DiGraph,
    topo::Vector{Int},
    nodes::Vector{Int},
)
    n = nv(g)

    # Mark S
    inS = falses(n)
    for u in nodes
        inS[u] = true
    end

    # DP ranks (0 = unreachable from S)
    rank = zeros(Int, n)

    # Activate frontier
    active = falses(n)
    for u in nodes
        active[u] = true
        rank[u] = 1
    end

    # Propagate only from active nodes
    for u in topo
        if !active[u]
            continue
        end

        # @show u

        ru = rank[u]
        for v in outneighbors(g, u)
            # If we improve v, propagate
            if rank[v] <= ru
                rank[v] = ru + 1
                active[v] = true
            end
        end
    end

    # Collect antichains
    layers = Dict{Int, Vector{Int}}()
    for u in nodes
        push!(get!(layers, rank[u], Int[]), u)
    end

    return collect(values(layers))
end
