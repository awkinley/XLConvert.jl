
mutable struct BroadcastedStatement <: AbstractStatement
    lhs_expr::ExcelExpr
    assigned_vars::Vector{CellDependency}
    func_expr::Any
    params::AbstractArray
    rhs_dependencies::Vector{CellDependency}
end


get_cell_deps(stmt::BroadcastedStatement) = stmt.rhs_dependencies
get_set_cells(stmt::BroadcastedStatement) = stmt.assigned_vars
apply_expr_transform!(stmt::BroadcastedStatement, transform) = stmt.func_expr = transform(stmt, stmt.func_expr)

get_set_table(stmt::BroadcastedStatement) = stmt.lhs_expr.args[1]

function to_string(exporter, statement::BroadcastedStatement)
    cell_ref = statement.assigned_vars[1]
    sheet = cell_ref.sheet_name
    lhs = convert(exporter, statement.lhs_expr, sheet)

    "BroadcastedStatement(lhs = $lhs)"
end

function Base.show(io::IO, stmt::BroadcastedStatement)
    assigned_cells = join(string.(stmt.assigned_vars), ", ")
    # assigned_cells = stmt.assigned_vars
    print(io, "BroadcastedStatement([$assigned_cells)]")
end

struct TableRef
    table::ExcelTable
    row::Any
    col::Any
end

get_table(t::TableRef) = t.table
get_rows(t::TableRef) = t.row
get_cols(t::TableRef) = t.col

function cell_dep(t::TableRef)
    @assert length(t.row) == 1 && length(t.col) == 1

    tbl = t.table
    CellDependency(tbl.sheet_name, startcol(tbl) + first(t.col) - 1, startrow(tbl) + first(t.row) - 1)
end

function TableRef(expr::ExcelExpr)
    @match expr begin
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => TableRef(table, row_idx, col_idx)
        _ => throw("tried to convert an invalid ExcelExpr to a TableRef $expr")
    end
end

function indent_str(str::AbstractString, indent::Int)
    string('\t'^indent, str)
end


function make_index_str(vars, coeffs, offset)
    out = ""
    for (name, coeff) in zip(vars, coeffs)
        if coeff == 0
            continue
        end
        
        if !isempty(out)
            out *= " + "
        end

        if coeff == 1
            out *= name
        else
            out *= "$coeff * $name"
        end
    end


    # normalize out slices of length 1
    if length(offset) == 1
        offset = first(offset)
    end

    if offset isa Integer
        if offset != 0
            if !isempty(out)
                out *= " + "
            end

            out *= "$offset"
        end
    else
        out = "($out + $(first(offset))):($out + $(last(offset)))"
    end


    out
end

function make_assertion_string(exporter::PythonExporter, statement::BroadcastedStatement, xf)
    table, lhs_row_idx, lhs_col_idx = @match statement.lhs_expr begin
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => (table, row_idx, col_idx)
        _ => (missing, missing, missing)
    end
    if is_transposed(table)
        (lhs_row_idx, lhs_col_idx) = (lhs_col_idx, lhs_row_idx)
    end

    sheet = string(table.sheet_name)

    out = ""

    if length(lhs_row_idx) > 1
        for c in lhs_col_idx
            first_cell = CellDependency(sheet, startcol(table) + c - 1, startrow(table) + first(lhs_row_idx) - 1)
            last_cell = CellDependency(sheet, startcol(table) + c - 1, startrow(table) + last(lhs_row_idx) - 1)
            cell_values = xf[sheet][first_cell.cell * ":" * last_cell.cell]
            value_strings = convert.(Ref(exporter), cell_values, Ref(sheet))
            value_str = string("[", join(value_strings, ", "), "]")

            lhs_expr = ExcelExpr(:table_ref, table, lhs_row_idx, c, (false, false), (false, false))
            lhs = convert(exporter, lhs_expr,sheet)
            out *= "assert xl.compare_list($lhs, $(value_str))\n"
        end
    else

        r = first(lhs_row_idx)
        first_cell = CellDependency(sheet, startcol(table) + first(lhs_col_idx) - 1, startrow(table) + r - 1)
        last_cell = CellDependency(sheet, startcol(table) + last(lhs_col_idx) - 1, startrow(table) + r - 1)
        cell_values = xf[sheet][first_cell.cell * ":" * last_cell.cell]
        value_strings = convert.(Ref(exporter), cell_values, Ref(sheet))
        value_str = string("[", join(value_strings, ", "), "]")
        lhs_expr = ExcelExpr(:table_ref, table, r, lhs_col_idx, (false, false), (false, false))
        lhs = convert(exporter, lhs_expr,sheet)
        out *= "assert xl.compare_list($lhs, $(value_str))\n"

        # for r in lhs_row_idx, c in lhs_col_idx
        #     cell = CellDependency(sheet, startcol(table) + c - 1, startrow(table) + r - 1)
        #     cell_value = xf[sheet][cell.cell]

        #     value_str = convert(exporter, cell_value, sheet)

        #     lhs_expr = ExcelExpr(:table_ref, table, r, c, (false, false), (false, false))
        #     lhs = convert(exporter, lhs_expr,sheet)

        #     if ismissing(cell_value)
        #         out *= "# assert xl.compare($lhs, $(value_str))\n"
        #     else
        #         out *= "assert xl.compare($lhs, $(value_str))\n"
        #     end
        # end
    end

    out
end

flatten_step_range_len(a) = a
function flatten_step_range_len(a::StepRangeLen) 
    @assert length(a) == 1
    first(a)
end

function can_py_broadcast(expr::FlatExpr)
    function is_broadcast_legal(expr::ExcelExpr)
        if expr.head in (:+, :-, :*, :/, :cell_ref, :table_ref, :func_param)
            return true
        end

        @match expr begin
            ExcelExpr(:call, ["COS", val]) => true
            ExcelExpr(:call, ["PI"]) => true
            _ => false
        end
    end
    all(is_broadcast_legal, expr.parts)
end

function export_statement_broadcasted(exporter::PythonExporter, wb::ExcelWorkbook, statement::BroadcastedStatement)
    table, lhs_row_idx, lhs_col_idx = @match statement.lhs_expr begin
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => (table, row_idx, col_idx)
        _ => (missing, missing, missing)
    end
    if is_transposed(table)
        (lhs_row_idx, lhs_col_idx) = (lhs_col_idx, lhs_row_idx)
    end
    # @show size(statement.params)

    num_rows, num_cols, num_params = size(statement.params)

    lines = Vector{String}()

    # push!(lines, "# $(convert(exporter, statement.lhs_expr, table.sheet_name))")

    if isempty(statement.params) 
        param_coords = zeros(Int64, (0, 0, 0, 0))
    else
        param_coords = stack(get_param_cell_coords, statement.params)
    end
    param_broadcast_behavior = Matrix{Tuple{Int, Int}}(undef, num_params, 2)
    for i in axes(param_coords, 4)
        coords = @view param_coords[:, :, :, i]

        if any(ismissing, coords)
            continue
        end

        rows = @view coords[1, :, :]
        cols = @view coords[2, :, :]

        row_behavior = flatten_step_range_len.(identify_broadcast_behavior(rows))
        col_behavior = flatten_step_range_len.(identify_broadcast_behavior(cols))

        param_broadcast_behavior[i, :] .= (row_behavior, col_behavior)
    end

    # @display param_broadcast_behavior

    unchanging = map(r -> all(==((0, 0)), r), eachrow(param_broadcast_behavior))
    # @show unchanging

    base_expr = replace_func_params(statement.func_expr, Dict(i => statement.params[1, 1, i] for i in findall(unchanging)))

    if !(can_py_broadcast(base_expr))
        return nothing
    end

    if all(unchanging)
        lhs = convert(exporter, statement.lhs_expr, table.sheet_name)
        rhs = convert(exporter, base_expr, table.sheet_name)

        push!(lines, "$lhs = $rhs")
        return lines
        # return join(lines, "\n")
    end

    changing = .!unchanging

    changing_params = statement.params[1, 1, changing]
    changing_param_behavior = param_broadcast_behavior[changing, :]

    param_table_refs = TableRef.(changing_params)

    lhs_table_ref = TableRef(statement.lhs_expr)

    num_rows, num_cols, num_params = size(statement.params)



    println("$(convert(exporter, statement.lhs_expr, table.sheet_name)) could be broadcast!")

    lhs = convert(exporter, statement.lhs_expr, table.sheet_name)

    param_index_strs = Dict{Int, String}()
    for (i, table_ref) in zip(findall(changing), param_table_refs)
        param = statement.params[1, 1, i]
        behavior = param_broadcast_behavior[i, :]
        # push!(lines, "# param = $(param) -- behavior = $(behavior)")
        row_behavior, col_behavior = behavior
        param_table = get_table(table_ref)
        row_idx = get_rows(table_ref)

        row_is_num = false

        row_loc = if row_behavior == (0, 0)
            row_names = row_name.(Ref(param_table), row_idx)
            if length(row_idx) == size(param_table)[1]
                ":"
            elseif length(row_idx) == 1 && num_rows == 1
                "$(repr(row_names))"
            elseif length(row_idx) == 1
                "$(repr(row_names)):$(repr(row_names))"
            else
                "$(repr(row_name(param_table, first(row_names)))):$(repr(row_name(param_table, last(row_names))))"
                # "$(repr(first(row_names))):$(repr(last(row_names)))"
            end
        elseif row_behavior == (1, 0) && length(row_idx) == 1
            # @show row_idx size(param_table)
            "$(repr(row_name(param_table, row_idx))):$(repr(row_name(param_table, row_idx + num_rows - 1)))"
        else
            @show param behavior row_idx
            println("trying to broadcast, but param row behavior wasn't expected")
            return nothing
        end

        col_idx = get_cols(table_ref)

        col_is_num = false
        col_loc = if col_behavior == (0, 0)
            col_names = column_name.(Ref(param_table), col_idx)
            if length(col_idx) == 1 && num_cols == 1
                "$(repr(col_names))"
            elseif length(col_idx) == size(param_table)[2]
                ":"
            else
                # "$(repr(first(col_names))):$(repr(last(col_names)))"
                "$(repr(column_name(param_table, first(col_idx)))):$(repr(column_name(param_table, last(col_idx))))"
            end
        elseif col_behavior == (0, 1) && length(col_idx) == 1
            "$(repr(column_name(param_table, col_idx))):$(repr(column_name(param_table, col_idx + num_cols - 1)))"
        else
            @show param behavior col_idx
            println("trying to broadcast, but param col behavior wasn't expected")
            return nothing
        end

        param_rows = length(row_idx)
        param_cols = length(col_idx)
        # push!(lines, "# col_loc = $col_loc")
        index_str = ".loc[$row_loc, $col_loc].values"

        param_index_strs[i] = getname(param_table) * index_str
    end

    function get_param_str_bcast(param_num, exporter, ctx)
        if changing[param_num]
            param_index_strs[param_num]
        else
            throw("Tried to get param_str for param_num $param_num, but it wasn't a changing param")
        end
    end

    typed_params = Dict{Int64, ExcelExpr}()
    for param_num in findall(changing)
        first_expr = statement.params[1, 1, param_num]
        # @show all_exprs
        param_type = get_type(first_expr, table.sheet_name, exporter.cell_types, exporter.named_values)
        typed_params[param_num] = ExcelExpr(:func_param, param_num, param_type)
    end
    base_expr = replace_func_params(base_expr, typed_params)

    custom_handler = CustomFuncParamHandler(get_param_str_bcast)
    custom_exporter = PythonExporter(exporter.wb, exporter.var_names, exporter.tables, exporter.named_values, [custom_handler, exporter.handlers...], exporter.cell_types)

    rhs = convert(custom_exporter, base_expr, table.sheet_name)
    wrap_na_to_zero = false
    if base_expr.parts[1].head == :table_ref
        wrap_na_to_zero = true
    elseif base_expr.parts[1].head == :func_param
        wrap_na_to_zero = true
    end
    if wrap_na_to_zero
            rhs = "xl.na_to_zero($rhs)"
    end
    push!(lines, "$lhs = $rhs")


    # join(lines, "\n") * "\n" * make_assertion_string(exporter, statement, wb.xf)
    lines
end

function export_statement(exporter::PythonExporter, wb::ExcelWorkbook, statement::BroadcastedStatement)
    # @show size(params)

    table, lhs_row_idx, lhs_col_idx = @match statement.lhs_expr begin
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => (table, row_idx, col_idx)
        _ => (missing, missing, missing)
    end
    if is_transposed(table)
        (lhs_row_idx, lhs_col_idx) = (lhs_col_idx, lhs_row_idx)
    end
    # @show size(statement.params)

    num_rows, num_cols, num_params = size(statement.params)

    lines = Vector{String}()

    push!(lines, "# $(convert(exporter, statement.lhs_expr, table.sheet_name))")


    if isempty(statement.params) 
        param_coords = zeros(Int64, (0, 0, 0, 0))
    else
        param_coords = stack(get_param_cell_coords, statement.params)
    end
    param_broadcast_behavior = Matrix{Tuple{Int, Int}}(undef, num_params, 2)
    for i in axes(param_coords, 4)
        coords = @view param_coords[:, :, :, i]

        if any(ismissing, coords)
            continue
        end

        rows = @view coords[1, :, :]
        cols = @view coords[2, :, :]

        row_behavior = flatten_step_range_len.(identify_broadcast_behavior(rows))
        col_behavior = flatten_step_range_len.(identify_broadcast_behavior(cols))

        param_broadcast_behavior[i, :] .= (row_behavior, col_behavior)
    end

    # @display param_broadcast_behavior

    unchanging = map(r -> all(==((0, 0)), r), eachrow(param_broadcast_behavior))
    # @show unchanging

    base_expr = replace_func_params(statement.func_expr, Dict(i => statement.params[1, 1, i] for i in findall(unchanging)))

    # if all(unchanging)
    #     lhs = convert(exporter, statement.lhs_expr, table.sheet_name)
    #     rhs = convert(exporter, base_expr, table.sheet_name)

    #     push!(lines, "$lhs = $rhs")
    #     return join(lines, "\n")
    # end

    changing = .!unchanging

    changing_params = statement.params[1, 1, changing]
    changing_param_behavior = param_broadcast_behavior[changing, :]

    param_table_refs = TableRef.(changing_params)

    lhs_table_ref = TableRef(statement.lhs_expr)

    all_tables_the_same = all([get_table(lhs_table_ref)] .== get_table.(param_table_refs))
    all_start_row_the_same = all([first(get_rows(lhs_table_ref))] .== get_rows.(param_table_refs))
    # @show get_rows(lhs_table_ref)
    # @display get_rows.(param_table_refs)


    # push!(lines, "# all_tables_the_same = $all_tables_the_same")
    # push!(lines, "# all_start_row_the_same = $all_start_row_the_same")
    sheet = table.sheet_name 

    top_left = CellDependency(sheet, startcol(table) + first(lhs_col_idx) - 1, startrow(table) + first(lhs_row_idx) - 1)
    bottom_right = CellDependency(sheet, startcol(table) + last(lhs_col_idx) - 1, startrow(table) + last(lhs_row_idx) - 1)
    set_region = WorkbookRegion(top_left, bottom_right)
    push!(lines, "# num_rows = $num_rows, num_cols = $num_cols, range = $(set_region)")

    broadcast_lines = export_statement_broadcasted(exporter, wb, statement)
    if !isnothing(broadcast_lines)
        append!(lines, broadcast_lines)
        return join(lines, "\n") * "\n" * make_assertion_string(exporter, statement, wb.xf)
    end
    # if can_py_broadcast(base_expr)

    # end


    needs_index = !(all_tables_the_same && all_start_row_the_same && num_cols == 1 && num_rows >= 1)
    # if all_tables_the_same && all_start_row_the_same && num_cols == 1 && num_rows >= 1
    #     println("All param tables are the same")
    # end


    tbl_name = getname(table)
    lhs_expr = ExcelExpr(:table_ref, [table, lhs_row_idx, lhs_col_idx, (false, false), (false, false)])
    lhs = convert(exporter, lhs_expr, table.sheet_name)
    column_iterator = if length(lhs_col_idx) == size(table)[2]
        "$tbl_name.columns"
    else
        start = column_name(table, first(lhs_col_idx)) |> repr
        stop = column_name(table, last(lhs_col_idx)) |> repr
        "$tbl_name.loc[:, $start:$stop].columns"
    end
    col_loop = "for i, col in enumerate($column_iterator):"

    row_loop = if needs_index
        "for j, row in enumerate($(lhs).index):"
    else
        "for row in $(lhs).index:"
    end
    indent = 0
    if num_cols > 1
        push!(lines, "$(indent_str(col_loop, indent))")
        indent += 1
    end
    if num_rows > 1
        push!(lines, "$(indent_str(row_loop, indent))")
        indent += 1
    end


    param_index_strs = Dict{Int, String}()
    for (i, table_ref) in zip(findall(changing), param_table_refs)
        param = statement.params[1, 1, i]
        behavior = param_broadcast_behavior[i, :]
        # push!(lines, "# param = $(param) -- behavior = $(behavior)")
        row_behavior, col_behavior = behavior
        param_table = get_table(table_ref)
        row_idx = get_rows(table_ref)

        row_is_num = false

        row_loc = if row_behavior == (0, 0)
            row_names = row_name.(Ref(param_table), row_idx)
            if length(row_idx) == size(param_table)[1]
                ":"
            elseif length(row_idx) == 1
                "$(repr(row_names))"
            else
                "$(repr(first(row_names))):$(repr(last(row_names)))"
            end
        elseif param_table == get_table(lhs_table_ref) && first(get_rows(lhs_table_ref)) == get_rows(table_ref)
            "row"
        elseif get_row_names(param_table) === get_row_names(get_table(lhs_table_ref)) && first(get_rows(lhs_table_ref)) == get_rows(table_ref)
            println("BroadcastedStatement, indexing on row in a different table, because row names are equal")
            "row"
        else
            row_is_num = true
            make_index_str(["j", "i"], row_behavior, row_idx .- 1)
        end

        col_idx = get_cols(table_ref)

        col_is_num = false
        col_loc = if col_behavior == (0, 0)
            col_names = column_name.(Ref(param_table), col_idx)
            if length(col_idx) == 1
                "$(repr(col_names))"
            elseif length(col_idx) == size(param_table)[2]
                ":"
            else
                "$(repr(first(col_names))):$(repr(last(col_names)))"
            end
        elseif param_table == get_table(lhs_table_ref) && first(get_cols(lhs_table_ref)) == get_cols(table_ref)
            # println("BroadcastedStatement, indexing on col because tables are equal and column offsets are equal")
            "col"
        elseif get_column_names(param_table) === get_column_names(get_table(lhs_table_ref)) && first(get_cols(lhs_table_ref)) == get_cols(table_ref)
            println("BroadcastedStatement, indexing on col in a different table, because col names are equal")
            "col"
        else
            i_coeff, j_coeff = col_behavior
            col_is_num = true
            make_index_str(["j", "i"], col_behavior, col_idx .- 1)
        end

        param_rows = length(row_idx)
        param_cols = length(col_idx)
        # push!(lines, "# col_loc = $col_loc")
        index_str = @match (row_is_num, col_is_num) begin
            (false, false) => begin
                if param_rows == param_cols == 1
                    ".at[$row_loc, $col_loc]"
                else
                    ".loc[$row_loc, $col_loc]"
                end
            end
            (true, false) => begin
                if size(param_table)[2] == 1
                    ".iloc[$row_loc, 0]"
                else
                    ".loc[:, $col_loc].iloc[$row_loc]"
                end
            end
            (false, true) => begin
                if param_rows == 1 && (size(param_table)[1] != 1)
                    ".loc[$row_loc].iloc[$col_loc]"
                elseif param_rows ==1 && (size(param_table)[1] == 1)
                    ".iloc[0, $col_loc]"
                else
                    ".loc[$row_loc].iloc[:, $col_loc]"
                end
            end
            (true, true) => ".iloc[$row_loc, $col_loc]"
        end

        param_index_strs[i] = getname(param_table) * index_str
    end

    function get_param_str(param_num, exporter, ctx)
        if changing[param_num]
            param_index_strs[param_num]
        else
            throw("Tried to get param_str for param_num $param_num, but it wasn't a changing param")
        end
    end

    typed_params = Dict{Int64, ExcelExpr}()
    for param_num in findall(changing)
        first_expr = statement.params[1, 1, param_num]
        # @show all_exprs
        param_type = get_type(first_expr, table.sheet_name, exporter.cell_types, exporter.named_values)
        typed_params[param_num] = ExcelExpr(:func_param, param_num, param_type)
    end
    base_expr = replace_func_params(base_expr, typed_params)

    custom_handler = CustomFuncParamHandler(get_param_str)
    custom_exporter = PythonExporter(exporter.wb, exporter.var_names, exporter.tables, exporter.named_values, [custom_handler, exporter.handlers...], exporter.cell_types)
    # xlookup_to_indexing!(base_expr)
    # @show lhs_row_idx lhs_col_idx
    lhs_row_str = num_rows == 1 ? repr(row_name(table, first(lhs_row_idx))) : "row"
    lhs_col_str = num_cols == 1 ? repr(column_name(table, first(lhs_col_idx))) : "col"
    lhs = getname(table) * ".loc[$lhs_row_str, $lhs_col_str]"

    name_handler = ColRowNameHandler(lhs_row_str, lhs_col_str)
    rhs = convert(with_handler(custom_exporter, name_handler), base_expr, table.sheet_name)
    wrap_na_to_zero = false
    if base_expr.parts[1].head == :table_ref
        wrap_na_to_zero = true
    elseif base_expr.parts[1].head == :func_param
        wrap_na_to_zero = true
    end
    if wrap_na_to_zero
            rhs = "xl.na_to_zero($rhs)"
    end
    push!(lines, indent_str("$lhs = $rhs", indent))


    join(lines, "\n") * "\n" * make_assertion_string(exporter, statement, wb.xf)
end