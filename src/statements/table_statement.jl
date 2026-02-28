
mutable struct TableStatement <: AbstractStatement
    lhs_expr::ExcelExpr
    assigned_vars::Vector{CellDependency}
    rhs_expr::Any
    rhs_dependencies::Vector{CellDependency}
    is_broadcast::Bool
end


get_cell_deps(stmt::TableStatement) = stmt.rhs_dependencies
get_set_cells(stmt::TableStatement) = stmt.assigned_vars
apply_expr_transform!(stmt::TableStatement, transform) = stmt.rhs_expr = transform(stmt, stmt.rhs_expr)

get_set_table(stmt::TableStatement) = stmt.lhs_expr.args[1]

function table_ref_transform!(statements::Vector{AbstractStatement}, tables::Vector{ExcelTable})
    # function transform(stmt::AbstractStatement, expr)
    #     set_cells = get_set_cells(stmt)
    #     set_sheets = map(c -> c.sheet_name, set_cells) |> unique |> collect
    #     if length(set_sheets) != 1
    #         throw("Trying to transform a statement that sets cells in multiple sheets is not possible!")
    #     end


    #     insert_table_refs(expr, set_sheets[1], tables)
    # end

    for i in eachindex(statements)
        s = statements[i]
        if s isa OutputStatement
            continue
        end
        @assert s isa StandardStatement

        lhs = s.assigned_var
        lhs_expr = ExcelExpr(:cell_ref, lhs.cell, lhs.sheet_name)
        lhs_expr = insert_table_refs(lhs_expr, tables)
        s.rhs_expr = insert_table_refs(s.rhs_expr, tables)

        if lhs_expr.head == :table_ref
            statements[i] = TableStatement(lhs_expr, [lhs], s.rhs_expr, s.rhs_dependencies, false)
        end
    end
end

function to_string(exporter, statement::TableStatement)
    cell_ref = statement.assigned_vars[1]
    sheet = cell_ref.sheet_name
    lhs = convert(exporter, statement.lhs_expr, sheet)

    "TableStatement(lhs = $lhs)"
end
function Base.show(io::IO, stmt::TableStatement)
    assigned_cells = join(string.(stmt.assigned_vars), ", ")
    # assigned_cells = stmt.assigned_vars
    print(io, "TableStatement([$assigned_cells)]")
end


function export_statement(exporter::JuliaExporter, wb::ExcelWorkbook, statement::TableStatement)
    cell_ref = statement.assigned_vars[1]
    sheet = cell_ref.sheet_name

    lhs = try
        convert(exporter, statement.lhs_expr, sheet)
    catch e
        @show statement.lhs_expr
        throw(e)
    end
    expr = statement.rhs_expr

    table, lhs_row_idx, lhs_col_idx = @match statement.lhs_expr begin
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => (table, row_idx, col_idx)
        _ => (missing, missing, missing)
    end
    if is_transposed(table)
        (lhs_row_idx, lhs_col_idx) = (lhs_col_idx, lhs_row_idx)
    end

    function replace_func_params(expr, params_dict)
        @match expr begin
            ExcelExpr(:func_param, [param_num]) => get(params_dict, param_num, expr)
            ExcelExpr(head, args) => ExcelExpr(head, map(e -> replace_func_params(e, params_dict), args)...)
            _ => expr
        end
    end

    if statement.is_broadcast

        run_cells = sort(statement.assigned_vars)
        if contains_if(expr)
            function_expr, params = functionalize(expr)
            typed_params = Dict{Int64, ExcelExpr}()
            for param_num in eachindex(params)
                param = params[param_num]
                param_type = Any
                try
                    param_type = get_type(param, sheet, exporter.cell_types, exporter.named_values)
                catch
                end
                typed_params[param_num] = ExcelExpr(:func_param, param_num, param_type)
            end

            function_expr = replace_func_params(function_expr, typed_params)

            function make_function_string(function_name, function_expr, num_params)
                params_str = join(["param_$i" for i in 1:num_params], ", ")
                expr_str = convert(exporter, function_expr, sheet)
                """
                function $function_name($params_str)
                    $expr_str
                end"""
            end
            function_name = "func_$(normalize_var_name(sheet))_$(run_cells[1].cell)_$(run_cells[end].cell)"
            func_str = make_function_string(function_name, function_expr, length(params))
            # convert(exporter, function_expr, sheet)
            # params_strings = [xl_expr_to_julia(param_expr, ctx, var_names, tables) for param_expr in params]
            params_strings = [convert(exporter, param_expr, sheet) for param_expr in params]
            func_params = join(params_strings, ", ")
            rhs = "$function_name($(func_params))"
            # line = "$func_str@. $lhs = $rhs\n"
            """
            $func_str
            @. $lhs = $rhs
            """
        else
            try
                rhs = convert(exporter, expr, sheet)
            catch e
                println("Failed to convert table rhs expr")
                @show lhs
                # @info "export table statement" sheet expr
                show(stdout, "text/plain", expr)
                throw(e)
            end
            """
            # $(to_string(run_cells[1])):$(to_string(run_cells[end]))
            @. $lhs = $rhs
            """
        end
    else
        row_str = ""
        if !ismissing(table) && !ismissing(lhs_row_idx)
            row_str = "Row: $(row_name(table, lhs_row_idx))"
        end
        try
            rhs = convert(exporter, expr, sheet)
        catch e
            println("Failed to convert table rhs expr")
            @show statement.assigned_vars
            # @show expr
            show(stdout, "text/plain", expr)
            @show e
            throw(e)
        end
        @assert length(statement.assigned_vars) == 1
        xf = wb.xf
        cell_ref = get_set_cells(statement)[1]

        # "@assert xl_compare($lhs, $(repr(xf[string(cell_ref.sheet_name)][cell_ref.cell]))) # $(to_string(cell_ref))"
        # "$lhs = $rhs # $(cell_ref.sheet_name) $(cell_ref.cell) $row_str\n"
        # """
        # $lhs = $rhs # $(cell_ref.sheet_name) $(cell_ref.cell) $row_str
        # @assert xl_compare($lhs, $(repr(xf[string(cell_ref.sheet_name)][cell_ref.cell]))) # $(to_string(cell_ref))
        # """
        """
        $lhs = $rhs # $(cell_ref.sheet_name) $(cell_ref.cell) $row_str
        """
    end
end

function param_is_scalar(param_expr)
    @match param_expr begin
        ExcelExpr(:cell_ref, [cell, sheet]) => true
        ExcelExpr(:named_range, [name]) => true
        ExcelExpr(:sheet_ref, [sheet_name, ref]) => param_is_scalar(ref)
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => begin
            if is_transposed(table)
                (row_idx, col_idx) = (col_idx, row_idx)
            end
            length(row_idx) == 1 && length(col_idx) == 1
        end
        ExcelExpr(:range, [ExcelExpr(:cell_ref, [cell, sheet]), ExcelExpr(:cell_ref, [cell, sheet])]) => true
        ExcelExpr(:range, args) => false
        _ => false
    end
end

function export_for_loop(exporter::PythonExporter, wb::ExcelWorkbook, statement::TableStatement, expr, sheet)
    formula_str = get_formula_str(statement, wb.xf)

    run_cells = sort(statement.assigned_vars)
    function_expr, params = functionalize(expr)
    rhs = convert(exporter, function_expr, sheet)

    loop_len = length(run_cells)
    loop_id = "$(normalize_var_name(sheet))_$(run_cells[1].cell)_$(run_cells[end].cell)"
    param_names = ["__tbl_$(loop_id)_v$(i)" for i in eachindex(params)]

    table, lhs_row_idx, lhs_col_idx = @match statement.lhs_expr begin
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => (table, row_idx, col_idx)
        _ => (missing, missing, missing)
    end
    if is_transposed(table)
        (lhs_row_idx, lhs_col_idx) = (lhs_col_idx, lhs_row_idx)
    end

    function replace_func_params(expr, params_dict)
        @match expr begin
            ExcelExpr(:func_param, [param_num]) => get(params_dict, param_num, expr)
            ExcelExpr(head, args) => ExcelExpr(head, map(e -> replace_func_params(e, params_dict), args)...)
            _ => expr
        end
    end

    param_lines = String[]
    for param_num in eachindex(params)
        param_expr = params[param_num]
        param_src = convert(exporter, param_expr, sheet)
        param_value_src = if param_is_scalar(param_expr)
            "np.repeat($param_src, $loop_len)"
        else
            "np.asarray($param_src).reshape(-1, order='F')"
        end
        push!(param_lines, "$(param_names[param_num]) = $param_value_src")
        rhs = replace(rhs, Regex("\\bparam_$(param_num)\\b") => "$(param_names[param_num])[i]")
    end

    if ismissing(table)
        """
        # $(to_string(run_cells[1])):$(to_string(run_cells[end]))
        $lhs = $rhs
        """
    else
        out_rows = String[]
        out_cols = String[]
        for c in run_cells
            row_idx = rownum(c) - startrow(table) + 1
            col_idx = colnum(c) - startcol(table) + 1
            if is_transposed(table)
                (row_idx, col_idx) = (col_idx, row_idx)
            end
            push!(out_rows, convert(exporter, row_name(table, row_idx), sheet))
            push!(out_cols, convert(exporter, string(column_name(table, col_idx)), sheet))
        end

        rows_var = "__tbl_$(loop_id)_rows"
        cols_var = "__tbl_$(loop_id)_cols"
        param_defs = isempty(param_lines) ? "" : join(param_lines, "\n")
        rows_src = join(out_rows, ", ")
        cols_src = join(out_cols, ", ")

        out = "$param_defs\n"
        loc_row = if length(lhs_row_idx) > 1
            out *= "$(rows_var) = [$rows_src]\n"
            "$(rows_var)[i]"
        else
            first(out_rows)
        end

        loc_col = if length(lhs_col_idx) > 1
            out *= "$(cols_var) = [$cols_src]\n"
            "$(cols_var)[i]"
        else
            first(out_cols)
        end
        out *= """
        # =$formula_str
        for i in range($loop_len):
            $(getname(table)).loc[$loc_row, $loc_col] = $rhs
        """

        out
    end
end

function get_formula_str(stmt::TableStatement, xf)
    cell_ref = stmt.assigned_vars[1]

    cell = getcell(xf, cell_ref)
    if !isempty(cell) && !(cell.formula isa XLSX.FormulaReference)
        replace(cell.formula.formula, "\n" => "\n# ")
    else
        ""
    end
end

function make_assertion_string(exporter::PythonExporter, statement::TableStatement, xf)
    table, lhs_row_idx, lhs_col_idx = @match statement.lhs_expr begin
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => (table, row_idx, col_idx)
        _ => (missing, missing, missing)
    end
    if is_transposed(table)
        (lhs_row_idx, lhs_col_idx) = (lhs_col_idx, lhs_row_idx)
    end

    sheet = string(table.sheet_name)

    out = ""

    for r in lhs_row_idx, c in lhs_col_idx
        cell = CellDependency(sheet, startcol(table) + c - 1, startrow(table) + r - 1)
        cell_value = xf[sheet][cell.cell]

        value_str = convert(exporter, cell_value, sheet)

        lhs_expr = ExcelExpr(:table_ref, table, r, c, (false, false), (false, false))
        lhs = convert(exporter, lhs_expr,sheet)
        if ismissing(cell_value)
            out *= "# assert xl.compare($lhs, $(value_str))\n"
        else
            out *= "assert xl.compare($lhs, $(value_str))\n"
        end

    end

    out
end

function export_statement(exporter::PythonExporter, wb::ExcelWorkbook, statement::TableStatement)
    cell_ref = statement.assigned_vars[1]
    sheet = cell_ref.sheet_name

    lhs = try
        convert(exporter, statement.lhs_expr, sheet)
    catch e
        @show statement.lhs_expr
        throw(e)
    end
    expr = statement.rhs_expr
    # xlookup_to_indexing!(expr)

    table, lhs_row_idx, lhs_col_idx = @match statement.lhs_expr begin
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => (table, row_idx, col_idx)
        _ => (missing, missing, missing)
    end
    if is_transposed(table)
        (lhs_row_idx, lhs_col_idx) = (lhs_col_idx, lhs_row_idx)
    end

    xf = wb.xf

    if statement.is_broadcast
        export_for_loop(exporter, wb, statement, expr, sheet) * make_assertion_string(exporter, statement, xf)
        # if contains_if(expr)
        #     export_for_loop(exporter, wb, statement, expr, sheet) * make_assertion_string(exporter, statement, xf)
        # else
        #     try
        #         rhs = convert(exporter, expr, sheet)
        #     catch e
        #         println("Failed to convert table rhs expr")
        #         @show lhs
        #         @display expr
        #         throw(e)
        #     end

        #     run_cells = sort(statement.assigned_vars)

        #     # when we're just directly copying data, do a fillna with zero,
        #     # which seems to basically be what excel does
        #     if expr.parts[1].head == :table_ref
        #         rhs = "as_array($rhs.fillna(0))"
        #     elseif length(lhs_row_idx) == 1 && length(lhs_col_idx) > 1
        #         rhs = "as_array($rhs)"
        #     end

        #     assertion_lines = make_assertion_string(exporter, statement, xf)

        #     """
        #     # $(to_string(run_cells[1])):$(to_string(run_cells[end]))
        #     # =$(get_formula_str(statement, xf))
        #     $lhs = $rhs
        #     $assertion_lines
        #     """
        # end
    else
        row_str = ""
        if !ismissing(table) && !ismissing(lhs_row_idx)
            row_str = "Row: $(row_name(table, lhs_row_idx))"
        end
        name_handler = ColRowNameHandler(repr(row_name(table, first(lhs_row_idx))), repr(column_name(table, first(lhs_col_idx))))
        rhs = try
            convert(with_handler(exporter, name_handler), expr, sheet)
        catch e
            println("Failed to convert table rhs expr")
            @show statement.assigned_vars
            @display expr
            @show e
            throw(e)
        end
        @assert length(statement.assigned_vars) == 1

        # "@assert xl_compare($lhs, $(repr(xf[string(cell_ref.sheet_name)][cell_ref.cell]))) # $(to_string(cell_ref))"
        # "$lhs = $rhs # $(cell_ref.sheet_name) $(cell_ref.cell) $row_str\n"
        cell_value = xf[string(cell_ref.sheet_name)][cell_ref.cell]
        value_str = convert(exporter, cell_value, sheet)

        wrap_na_to_zero = false
        if expr isa FlatExpr
            if expr.parts[1].head == :table_ref
                wrap_na_to_zero = true
            elseif expr.parts[1].head == :cell_ref
                wrap_na_to_zero = true
            end
        end
        if wrap_na_to_zero
                rhs = "xl.na_to_zero($rhs)"
        end

        """
        # =$(get_formula_str(statement, xf))
        $lhs = $rhs # $(cell_ref.sheet_name) $(cell_ref.cell) $row_str
        $(make_assertion_string(exporter, statement, xf))
        """
        # """
        # $lhs = $rhs # $(cell_ref.sheet_name) $(cell_ref.cell) $row_str
        # """
    end
end
