
mutable struct TableStatement <: AbstractStatement
    lhs_expr::ExcelExpr
    assigned_vars::Vector{CellDependency}
    rhs_expr
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

        set_cells = get_set_cells(s)
        lhs_cell = set_cells[1]
        sheet = lhs_cell.sheet_name

        s.rhs_expr = insert_table_refs(s.rhs_expr, tables)

        lhs = s.assigned_var
        lhs_expr = ExcelExpr(:cell_ref, lhs.cell, sheet)
        lhs_expr = insert_table_refs(lhs_expr, tables)

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
    assigned_cells = string.(stmt.assigned_vars)
    print(io, "TableStatement($assigned_cells)")
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

    table, lhs_row_idx = @match statement.lhs_expr begin
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => (table, row_idx)
        _ => (missing, missing)
    end

    function replace_func_params(expr, params_dict)
        @match expr begin
            ExcelExpr(:func_param, [param_num,]) => get(params_dict, param_num, expr)
            ExcelExpr(head, args) => ExcelExpr(head, map(e -> replace_func_params(e, params_dict), args)...)
            _ => expr
        end
    end

    if statement.is_broadcast

        run_cells = sort(statement.assigned_vars)
        if contains_if(expr)
            function_expr, params = functionalize(expr)
            typed_params = Dict{Int64,ExcelExpr}()
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
            # @show lhs
            # @info "export table statement" sheet expr
            rhs = convert(exporter, expr, sheet)
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
        rhs = convert(exporter, expr, sheet)
        @assert length(statement.assigned_vars) == 1
        "$lhs = $rhs # $(cell_ref.sheet_name) $(cell_ref.cell) $row_str\n"
    end
end
