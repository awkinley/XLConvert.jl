
function normalize_var_name(name)
    replace(name,
        ' ' => '_',
        '\'' => "",
        '\"' => "",
        '.' => "",
        ',' => "",
        '+' => "_plus_",
        "-" => "_",
        "–" => "_",
        '/' => "_per_",
        '*' => "_",
        '^' => "",
        '%' => "pcnt",
        '&' => "_and_",
        '$' => "_dollar",
        '=' => "",
        '(' => "",
        ')' => "",
        '#' => "num",
        ';' => "",
        ':' => "_",
        '?' => "_",
    )
end

function simple_variable_name(cell_ref::CellDependency)
    # new_sheet_name = replace(cell_ref.sheet_name, ' ' => '_')
    "s_$(normalize_var_name(cell_ref.sheet_name))_$(normalize_var_name(cell_ref.cell))"
end


function get_name_to_left(cell_dependency::CellDependency, xf)
    to_left = offset(cell_dependency, 0, -1)
    if ismissing(to_left)
        return missing
    end
    sheet = xf[String(to_left.sheet_name)]
    value = sheet[to_left.cell]
    # left_cell = XLSX.getcell(sheet, to_left.cell)
    if value isa AbstractString
        new_sheet_name = normalize_var_name(to_left.sheet_name)
        "c_$(new_sheet_name)_$(normalize_var_name(value))"
    else
        missing
    end
end

function get_name_to_left_no_sheet(cell_dependency::CellDependency, xf::XLSX.XLSXFile)
    to_left = offset(cell_dependency, 0, -1)
    if ismissing(to_left)
        return missing
    end
    sheet = xf[String(to_left.sheet_name)]
    c, r = parse_cell(to_left.cell)
    # value = sheet[to_left.cell]
    value = XLSX.getdata(sheet, r, c)
    if value isa AbstractString
        normalize_var_name(value)
    else
        missing
    end
end

function make_var_names_map(cell_dependencies::Vector{CellDependency}, wb::ExcelWorkbook)
    xf = wb.xf
    workbook = XLSX.get_workbook(xf)
    # named_values = Dict((p[1] => FormulaParser.toexpr(repr(p[2]))) for p in workbook.workbook_names)
    # named_values = Dict((p[1] => toexpr(repr(p[2]))) for p in workbook.workbook_names)
    named_values = wb.key_values
    named_cells = Dict()
    for (name, expr) in named_values
        referenced_cell = @match expr begin
            ExcelExpr(:sheet_ref, (sheet_name, ExcelExpr(:cell_ref, (cell,)))) => CellDependency(sheet_name, cell)
            _ => missing
        end
        if !ismissing(referenced_cell)
            named_cells[referenced_cell] = name
        end
    end

    output = Dict{CellDependency, String}()
    rev_dict = Dict{String, Vector{CellDependency}}()
    for cell in cell_dependencies
        if cell in keys(named_cells)
            name = named_cells[cell]
            output[cell] = name
            rev_dict[name] = push!(get(rev_dict, name, Vector{CellDependency}()), cell)
        else
            name_to_left = get_name_to_left_no_sheet(cell, xf)
            # name_available = name_to_left ∉ values(output)

            name = if !ismissing(name_to_left)
                name_to_left
            else
                simple_variable_name(cell)
            end

            # if !ismissing(name_to_left)
            #     if name_available
            #         name = name_to_left
            #         # output[cell] = name_to_left
            #     else
            #         name = name_to_left * "_" * normalize_var_name(cell.cell)
            #         # output[cell] = name_to_left * "_" * normalize_var_name(cell.cell)
            #     end
            # else
            #     name = simple_variable_name(cell)
            #     # output[cell] = simple_variable_name(cell)
            # end

            output[cell] = name
            rev_dict[name] = push!(get(rev_dict, name, Vector{CellDependency}()), cell)
        end
    end

    for (name, cells) in pairs(rev_dict)
        if length(cells) <= 1
            continue
        end

        for c in cells
            output[c] *= "_" * normalize_var_name(c.cell)
        end
    end
    # @show named_cells
    # ExcelContext(current_sheet,
    #     Dict([(name => xl[name]) for name in XLSX.sheetnames(xl)]),
    # )
    output
end

function make_var_names_map(cell_dependencies, xf)
    workbook = XLSX.get_workbook(xf)
    # named_values = Dict((p[1] => FormulaParser.toexpr(repr(p[2]))) for p in workbook.workbook_names)
    # named_values = Dict((p[1] => toexpr(repr(p[2]))) for p in workbook.workbook_names)
    named_values = Dict((p[1] => toexpr(repr(p[2]))) for p in workbook.workbook_names)
    named_cells = Dict()
    for (name, expr) in named_values
        referenced_cell = @match expr begin
            ExcelExpr(:sheet_ref, (sheet_name, ExcelExpr(:cell_ref, (cell,)))) => CellDependency(sheet_name, cell)
            _ => missing
        end
        if !ismissing(referenced_cell)
            named_cells[referenced_cell] = name
        end
    end

    output = Dict()
    rev_dict = Dict()
    for cell in cell_dependencies
        if cell in keys(named_cells)
            name = named_cells[cell]
            output[cell] = name
            rev_dict[name] = push!(get(rev_dict, name, []), cell)
        else
            name_to_left = get_name_to_left_no_sheet(cell, xf)
            # name_available = name_to_left ∉ values(output)

            name = if !ismissing(name_to_left)
                name_to_left
            else
                simple_variable_name(cell)
            end

            # if !ismissing(name_to_left)
            #     if name_available
            #         name = name_to_left
            #         # output[cell] = name_to_left
            #     else
            #         name = name_to_left * "_" * normalize_var_name(cell.cell)
            #         # output[cell] = name_to_left * "_" * normalize_var_name(cell.cell)
            #     end
            # else
            #     name = simple_variable_name(cell)
            #     # output[cell] = simple_variable_name(cell)
            # end

            output[cell] = name
            rev_dict[name] = push!(get(rev_dict, name, []), cell)
        end
    end

    for (name, cells) in pairs(rev_dict)
        if length(cells) <= 1
            continue
        end

        for c in cells
            output[c] *= "_" * normalize_var_name(c.cell)
        end
    end
    # @show named_cells
    # ExcelContext(current_sheet,
    #     Dict([(name => xl[name]) for name in XLSX.sheetnames(xl)]),
    # )
    output
end