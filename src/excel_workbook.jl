
struct MissingCell
end

struct ValueCell1
    cell::XLSX.Cell
    value
end

ValueCell = ValueCell1

struct FormulaCell2
    cell::XLSX.Cell
    expr::Union{ExcelExpr,Float64,Int64,String,Missing}
end

FormulaCell = FormulaCell2

CellTypes = Union{ValueCell,FormulaCell}

function get_expr(cell::FormulaCell)
    cell.expr
end

function get_expr(cell::ValueCell)
    cell.value
end

function get_expr(::MissingCell)
    missing
end



struct ExcelWorkbook2
    xf::XLSX.XLSXFile
    cell_dict::Dict{CellDependency,CellTypes}
    cell_dependencies::Dict{CellDependency,Vector{CellDependency}}
    key_values::Dict{String,Any}
end

ExcelWorkbook = ExcelWorkbook2

has_formula(cell) = false
has_formula(cell::XLSX.Cell) = !isempty(cell.formula)

is_ref_formula(c) = c.formula isa XLSX.ReferencedFormula

function lower_sheet_names(expr, current_sheet::AbstractString)
    return expr
end

function lower_sheet_names(expr::ExcelExpr, current_sheet::AbstractString)
    @match expr begin
        ExcelExpr(:cell_ref, (cell,)) => ExcelExpr(:cell_ref, cell, current_sheet)
        ExcelExpr(:cols, (columns,)) => begin
            println("Got column expr")
            @show expr
            ExcelExpr(:cols, current_sheet, columns)
        end
        ExcelExpr(:sheet_ref, (sheet_name, ref)) => lower_sheet_names(ref, sheet_name)
        ExcelExpr(op, args) => begin
            ExcelExpr(op, lower_sheet_names.(args, current_sheet)...)
        end
    end
end

function get_cell_dict(xl)
    cell_dict = Dict{CellDependency,CellTypes}()

    for sheet_name in XLSX.sheetnames(xl)
        sheet = xl[sheet_name]
        all_cells = filter(!isempty, get_all_cells(sheet))

        ref_cells_dict = Dict{Int64,CellDependency}()

        # Do referenced cell first, so that we can look them up later
        for ref_formula_cell in filter(is_ref_formula, all_cells)
            cell_dep = CellDependency(sheet_name, ref_formula_cell.ref.name)

            try
                # expression = toexpr(parse_formula(ref_formula_cell.formula.formula))
                # expression = FormulaParser.toexpr(ref_formula_cell.formula.formula)
                expression = toexpr(ref_formula_cell.formula.formula)
                cell_dict[cell_dep] = FormulaCell(ref_formula_cell, lower_sheet_names(expression, sheet_name))

                ref_cells_dict[ref_formula_cell.formula.id] = cell_dep

            catch e
                @show cell_dep
                @show ref_formula_cell.formula.formula
            end
        end

        for cell in filter(!is_ref_formula, all_cells)
            cell_dep = CellDependency(sheet_name, cell.ref.name)
            if has_formula(cell)
                if cell.formula isa XLSX.FormulaReference
                    formula = cell.formula
                    base_formula_ref = ref_cells_dict[formula.id]
                    base_formula_cell = cell_dict[base_formula_ref].cell

                    start_row = base_formula_cell.ref.row_number
                    start_col = base_formula_cell.ref.column_number
                    end_row = cell.ref.row_number
                    end_col = cell.ref.column_number
                    delta_x = end_col - start_col
                    delta_y = end_row - start_row

                    expression = offset(cell_dict[base_formula_ref].expr, delta_y, delta_x)
                    cell_dict[cell_dep] = FormulaCell(cell, expression)
                else
                    try
                        # expression = FormulaParser.toexpr(cell.formula.formula)
                        expression = toexpr(cell.formula.formula)
                        # expression = toexpr(parse_formula(cell.formula.formula))
                        cell_dict[cell_dep] = FormulaCell(cell, lower_sheet_names(expression, sheet_name))
                    catch e
                        # throw(e)
                        @show e
                        @show cell_dep
                        @show cell.formula.formula
                    end
                end
            else
                cell_dict[cell_dep] = ValueCell(cell, XLSX.getdata(sheet, cell))
            end
        end
    end

    cell_dict
end

function get_all_dependencies(cell_dict::Dict{CellDependency,CellTypes}, key_values)
    output = Dict{CellDependency,Vector{CellDependency}}()

    for (cell, content) in cell_dict
        if content isa FormulaCell
            try
                output[cell] = get_expr_dependencies(content.expr, key_values)
            catch e
                @show cell
                @show content
                # throw(e)
            end
        end
    end

    output
end

function get_extra_named_values!(xf::XLSX.XLSXFile)
    xroot = XLSX.xmlroot(xf, "xl/workbook.xml")
    @assert EzXML.nodename(xroot) == "workbook" "Malformed xl/workbook.xml. Root node name should be 'workbook'. Got '$(EzXML.nodename(xroot))'."

    # workbook to be parsed
    workbook = XLSX.get_workbook(xf)

    existing_keys = keys(workbook.workbook_names)
    # named ranges
    for node in EzXML.eachelement(xroot)
        if EzXML.nodename(node) == "definedNames"
            for defined_name_node in EzXML.eachelement(node)
                @assert EzXML.nodename(defined_name_node) == "definedName"
                defined_value_string = EzXML.nodecontent(defined_name_node)
                name = defined_name_node["name"]

                local defined_value::XLSX.DefinedNameValueTypes
                if !(name in existing_keys)
                    defined_value = defined_value_string
                else
                    continue
                end


                if haskey(defined_name_node, "localSheetId")
                    # is a Worksheet level name

                    # localSheetId is the 0-based index of the Worksheet in the order
                    # that it is displayed on screen.
                    # Which is the order of the elements under <sheets> element in workbook.xml .
                    localSheetId = parse(Int, defined_name_node["localSheetId"]) + 1
                    sheetId = workbook.sheets[localSheetId].sheetId
                    # println("worksheet_names ($sheetId, $name) = $defined_value")
                    workbook.worksheet_names[(sheetId, name)] = defined_value
                else
                    # is a Workbook level name
                    # println("worksheet_names $name = $defined_value")
                    workbook.workbook_names[name] = defined_value
                end
            end

            break
        end
    end

    nothing
end

function parse_workbook(filepath::AbstractString)
    println("Opening excel file...")
    @time xf = XLSX.readxlsx(filepath)
    get_extra_named_values!(xf)
    println("Getting cell dict...")
    @time cell_dict = get_cell_dict(xf)
    println("Getting cell dependencies...")

    # parsed_key_values = Dict((p[1] => lower_sheet_names(FormulaParser.toexpr(repr(p[2])), "")) for p in XLSX.get_workbook(xf).workbook_names)
    # parsed_key_values = Dict((p[1] => lower_sheet_names(toexpr(repr(p[2])), "")) for p in XLSX.get_workbook(xf).workbook_names)
    parsed_key_values = Dict((p[1] => lower_sheet_names(toexpr(string(p[2])), "")) for p in XLSX.get_workbook(xf).workbook_names)
    # @show keys(XLSX.get_workbook(xf).workbook_names)

    @time cell_dependencies = get_all_dependencies(cell_dict, parsed_key_values)

    ExcelWorkbook(xf, cell_dict, cell_dependencies, parsed_key_values)
end

function get_all_referenced_cells(workbook::ExcelWorkbook)
    unioned = Set(keys(workbook.cell_dependencies))
    for cells in values(workbook.cell_dependencies)
        union!(unioned, cells)
    end

    # collect(union(keys(workbook.cell_dependencies), values(workbook.cell_dependencies)...))
    collect(unioned)
end