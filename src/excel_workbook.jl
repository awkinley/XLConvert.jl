
struct MissingCell
end

struct ValueCell1
    cell::XLSX.Cell
    value::Any
end

ValueCell = ValueCell1

struct FormulaCell
    cell::XLSX.Cell
    expr::Union{FlatExpr, ExcelExpr, Float64, Int64, String, Missing}
end

# FormulaCell = FormulaCell2

CellTypes = Union{ValueCell, FormulaCell}

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
    cell_dict::Dict{CellDependency, CellTypes}
    cell_dependencies::Dict{CellDependency, Vector{CellDependency}}
    key_values::Dict{String, Any}
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
        ExcelExpr(:cell_ref, [cell]) => ExcelExpr(:cell_ref, cell, current_sheet)
        ExcelExpr(:cols, [columns]) => begin
            println("Got column expr")
            @show expr
            ExcelExpr(:cols, current_sheet, columns)
        end
        ExcelExpr(:sheet_ref, [sheet_name, ref]) => lower_sheet_names(ref, sheet_name)
        ExcelExpr(op, args) => begin
            ExcelExpr(op, lower_sheet_names.(args, current_sheet))
        end
    end
end

function lower_sheet_names!(expr, current_sheet::AbstractString)
end

function lower_sheet_names!(expr::ExcelExpr, current_sheet::AbstractString)
    @match expr begin
        ExcelExpr(:cell_ref, [cell]) => begin
            push!(expr.args, current_sheet)
        end
        ExcelExpr(:cols, [columns]) => begin
            println("Got column expr")
            @show expr
            push!(expr.args, current_sheet)
            # ExcelExpr(:cols, current_sheet, columns)
        end
        ExcelExpr(:sheet_ref, [sheet_name, ref]) => lower_sheet_names!(ref, sheet_name)
        ExcelExpr(op, args) => begin
            for arg in args
                lower_sheet_names!(arg, current_sheet)
            end
            # ExcelExpr(op, lower_sheet_names.(args, current_sheet))
        end
    end
end

function convert_cell(sheet, sheet_name, cell::XLSX.Cell)
    if has_formula(cell)
        @assert !(cell.formula isa XLSX.FormulaReference)
        expr = toexpr(cell.formula.formula)
        lower_sheet_names!(expr, sheet_name)
        FormulaCell(cell, expr)
    else
        ValueCell(cell, XLSX.getdata(sheet, cell))
    end
end

function offset_formula_cell(new_cell::XLSX.Cell, formula_cell::FormulaCell)
    cell_ref = formula_cell.cell.ref
    start_row = cell_ref.row_number
    start_col = cell_ref.column_number

    end_col, end_row = parse_cell(new_cell.ref.name)
    delta_x = end_col - start_col
    delta_y = end_row - start_row

    expression = offset(formula_cell.expr, delta_y, delta_x)
    FormulaCell(new_cell, expression)
end

function get_cell_dict(xl)
    cell_dict = Dict{CellDependency, CellTypes}()

    for sheet_name in XLSX.sheetnames(xl)
        sheet = xl[sheet_name]
        all_cells = filter(!isempty, get_all_cells(sheet))

        ref_cells_dict = Dict{Int64, CellDependency}()
        formula_refs_to_handle = Vector{XLSX.Cell}()

        for row in XLSX.eachrow(sheet)
            for cell in values(row.rowcells)
                cell_dep = CellDependency(sheet_name, cell.ref.name)

                if cell.formula isa XLSX.FormulaReference
                    formula = cell.formula
                    if formula.id in keys(ref_cells_dict)
                        formula_cell = cell_dict[ref_cells_dict[formula.id]]
                        cell_dict[cell_dep] = offset_formula_cell(cell, formula_cell)
                    else
                        # This would happen if a formula is referenced before
                        # it's defined, I don't know if this actually happens in
                        # practice. I haven't seen it happen, but haven't looked too hard.
                        push!(formula_refs_to_handle, cell)
                    end
                else
                    cell_dict[cell_dep] = convert_cell(sheet, sheet_name, cell)

                    if cell.formula isa XLSX.ReferencedFormula
                        ref_cells_dict[cell.formula.id] = cell_dep
                    end
                end

            end
        end

        # @show length(formula_refs_to_handle)
        for cell in formula_refs_to_handle
            cell_dep = CellDependency(sheet_name, cell.ref.name)
            formula_cell = cell_dict[ref_cells_dict[cell.formula.id]]
            cell_dict[cell_dep] = offset_formula_cell(cell, formula_cell)
        end
    end

    cell_dict
end

function get_all_dependencies(cell_dict::Dict{CellDependency, CellTypes}, key_values)
    output = Dict{CellDependency, Vector{CellDependency}}()

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