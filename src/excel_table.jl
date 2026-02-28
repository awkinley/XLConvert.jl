
using AutoHashEquals
struct ExcelTable
    sheet_name::String
    table_name::String
    top_left::String
    bottom_right::String
    column_names_range::String
    row_names_range::Union{String, Missing}
    is_transposed::Bool

    _col_names::Any
    _row_names::Any

    __startcol::Int64
    __startrow::Int64
    __endcol::Int64
    __endrow::Int64

    __name::String
end

function ExcelTable(sheet_name::String,
    table_name::String,
    top_left::String,
    bottom_right::String,
    column_names_range::String,
    row_names_range::Union{String, Missing},
    _col_names::Any,
    _row_names::Any,
)

    start_col, start_row = parse_cell(top_left)
    end_col, end_row = parse_cell(bottom_right)
    name = "tab_$(normalize_var_name(sheet_name))_$(normalize_var_name(table_name))"

    ExcelTable(sheet_name, table_name, top_left, bottom_right, column_names_range, row_names_range, false, _col_names, _row_names, start_col, start_row, end_col, end_row, name)

end

function transpose(tbl::ExcelTable)
    ExcelTable(
        tbl.sheet_name,
        tbl.table_name,
        tbl.top_left,
        tbl.bottom_right,
        tbl.column_names_range,
        tbl.row_names_range,
        true,
        tbl._col_names,
        tbl._row_names,
        tbl.__startcol,
        tbl.__startrow,
        tbl.__endcol,
        tbl.__endrow,
        tbl.__name,
    )
end

is_transposed(tbl::ExcelTable) = tbl.is_transposed

function Base.:(==)(a::ExcelTable, b::ExcelTable)
    a.sheet_name == b.sheet_name && a.top_left == b.top_left && a.bottom_right == b.bottom_right
end

function Base.hash(a::ExcelTable)
    hash((a.sheet_name, a.top_left, a.bottom_right))
end
# ExcelTable = ExcelTable5

# getname(table::ExcelTable) = "tab_$(normalize_var_name(table.sheet_name))_$(normalize_var_name(table.table_name))"
getname(table::ExcelTable) = table.__name
Base.size(table::ExcelTable) = (endrow(table) - startrow(table) + 1, endcol(table) - startcol(table) + 1)

startcol(table::ExcelTable) = table.__startcol
startrow(table::ExcelTable) = table.__startrow
endcol(table::ExcelTable) = table.__endcol
endrow(table::ExcelTable) = table.__endrow

# startcol(table::ExcelTable) = parse_cell(table.top_left)[1]
# startrow(table::ExcelTable) = parse_cell(table.top_left)[2]
# endcol(table::ExcelTable) = parse_cell(table.bottom_right)[1]
# endrow(table::ExcelTable) = parse_cell(table.bottom_right)[2]
function column_name(table::ExcelTable, col_idx)
    if col_idx > length(table._col_names)
        @show table
        @show table.top_left table.bottom_right
        @show table._col_names
    end
    table._col_names[col_idx]
    # string(xf[table.sheet_name][table.column_names_range][col_idx])
end
function row_name(table::ExcelTable, row_idx)
    if ismissing(table._row_names)
        row_idx
    else
        table._row_names[row_idx]
    end
    # string(xf[table.sheet_name][table.column_names_range][col_idx])
end

get_column_names(table::ExcelTable) = table._col_names
get_row_names(table::ExcelTable) = table._row_names

function region(table::ExcelTable)
    sheet = table.sheet_name
    WorkbookRegion(CellDependency(sheet, table.__startcol, table.__startrow), CellDependency(sheet, table.__endcol, table.__endrow))
end

function column_name_region(table::ExcelTable)
    if ismissing(table.column_names_range) || isempty(table.column_names_range)
        return nothing
    end

    first, last = split(table.column_names_range, ":")
    WorkbookRegion(table.sheet_name, first, last)
end

function row_name_region(table::ExcelTable)
    if ismissing(table.row_names_range) || isempty(table.row_names_range)
        return nothing
    end

    first, last = split(table.row_names_range, ":")
    WorkbookRegion(table.sheet_name, first, last)
end

# Base.in((row, col), table::ExcelTable) = (row >= startrow(table) && row <= endrow(table) && col >= startcol(table) && col <= endcol(table))
function Base.in((row, col), table::ExcelTable)
    (startrow(table) <= row <= endrow(table)) && (startcol(table) <= col <= endcol(table))
end
Base.in(cell::CellDependency, table::ExcelTable) = ((rownum(cell), colnum(cell)) in table) && (cell.sheet_name == table.sheet_name)

function Base.show(io::IO, table::ExcelTable)
    print(io, "ExcelTable($(getname(table)))")
end
