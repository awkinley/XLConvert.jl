
using AutoHashEquals
@auto_hash_equals struct ExcelTable5
    sheet_name::String
    table_name::String
    top_left::String
    bottom_right::String
    column_names_range::String
    row_names_range::Union{String,Missing}

    _col_names
    _row_names
end

ExcelTable = ExcelTable5

getname(table::ExcelTable) = "tab_$(normalize_var_name(table.sheet_name))_$(normalize_var_name(table.table_name))"
Base.size(table::ExcelTable) = (endrow(table) - startrow(table) + 1, endcol(table) - startcol(table) + 1)

startcol(table::ExcelTable) = parse_cell(table.top_left)[1]
startrow(table::ExcelTable) = parse_cell(table.top_left)[2]
endcol(table::ExcelTable) = parse_cell(table.bottom_right)[1]
endrow(table::ExcelTable) = parse_cell(table.bottom_right)[2]
function column_name(table::ExcelTable, col_idx)
    table._col_names[col_idx]
    # string(xf[table.sheet_name][table.column_names_range][col_idx])
end
function row_name(table::ExcelTable, row_idx)
    if ismissing(table._row_names)
        "$row_idx"
    else
        table._row_names[row_idx]
    end
    # string(xf[table.sheet_name][table.column_names_range][col_idx])
end

# Base.in((row, col), table::ExcelTable) = (row >= startrow(table) && row <= endrow(table) && col >= startcol(table) && col <= endcol(table))
function Base.in((row, col), table::ExcelTable) 
    start_c, start_r = parse_cell(table.top_left)
    end_c, end_r = parse_cell(table.bottom_right)

    (start_r <= row <= end_r) && (start_c <= col <= end_c)
    # (row >= startrow(table) && row <= endrow(table) && col >= startcol(table) && col <= endcol(table))
end
Base.in(cell::CellDependency, table::ExcelTable) = (cell.sheet_name == table.sheet_name) && ((rownum(cell), colnum(cell)) in table)

function Base.show(io::IO, table::ExcelTable)
    print(io, "ExcelTable($(getname(table)))")
end