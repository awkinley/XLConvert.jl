@auto_hash_equals struct CellDependency
    sheet_name::String
    cell::String

    function CellDependency(sheet_name, cell)
        new_cell = replace(cell, "\$" => "")
        new(sheet_name, new_cell)
    end
end

function Base.isless(a::CellDependency, b::CellDependency)
    as_tuple = c -> (c.sheet_name, colnum(c), rownum(c))
    Base.isless(as_tuple(a), as_tuple(b))
end



cell_parse_rgx = r"[$]?([A-Z]+)[$]?([0-9]+)"
function offset(cell::CellDependency, rows::Int, cols::Int)
    cell_match = match(cell_parse_rgx, cell.cell)
    @assert cell_match.match == cell.cell "Cell didn't parse properly"
    col_str = cell_match[1]
    row_str = cell_match[2]

    new_col_num = XLSX.decode_column_number(col_str[1:end]) + cols
    new_row_num = parse(Int, row_str) + rows
    if new_col_num < 1 || new_col_num < 1
        return missing
    end
    new_col = XLSX.encode_column_number(new_col_num)
    new_row = string(new_row_num)

    CellDependency(cell.sheet_name, string(new_col, new_row))
end

get_coords(cell::CellDependency) = parse_cell(cell.cell)

function rownum(cell::CellDependency)
    cell_match = match(cell_parse_rgx, cell.cell)
    @assert cell_match.match == cell.cell "Cell didn't parse properly"
    row_str = cell_match[2]
    parse(Int, row_str)
end

function colnum(cell::CellDependency)
    cell_match = match(cell_parse_rgx, cell.cell)
    @assert cell_match.match == cell.cell "Cell didn't parse properly"
    col_str = cell_match[1]
    XLSX.decode_column_number(col_str[1:end])
end

function to_string(cell_ref::CellDependency)
    "\"$(cell_ref.sheet_name)!$(cell_ref.cell)\""
end