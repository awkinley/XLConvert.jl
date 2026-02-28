@auto_hash_equals struct CellDependency
    sheet_name::String
    col::Int64
    row::Int64
end

function cell_str(cell::CellDependency)
    XLSX.encode_column_number(cell.col) * string(cell.row)
end

function CellDependency(sheet_name, cell::AbstractString)
    col, row = parse_cell(cell)
    CellDependency(sheet_name, col, row)
end

function Base.getproperty(cell::CellDependency, sym::Symbol)
    if sym === :cell
        cell_str(cell)
    else # fallback to getfield
        getfield(cell, sym)
    end
end

function Base.isless(a::CellDependency, b::CellDependency)
    as_tuple = c -> (c.sheet_name, colnum(c), rownum(c))
    Base.isless(as_tuple(a), as_tuple(b))
end

function Base.show(io::IO, cell_dep::CellDependency)
    print(io, "CellDep(", cell_dep.sheet_name, "!", cell_dep.cell, ")")
end



cell_parse_rgx = r"[$]?([A-Z]+)[$]?([0-9]+)"
function offset(cell::CellDependency, rows::Int, cols::Int)
    # new_cell = offset_cell_str(cell.cell, rows, cols, false)
    # isnothing(new_cell) && return missing

    # CellDependency(cell.sheet_name, new_cell)
    new_col = cell.col + cols
    new_row = cell.row + rows
    if new_col < 1 || new_row < 1
        return missing
    end

    CellDependency(cell.sheet_name, new_col, new_row)
    # cell_match = match(cell_parse_rgx, cell.cell)
    # @assert cell_match.match == cell.cell "Cell didn't parse properly"
    # col_str = cell_match[1]
    # row_str = cell_match[2]

    # new_col_num = XLSX.decode_column_number(col_str[1:end]) + cols
    # new_row_num = parse(Int, row_str) + rows
    # if new_col_num < 1 || new_col_num < 1
    #     return missing
    # end
    # new_col = XLSX.encode_column_number(new_col_num)
    # new_row = string(new_row_num)
    # offset_cell_str(cell.cell, rows, cols)
    # CellDependency(cell.sheet_name, string(new_col, new_row))
end

get_coords(cell::CellDependency) = (cell.col, cell.row)

function rownum(cell::CellDependency)
    cell.row
end

function colnum(cell::CellDependency)
    cell.col
end

function to_string(cell_ref::CellDependency)
    "\"$(cell_ref.sheet_name)!$(cell_ref.cell)\""
end