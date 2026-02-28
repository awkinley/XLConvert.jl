@auto_hash_equals struct WorkbookRegion
    first::CellDependency
    last::CellDependency
end

function WorkbookRegion(sheet::AbstractString, first_cell::AbstractString, last_cell::AbstractString)
    WorkbookRegion(CellDependency(sheet, first_cell), CellDependency(sheet, last_cell))
end

function cells(region::WorkbookRegion)
    start_c, start_r = start_coord(region)
    end_c, end_r = end_coord(region)

    sheet = region.first.sheet_name

    [CellDependency(sheet, c, r) for c in start_c:end_c, r in start_r:end_r]
end

function Base.getindex(region::WorkbookRegion, row::Integer, col::Integer)
    @assert all((row, col) .<= size(region))

    offset(region.first, row - 1, col - 1)
end
"""
(start_column, start_row) of the range
"""
start_coord(region::WorkbookRegion) = XLConvert.get_coords(region.first)
"""
(end_column, end_row) of the range
"""
end_coord(region::WorkbookRegion) = XLConvert.get_coords(region.last)

function Base.size(region::WorkbookRegion)
    (first_col, first_row) = start_coord(region)
    (last_col, last_row) = end_coord(region)
    num_rows = last_row - first_row + 1
    num_cols = last_col - first_col + 1

    (num_rows, num_cols)
end

function Base.in(cell::CellDependency, region::WorkbookRegion) 
    if cell.sheet_name != region.first.sheet_name
        return false
    end

    cell_coords = get_coords(cell)

    all(cell_coords .>= start_coord(region)) && all(cell_coords .<= end_coord(region)) 
end

Base.in(in_region::WorkbookRegion, region::WorkbookRegion)  = (in_region.first in region) && (in_region.last in region)

function Base.show(io::IO, region::WorkbookRegion)
    first = region.first
    last = region.last
    if first.sheet_name == last.sheet_name
        sheet = first.sheet_name
        (num_rows, num_cols) = size(region)
        print(io, "WorkbookRegion($(sheet)!$(first.cell):$(last.cell), $(num_rows) x $(num_cols))")
    else
        print(io, "WorkbookRegion($(region.first):$(region.last)")
    end
end

function get_cell_values(region::WorkbookRegion, xf)
    xf[region.first.sheet_name][string(region.first.cell, ":", region.last.cell)]
end