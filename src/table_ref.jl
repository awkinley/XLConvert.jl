
struct TableRef
    table::ExcelTable
    row::Any
    col::Any
end

get_table(t::TableRef) = t.table
get_rows(t::TableRef) = t.row
get_cols(t::TableRef) = t.col

function cell_dep(t::TableRef)
    @assert length(t.row) == 1 && length(t.col) == 1

    tbl = t.table
    CellDependency(tbl.sheet_name, startcol(tbl) + first(t.col) - 1, startrow(tbl) + first(t.row) - 1)
end

function TableRef(expr::ExcelExpr)
    @match expr begin
        ExcelExpr(:table_ref, [table, row_idx, col_idx, _, _]) => TableRef(table, row_idx, col_idx)
        _ => throw("tried to convert an invalid ExcelExpr to a TableRef $expr")
    end
end

"""
Table indexing system

Want to simplify handling and rendering of table indexing.
Also remove as much exporter-specific code in statements as possible.


Example of code I want to simplify:

for (i, table_ref) in zip(findall(changing), param_table_refs)
    param = statement.params[1, 1, i]
    behavior = param_broadcast_behavior[i, :]

    row_behavior, col_behavior = behavior
    param_table = get_table(table_ref)
    row_idx = get_rows(table_ref)

    row_is_num = false

    row_loc = if row_behavior == (0, 0)
        row_names = row_name.(Ref(param_table), row_idx)
        if length(row_idx) == size(param_table)[1]
            ":"
        elseif length(row_idx) == 1
            "\$(repr(row_names))"
        else
            "\$(repr(first(row_names))):\$(repr(last(row_names)))"
        end
    elseif param_table == get_table(lhs_table_ref) && first(get_rows(lhs_table_ref)) == get_rows(table_ref)
        "row"
    elseif get_row_names(param_table) === get_row_names(get_table(lhs_table_ref)) && first(get_rows(lhs_table_ref)) == get_rows(table_ref)
        println("BroadcastedStatement, indexing on row in a different table, because row names are equal")
        "row"
    else
        row_is_num = true
        make_index_str(["j", "i"], row_behavior, row_idx .- 1)
    end

    col_idx = get_cols(table_ref)

    col_is_num = false
    col_loc = if col_behavior == (0, 0)
        col_names = column_name.(Ref(param_table), col_idx)
        if length(col_idx) == 1
            "\$(repr(col_names))"
        elseif length(col_idx) == size(param_table)[2]
            ":"
        else
            "\$(repr(first(col_names))):\$(repr(last(col_names)))"
        end
    elseif param_table == get_table(lhs_table_ref) && first(get_cols(lhs_table_ref)) == get_cols(table_ref)
        # println("BroadcastedStatement, indexing on col because tables are equal and column offsets are equal")
        "col"
    elseif get_column_names(param_table) === get_column_names(get_table(lhs_table_ref)) && first(get_cols(lhs_table_ref)) == get_cols(table_ref)
        println("BroadcastedStatement, indexing on col in a different table, because col names are equal")
        "col"
    else
        i_coeff, j_coeff = col_behavior
        col_is_num = true
        make_index_str(["j", "i"], col_behavior, col_idx .- 1)
    end

    index_str = @match (row_is_num, col_is_num) begin
        (false, false) => begin
            if param_rows == param_cols == 1
                ".at[\$row_loc, \$col_loc]"
            else
                ".loc[\$row_loc, \$col_loc]"
            end
        end
        (true, false) => begin
            if size(param_table)[2] == 1
                ".iloc[\$row_loc, 0]"
            else
                ".loc[:, \$col_loc].iloc[\$row_loc]"
            end
        end
        (false, true) => begin
            if param_rows == 1 && (size(param_table)[1] != 1)
                ".loc[\$row_loc].iloc[\$col_loc]"
            elseif param_rows ==1 && (size(param_table)[1] == 1)
                ".iloc[0, \$col_loc]"
            else
                ".loc[\$row_loc].iloc[:, \$col_loc]"
            end
        end
        (true, true) => ".iloc[\$row_loc, \$col_loc]"
    end

    param_index_strs[i] = getname(param_table) * index_str
end


The complexity:
- Julia doesn't support row index names, python does
- There's a distinction between referencing a chunk of a table, and looping over it
    - further complicated by potential loop -> broadcast transforms
- Need to keep things as the correct type in certain situations
- Desire to use label based indexing when possible, even between tables

Solution outline:

Use types (new or julia builtin) to generically define table indexing that can
then be rendered by the exporter.

What does this IR look like so that both python and julia can generate nice code?
What should the above code look like?

Represent iterating over table dimensions:

# Loosely this is in export_statement for broadcasted_statement
lhs_table_ref = TableRef(statement.lhs_expr)
col_iter = DimIterator(lhs_table_ref, 2, "i", "col")
row_iter = DimIterator(lhs_table_ref, 1, "j", "row")
# col_iter = make_col_iterator(lhs_table_ref, "i", "col")
# row_iter = make_row_iterator(lhs_table_ref, "j", "row")


for (i, table_ref) in zip(findall(changing), param_table_refs)
    # row_behavior = (change when incrementing the lhs row, change when incrementing the lhs col)
    row_behavior, col_behavior = param_broadcast_behavior[i, :]

    row_index = make_row_index(table_ref, row_behavior, row_iter, col_iter)
    col_index = make_col_index(table_ref, col_behavior, row_iter, col_iter)

    param_index_types[i] = TableIndex(get_table(table_ref), row_index, col_index)
end


Then we need some sort of export_for_loops to actually handle the result of this.
Because somewhere in there we want simplifications like:
- In python, don't enumerate if we never use the index value

"""

struct DimIterator
    table_ref::TableRef
    dim::Int
    """1 = row, 2 = col"""
    index_var::String
    label_var::String
end

function ref_indices(dim_iter::DimIterator)
    if dim_iter.dim == 1
        dim_iter.table_ref.row
    elseif dim_iter.dim == 2
        dim_iter.table_ref.col
    else
        throw("DimIterator dim must be 1 or 2, was $(dim_iter.dim)")
    end
end

function Base.length(dim_iter::DimIterator)
    length(ref_indices(dim_iter))
end

function labels(dim_iter::DimIterator)
    tbl = get_table(dim_iter.table_ref)
    indices = ref_indices(dim_iter)

    if dim_iter.dim == 1
        row_name.(Ref(tbl), indices)
    elseif dim_iter.dim == 2
        column_name.(Ref(tbl), indices)
    else
        throw("DimIterator dim must be 1 or 2, was $(dim_iter.dim)")
    end
end

struct IterOffsetIndex
    base_index::Union{Int, UnitRange}
    iters::Vector{DimIterator}
    offsets::Vector{Int}
end

Base.length(idx::IterOffsetIndex) = length(idx.base_index)

struct IndexRange
    start::Any
    stop::Any
end

struct TableIndex
    table::ExcelTable
    "An index can be of types such as:
        - Int (for single index)
        - UnitRange{Int, Int} (for slice)
        - IterOffsetIndex
        - IndexRange
    "
    row_index::Any
    col_index::Any
end

function make_iter_offset_index(base_index::Union{Int, UnitRange{Int}}, iters::Vector{DimIterator}, offsets::Vector)
    # if any of the offsets are UnitRanges, we probably need to make it an IndexRange
    if any(isa.(offsets, Ref(UnitRange)))
        @show table_ref behavior row_iter col_iter
        throw("make_row_index with an offset that's a unit range isn't yet supported")
    end

    is_zero_offset = offsets .== 0

    if all(is_zero_offset)
        return base_index
    end


    non_zero_offset = .!is_zero_offset

    IterOffsetIndex(base_index, iters[non_zero_offset], offsets[non_zero_offset])
end

function make_row_index(table_ref::TableRef, behavior, row_iter::DimIterator, col_iter::DimIterator)
    iters = DimIterator[row_iter, col_iter]
    offsets = collect(behavior)

    base_rows = get_rows(table_ref)

    make_iter_offset_index(base_rows, iters, offsets)
end

function make_col_index(table_ref::TableRef, behavior, row_iter::DimIterator, col_iter::DimIterator)
    iters = DimIterator[row_iter, col_iter]
    offsets = collect(behavior)

    base_cols = get_cols(table_ref)


    make_iter_offset_index(base_cols, iters, offsets)
end
