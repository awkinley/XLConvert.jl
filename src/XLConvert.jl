module XLConvert

using AutoHashEquals
using XLSX
using JSON
using Graphs
using Match
using Random
using Dates
using DataFrames
using EzXML: EzXML

macro display(val)
    return :(
        println($(sprint(Base.show_unquoted, val)*" = "), "\n",
        repr("text/plain", begin
            local value = $(esc(val))
        end)))
end

export CellDependency,
    WorkbookRegion,
    MissingCell,
    AbstractHandler,
    ExcelExpr,
    FlatExpr,
    ExcelTable,
    CellTypes,
    AbstractStatement,
    parse_workbook,
    get_cell,
    get_num,
    get_workbook_subset,
    get_all_referenced_cells,
    get_cell_value,
    get_topo_levels_bottom_up,
    get_topo_levels_top_down,
    get_expr,
    get_type,
    group_to_dict,
    get_set_cells,
    get_cell_deps,
    DefTable,
    make_statements,
    if_multiple_transform!,
    if_toggle_transform!,
    round_if_transform!,
    table_ref_transform!,
    table_broadcast_transform_2d!,
    new_broadcast,
    make_statement_graph,
    group_statements,
    add_functions,
    get_all_referenced_cells,
    make_var_names_map,
    set_names_from_table!,
    BasicOpHandler,
    TableRefHandler,
    EverythingElseHandler,
    JuliaExporter,
    PythonExporter,
    write_file,
    getdatatype,
    named_range_to_cell,
    rownum,
    colnum,
    startcol,
    endcol,
    startrow,
    endrow,
    getname,
    @display,
    find_tables,
    find_tables!,
    get_statements,
    export_julia,
    infer_types,
    xl_sum,
    xl_eq,
    xl_lt,
    xl_gt,
    xl_leq,
    xl_geq,
    xl_logical,
    xl_min,
    xl_max,
    xl_lookup,
    xl_xlookup,
    xl_vlookup,
    xl_index,
    xl_match,
    xl_add,
    xl_sub,
    xl_mul,
    xl_div,
    xl_pmt,
    xl_npv,
    xl_compare,
    xl_average,
    xl_iferror,
    xl_convert,
    xl_isnumber,
    @flat_match,
    handle,
    change_expr_part!,
    change_expr_part,
    find_cycle,
    lookup_const_propagate!



include("object_numbering.jl")
include("graph_utils.jl")
include("excel_expr.jl")
include("formula_parser.jl")
include("excel_formula.jl")
include("flat_expr.jl")
include("flat_match.jl")
include("type_infer.jl")
include("cell_dependency.jl")
include("workbook_region.jl")
include("excel_table.jl")
include("excel_workbook.jl")
include("export_julia.jl")
include("export_python.jl")
include("common_subexpr_elim.jl")
include("statement.jl")
include("workbook_subset.jl")
include("variable_naming.jl")
include("expr_utils.jl")
include("read_excel.jl")

include("transforms/if_multiple.jl")
include("transforms/if_toggle.jl")
include("transforms/round_if.jl")
include("transforms/table_broadcast.jl")
include("transforms/group_statements.jl")
include("transforms/new_broadcast.jl")

include("high_level_api.jl")
include("constant_propagation.jl")


end # module XLConvert
