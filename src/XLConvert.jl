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

export CellDependency,
    MissingCell,
    AbstractHandler,
    ExcelExpr,
    ExcelTable,
    CellTypes,
    AbstractStatement,
    parse_workbook,
    get_workbook_subset,
    get_all_referenced_cells,
    get_cell_value,
    get_topo_levels_bottom_up,
    get_topo_levels_top_down,
    get_expr,
    get_type,
    DefTable,
    make_statements,
    if_multiple_transform!,
    if_toggle_transform!,
    round_if_transform!,
    table_ref_transform!,
    table_broadcast_transform_2d!,
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
    write_file,
    getdatatype,
    named_range_to_cell,
    rownum,
    colnum,
    startcol,
    endcol,
    startrow,
    endrow,
    getname


include("excel_expr.jl")
include("formula_parser.jl")
include("excel_formula.jl")
include("flat_expr.jl")
include("type_infer.jl")
include("cell_dependency.jl")
include("excel_table.jl")
include("excel_workbook.jl")
include("export_julia.jl")
include("statement.jl")
include("workbook_subset.jl")
include("variable_naming.jl")
include("read_excel.jl")

include("transforms/if_multiple.jl")
include("transforms/if_toggle.jl")
include("transforms/round_if.jl")
include("transforms/table_broadcast.jl")
include("transforms/group_statements.jl")


end # module XLConvert
