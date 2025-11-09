if false
    include("./src/XLConvert.jl")
end
using XLConvert
using XLConvert: FlatExpr, FlatIdx
using XLSX
using Graphs
using Match

function read_wb()
    file = "test_workbooks/linear_programming.xlsx"
    parse_workbook(file)
end

function run(wb_in::XLConvert.ExcelWorkbook2)
    all_target_outputs = [CellDependency("Sheet1", "E10"), CellDependency("Sheet1", "E11"), CellDependency("Sheet1", "E12")]

    tables = [XLConvert.DefTable(wb.xf, "Sheet1", "coeffs", "C10", "D12", "C8:D8", "B10:B12")]

    export_julia(wb, all_target_outputs, "test_workbooks/linear_programming.jl", tables)
end
