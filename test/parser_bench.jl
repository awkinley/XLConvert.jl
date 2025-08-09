using Test

using XLConvert
using XLConvert: toexpr, lower_sheet_names, get_expr_dependencies, convert_to_flat_expr, FlatExpr, FlatIdx
using XLSX
using BenchmarkTools


function test_tokenize()

    suite = BenchmarkGroup()

    test_strs = [
        "short_num" => "1.0",
        "short_math" => "1 + 4 * 12",
        "long_math" => "0.4 * (1.23e4 / 12)^(2 - 1) + 1.2/9.45",
        "func_call" => "SUM(A1, A2, A3, A4, A5)",
        "nested_if" => "IF(A1, IF(A2, IF(A3, IF(A4, IF(A5, IF(A6, IF(A7, IF(A8, IF(A9, IF(A10, 1, 2), 2), 2), 2), 2), 2), 2), 2), 2), 2)",
    ]

    for (name, str) in test_strs
        @test length(XLConvert.tokenize(str)) > 0
        suite[name] = @benchmarkable XLConvert.tokenize($str)
    end

    results = run(suite, verbose = true)

    results
end

function get_field_test(lexer::XLConvert.Lexer)
    lexer.base_string
end

function read_no_cache(workbook_path::String)
    XLSX.openxlsx(workbook_path, enable_cache = false) do f
        num_cells = 0
        for sheet_name in XLSX.sheetnames(f)
            sheet = f[sheet_name]
            for r in XLSX.eachrow(sheet)
                num_cells += length(values(r.rowcells))
            end
        end
    end

end

function time_read_workbook(workbook_path::String)

    println("XLSX.readxlsx")
    @time xf = XLSX.readxlsx(workbook_path)


    println("XLSX.openxlsx")
    @time read_no_cache(workbook_path)
end
