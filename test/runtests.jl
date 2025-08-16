using Test
# if false
#     include("../src/XLConvert.jl")

using XLConvert
using XLConvert: toexpr, lower_sheet_names, get_expr_dependencies, convert_to_flat_expr, FlatExpr, FlatIdx


function test_to_expr()
#! format: off
@test toexpr("A1:B10") == ExcelExpr(:range, ExcelExpr(:cell_ref, "A1"), ExcelExpr(:cell_ref, "B10"))
@test toexpr("SUM(A1:B20, B2)") == ExcelExpr(:call, "SUM", ExcelExpr(:range, ExcelExpr(:cell_ref, "A1"), ExcelExpr(:cell_ref, "B20")), ExcelExpr(:cell_ref, "B2"))
@test toexpr("1 + 2") == ExcelExpr(:+, 1.0, 2.0)
@test toexpr("'Sheet 1'!A1:A10") == ExcelExpr(:sheet_ref, "Sheet 1", ExcelExpr(:range, ExcelExpr(:cell_ref, "A1"), ExcelExpr(:cell_ref, "A10")))
@test toexpr("'Sheet 1'!A1") == ExcelExpr(:sheet_ref, "Sheet 1", ExcelExpr(:cell_ref, "A1"))
@test toexpr("'Sheet 1'!A1 - B10") == ExcelExpr(:-, ExcelExpr(:sheet_ref, "Sheet 1", ExcelExpr(:cell_ref, "A1")), ExcelExpr(:cell_ref, "B10"))
@test toexpr("(C4/2204.62)*F38") == ExcelExpr(:*, ExcelExpr(:/, ExcelExpr(:cell_ref, "C4"), 2204.62), ExcelExpr(:cell_ref, "F38"))
@test toexpr("total_harvest") == ExcelExpr(:named_range, "total_harvest")
@test toexpr("1 * 2 - 4") == ExcelExpr(:-, ExcelExpr(:*, 1.0, 2.0), 4.0)
@test toexpr("1 - 2 * 4") == ExcelExpr(:-, 1, ExcelExpr(:*, 2, 4))
@test toexpr("IF(salary_worker_toggle=1,((C309*(I143-I142))+(I142*C310))*F136,I136*I140*I142*(((F138-1)*'1. Assumptions'!F105)+('1. Assumptions'!F106*1)))*(1+'1. Assumptions'!F322)*F313*array_deploy_scalar") == ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:call, "IF", ExcelExpr(:eq, ExcelExpr(:named_range, "salary_worker_toggle"), 1.0), ExcelExpr(:*, ExcelExpr(:+, ExcelExpr(:*, ExcelExpr(:cell_ref, "C309"), ExcelExpr(:-, ExcelExpr(:cell_ref, "I143"), ExcelExpr(:cell_ref, "I142"))), ExcelExpr(:*, ExcelExpr(:cell_ref, "I142"), ExcelExpr(:cell_ref, "C310"))), ExcelExpr(:cell_ref, "F136")), ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:cell_ref, "I136"), ExcelExpr(:cell_ref, "I140")), ExcelExpr(:cell_ref, "I142")), ExcelExpr(:+, ExcelExpr(:*, ExcelExpr(:-, ExcelExpr(:cell_ref, "F138"), 1.0), ExcelExpr(:sheet_ref, "1. Assumptions", ExcelExpr(:cell_ref, "F105"))), ExcelExpr(:*, ExcelExpr(:sheet_ref, "1. Assumptions", ExcelExpr(:cell_ref, "F106")), 1.0)))), ExcelExpr(:+, 1.0, ExcelExpr(:sheet_ref, "1. Assumptions", ExcelExpr(:cell_ref, "F322")))), ExcelExpr(:cell_ref, "F313")), ExcelExpr(:named_range, "array_deploy_scalar"))
@test toexpr("(N15-#REF!)/#REF!") == ExcelExpr(:/, ExcelExpr(:-, ExcelExpr(:cell_ref, "N15"), ExcelExpr(:error_ref, "#REF!")), ExcelExpr(:error_ref, "#REF!"))
@test toexpr("(F262*F263*F264)/(3.28^3)*(1+F240)") == ExcelExpr(:*, ExcelExpr(:/, ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:cell_ref, "F262"), ExcelExpr(:cell_ref, "F263")), ExcelExpr(:cell_ref, "F264")), ExcelExpr(:^, 3.28, 3.0)), ExcelExpr(:+, 1.0, ExcelExpr(:cell_ref, "F240")))
@test toexpr("RAND()") == ExcelExpr(:call, "RAND")
@test toexpr("IF(H75<\$I\$71,H75+1,\"\")") == ExcelExpr(:call, "IF", ExcelExpr(:lt, ExcelExpr(:cell_ref, "H75"), ExcelExpr(:cell_ref, "\$I\$71")), ExcelExpr(:+, ExcelExpr(:cell_ref, "H75"), 1.0), "")
@test toexpr("IF(H75<>\"\",COUNTIFS(\$M\$4:\$ZZ\$4,\">=\"&I75,\$M\$4:\$ZZ\$4,\"<=\"&J75),\"\")") == ExcelExpr(:call, "IF", ExcelExpr(:neq, ExcelExpr(:cell_ref, "H75"), ""), ExcelExpr(:call, "COUNTIFS", ExcelExpr(:range, ExcelExpr(:cell_ref, "\$M\$4"), ExcelExpr(:cell_ref, "\$ZZ\$4")), ExcelExpr(:&, ">=", ExcelExpr(:cell_ref, "I75")), ExcelExpr(:range, ExcelExpr(:cell_ref, "\$M\$4"), ExcelExpr(:cell_ref, "\$ZZ\$4")), ExcelExpr(:&, "<=", ExcelExpr(:cell_ref, "J75"))), "")
@test toexpr("IF(H75<>\"\",AVERAGE(I75:J75),FALSE)") == ExcelExpr(:call, "IF", ExcelExpr(:neq, ExcelExpr(:cell_ref, "H75"), ""), ExcelExpr(:call, "AVERAGE", ExcelExpr(:range, ExcelExpr(:cell_ref, "I75"), ExcelExpr(:cell_ref, "J75"))), false)
@test toexpr("-A1") == ExcelExpr(:-, ExcelExpr(:cell_ref, "A1"))
@test toexpr("100%-equity") == ExcelExpr(:-, ExcelExpr(:%, 100.0), ExcelExpr(:named_range, "equity"))
@test isequal(toexpr("IF(G131=\"\",,1)"), ExcelExpr(:call, "IF", ExcelExpr(:eq, ExcelExpr(:cell_ref, "G131"), ""), missing, 1.0))
@test toexpr("OFFSET(Lists!\$G\$4,0,0,COUNTA(Lists!\$G:\$G),2)") == ExcelExpr(:call, "OFFSET", ExcelExpr(:sheet_ref, "Lists", ExcelExpr(:cell_ref, "\$G\$4")), 0.0, 0.0, ExcelExpr(:call, "COUNTA", ExcelExpr(:sheet_ref, "Lists", ExcelExpr(:cols, "\$G:\$G"))), 2.0)
@test isequal(toexpr("SUM(1, 2,)"), ExcelExpr(:call, "SUM", 1.0, 2.0, missing))

@test toexpr(".95") == 0.95

@test toexpr("SUM(A1:(B1))") == ExcelExpr(:call, "SUM", ExcelExpr(:range, ExcelExpr(:cell_ref, "A1"), ExcelExpr(:cell_ref, "B1")))

@test toexpr("SUM(A:B)") == ExcelExpr(:call, "SUM", ExcelExpr(:cols, "A:B"))

@test toexpr("'[1]OtherSheet'!A1:B10") isa ExcelExpr
#! format: on
end


function test_get_expr_deps()

    sheet = "Sheet 1"
    cell = Base.Fix1(CellDependency, sheet)
    expr = lower_sheet_names(toexpr("A1 + B2"), sheet)

    key_values = Dict()
    deps = get_expr_dependencies(expr, key_values)
    @test deps == [cell("A1"), cell("B2")]

    flat_expr = XLConvert.convert_to_flat_expr(expr)
    deps = get_expr_dependencies(flat_expr, key_values)
    @test deps == [cell("A1"), cell("B2")]

    expr = lower_sheet_names(toexpr("SUM(A1:B2)"), sheet)
    expected_deps = [cell("A1"), cell("A2"), cell("B1"), cell("B2")]

    deps = get_expr_dependencies(expr, key_values)
    @test deps == expected_deps

    flat_expr = XLConvert.convert_to_flat_expr(expr)
    deps = get_expr_dependencies(flat_expr, key_values)
    @test deps == expected_deps

end

function test_flat_expr_convert()
    expr = toexpr("A1:B10")
    flat_expr = convert_to_flat_expr(expr)
    @test flat_expr == FlatExpr([
        ExcelExpr(:range, FlatIdx(2), FlatIdx(3)),
        ExcelExpr(:cell_ref, "A1"),
        ExcelExpr(:cell_ref, "B10"),
    ])
    @test XLConvert.convert_to_expr(flat_expr) == expr

    test_expr_strs = [
        "A1:B10",
        "SUM(A1:B20, B2)",
        "1 + 2",
        "'Sheet 1'!A1:A10",
        "'Sheet 1'!A1",
        "'Sheet 1'!A1 - B10",
        "(C4/2204.62)*F38",
        "total_harvest",
        "1 * 2 - 4",
        "1 - 2 * 4",
        "IF(salary_worker_toggle=1,((C309*(I143-I142))+(I142*C310))*F136,I136*I140*I142*(((F138-1)*'1. Assumptions'!F105)+('1. Assumptions'!F106*1)))*(1+'1. Assumptions'!F322)*F313*array_deploy_scalar",
        "(N15-#REF!)/#REF!",
        "(F262*F263*F264)/(3.28^3)*(1+F240)",
        "RAND()",
        "IF(H75<\$I\$71,H75+1,\"\")",
        "IF(H75<>\"\",COUNTIFS(\$M\$4:\$ZZ\$4,\">=\"&I75,\$M\$4:\$ZZ\$4,\"<=\"&J75),\"\")",
        "IF(H75<>\"\",AVERAGE(I75:J75),FALSE)",
        "-A1",
        "100%-equity",
        "IF(G131=\"\",,1)",
        "OFFSET(Lists!\$G\$4,0,0,COUNTA(Lists!\$G:\$G),2)",
        "SUM(1, 2,)",
    ]

    for s in test_expr_strs
        start_expr = toexpr(s)
        flat_expr = convert_to_flat_expr(start_expr)
        new_expr = XLConvert.convert_to_expr(flat_expr)

        @test isequal(start_expr, new_expr)
    end
end


function test_functionalized()
    test_expr_strs = [
        "A1:B10",
        "SUM(A1:B20, B2)",
        "1 + 2",
        "'Sheet 1'!A1:A10",
        "'Sheet 1'!A1",
        "'Sheet 1'!A1 - B10",
        "(C4/2204.62)*F38",
        "total_harvest",
        "1 * 2 - 4",
        "1 - 2 * 4",
        "IF(salary_worker_toggle=1,((C309*(I143-I142))+(I142*C310))*F136,I136*I140*I142*(((F138-1)*'1. Assumptions'!F105)+('1. Assumptions'!F106*1)))*(1+'1. Assumptions'!F322)*F313*array_deploy_scalar",
        "(N15-#REF!)/#REF!",
        "(F262*F263*F264)/(3.28^3)*(1+F240)",
        "RAND()",
        "IF(H75<\$I\$71,H75+1,\"\")",
        "IF(H75<>\"\",COUNTIFS(\$M\$4:\$ZZ\$4,\">=\"&I75,\$M\$4:\$ZZ\$4,\"<=\"&J75),\"\")",
        "IF(H75<>\"\",AVERAGE(I75:J75),FALSE)",
        "-A1",
        "100%-equity",
        "IF(G131=\"\",,1)",
        "OFFSET(Lists!\$G\$4,0,0,COUNTA(Lists!\$G:\$G),2)",
        "SUM(1, 2,)",
    ]

    for s in test_expr_strs
        expr = lower_sheet_names(toexpr(s), "sheet")
        expr_func, expr_params = XLConvert.functionalize(expr)
        flat_expr = convert_to_flat_expr(expr)
        flat_func, flat_params = XLConvert.functionalize(flat_expr)
        if !isequal(XLConvert.convert_to_expr(flat_func), expr_func)
            @show s
            @show expr
            @show expr_func
            show(stdout, "text/plain", flat_expr)
            show(stdout, "text/plain", flat_func)
        end
        @test isequal(XLConvert.convert_to_expr(flat_func), expr_func)
        @test expr_params == flat_params
        # new_expr = XLConvert.convert_to_expr(flat_expr)

        # @test isequal(start_expr, new_expr)
    end
end

function make_flat_expr(str::String, sheet::String)
    expr = toexpr(value)
    XLConvert.lower_sheet_names!(expr, sheet_name)
    expr = convert_to_flat_expr(expr)
    expr
end


function create_test_statements(workbook::Dict{CellDependency, String}, key_values::Dict{String, Any})
    statements = AbstractStatement[]

    handled_deps = Set(keys(workbook))

    for (cell, value) in workbook
        expr = toexpr(value)
        XLConvert.lower_sheet_names!(expr, cell.sheet_name)
        expr = convert_to_flat_expr(expr)
        deps = XLConvert.get_expr_dependencies(expr, key_values)
        for dep in deps
            if dep ∉ handled_deps
                push!(statements, XLConvert.StandardStatement(dep, "", []))
                push!(handled_deps, dep)
            end
        end
        push!(statements, XLConvert.StandardStatement(cell, expr, deps))
    end

    statements
end

function test_insert_table_refs()
    tables = [
        XLConvert.ExcelTable("Sheet1", "Table1", "B1", "B2", "", "", "Column1", missing),
    ]

    expr_str = "SUM(B1:B2)"
    expr = toexpr(expr_str)
    XLConvert.lower_sheet_names!(expr, "Sheet1")
    flat_expr = convert_to_flat_expr(expr)
    # function call, range, two cells
    @test length(flat_expr.parts) == 4


    expr = XLConvert.insert_table_refs(expr, tables)
    flat_expr = XLConvert.insert_table_refs(flat_expr, tables)

    # function call, table reference
    @show flat_expr.parts
    @test length(flat_expr.parts) == 2

    @test isequal(XLConvert.convert_to_expr(flat_expr), expr)

    expr_str = "C3 - SUM(B1:B2) + C2"
    expr = toexpr(expr_str)
    XLConvert.lower_sheet_names!(expr, "Sheet1")
    flat_expr = convert_to_flat_expr(expr)
    println("Before:")
    show(stdout, "text/plain", flat_expr)

    expr = XLConvert.insert_table_refs(expr, tables)
    flat_expr = XLConvert.insert_table_refs(flat_expr, tables)

    # @show flat_expr.parts
    println("After:")
    show(stdout, "text/plain", flat_expr)

    @test isequal(XLConvert.convert_to_expr(flat_expr), expr)
end

function test_table_ref_transform()
    cd = cell -> CellDependency("Sheet1", cell)
    begin
        workbook = Dict{CellDependency, String}(
            cd("A1") => "1",
            cd("A2") => "2",
            cd("B1") => "A1 * 2",
            cd("B2") => "A2 * 2",
        )

        key_values = Dict{String, Any}()

        statements = create_test_statements(workbook, key_values)

        @test length(statements) == 4

        tables = [
            XLConvert.ExcelTable("Sheet1", "Table1", "B1", "B2", "", "", "Column1", missing),
        ]


        @time "table_ref_transform" table_ref_transform!(statements, tables)
        table_stmts = filter(s -> s isa XLConvert.TableStatement, statements)
        @test length(table_stmts) == 2


        println("-"^40)
        println("Table Broadcast Transform")
        println("-"^40)
        @time statements = table_broadcast_transform_2d!(statements)

        for statement in statements
            println(statement)
            if statement isa XLConvert.TableStatement
                show(stdout, "text/plain", statement.rhs_expr)
            end
        end

        @test length(statements) == 3

    end

    println("\n", "-"^40)
    println("More Complicated Case")
    println("-"^40, "\n")


    begin
        workbook = Dict{CellDependency, String}(
            cd("B139") => raw"IF($A139=\"\",\"\",IF($B$131=\"\",\"\",IF($B$135=\"no\",INFLATION_FACTOR*$B$133*$B$134,INFLATION_FACTOR*$B$133*$B$132*INDEX(FEEDPRICETABLE,MATCH($B$131,List_Energy_Feedstocks,0),MATCH($A139,Feedstock_Year,0)))))",
            cd("B140") => raw"IF($A140=\"\",\"\",IF($B$131=\"\",\"\",IF($B$135=\"no\",INFLATION_FACTOR*$B$133*$B$134,INFLATION_FACTOR*$B$133*$B$132*INDEX(FEEDPRICETABLE,MATCH($B$131,List_Energy_Feedstocks,0),MATCH($A140,Feedstock_Year,0)))))",
        )

        key_values = Dict{String, Any}(
            "INFLATION_FACTOR" => 1.0,
            "FEEDPRICETABLE" => 0,
            "List_Energy_Feedstocks" => 0,
            "Feedstock_Year" => 0,
        )

        statements = create_test_statements(workbook, key_values)

        # @test length(statements) == 8
        println("Initial Statements:")

        for statement in statements
            println(statement)
            if statement isa XLConvert.StandardStatement && statement.rhs_expr isa XLConvert.FlatExpr
                show(stdout, "text/plain", statement.rhs_expr)
            end
            if statement isa XLConvert.TableStatement
                show(stdout, "text/plain", statement.rhs_expr)
            end
        end

        col_names = ["Column$i" for i in 1:14]
        tables = [
            XLConvert.ExcelTable("Sheet1", "tab_Cash_Flow_Analysis_A139_N182", "A139", "N182", "", "", col_names, missing),
        ]


        @time "table_ref_transform" table_ref_transform!(statements, tables)
        table_stmts = filter(s -> s isa XLConvert.TableStatement, statements)
        @test length(table_stmts) == 4

        println("After table ref transform")
        for statement in statements
            println(statement)
            if statement isa XLConvert.TableStatement
                show(stdout, "text/plain", statement.rhs_expr)
            end
        end


        println("-"^40)
        println("Table Broadcast Transform")
        println("-"^40)
        @time statements = table_broadcast_transform_2d!(statements)

        for statement in statements
            println(statement)
            if statement isa XLConvert.TableStatement
                show(stdout, "text/plain", statement.rhs_expr)
            end
        end

        @test length(statements) == 7

    end

end
