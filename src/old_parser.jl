
using RBNF
using Test
using PrettyPrint
using AutoHashEquals
using Match
using Dates
using Statistics
using XLSX
using DataFrames
using .XLExpr

second((a, b)) = b
second(vec::V) where {V<:AbstractArray} = vec[2]
parseReal(x::RBNF.Token) = parse(Float64, x.str)
getstr(c::RBNF.Token) = c.str
getstr(c::String) = c

function unwrap_if_singleton(container)
    println(container)
    if length(container.children) == 1
        container.children[1]
    else
        container
    end
end

struct ExcelFormula3 end

XLFormula = ExcelFormula3

flatten(arr::AbstractArray) = reduce(vcat, arr)
flatten(arr) = arr

RBNF.@parser XLFormula begin

    ignore{space}

    @grammar
    Start := [Formula = Expression]
    Expression = BinOpFormula
    # atom = Constant | paren | ReferenceItem | FunctionCall | UnaryExpr
    # atom = Constant | paren | ReferenceItem | FunctionCall | UnaryExpr
    PercentExpr := [expr = formulaWithUnaryOp, "%"]
    UnaryExpr := [prefix = UnaryOp % getstr, expr = RangeFormula]
    UnaryOp = "+" | "-"
    BinOpFormula := [children = Formula]
    Formula = @direct_recur begin
        init = [formulaWithPercentOp]
        prefix = [recur..., BinOp, formulaWithPercentOp]
    end
    formulaWithPercentOp = PercentExpr | formulaWithUnaryOp
    formulaWithUnaryOp = UnaryExpr | formulaWithRange
    formulaWithRange = RangeFormula

    # UnaryOrRange = UnaryExpr | PercentExpr | RangeFormula

    FunctionCall := [fn = excel_function % getstr, [children = Arguments].?, ")"]
    Arguments = @direct_recur begin
        init = [Expression.?]
        prefix = [recur..., (',', Expression.?) % second]
    end
    paren = (["(", BinOpFormula, ")"] % second)
    formula = referenceWithoutInfix | paren | Constant | FunctionCall
    referenceWithoutInfix = ReferenceItem | PrefixedRef

    # Reference = ReferenceItem | RefFunctionCall | PrefixedRef
    # ReferenceRange := [first = Reference, ":", second = Reference]

    RangeFormula := [children = range_formula]
    range_formula = @direct_recur begin
        init = [formula]
        prefix = [recur..., (":", formula) % second]
    end

    Constant = xl_float | xl_string | Boolean
    Boolean := [value = boolean]
    xl_float := [value = number % parseReal]
    xl_string := [value = string_literal % (s -> string(strip(s.str, '"')))]

    PrefixedRef := [sheet = sheet_prefix % (s -> strip(s.str, ['\'', '!'])), cell = RangeFormula]
    sheet_prefix = unquoted_sheet | quoted_sheet
    BinOp := [Op = binary_operator % getstr]
    binary_operator = "+" | "-" | "*" | "/" | "^" | "<>" | "<=" | ">=" | "<" | ">" | "=" | "&"
    ReferenceItem = CellName | NamedRange | VerticalRange | error_ref
    # RefFunctionCall = ReferenceRange
    RefFunctionCall = [fn = "IF(", args = Arguments, ")"]

    CellName := [cell = cell % getstr]
    VerticalRange := [cell = vertical_range % getstr]
    NamedRange := [name = (named_range | named_range_prefixed) % getstr]

    @token

    excel_function := r"\G(_xlfn\.)?[A-Z][A-Z0-9\.]*\("
    unquoted_sheet := r"\G'[^\'\*\[\]\\:/\?]+'!"
    quoted_sheet := r"\G[^\'\*\[\]\\:/\?\\(\);\{\}#\"=<>&\+\-\*^%, ]+!"
    named_range_prefixed := r"\G(TRUE|FALSE|[A-Z]+[0-9]++)[A-Za-z0-9\\_]+"
    cell := r"\G[\$]?[A-Z]+[\$]?\d+"
    error_ref := "#REF!"
    boolean := "TRUE" | "FALSE"
    horizontal_range := r"\G[\$]?\d+:[\$]?\d+"
    vertical_range := r"\G\$?[A-Z]+:\$?[A-Z]+"
    number := r"\G\d+[\.]?\d*(e\d+)?"
    string_literal := r"\G\"([^\"]|(?:\"\"))*\""
    space := r"\G\s+"
    named_range := r"\G[A-Za-z_\\][A-Za-z0-9\\_]*"
end

function parse_XLFormula(start_rule, formula)
    tokens = RBNF.runlexer(XLFormula, formula)
    ast, ctx = RBNF.runparser(start_rule, tokens)
    ast
end

function parse_formula(formula)
    tokens = RBNF.runlexer(XLFormula, formula)
    # @show tokens
    ast, ctx = RBNF.runparser(Start, tokens)
    # @show ast
    return ast
end


function offset(value, ::Int, ::Int)
    value
end
offset_cell_parse_rgx = r"([$]?[A-Z]+)([$]?[0-9]+)"
function offset(expr::ExcelExpr, rows::Int, cols::Int)
    head = expr.head
    args = expr.args

    @match expr begin
        ExcelExpr(:cell_ref, args) => begin
            cell_match = match(offset_cell_parse_rgx, args[1])
            @assert cell_match.match == args[1] "Cell didn't parse properly"
            col_str = cell_match[1]
            row_str = cell_match[2]

            if col_str[1] == '$'
                new_col = col_str
            else
                new_col_num = XLSX.decode_column_number(col_str[1:end]) + cols
                # This allows for (invalid) offsetting of expressions into negative columns
                if new_col_num >= 1
                    new_col = XLSX.encode_column_number(new_col_num)
                else
                    new_col = string(new_col_num)
                end
            end
            if row_str[1] == '$'
                new_row = row_str
            else
                new_row = string(parse(Int, row_str) + rows)
            end
            ExcelExpr(:cell_ref, String(new_col * new_row), args[2:end]...)
        end
        ExcelExpr(:table_ref, (table, row_idx, col_idx, fixed_row, fixed_col)) => begin
            row_idx = @match fixed_row begin
                (true, true) => row_idx
                (false, false) => row_idx .+ rows
                (true, false) => first(row_idx):last(row_idx)+rows
                (false, true) => first(row_idx)+rows:last(row_idx)
            end
            col_idx = @match fixed_col begin
                (true, true) => col_idx
                (false, false) => col_idx .+ cols
                (true, false) => first(col_idx):last(col_idx)+cols
                (false, true) => first(col_idx)+cols:last(col_idx)
            end

            ExcelExpr(:table_ref, table, row_idx, col_idx, fixed_row, fixed_col)
        end
        _ => ExcelExpr(head, map(x -> offset(x, rows, cols), args)...)
    end
    # if head == :cell_ref
    #     cell_match = match(offset_cell_parse_rgx, args[1])
    #     @assert cell_match.match == args[1] "Cell didn't parse properly"
    #     col_str = cell_match[1]
    #     row_str = cell_match[2]

    #     if col_str[1] == '$'
    #         new_col = col_str
    #     else
    #         new_col = XLSX.encode_column_number(XLSX.decode_column_number(col_str[1:end]) + cols)
    #     end
    #     if row_str[1] == '$'
    #         new_row = row_str
    #     else
    #         new_row = string(parse(Int, row_str) + rows)
    #     end
    #     ExcelExpr(:cell_ref, String(new_col * new_row))
    # else
    #     ExcelExpr(head, map(x -> offset(x, rows, cols), args)...)
    # end
end

function toexpr(node::Struct_Start)
    toexpr(node.Formula)
end

function expr_bp(tokens, min_bp)
    lhs = popat!(tokens, 1)

    while length(tokens) > 0
        op = tokens[1]
        # op = popat!(tokens, 1)
        l_bp, r_bp = infix_binding_power(op)
        if l_bp < min_bp
            break
        end
        popat!(tokens, 1)
        rhs = expr_bp(tokens, r_bp)
        lhs = (op, (lhs, rhs))
    end
    lhs
end

function infix_binding_power(op)
    lookup = Dict(
        :+ => (2, 3),
        :- => (2, 3),
        :* => (4, 5),
        :/ => (4, 5),
        :^ => (6, 7),
        :eq => (0, 0),
        :neq => (0, 0),
        :leq => (0, 0),
        :geq => (0, 0),
        :lt => (0, 0),
        :gt => (0, 0),
        :& => (1, 2),
    )
    lookup[op]
end

toexpr(node::Struct_BinOpFormula) = toexpr(node, 0)

function toexpr(node::Struct_BinOp)
    if node.Op == "="
        Symbol(:eq)
    elseif node.Op == "<>"
        Symbol(:neq)
    elseif node.Op == "<="
        Symbol(:leq)
    elseif node.Op == ">="
        Symbol(:geq)
    elseif node.Op == ">"
        Symbol(:gt)
    elseif node.Op == "<"
        Symbol(:lt)
    else
        Symbol(node.Op)
    end
end
function parse_binops(tokens, min_bp=0)
    lhs = popat!(tokens, 1)

    while length(tokens) > 0
        op = tokens[1]
        # op = popat!(tokens, 1)
        l_bp, r_bp = infix_binding_power(op)
        if l_bp < min_bp
            break
        end
        popat!(tokens, 1)
        rhs = parse_binops(tokens, r_bp)
        lhs = ExcelExpr(op, lhs, rhs)
    end
    lhs
end
function toexpr(node::Struct_BinOpFormula, min_bp)
    tokens = collect(map(toexpr, node.children))
    parse_binops(tokens)
end
toexpr(node::RBNF.Token{:error_ref}) = ExcelExpr(:error_ref, node.str)
toexpr(node::Struct_CellName) = ExcelExpr(:cell_ref, node.cell)
toexpr(node::Struct_VerticalRange) = ExcelExpr(:cols, node.cell)
function toexpr(node::Struct_RangeFormula)
    if length(node.children) == 1
        toexpr(node.children[1])
    else
        children = collect(map(toexpr, node.children))
        ExcelExpr(:range, children...)
    end
end

function toexpr(node::Array)
    @assert length(node) == 1 "Can only unwrap an array with a single element. $(node)"
    toexpr(node[1])
end

function toexpr(node::Struct_FunctionCall)
    if node.children === nothing || node.children == [Some(nothing)]
        ExcelExpr(:call, string(node.fn[1:end-1]))
    else
        function process_function_arg(arg::Some)
            arg = something(arg)
            isnothing(arg) ? missing : toexpr(arg)
        end

        # args = collect(map(toexpr, node.children))
        args = collect(map(process_function_arg, node.children))
        ExcelExpr(:call, string(node.fn[1:end-1]), args...)
    end
end
function toexpr(node::Struct_xl_float)
    node.value
end
function toexpr(node::Struct_xl_string)
    node.value
end
function toexpr(node::Struct_Boolean)
    bool = node.value.str
    if bool == "TRUE"
        true
    elseif bool == "FALSE"
        false
    else
        throw("\"$(bool)\" not recognized as a boolean value")
    end
end
function toexpr(node::Struct_PrefixedRef)
    ExcelExpr(:sheet_ref, node.sheet, toexpr(node.cell))
end
function toexpr(node::Struct_NamedRange)
    ExcelExpr(:named_range, node.name)
end
function toexpr(node::Struct_UnaryExpr)
    ExcelExpr(Symbol(node.prefix), toexpr(node.expr))
end

function toexpr(node::Struct_PercentExpr)
    ExcelExpr(:%, toexpr(node.expr))
end

# e = toexpr(parse_XLFormula(RangeFormula, "A1:B10"))
# new_e = ExcelExpr(:range, ExcelExpr(:cell_ref, "A1"), ExcelExpr(:cell_ref, "B10"))
# toexpr(missing_args_formula)

# test_formula =  "OFFSET(Lists!\$G\$4,0,0,COUNTA(Lists!\$G:\$G),2)"
# # test_formula =  "\$G:\$G"
# tokens = RBNF.runlexer(XLFormula, test_formula)
# # @show tokens
# ast, ctx = RBNF.runparser(Start, tokens)
parse_formula("OFFSET(Lists!\$G\$4,0,0,COUNTA(Lists!\$G:\$G),2)") |> toexpr

# percent_formula = parse_formula("100%-equity")
# toexpr(percent_formula)
@test toexpr(parse_XLFormula(RangeFormula, "A1:B10")) == ExcelExpr(:range, ExcelExpr(:cell_ref, "A1"), ExcelExpr(:cell_ref, "B10"))
@test toexpr(parse_formula("A1:B10")) == ExcelExpr(:range, ExcelExpr(:cell_ref, "A1"), ExcelExpr(:cell_ref, "B10"))
@test toexpr(parse_formula("SUM(A1:B20, B2)")) == ExcelExpr(:call, "SUM", ExcelExpr(:range, ExcelExpr(:cell_ref, "A1"), ExcelExpr(:cell_ref, "B20")), ExcelExpr(:cell_ref, "B2"))
@test toexpr(parse_formula("1 + 2")) == ExcelExpr(:+, 1.0, 2.0)
@test toexpr(parse_formula("'Sheet 1'!A1:A10")) == ExcelExpr(:sheet_ref, "Sheet 1", ExcelExpr(:range, ExcelExpr(:cell_ref, "A1"), ExcelExpr(:cell_ref, "A10")))
@test toexpr(parse_formula("'Sheet 1'!A1")) == ExcelExpr(:sheet_ref, "Sheet 1", ExcelExpr(:cell_ref, "A1"))
@test toexpr(parse_formula("'Sheet 1'!A1 - B10")) == ExcelExpr(:-, ExcelExpr(:sheet_ref, "Sheet 1", ExcelExpr(:cell_ref, "A1")), ExcelExpr(:cell_ref, "B10"))
@test toexpr(parse_formula("(C4/2204.62)*F38")) == ExcelExpr(:*, ExcelExpr(:/, ExcelExpr(:cell_ref, "C4"), 2204.62), ExcelExpr(:cell_ref, "F38"))
@test toexpr(parse_formula("total_harvest")) == ExcelExpr(:named_range, "total_harvest")
@test toexpr(parse_formula("1 * 2 - 4")) == ExcelExpr(:-, ExcelExpr(:*, 1.0, 2.0), 4.0)
@test toexpr(parse_formula("1 - 2 * 4")) == ExcelExpr(:-, 1, ExcelExpr(:*, 2, 4))
@test toexpr(parse_formula("IF(salary_worker_toggle=1,((C309*(I143-I142))+(I142*C310))*F136,I136*I140*I142*(((F138-1)*'1. Assumptions'!F105)+('1. Assumptions'!F106*1)))*(1+'1. Assumptions'!F322)*F313*array_deploy_scalar")) == ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:call, "IF", ExcelExpr(:eq, ExcelExpr(:named_range, "salary_worker_toggle"), 1.0), ExcelExpr(:*, ExcelExpr(:+, ExcelExpr(:*, ExcelExpr(:cell_ref, "C309"), ExcelExpr(:-, ExcelExpr(:cell_ref, "I143"), ExcelExpr(:cell_ref, "I142"))), ExcelExpr(:*, ExcelExpr(:cell_ref, "I142"), ExcelExpr(:cell_ref, "C310"))), ExcelExpr(:cell_ref, "F136")), ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:cell_ref, "I136"), ExcelExpr(:cell_ref, "I140")), ExcelExpr(:cell_ref, "I142")), ExcelExpr(:+, ExcelExpr(:*, ExcelExpr(:-, ExcelExpr(:cell_ref, "F138"), 1.0), ExcelExpr(:sheet_ref, "1. Assumptions", ExcelExpr(:cell_ref, "F105"))), ExcelExpr(:*, ExcelExpr(:sheet_ref, "1. Assumptions", ExcelExpr(:cell_ref, "F106")), 1.0)))), ExcelExpr(:+, 1.0, ExcelExpr(:sheet_ref, "1. Assumptions", ExcelExpr(:cell_ref, "F322")))), ExcelExpr(:cell_ref, "F313")), ExcelExpr(:named_range, "array_deploy_scalar"))
@test toexpr(parse_formula("(N15-#REF!)/#REF!")) == ExcelExpr(:/, ExcelExpr(:-, ExcelExpr(:cell_ref, "N15"), ExcelExpr(:error_ref, "#REF!")), ExcelExpr(:error_ref, "#REF!"))
@test toexpr(parse_formula("(F262*F263*F264)/(3.28^3)*(1+F240)")) == ExcelExpr(:*, ExcelExpr(:/, ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:cell_ref, "F262"), ExcelExpr(:cell_ref, "F263")), ExcelExpr(:cell_ref, "F264")), ExcelExpr(:^, 3.28, 3.0)), ExcelExpr(:+, 1.0, ExcelExpr(:cell_ref, "F240")))
@test toexpr(parse_formula("RAND()")) == ExcelExpr(:call, "RAND")
@test toexpr(parse_formula("IF(H75<\$I\$71,H75+1,\"\")")) == ExcelExpr(:call, "IF", ExcelExpr(:lt, ExcelExpr(:cell_ref, "H75"), ExcelExpr(:cell_ref, "\$I\$71")), ExcelExpr(:+, ExcelExpr(:cell_ref, "H75"), 1.0), "")
@test toexpr(parse_formula("IF(H75<>\"\",COUNTIFS(\$M\$4:\$ZZ\$4,\">=\"&I75,\$M\$4:\$ZZ\$4,\"<=\"&J75),\"\")")) == ExcelExpr(:call, "IF", ExcelExpr(:neq, ExcelExpr(:cell_ref, "H75"), ""), ExcelExpr(:call, "COUNTIFS", ExcelExpr(:range, ExcelExpr(:cell_ref, "\$M\$4"), ExcelExpr(:cell_ref, "\$ZZ\$4")), ExcelExpr(:&, ">=", ExcelExpr(:cell_ref, "I75")), ExcelExpr(:range, ExcelExpr(:cell_ref, "\$M\$4"), ExcelExpr(:cell_ref, "\$ZZ\$4")), ExcelExpr(:&, "<=", ExcelExpr(:cell_ref, "J75"))), "")
@test toexpr(parse_formula("IF(H75<>\"\",AVERAGE(I75:J75),FALSE)")) == ExcelExpr(:call, "IF", ExcelExpr(:neq, ExcelExpr(:cell_ref, "H75"), ""), ExcelExpr(:call, "AVERAGE", ExcelExpr(:range, ExcelExpr(:cell_ref, "I75"), ExcelExpr(:cell_ref, "J75"))), false)
@test toexpr(parse_formula("-A1")) == ExcelExpr(:-, ExcelExpr(:cell_ref, "A1"))
@test toexpr(parse_formula("100%-equity")) == ExcelExpr(:-, ExcelExpr(:%, 100.0), ExcelExpr(:named_range, "equity"))
@test toexpr(parse_formula("IF(G131=\"\",,1)")) === ExcelExpr(:call, "IF", ExcelExpr(:eq, ExcelExpr(:cell_ref, "G131"), ""), missing, 1.0)
@test toexpr(parse_formula("OFFSET(Lists!\$G\$4,0,0,COUNTA(Lists!\$G:\$G),2)")) == ExcelExpr(:call, "OFFSET", ExcelExpr(:sheet_ref, "Lists", ExcelExpr(:cell_ref, "\$G\$4")), 0.0, 0.0, ExcelExpr(:call, "COUNTA", ExcelExpr(:sheet_ref, "Lists", ExcelExpr(:cols, "\$G:\$G"))), 2.0)
@test offset(toexpr(parse_formula("SUM(A1:B10)")), 1, 1) == ExcelExpr(:call, "SUM", ExcelExpr(:range, ExcelExpr(:cell_ref, "B2"), ExcelExpr(:cell_ref, "C11")))
