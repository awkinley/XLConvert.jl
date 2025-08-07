# module FormulaParser
# using Test
# using Match
# using ..XLExpr
using Automa

@enum TokenKind begin
    EXCEL_FUNCTION
    UNQUOTED_SHEET
    QUOTED_SHEET
    NAMED_RANGE_PREFIXED
    CELL
    ERROR_REF
    ERROR
    BOOLEAN
    HORIZONTAL_RANGE
    VERTICAL_RANGE
    NUMBER
    STRING_LITERAL
    SPACE
    NAMED_RANGE
    OTHER
    TOK_ERR
end


struct Token
    kind::TokenKind
    val::SubString{String}
end

val(token::Token) = token.val
kind(token::Token) = token.kind


# function Token(kind::TokenKind, val::SubString{String})
#     Token(kind, val)
# end
function Token(kind::TokenKind, val::String)
    Token(kind, @view val[begin:end])
end
function Token(kind::TokenKind, val::T) where {T<:AbstractChar}
    Token(kind, string(val))
end


mutable struct Lexer
    base_string::String
    idx::Int
end

function Lexer(str::AbstractString)
    base_string = string(str)
    idx = firstindex(base_string)

    Lexer(base_string, idx)
end

const single_char_toks::Vector{Char} = ['+', '-', '*', '/', '^', '(', ')', ',', '%', '&', ':', '=']
function match_other_token(str::T) where {T<:AbstractString}

    for c in single_char_toks
        if str[1] == c
            return Token(OTHER, @view str[1:1])
        end
    end

    first_char = str[1]
    if first_char == '<'
        second_char = if length(str) > 1
            str[2]
        else
            nothing
        end

        if second_char == '='
            Token(OTHER, @view str[1:2])
        elseif second_char == '>'
            Token(OTHER, @view str[1:2])
        else
            Token(OTHER, @view str[1:1])
        end
    elseif first_char == '>'
        second_char = if length(str) > 1
            str[2]
        else
            nothing
        end
        if second_char == '='
            Token(OTHER, @view str[1:2])
        else
            Token(OTHER, @view str[1:1])
        end
    else
        nothing
    end

end

function define_tokenizer()

    tokens = [
        EXCEL_FUNCTION => re"(_xlfn\.)?[A-Z][A-Z0-9\.]*\(",
        UNQUOTED_SHEET => re"[^'*\[\]\\:/\?\(\);{}#\"=<>&\+\-\*^%, ]+!",
        QUOTED_SHEET => re"'[^'*\[\]\\:/\?]+'!",
        CELL => re"[$]?[A-Z]+[$]?[0-9]+",
        ERROR_REF => re"#REF!",
        ERROR => re"#NULL!" | re"#DIV/0!" | re"#VALUE!" | re"#NAME?" | re"#NUM!" | re"#N/A",
        BOOLEAN => re"TRUE|FALSE",
        HORIZONTAL_RANGE => re"[$]?[0-9]+:[$]?[0-9]+",
        VERTICAL_RANGE => re"[$]?[A-Z]+:[$]?[A-Z]+",
        NUMBER => re"\.[0-9]+(e[0-9]+)?" | re"[0-9]+(\.[0-9]+)?(e[0-9]+)?",
        STRING_LITERAL => re"\"([^\"]|\"\")*\"",
        NAMED_RANGE_PREFIXED => re"(TRUE|FALSE|([A-Z]+[0-9]+))[A-Za-z0-9\\_]+",
        # (space, SPACE),
        NAMED_RANGE => re"[A-Za-z_\\][A-Za-z0-9\\_]*",
        SPACE => re" ",
    ]

    other_res = [
        re">=",
        re"<=",
        re"<>",
        re">",
        re"<",
    ]
    for c in single_char_toks
        push!(other_res, RE(c))
    end
    push!(tokens, OTHER => reduce(|, other_res))

    # Automa uses higher index to be higher priority, but the token list uses lower index for higher priority
    reverse!(tokens)

    Automa.make_tokenizer((TOK_ERR, tokens)) |> eval
end

define_tokenizer()
const rg_excel_function = r"\G(_xlfn\.)?[A-Z][A-Z0-9\.]*\("
const rg_unquoted_sheet = r"\G'[^\'\*\[\]\\:/\?]+'!"
const rg_quoted_sheet = r"\G[^\'\*\[\]\\:/\?\\(\);\{\}#\"=<>&\+\-\*^%, ]+!"
const rg_named_range_prefixed = r"\G(TRUE|FALSE|[A-Z]+[0-9]++)[A-Za-z0-9\\_]+"
const rg_cell = r"\G[\$]?[A-Z]+[\$]?\d+"
const rg_error_ref = r"#REF!"
const rg_boolean = r"\G(TRUE|FALSE)"
const rg_horizontal_range = r"\G[\$]?\d+:[\$]?\d+"
const rg_vertical_range = r"\G\$?[A-Z]+:\$?[A-Z]+"
const rg_number = r"\G\d+[\.]?\d*(e\d+)?"
const rg_string_literal = r"\G\"([^\"]|(?:\"\"))*\""
const rg_space = r"\G\s+"
const rg_named_range = r"\G[A-Za-z_\\][A-Za-z0-9\\_]*"

function tokenize_test(lexer)

    # token_kinds = [
    #     (rg_excel_function, EXCEL_FUNCTION),
    #     (rg_unquoted_sheet, UNQUOTED_SHEET),
    #     (rg_quoted_sheet, QUOTED_SHEET),
    #     (rg_named_range_prefixed, NAMED_RANGE_PREFIXED),
    #     (rg_cell, CELL),
    #     (rg_error_ref, ERROR_REF),
    #     (rg_boolean, BOOLEAN),
    #     (rg_horizontal_range, HORIZONTAL_RANGE),
    #     (rg_vertical_range, VERTICAL_RANGE),
    #     (rg_number, NUMBER),
    #     (rg_string_literal, STRING_LITERAL),
    #     # (space, SPACE),
    #     (rg_named_range, NAMED_RANGE),
    # ]

    tokens = Vector{Token}()

    idx = lexer.idx
    max_idx = length(lexer.base_string)
    # Remove any leading whitespace, which sometimes shows up for some reason
    while idx <= max_idx && Base.isspace(lexer.base_string[idx])
        idx += 1
    end

    current_str = lexer.base_string[idx:end]

    iter = Automa.tokenize(TokenKind, current_str)
    for (start_idx, tok_length, kind) in iter
        if kind == TOK_ERR
            println("Tokenizer encountered an error")
            println("String = $current_str")
            println("Error occurred at index $(start_idx) with length $(tok_length)")
            if start_idx > 20
                end_idx = min(length(current_str), start_idx + tok_length + 20)
                println(current_str[start_idx - 20:end_idx])
                println(" "^20 * "^"*tok_length)
            else
                end_idx = min(length(current_str), start_idx + tok_length + 20)
                println(current_str[begin:end_idx])
                println(" "^start_idx * "^"*tok_length)
            end
            throw("Error during tokenization")
        end
        if kind != SPACE
            push!(tokens, Token(kind, SubString(current_str, start_idx, start_idx + tok_length - 1)))
        end
    end


    # while !isempty(current_str)
    #     did_match = false
    #     other = match_other_token(current_str)
    #     if !isnothing(other)
    #         # idx += length(val(other))
    #         push!(tokens, other)
    #         did_match = true
    #     else
    #         iter = @show Automa.tokenize(TokenKind, current_str)
    #         m = collect(iter)
    #         @show m
    #         @show current_str
    #         did_match = true
    #         return tokens
    #         # for (rgx, kind) in token_kinds
    #         #     m = match(rgx, current_str)
    #         #     if !isnothing(m)
    #         #         t = Token(kind, m.match)
    #         #         push!(tokens, t)
    #         #         idx += length(m.match)
    #         #         did_match = true
    #         #         break
    #         #     end
    #         # end
    #     end

    #     # if !did_match
    #     #     @show lexer.base_string
    #     #     @show current_str
    #     #     throw("Failed to match $(repr(current_str)) against any token")
    #     # end

    #     while idx <= max_idx && Base.isspace(lexer.base_string[idx])
    #         idx += 1
    #     end
    #     current_str = lexer.base_string[idx:end]

    # end


    tokens
end

function tokenize(lexer::Lexer; dbg=false)


    # token_kinds = [
    #     (rg_excel_function, EXCEL_FUNCTION),
    #     (rg_unquoted_sheet, UNQUOTED_SHEET),
    #     (rg_quoted_sheet, QUOTED_SHEET),
    #     (rg_named_range_prefixed, NAMED_RANGE_PREFIXED),
    #     (rg_cell, CELL),
    #     (rg_error_ref, ERROR_REF),
    #     (rg_boolean, BOOLEAN),
    #     (rg_horizontal_range, HORIZONTAL_RANGE),
    #     (rg_vertical_range, VERTICAL_RANGE),
    #     (rg_number, NUMBER),
    #     (rg_string_literal, STRING_LITERAL),
    #     # (space, SPACE),
    #     (rg_named_range, NAMED_RANGE),
    # ]


    idx = lexer.idx
    max_idx = length(lexer.base_string)
    # Remove any leading whitespace, which sometimes shows up for some reason
    while idx <= max_idx && Base.isspace(lexer.base_string[idx])
        idx += 1
    end

    current_str = @view lexer.base_string[idx:end]

    tokens = Vector{Token}()
    if length(current_str) > 8
        sizehint!(tokens, length(current_str) / 8)
    end
    iter = Automa.tokenize(TokenKind, current_str)
    for (start_idx, tok_length, kind) in iter
        if kind == TOK_ERR
            println("Tokenizer encountered an error")
            println("String = $current_str")
            println("Error occurred at index $(start_idx) with length $(tok_length)")
            if start_idx > 20
                end_idx = min(length(current_str), start_idx + tok_length + 20)
                println(current_str[start_idx - 20:end_idx])
                println(" "^20 * "^"*tok_length)
            else
                end_idx = min(length(current_str), start_idx + tok_length + 20)
                println(current_str[begin:end_idx])
                println(" "^start_idx * "^"^tok_length)
            end
            throw("Error during tokenization")
        end

        if kind != SPACE
            push!(tokens, Token(kind, SubString(current_str, start_idx, start_idx + tok_length - 1)))
        end
    end


    # while !isempty(current_str)
    #     dbg && @show current_str
    #     did_match = false
    #     other = match_other_token(current_str)
    #     if !isnothing(other)
    #         dbg && @show other
    #         idx += length(val(other))
    #         push!(tokens, other)
    #         did_match = true
    #     else
    #         for (rgx, kind) in token_kinds
    #             m = match(rgx, current_str)
    #             if !isnothing(m)
    #                 dbg && @show m
    #                 dbg && @show length(m.match)
    #                 t = Token(kind, m.match)
    #                 dbg && @show t
    #                 push!(tokens, t)
    #                 idx += length(m.match)
    #                 did_match = true
    #                 break
    #             end
    #         end
    #     end
    #     dbg && @show idx

    #     if !did_match
    #         # @show lexer.base_string
    #         # @show current_str
    #         throw("Failed to match $(repr(current_str)) against any token")
    #     end

    #     while idx <= max_idx && Base.isspace(lexer.base_string[idx])
    #         idx += 1
    #     end
    #     current_str = @view lexer.base_string[idx:end]

    # end


    tokens
end

tokenize(str::AbstractString; dbg=false) = tokenize(Lexer(str), dbg=dbg)


mutable struct Parser
    tokens::Vector{Token}
    idx::Int
end

function Parser(tokens::Vector{Token})
    Parser(tokens, 1)
end
function Parser(str::AbstractString)
    Parser(tokenize(str))
end


function decrement!(parser::Parser)
    parser.idx -= 1
end

function finished(parser::Parser)
    parser.idx > length(parser.tokens)
end

function peek(parser::Parser)
    parser.tokens[parser.idx]
end

function accept!(parser::Parser, tok_kind::TokenKind)
    finished(parser) && return nothing

    t = parser.tokens[parser.idx]
    if kind(t) == tok_kind
        parser.idx += 1
        t
    else
        nothing
    end
end

function accept_other!(parser::Parser, tok_str::AbstractString)
    t = accept!(parser, OTHER)

    isnothing(t) && return nothing

    if val(t) == tok_str
        t
    else
        decrement!(parser)
        nothing
    end
end


function cellname(parser::Parser)
    t = accept!(parser, CELL)
    isnothing(t) && return nothing

    ExcelExpr(:cell_ref, Any[val(t)])
end

# @test cellname(Parser("A10")) == ExcelExpr(:cell_ref, "A10")

function named_range(parser::Parser)
    t = accept!(parser, NAMED_RANGE_PREFIXED)
    if !isnothing(t)
        return ExcelExpr(:named_range, val(t))
    end
    t = accept!(parser, NAMED_RANGE)
    if !isnothing(t)
        return ExcelExpr(:named_range, val(t))
    end

    nothing
end

# @test named_range(Parser("CO2_concentration")) == ExcelExpr(:named_range, "CO2_concentration")

function vertical_range(parser::Parser)
    t = accept!(parser, VERTICAL_RANGE)
    isnothing(t) && return nothing

    ExcelExpr(:cols, val(t))
end

function error_ref(parser::Parser)
    t = accept!(parser, ERROR_REF)
    isnothing(t) && return nothing

    ExcelExpr(:error_ref, val(t))
end

function reference_item(parser::Parser)
    e = cellname(parser)
    !isnothing(e) && return e
    e = named_range(parser)
    !isnothing(e) && return e
    e = vertical_range(parser)
    !isnothing(e) && return e
    e = error_ref(parser)
    !isnothing(e) && return e

    nothing
end

function xl_float(parser::Parser)
    t = accept!(parser, NUMBER)
    isnothing(t) && return nothing

    parse(Float64, val(t))
end

function xl_bool(parser::Parser)
    t = accept!(parser, BOOLEAN)
    isnothing(t) && return nothing

    if val(t) == "TRUE"
        true
    elseif val(t) == "FALSE"
        false
    else
        throw("Unexpected boolean value $(val(t))")
    end
end

function xl_string(parser::Parser)
    t = accept!(parser, STRING_LITERAL)
    isnothing(t) && return nothing

    val(t)[2:end-1]
end

function xl_error(parser::Parser)
    t = accept!(parser, ERROR)
    isnothing(t) && return nothing

    val(t)
end

function constant(parser::Parser)
    e = xl_float(parser)
    !isnothing(e) && return e
    e = xl_bool(parser)
    !isnothing(e) && return e
    e = xl_string(parser)
    !isnothing(e) && return e
    e = xl_error(parser)
    !isnothing(e) && return e

    nothing
end

function paren(parser::Parser)
    t = accept_other!(parser, "(")
    isnothing(t) && return nothing

    inner = binop_formula(parser)
    if isnothing(inner)
        throw("No valid formula inside parentheses")
    end

    t = accept_other!(parser, ")")
    if isnothing(inner)
        throw("Missing close paren!")
    end

    inner
end

function sheet_prefix(parser::Parser)
    t = accept!(parser, UNQUOTED_SHEET)
    if !isnothing(t)
        return strip(val(t), ['\'', '!'])
    end

    t = accept!(parser, QUOTED_SHEET)
    if !isnothing(t)
        return strip(val(t), ['\'', '!'])
    end

    nothing
end

function prefixed_ref(parser::Parser)
    sheet = sheet_prefix(parser)
    isnothing(sheet) && return nothing

    first = formula(parser)
    if isnothing(first)
        throw("Missing formula after sheet prefix!")
    end

    rest = []
    while !isnothing(accept_other!(parser, ":"))
        e = formula(parser)
        if isnothing(e)
            throw("Missing formula after :")
        end

        push!(rest, e)
    end

    if isempty(rest)
        ExcelExpr(:sheet_ref, Any[sheet, first])
    else
        ExcelExpr(:sheet_ref, Any[sheet, ExcelExpr(:range, first, rest...)])
    end
end

function formula(parser::Parser)
    e = reference_item(parser)
    !isnothing(e) && return e
    e = prefixed_ref(parser)
    !isnothing(e) && return e
    e = constant(parser)
    !isnothing(e) && return e
    e = paren(parser)
    !isnothing(e) && return e
    e = function_call(parser)
    !isnothing(e) && return e

    nothing

end

function accept_other_any!(parser::Parser, tok_strs::Vector{String})
    t = accept!(parser, OTHER)
    isnothing(t) && return nothing

    if val(t) in tok_strs
        t
    else
        decrement!(parser)
        nothing
    end
end

isprefixop(::Any) = false
function isprefixop(str::AbstractString)
    str in ("+", "-")
end

ispostfixop(::Any) = false
function ispostfixop(str::AbstractString)
    str in ("%",)
end

function prefix_binding_power(op)
    @match op begin
        "+" | "-" => (nothing, 6)
        _ => throw("bad prefix op $(op)")
    end
end

function postfix_binding_power(op)
    @match op begin
        "%" => (11, nothing)
        _ => throw("bad postfix op $(op)")
    end
end

function infix_binding_power(op)
    @match op begin
        "+" => (2, 3)
        "-" => (2, 3)
        "*" => (4, 5)
        "/" => (4, 5)
        "^" => (6, 7)
        ":" => (10, 11)
        "=" => (0, 0)
        "<>" => (0, 0)
        "<=" => (0, 0)
        ">=" => (0, 0)
        "<" => (0, 0)
        ">" => (0, 0)
        "&" => (1, 2)
        _ => throw("Bad infix op %op")
    end
end

function opsymbol(op::AbstractString)::Symbol
    @match op begin
        "=" => :eq
        "<>" => :neq
        "<=" => :leq
        ">=" => :geq
        ">" => :gt
        "<" => :lt
        ":" => :range
        _ => Symbol(op)
    end
end

function parse_binops(tokens::Vector{Any}, min_bp::Int)
    # @show tokens
    lhs = popat!(tokens, 1)
    if isprefixop(lhs)
        _, r_bp = prefix_binding_power(lhs)
        rhs = parse_binops(tokens, r_bp)
        lhs = ExcelExpr(opsymbol(lhs), rhs)
    end

    while length(tokens) > 0
        op = tokens[1]
        if ispostfixop(op)
            l_bp, _ = postfix_binding_power(op)
            if l_bp < min_bp
                break
            end

            popfirst!(tokens)

            lhs = ExcelExpr(opsymbol(op), lhs)
            continue
        end

        l_bp, r_bp = infix_binding_power(op)
        if l_bp < min_bp
            break
        end
        popfirst!(tokens)

        rhs = parse_binops(tokens, r_bp)
        lhs = ExcelExpr(opsymbol(op), Any[lhs, rhs])
    end
    lhs
end
parse_binops(tokens::Vector{Any}) = parse_binops(tokens, 0)

const valid_symbols = String["+", "-", "*", "/", "%", "^", "&", "=", "<", "<=", ">", ">=", "<>", ":"]

function binop_formula_inner(parser::Parser)
    parts = Any[]
    t = accept_other_any!(parser, valid_symbols)
    if !isnothing(t)
        push!(parts, val(t))
    end

    # @show parser t
    e = formula(parser)
    # @show parser e
    isnothing(e) && return nothing
    push!(parts, e)



    while true
        t = accept_other_any!(parser, valid_symbols)
        # @show parser t
        isnothing(t) && return parts

        while !isnothing(t)
            push!(parts, val(t))
            t = accept_other_any!(parser, valid_symbols)
            # @show parser t
        end

        e = formula(parser)
        # @show parser e
        isnothing(e) && return parts
        push!(parts, e)
    end

    parts
end

function binop_formula(parser::Parser)
    parts = binop_formula_inner(parser)
    isnothing(parts) && return nothing

    parse_binops(parts)
end


function function_call(parser::Parser)
    t = accept!(parser, EXCEL_FUNCTION)
    isnothing(t) && return nothing

    fn_name = val(t)[begin:end-1]
    # @show fn_name

    args = Any[fn_name]
    right_paren = accept_other!(parser, ")")
    if !isnothing(right_paren)
        # @info "Found right parent"
        return ExcelExpr(:call, args)
    end

    while true
        # @show parser
        # tok = peek(parser)
        # e = if kind(tok) == OTHER && val(tok) == ","
        #     missing
        # else
        #     e = binop_formula(parser)
        # end

        # last_arg = accept_other!(parser, ",")
        # if isnothing(last_arg)
        #     last_arg = binop_formula(parser)
        # end

        e = binop_formula(parser)
        if isnothing(e)
            e = missing
        end
        # if isnothing(e)
        #     e = missing
        #     # if !isnothing(accept_other!(parser, ","))
        #     #     e = missing
        #     #     decrement!(parser)
        #     #     push!(args, e)
        #     # # else
        #     # #     throw("Can this happen in valid formulas?")
        #     # end
        # # else
        # end
        push!(args, e)

        right_paren = accept_other!(parser, ")")
        if !isnothing(right_paren)
            # @info "Found right parent"
            return ExcelExpr(:call, args)
        end
        # while !isnothing(accept_other!(parser, ","))
        #     # @info "Pusing missing"
        #     push!(args, missing)
        # end

        # e = binop_formula(parser)
        # if isnothing(e)
        #     throw("Can this happen in valid formulas?")
        # end
        (accept_other!(parser, ","))
    end

end

function toexpr(str::AbstractString)
    binop_formula(Parser(str))
end

# binop_formula(Parser("1% + 2 / -4 + SUM(1 / (3^5))"))
# binop_formula(Parser("IF(A10=1, \"YES\", \"NO\")"))
# tokenize("IF(H75<>\"\",AVERAGE(I75:J75),FALSE)")
# binop_formula(Parser("'Sheet 1'!A1:A2 - B10"))
# toexpr("-A1")
# function_call(Parser("SUM(1)"))

# @test toexpr("A1:B10") == ExcelExpr(:range, ExcelExpr(:cell_ref, "A1"), ExcelExpr(:cell_ref, "B10"))
# @test toexpr("SUM(A1:B20, B2)") == ExcelExpr(:call, "SUM", ExcelExpr(:range, ExcelExpr(:cell_ref, "A1"), ExcelExpr(:cell_ref, "B20")), ExcelExpr(:cell_ref, "B2"))
# @test toexpr("1 + 2") == ExcelExpr(:+, 1.0, 2.0)
# @test toexpr("'Sheet 1'!A1:A10") == ExcelExpr(:sheet_ref, "Sheet 1", ExcelExpr(:range, ExcelExpr(:cell_ref, "A1"), ExcelExpr(:cell_ref, "A10")))
# @test toexpr("'Sheet 1'!A1") == ExcelExpr(:sheet_ref, "Sheet 1", ExcelExpr(:cell_ref, "A1"))
# @test toexpr("'Sheet 1'!A1 - B10") == ExcelExpr(:-, ExcelExpr(:sheet_ref, "Sheet 1", ExcelExpr(:cell_ref, "A1")), ExcelExpr(:cell_ref, "B10"))
# @test toexpr("(C4/2204.62)*F38") == ExcelExpr(:*, ExcelExpr(:/, ExcelExpr(:cell_ref, "C4"), 2204.62), ExcelExpr(:cell_ref, "F38"))
# @test toexpr("total_harvest") == ExcelExpr(:named_range, "total_harvest")
# @test toexpr("1 * 2 - 4") == ExcelExpr(:-, ExcelExpr(:*, 1.0, 2.0), 4.0)
# @test toexpr("1 - 2 * 4") == ExcelExpr(:-, 1, ExcelExpr(:*, 2, 4))
# @test toexpr("IF(salary_worker_toggle=1,((C309*(I143-I142))+(I142*C310))*F136,I136*I140*I142*(((F138-1)*'1. Assumptions'!F105)+('1. Assumptions'!F106*1)))*(1+'1. Assumptions'!F322)*F313*array_deploy_scalar") == ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:call, "IF", ExcelExpr(:eq, ExcelExpr(:named_range, "salary_worker_toggle"), 1.0), ExcelExpr(:*, ExcelExpr(:+, ExcelExpr(:*, ExcelExpr(:cell_ref, "C309"), ExcelExpr(:-, ExcelExpr(:cell_ref, "I143"), ExcelExpr(:cell_ref, "I142"))), ExcelExpr(:*, ExcelExpr(:cell_ref, "I142"), ExcelExpr(:cell_ref, "C310"))), ExcelExpr(:cell_ref, "F136")), ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:cell_ref, "I136"), ExcelExpr(:cell_ref, "I140")), ExcelExpr(:cell_ref, "I142")), ExcelExpr(:+, ExcelExpr(:*, ExcelExpr(:-, ExcelExpr(:cell_ref, "F138"), 1.0), ExcelExpr(:sheet_ref, "1. Assumptions", ExcelExpr(:cell_ref, "F105"))), ExcelExpr(:*, ExcelExpr(:sheet_ref, "1. Assumptions", ExcelExpr(:cell_ref, "F106")), 1.0)))), ExcelExpr(:+, 1.0, ExcelExpr(:sheet_ref, "1. Assumptions", ExcelExpr(:cell_ref, "F322")))), ExcelExpr(:cell_ref, "F313")), ExcelExpr(:named_range, "array_deploy_scalar"))
# @test toexpr("(N15-#REF!)/#REF!") == ExcelExpr(:/, ExcelExpr(:-, ExcelExpr(:cell_ref, "N15"), ExcelExpr(:error_ref, "#REF!")), ExcelExpr(:error_ref, "#REF!"))
# @test toexpr("(F262*F263*F264)/(3.28^3)*(1+F240)") == ExcelExpr(:*, ExcelExpr(:/, ExcelExpr(:*, ExcelExpr(:*, ExcelExpr(:cell_ref, "F262"), ExcelExpr(:cell_ref, "F263")), ExcelExpr(:cell_ref, "F264")), ExcelExpr(:^, 3.28, 3.0)), ExcelExpr(:+, 1.0, ExcelExpr(:cell_ref, "F240")))
# @test toexpr("RAND()") == ExcelExpr(:call, "RAND")
# @test toexpr("IF(H75<\$I\$71,H75+1,\"\")") == ExcelExpr(:call, "IF", ExcelExpr(:lt, ExcelExpr(:cell_ref, "H75"), ExcelExpr(:cell_ref, "\$I\$71")), ExcelExpr(:+, ExcelExpr(:cell_ref, "H75"), 1.0), "")
# @test toexpr("IF(H75<>\"\",COUNTIFS(\$M\$4:\$ZZ\$4,\">=\"&I75,\$M\$4:\$ZZ\$4,\"<=\"&J75),\"\")") == ExcelExpr(:call, "IF", ExcelExpr(:neq, ExcelExpr(:cell_ref, "H75"), ""), ExcelExpr(:call, "COUNTIFS", ExcelExpr(:range, ExcelExpr(:cell_ref, "\$M\$4"), ExcelExpr(:cell_ref, "\$ZZ\$4")), ExcelExpr(:&, ">=", ExcelExpr(:cell_ref, "I75")), ExcelExpr(:range, ExcelExpr(:cell_ref, "\$M\$4"), ExcelExpr(:cell_ref, "\$ZZ\$4")), ExcelExpr(:&, "<=", ExcelExpr(:cell_ref, "J75"))), "")
# @test toexpr("IF(H75<>\"\",AVERAGE(I75:J75),FALSE)") == ExcelExpr(:call, "IF", ExcelExpr(:neq, ExcelExpr(:cell_ref, "H75"), ""), ExcelExpr(:call, "AVERAGE", ExcelExpr(:range, ExcelExpr(:cell_ref, "I75"), ExcelExpr(:cell_ref, "J75"))), false)
# @test toexpr("-A1") == ExcelExpr(:-, ExcelExpr(:cell_ref, "A1"))
# @test toexpr("100%-equity") == ExcelExpr(:-, ExcelExpr(:%, 100.0), ExcelExpr(:named_range, "equity"))
# @test toexpr("IF(G131=\"\",,1)") === ExcelExpr(:call, "IF", ExcelExpr(:eq, ExcelExpr(:cell_ref, "G131"), ""), missing, 1.0)
# @test toexpr("OFFSET(Lists!\$G\$4,0,0,COUNTA(Lists!\$G:\$G),2)") == ExcelExpr(:call, "OFFSET", ExcelExpr(:sheet_ref, "Lists", ExcelExpr(:cell_ref, "\$G\$4")), 0.0, 0.0, ExcelExpr(:call, "COUNTA", ExcelExpr(:sheet_ref, "Lists", ExcelExpr(:cols, "\$G:\$G"))), 2.0)
# @test toexpr("SUM(1, 2,)") === ExcelExpr(:call, "SUM", 1.0, 2.0, missing)
# end
