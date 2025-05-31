"""
The base type of all statements. 

A statement is basically the building block of the internal intermediate representation of some part of the output code.

The basic interface that a statement should implement is

get_cell_deps(stmt::AbstractStatement) -> Vector{CellDependency}
get_set_cells(stmt::TableStatement) -> Vector{CellDependency}

Statements should also probably implement 

export_statement(exporter::JuliaExporter, wb::ExcelWorkbook, statement::AbstractStatement) -> String

If you're supporting multiple exporters, you probably want a different definition for each exporter.

"""
abstract type AbstractStatement end;

include("./statements/standard_statement.jl")
include("./statements/output_statement.jl")
include("./statements/table_statement.jl")
include("./statements/grouped_statement.jl")
include("./statements/function_statement.jl")

