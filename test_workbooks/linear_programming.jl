using XLConvert
using DataFrames

using Dates

function if_multiple(dividend, divisor, value)
xl_compare(xl_mod(dividend, divisor), 0) ? value : 0.0
end

struct Outputs
	solver_opt
	solver_lhs1
	solver_lhs2
end
@kwdef mutable struct Inputs
	# used in 3 statements, [solver_lhs1], [solver_opt], [solver_lhs2]
	X::Float64 = 4.25
	# used in 3 statements, [solver_lhs1], [solver_opt], [solver_lhs2]
	Y::Float64 = 5.5
end
struct Tables
	tab_Sheet1_coeffs::DataFrame
end

function make_input_tables()
	tab_Sheet1_coeffs = DataFrame("X" => zeros(3), "Y" => zeros(3))

	tab_Sheet1_coeffs[!, Between("X", "Y")] .= [2.0 3.0;4.0 6.0;2.0 5.0]

	Tables(
		tab_Sheet1_coeffs,
	)
end
function calculate(inputs::Inputs, tables::Tables)
tab_Sheet1_coeffs = tables.tab_Sheet1_coeffs
# Level 0


# Level 1
# Used in 1 places: [OutputStatement]
# =C10*X+D10*Y
solver_opt = tab_Sheet1_coeffs[1, "X"] * inputs.X + tab_Sheet1_coeffs[1, "Y"] * inputs.Y # Sheet1 E10
@assert xl_compare(solver_opt, 25) # "Sheet1!E10"
# Used in 1 places: [OutputStatement]
# =C11*X+D11*Y
solver_lhs1 = tab_Sheet1_coeffs[2, "X"] * inputs.X + tab_Sheet1_coeffs[2, "Y"] * inputs.Y # Sheet1 E11
@assert xl_compare(solver_lhs1, 50) # "Sheet1!E11"
# Used in 1 places: [OutputStatement]
# =C12*X+D12*Y
solver_lhs2 = tab_Sheet1_coeffs[3, "X"] * inputs.X + tab_Sheet1_coeffs[3, "Y"] * inputs.Y # Sheet1 E12
@assert xl_compare(solver_lhs2, 36) # "Sheet1!E12"


# Level 2
Outputs(
    solver_opt,
	solver_lhs1,
	solver_lhs2    
)


end

function run_crest_solar()
    inputs = Inputs()
    tables = make_input_tables()
    calculate(inputs, tables)
end
