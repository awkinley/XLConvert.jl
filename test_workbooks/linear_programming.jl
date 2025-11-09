using XLConvert
using DataFrames

using Dates

function if_multiple(dividend, divisor, value)
xl_compare(xl_mod(dividend, divisor), 0) ? value : 0.0
end

struct Outputs
	s_Sheet1_E10
	s_Sheet1_E11
	s_Sheet1_E12
end
@kwdef mutable struct Inputs
	# used in 1 statements, [tab_Sheet1_E10_E12[1, "E"], tab_Sheet1_E10_E12[2, "E"], tab_Sheet1_E10_E12[3,...]
	X::Float64 = 4.25
	# used in 1 statements, [tab_Sheet1_E10_E12[1, "E"], tab_Sheet1_E10_E12[2, "E"], tab_Sheet1_E10_E12[3,...]
	Y::Float64 = 5.5
end
struct Tables
	tab_Sheet1_coeffs::DataFrame
	tab_Sheet1_E10_E12::DataFrame
end

function make_input_tables()
	tab_Sheet1_coeffs = DataFrame("X" => zeros(3), "Y" => zeros(3))
	tab_Sheet1_E10_E12 = DataFrame("E" => zeros(3))

	tab_Sheet1_coeffs[!, Between("X", "Y")] .= [2.0 3.0;4.0 6.0;2.0 5.0]

	Tables(
		tab_Sheet1_coeffs,
		tab_Sheet1_E10_E12,
	)
end
function calculate(inputs::Inputs, tables::Tables)
tab_Sheet1_coeffs = tables.tab_Sheet1_coeffs
tab_Sheet1_E10_E12 = tables.tab_Sheet1_E10_E12
# Level 0


# Level 1
# Used in 1 places: [OutputStatement]
# "Sheet1!E10":"Sheet1!E12"
@. tab_Sheet1_E10_E12[!, "E"] = tab_Sheet1_coeffs[!, "X"] * inputs.X + tab_Sheet1_coeffs[!, "Y"] * inputs.Y


# Level 2
Outputs(
    tab_Sheet1_E10_E12[1, "E"],
	tab_Sheet1_E10_E12[2, "E"],
	tab_Sheet1_E10_E12[3, "E"]    
)


end

function run_crest_solar()
    inputs = Inputs()
    tables = make_input_tables()
    calculate(inputs, tables)
end
