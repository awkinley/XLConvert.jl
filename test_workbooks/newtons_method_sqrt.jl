using XLConvert
using DataFrames

using Dates

function if_multiple(dividend, divisor, value)
xl_compare(xl_mod(dividend, divisor), 0) ? value : 0.0
end

struct Outputs
	s_Sheet1_C20
end
@kwdef mutable struct Inputs
	# used in 1 statements, [tab_Sheet1_value_series[1, "values"], tab_Sheet1_value_series[2, "values"], t...]
	radicand::Float64 = 1025.0
end
struct Tables
	tab_Sheet1_value_series::DataFrame
end

function make_input_tables()
	tab_Sheet1_value_series = DataFrame("values" => zeros(12))


	Tables(
		tab_Sheet1_value_series,
	)
end
function calculate(inputs::Inputs, tables::Tables)
tab_Sheet1_value_series = tables.tab_Sheet1_value_series
# Level 0


# Level 1
# Used in 1 places: [OutputStatement]
# Group of 13 statements
begin
tab_Sheet1_value_series[1, "values"] = inputs.radicand / 2.0 # Sheet1 C6 Row: 1
@assert xl_compare(tab_Sheet1_value_series[1, "values"], 512.5) # "Sheet1!C6"
for i in 0:10
	tab_Sheet1_value_series[2 + i, 1] = tab_Sheet1_value_series[1 + i, 1] + (inputs.radicand - tab_Sheet1_value_series[1 + i, 1] * tab_Sheet1_value_series[1 + i, 1]) / (2.0 * tab_Sheet1_value_series[1 + i, 1])
end
# =C17*C17 - $C$4
err = tab_Sheet1_value_series[12, "values"] * tab_Sheet1_value_series[12, "values"] - inputs.radicand # Sheet1 C20
@assert xl_compare(err, 0.0) # "Sheet1!C20"
end


# Level 2
Outputs(
    err    
)


end

function run_crest_solar()
    inputs = Inputs()
    tables = make_input_tables()
    calculate(inputs, tables)
end
