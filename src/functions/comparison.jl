xl_eq(a, b) = a == b
xl_eq(::Missing, b::String) = b == ""
xl_eq(a::String, ::Missing) = a == ""
xl_eq(a::Bool, b::Bool) = a == b
xl_eq(::Bool, ::Number) = false
xl_eq(::Number, ::Bool) = false

xl_gt(a, b) = a > b
xl_gt(a::Bool, b::Bool) = a > b
xl_gt(::Bool, ::Number) = true
xl_gt(::Number, ::Bool) = false

xl_lt(a, b) = a < b
xl_lt(a::Bool, b::Bool) = a < b
xl_lt(::Bool, ::Number) = false
xl_lt(::Number, ::Bool) = true

xl_geq(a, b) = a >= b
xl_geq(a::Bool, b::Bool) = a >= b
xl_geq(::Bool, ::Number) = true
xl_geq(::Number, ::Bool) = false

xl_leq(a, b) = a <= b
xl_leq(a::Bool, b::Bool) = a <= b
xl_leq(::Bool, ::Number) = false
xl_leq(::Number, ::Bool) = true
