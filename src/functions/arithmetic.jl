xl_add(lhs, rhs) = lhs + rhs
xl_add(::Missing, ::Missing) = missing
xl_add(lhs, ::Missing) = lhs
xl_add(::Missing, rhs) = rhs
# Adding to a date adds days
xl_add(lhs::Date, rhs::T) where {T<:Number} = lhs + Dates.Day(rhs)
xl_add(lhs::T, rhs::Date) where {T<:Number} = Dates.Day(lhs) + rhs
# This exist because Date - Date gives Dates.Period, instead of interger
xl_add(lhs::Dates.Period, rhs::T) where {T<:Number} = Dates.value(Dates.days(lhs) + rhs)
xl_add(lhs::T, rhs::Dates.Period) where {T<:Number} = Dates.value(lhs + Dates.days(rhs))


xl_sub(lhs, rhs) = lhs - rhs
xl_sub(::Missing, ::Missing) = missing
xl_sub(lhs, ::Missing) = lhs
xl_sub(::Missing, rhs) = rhs
# subing to a date subs days
xl_sub(lhs::Date, rhs::T) where {T<:Number} = lhs - Dates.Day(rhs)
xl_sub(lhs::T, rhs::Date) where {T<:Number} = Dates.Day(lhs) - rhs
# This exist because Date - Date gives Dates.Period, instead of interger
xl_sub(lhs::Dates.Period, rhs::T) where {T<:Number} = Dates.value(Dates.days(lhs) - rhs)
xl_sub(lhs::T, rhs::Dates.Period) where {T<:Number} = Dates.value(lhs - Dates.days(rhs))
# This might actually address the issue of having to handle Periods?
xl_sub(lhs::Date, rhs::Date) = Dates.value(lhs - rhs)


xl_mul(a, b) = a * b
xl_mul(lhs::Dates.Day, rhs::T) where {T<:Number} = Dates.value(lhs) * rhs
xl_mul(lhs::T, rhs::Dates.Day) where {T<:Number} = lhs * Dates.value(rhs)
xl_mul(::Missing, ::Missing) = 0.0
xl_mul(a, ::Missing) = 0.0
xl_mul(::Missing, a) = 0.0

xl_div(a, b) = a / b
xl_div(lhs::Dates.Day, rhs::T) where {T<:Number} = Dates.value(lhs) / rhs
xl_div(lhs::T, rhs::Dates.Day) where {T<:Number} = lhs / Dates.value(rhs)
