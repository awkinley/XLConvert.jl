
struct ObjectNumbering{T}
    objs::Vector{T}
    obj_nums::Dict{T, Int64}
end

get_num(numbering::ObjectNumbering{T}, obj::T) where {T}  = numbering.obj_nums[obj]

# This variant will insert if it doesn't exist
function get_num!(numbering::ObjectNumbering{T}, obj::T) where {T} 
    num = get(numbering.obj_nums, obj, length(numbering) + 1) 
    if num == length(numbering) + 1
        push!(numbering.objs, obj)
        numbering.obj_nums[obj] = num
    end
    num
end

get_obj(numbering::ObjectNumbering{T}, i::Int64) where {T} = numbering.objs[i]

Base.length(numbering::ObjectNumbering{T}) where {T} = length(numbering.objs)

function ObjectNumbering(objs::Vector{T}) where {T}
    numbers = Dict{T, Int64}(n => i for (i, n) in enumerate(objs))

    ObjectNumbering(objs, numbers)
end