--find lab texture where psda is not  provided
select * from result 
--delete result
where result_source_key in (
select layer_key
from (
select b.layer_key, b.source_system_id, c, texture

from (
select * from bridge_layer_ss --where source_system_key = 26
) b 
left join (
select result_source_key, count(*) c
from result
where 
(method_code like 'clay_tot%'
or method_code like 'sand_d%'
or method_code like 'silt_d%')
--and source_system_key = 26
group by result_source_key
) psda on b.layer_key = psda.result_source_key

left join (
select result_source_key, result_value texture
from result
where 
method_code like 'text%'
--and source_system_key = 26
) tex on b.layer_key = tex.result_source_key

where c is null and not texture is null
) a
)
and method_code like 'text%'




select *

from result

where result_source_key in (

select result_source_key
 
from result
where source_system_key = 13 and analyte_key = 569
and result_value like '%-%'

)

select analyte_key, method_code, count(*) c

from result

where source_system_key = 19
group by analyte_key, method_code
order by method_code

select result_value, count(*) c

from result

where source_system_key = 22
and method_code like 'text%'
group by result_value
order by result_value

select *
from result
where result_source_key in (
select result_source_key
from result
where source_system_key = 22
and method_code like 'text%'
and result_value in ('muck', 'clay loam')
)



