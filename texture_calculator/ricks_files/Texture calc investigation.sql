
select * from (
select x.*, sa.total_sand, si.total_silt, cl.total_clay, 
case when (cast(total_silt as float) + 1.5 * cast(total_clay as float)) < 15.0 then 's'  
     when (cast(total_silt as float) + 1.5*cast(total_clay as float) >= 15.0) and (cast(total_silt as float) + 2*cast(total_clay as float) < 30.0) then 'ls'
	 when (cast(total_clay as float) >= 7.0 and cast(total_clay as float) < 20.0) and (cast(total_sand as float) > 52.0) and (((cast(total_silt as float) + 2*cast(total_clay as float)) >= 30.0) or (cast(total_clay as float) < 7.0 and cast(total_silt as float) < 50.0 and (cast(total_silt as float) + 2*cast(total_clay as float))>=30.0)) then 'sl'
     else null end new_result
from (

select r.source_system_key, a.analyte_key, a.analyte_code, r.result_value, r.result_source_key
from result r
inner join (
select analyte_key, analyte_name, analyte_code
from analyte 
where analyte_name like '%text%'
) a on r.analyte_key = a.analyte_key

where r.source_system_key > 2 and r.analyte_key = 569

--group by r.source_system_key, a.analyte_code
) x
left join (
select result_source_key, count(result_key) c
from result r
where analyte_key in (321,318,319,317,320)
and r.source_system_key > 2
group by result_source_key
) z on x.result_source_key = z.result_source_key

inner join (
select result_source_key, result_value as total_sand
from result r
where analyte_key in (565)
and r.source_system_key > 2
) sa on x.result_source_key = sa.result_source_key
inner join (
select result_source_key, result_value as total_silt
from result r
where analyte_key in (567)
and r.source_system_key > 2
) si on x.result_source_key = si.result_source_key
inner join (
select result_source_key, result_value as total_clay
from result r
where analyte_key in (313)
and r.source_system_key > 2
) cl on x.result_source_key = cl.result_source_key
--"565,567,313" 
where  c is null
and (isnumeric(total_sand) <> 0 and isnumeric(total_silt) <> 0 and isnumeric(total_clay) <> 0)
) a

where result_value <> new_result and not new_result is null

order by new_result, source_system_key




-- select * from source_system where source_system_key  = 15
