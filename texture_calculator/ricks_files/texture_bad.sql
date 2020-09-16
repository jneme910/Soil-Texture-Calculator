select source_system_name, analyte_key, analyte_abbrev, result_value, c number_occurrences

from source_system

inner join (
select source_system_key, r.analyte_key, analyte_abbrev, result_value, count(*) c
from result r
inner join analyte a on r.analyte_key = a.analyte_key
where method_code like 'text%'
and result_value in (
'{FS}[MUCK]',
'{LS}[SL]',
'error. total = 26',
'MK',
'S PROB',
'Sandstone',
'Sandstone and Shale',
'SFL',
'GRFSL',
'CS',
'CS',
'fs1',
'LOCS',
'LCS',
'Limestone-shale-siltstone',
'SAND',
'Siltstone, sandstone',
'clay loam',
'GYPSIC',
'Limestone',
'PT',
'Muck',
'sc1',
'Shale',
'SL PROB',
'[MUCK]',
'CS',
'error. total = 31.5',
'error. total = 34',
'error. total = 40.5',
'GRC',
'Limestone-Shale',
'LOCS',
'MUCKY CLAY',
'Mucky fine sand',
'MUCKY PEAT',
'Siltstone',
'{FS}[MK-FS]',
'{LFS}[FSL]',
'Chert rock',
'error. total ? 100.',
'CS',
'LD',
'MARL',
'mucky sand',
'peat',
'S PR',
' c ',
'[MK-SC]',
'{FS}[LFS]',
'CSL',
'error. total = 15.5',
'MK',
'mucky sandy loam',
'Sandstone, siltstone, shale',
'Shale and Siltstone',
'[MK-LFS]',
'error. total = 27.5',
'error. total = 28.5',
'G-LCOS',
'C0SL',
'Mucky Marl',
'MUCKY-PEAT'
) group by source_system_key, r.analyte_key, analyte_abbrev, result_value
) a on source_system.source_system_key = a.source_system_key

order by a.source_system_key, analyte_key, result_value

