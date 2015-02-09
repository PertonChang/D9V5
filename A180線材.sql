select * from ieod26h
where rfq_list='A14010291'		-- 匯總單號

select * from ieod25h 
where rfq_list='A14010291' and rfq_inquiry='8-A1800-0479-006'	-- ==> Q08120057R4

select * from ieod25h 
where rfq_inquiry like '8-A1800-0479-00%'

// D10/Q1 線材
select rfq_price,rfq_cuqty,* from ieod25d3
where rfq_no='Q08120057R4'

--S11F0664	0001	309442-C02	8-A1800-0479-003
--S11G0487	0002	309442-C02	8-A1800-0479-003
--S11H0782	0003	309442-C02	8-A1800-0479-003
--S11I0271	0002	309442-C02	8-A1800-0479-003
--S11I1355	0003	309442-C02	8-A1800-0479-003
select od_serno,od_seq,cu_elnoold,qu_no,* from ieod01d1
where od_serno in ('S11F0664','S11G0487','S11H0782','S11I0271','S11I1355') and cu_elnoold='309442-C02'
