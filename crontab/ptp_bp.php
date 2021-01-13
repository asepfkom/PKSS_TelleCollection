<?php
	ini_set("display_errors",true);
	
	include("conn.php");
	$conn_crm=new open_conn();
	
	pg_query($conn_crm->get_conn(),"UPDATE mandiri.mgm SET statuscall='BP' WHERE custid IN 
	(SELECT x.custid FROM (SELECT custid, max(date(promisedate)) ptp FROM mandiri.tblnegoptp GROUP BY 1) x LEFT JOIN (SELECT custid,max(date(paydate)) as tglbayar FROM mandiri.tbllunas GROUP BY 1) y 
	ON x.custid=y.custid WHERE ptp>tglbayar AND ptp+interval '3 day' <=date(now()));");
	
	pg_query($conn_crm->get_conn(),"INSERT INTO mandiri.mgm_hst (custid,agent,f_cek,statuscall,f_cek_new,tgl,hst,user_log) 
	select custid,agent,f_cek_new,f_cek_new,f_cek_new,now(),'MELEWATI TANGGAL PTP','SYSTEM' from mandiri.mgm where custid in (
	SELECT x.custid FROM (SELECT custid, max(date(promisedate)) ptp FROM mandiri.tblnegoptp GROUP BY 1) x LEFT JOIN 
	(SELECT custid,max(date(paydate)) as tglbayar FROM mandiri.tbllunas GROUP BY 1) y 
	ON x.custid=y.custid WHERE ptp>tglbayar AND ptp+interval '3 day' <=date(now()));");

?>