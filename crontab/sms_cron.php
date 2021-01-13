<?php
	ini_set("display_errors",true);
	
	include("conn.php");
	$conn_sms=new open_conn("sms");
	$conn_crm=new open_conn();
	
	$rs_sms=pg_query($conn_sms->get_conn(),"SELECT \"ID\",\"ReceivingDateTime\",\"SenderNumber\",\"TextDecoded\" FROM inbox WHERE \"Processed\"='f' AND f_notif=0 ORDER BY \"ReceivingDateTime\" DESC");
	while($row = pg_fetch_array($rs_sms)){
		$id=$row["ID"];
		$receive_date=$row["ReceivingDateTime"];
		$sender_number=$row["SenderNumber"];
		$sender_number=str_replace("+62","0",$sender_number);
		$text=$row["TextDecoded"];
		// SET FLAG = NOTIF
		pg_query($conn_sms->get_conn(),"UPDATE inbox SET f_notif=1 WHERE \"ID\"= $id");
		// GET CUST ID AND AGENT
		$rs_query=pg_query($conn_crm->get_conn(),"SELECT * FROM (SELECT distinct custid,phone_no FROM (SELECT custid,contact1 as phone_no FROM mandiri.tbl_address UNION ALL
												SELECT custid,contact2 as phone_no FROM mandiri.tbl_address UNION ALL
												SELECT custid,mobileno as phone_no FROM mandiri.tbl_address) x WHERE coalesce(phone_no,'')<>'' ORDER BY 1) x,(SELECT custid,agent FROM mandiri.mgm) y WHERE x.custid=y.custid AND phone_no='$sender_number' AND agent IS NOT NULL LIMIT 1");
		$rs_result=pg_fetch_assoc($rs_query);
		if (count($rs_result)>0){
			pg_query($conn_crm->get_conn(),"DELETE FROM mandiri.tbl_notif_sms WHERE id_sms=$id");
			pg_query($conn_crm->get_conn(),"INSERT INTO mandiri.tbl_notif_sms(id_sms,received_sms_date,sender_number,text_sms,agent,custid) 
			VALUES ($id,'$receive_date','$sender_number','$text','$rs_result[agent]','$rs_result[custid]')");
		}
	}
?>