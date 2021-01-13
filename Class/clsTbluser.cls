VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTbluser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class untuk semua modul user
' 1) Modul Manager
Public Function saveManager(sUserid As String, sName As String, sAktif As String, sPwd As String, sNminput As String) As Boolean
Dim M_objrs  As New ADODB.Recordset
Dim dExpired As Date
On Error GoTo out
    saveManager = True
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        sStrsql = " select current_date + cast('30 day' as interval) as tglexpired "
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        If Not M_objrs.EOF Then
          dExpired = IIf(IsNull(M_objrs!tglexpired), Null, M_objrs!tglexpired)
        End If
    Set M_objrs = Nothing
        
    sStrsql = "insert into   usertbl (userid,agent,usertype,lvl ,aktif ,accrec, tbluser_tglexpired) "
    sStrsql = sStrsql + " values    ('" + sUserid + "','" + sName + "','5','Manager','1','" + sPwd + "','" + CStr(Format(dExpired, "yyyy-mm-dd")) + "')"
    M_OBJCONN.Execute (sStrsql)
    
    
    STRSQL = " UPDATE TBLUSER SET web_password=MD5('PASS!@#$%') WHERE tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (STRSQL)
    
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'New User','" + sNminput + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)

      STRSQL = "insert into  tbluser_log (aksi,field_nama,field_nilai,user_by ,sts_form ) values ('New','Userid,Nama,Status','" + sUserid + "," + sName + ",Active " + " ','" + MDIForm1.TxtUsername.Text + "','5')"
      M_OBJCONN.Execute (STRSQL)
      
    Exit Function
out:
    MsgBox err.Description, vbInformation + vbOKOnly, App.Title
    saveManager = False
End Function
Public Function updateManager(sUserid As String, sName As String, sAktif As String, sPwd As String, sNminput As String, field_awal As String, field_nilai As String) As Boolean
On Error GoTo out
    updateManager = True
    
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_tglupdate,tbluser_hst_tglentry ,tbluser_hst_tglregister ,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Update User',tbluser_tglupdate,tbluser_tglentry ,tbluser_tglregister,'" + sNminput + "' as nminput "
'    sstrsql = sstrsql + " from  tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
    
    
    sStrsql = "update tbluser set  tbluser_name='" + sName + "', tbluser_tglupdate =now(),tbluser_kdstatus ='" + sAktif + "'"
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    sStrsql = "update tbluser set tbluser_mgrname='" + sName + "'  where tbluser_userid='" + sUserid + "' "
    M_OBJCONN.Execute (sStrsql)
       
     STRSQL = "insert into  tbluser_log (userid,aksi,field_nama,field_nilai,user_by ,sts_form ) values ('" + sUserid + "','Edit','" + field_awal + "','" + field_nilai + "','" + MDIForm1.TxtUsername.Text + "','5')"
      M_OBJCONN.Execute (STRSQL)
    
    
   
    Exit Function
out:
    MsgBox err.Description
    updateManager = False
End Function
Public Function deleteManager(sUserid As String, sNminput As String, field_awal As String, field_nilai As String) As Boolean
    On Error GoTo out
    deleteManager = True
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_tglupdate,tbluser_hst_tglentry ,tbluser_hst_tglregister,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Delete User',tbluser_tglupdate,tbluser_tglentry ,tbluser_tglregister,'" + sNminput + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserID + "'"
'    M_OBJCONN.Execute (sstrsql)
    
    
    sStrsql = "DELETE FROM tbluser "
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
      
    STRSQL = "INSERT INTO tbluser_log (userid ,aksi,field_nama,field_nilai,user_by ,sts_form ) values ('" + sUserid + "','Delete','" + field_awal + "','" + field_nilai + "','" + MDIForm1.TxtUsername.Text + "','5')"
    M_OBJCONN.Execute (STRSQL)
   
    Exit Function
out:
    MsgBox err.Description
    deleteManager = False
End Function
Public Function finduseridManager(sUserid As String) As ADODB.Recordset
On Error GoTo out
    sStrsql = "select * from  tbluser where tbluser_userid='" + sUserid + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set finduseridManager = M_objrs
    Set M_objrs = Nothing
    Exit Function
out:
    MsgBox err.Description
End Function
Public Function getrecordallManager() As ADODB.Recordset
On Error GoTo out
    sStrsql = "select * from  tbluser,tbllevel,tblstatus where tbluser.tbluser_kdlevel=tbllevel.tbllevel_kdlevel"
    sStrsql = sStrsql + " and  tbluser.tbluser_kdstatus= tblstatus.tblstatus_kdstatus  and  tbluser_lockkdstatus=tblstatus.tblstatus_kdstatus "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set getrecordallManager = M_objrs
    Set M_objrs = Nothing
    Exit Function
out:
    MsgBox err.Description
End Function
'2 Modul Asisten Manager
Public Function saveAM(sUserid As String, sName As String, sAktif As String, sPwd As String, scodeManager As String, snameManager As String, sNminput As String) As Boolean
Dim M_objrs  As New ADODB.Recordset
Dim dExpired As Date
On Error GoTo out
    saveAM = True
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        sStrsql = " select current_date + cast('30 day' as interval) as tglexpired "
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        If Not M_objrs.EOF Then
          dExpired = IIf(IsNull(M_objrs!tglexpired), Null, M_objrs!tglexpired)
        End If
    Set M_objrs = Nothing
        
    sStrsql = "insert into   tbluser (tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired) "
    sStrsql = sStrsql + " values    ('" + sUserid + "','" + sName + "','4','Assisten Manager','1','9','" + sPwd + "','" + CStr(Format(dExpired, "yyyy-mm-dd")) + "')"
    M_OBJCONN.Execute (sStrsql)
    
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid, tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'New User','" + sNminput + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
    
    STRSQL = "insert into  tbluser_log (aksi,field_nama,field_nilai,user_by ,sts_form ) values ('New','Userid,Nama,Status','" + sUserid + "," + sName + ",Active " + " ','" + MDIForm1.TxtUsername.Text + "','4')"
      M_OBJCONN.Execute (STRSQL)
 
 
    Exit Function
out:
    MsgBox err.Description
    saveAM = False
End Function
Public Function updateAM(sUserid As String, sName As String, sAktif As String, sPwd As String, scodeManager As String, snameManager As String, sNminput As String, field_awal As String, field_nilai As String) As Boolean
On Error GoTo out
    updateAM = True
    
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Update User','" + sNminput + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
    
    
    sStrsql = "update tbluser set  tbluser_name='" + sName + "', tbluser_tglupdate =now(),tbluser_kdstatus ='" + sAktif + "',tbluser_mgrcode='" + scodeManager + "',tbluser_mgrname='" + snameManager + "' "
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
      STRSQL = "insert into  tbluser_log (userid,aksi,field_nama,field_nilai,user_by ,sts_form ) values ('" + sUserid + "','Edit','" + field_awal + "','" + field_nilai + "','" + MDIForm1.TxtUsername.Text + "','4')"
      M_OBJCONN.Execute (STRSQL)
    
    
   
    Exit Function
out:
    MsgBox err.Description
    updateAM = False
End Function
Public Function deleteAM(sUserid As String, sNminput As String, field_awal As String, field_nilai As String) As Boolean
On Error GoTo out
    deleteAM = True
  
    
    
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Delete User','" + sNminput + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
'
    STRSQL = "insert into  tbluser_log (userid ,aksi,field_nama,field_nilai,user_by ,sts_form ) values ('" + sUserid + "','Delete','" + field_awal + "','" + field_nilai + "','" + MDIForm1.TxtUsername.Text + "','4')"
    M_OBJCONN.Execute (STRSQL)
    
    
    sStrsql = "delete  from  tbluser "
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
   
    Exit Function
out:
    MsgBox err.Description
    deleteAM = False
End Function
Public Function finduseridAM(sUserid As String) As ADODB.Recordset
On Error GoTo out
    sStrsql = "select * from  tbluser where tbluser_userid='" + sUserid + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set finduseridAM = M_objrs
    Set M_objrs = Nothing
    Exit Function
out:
    MsgBox err.Description
End Function
Public Function getrecordallAM() As ADODB.Recordset
On Error GoTo out
    sStrsql = "select * from  tbluser,tbllevel,tblstatus where tbluser.tbluser_kdlevel=tbllevel.tbllevel_kdlevel"
    sStrsql = sStrsql + " and  tbluser.tbluser_kdstatus= tblstatus.tblstatus_kdstatus  and  tbluser_lockkdstatus=tblstatus.tblstatus_kdstatus "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set getrecordallAM = M_objrs
    Set M_objrs = Nothing
    Exit Function
out:
    MsgBox err.Description
End Function
'3 Modul Supervisor
Public Function saveSpv(sUserid As String, sName As String, sAktif As String, sPwd As String, amcode As String, amname As String, sNminput As String, JenisQa As String) As Boolean
    Dim M_objrs  As New ADODB.Recordset
    Dim dExpired As Date
    On Error GoTo out
    saveSpv = True
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    sStrsql = " select current_date + cast('30 day' as interval) as tglexpired "
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
      dExpired = IIf(IsNull(M_objrs!tglexpired), Null, M_objrs!tglexpired)
    End If
    Set M_objrs = Nothing

    If JenisQa = "QA" Then
        sStrsql = "insert into   tbluser (tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired,tbluser_amcode,tbluser_amname,tbluser_lastuser) "
        sStrsql = sStrsql + " values    ('" + sUserid + "','" + sName + "','8','Head Quality Monitoring','1','9','" + sPwd + "','" + CStr(Format(dExpired, "yyyy-mm-dd")) + "','" + amcode + "','" + amname + "','" + sNminput + "')"
        M_OBJCONN.Execute (sStrsql)
        STRSQL = "insert into  tbluser_log (aksi,field_nama,field_nilai,user_by ,sts_form ) values ('New','Userid,Nama,Status','" + sUserid + "," + sName + ",Active " + " ','" + MDIForm1.TxtUsername.Text + "','8')"
        M_OBJCONN.Execute (STRSQL)
    Else
        sStrsql = "insert into   tbluser (tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired,tbluser_amcode,tbluser_amname,tbluser_lastuser) "
        sStrsql = sStrsql + " values    ('" + sUserid + "','" + sName + "','2','Supervisor','1','9','" + sPwd + "','" + CStr(Format(dExpired, "yyyy-mm-dd")) + "','" + amcode + "','" + amname + "','" + sNminput + "')"
        M_OBJCONN.Execute (sStrsql)
        STRSQL = "insert into  tbluser_log (aksi,field_nama,field_nilai,user_by ,sts_form ) values ('New','Userid,Nama,Status','" + sUserid + "," + sName + ",Active " + " ','" + MDIForm1.TxtUsername.Text + "','2')"
        M_OBJCONN.Execute (STRSQL)
    End If
    
    Set M_objrs = New ADODB.Recordset
    sStrsql = "select * from tbluser where  tbluser_userid ='" + amcode + "'"
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
    mngrcode = IIf(IsNull(M_objrs!tbluser_mgrcode), "", M_objrs!tbluser_mgrcode)
    mngrname = IIf(IsNull(M_objrs!tbluser_mgrname), "", M_objrs!tbluser_mgrname)
    End If
    
    Set M_objrs = Nothing
    sStrsql = "update tbluser set  web_password=MD5('PASS!@#$%'),tbluser_mgrcode='" + mngrcode + "',tbluser_mgrname='" + mngrname + "' "
    sStrsql = sStrsql + " where tbluser_userid ='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    Exit Function
out:
    MsgBox err.Description
    saveSpv = False
End Function
Public Function updateSpv(sUserid As String, sName As String, sAktif As String, sPwd As String, amcode As String, amname As String, Optional snameInput As String, Optional field_awal As String, Optional field_nilai As String, Optional JenisQa As String) As Boolean

On Error GoTo out
    updateSpv = True
    
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Update User','" + snameInput + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
    
    
    
    sStrsql = "update tbluser set  tbluser_name='" + sName + "', tbluser_tglupdate =now(),tbluser_kdstatus ='" + sAktif + "',tbluser_amcode ='" + amcode + "',tbluser_amname='" + amname + "' ,tbluser_lastuser ='" + snameInput + "'"
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    Set M_objrs = New ADODB.Recordset
    sStrsql = "select * from tbluser where  tbluser_userid ='" + amcode + "'"
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
        mngrcode = IIf(IsNull(M_objrs!tbluser_mgrcode), "", M_objrs!tbluser_mgrcode)
        mngrname = IIf(IsNull(M_objrs!tbluser_mgrname), "", M_objrs!tbluser_mgrname)
    End If
       
    Set M_objrs = Nothing
    sStrsql = "update tbluser set  tbluser_mgrcode='" + mngrcode + "',tbluser_mgrname='" + mngrname + "' "
    sStrsql = sStrsql + " where tbluser_userid ='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
        
        
     If JenisQa = "QA" Then
            
            
            sStrsql = "update tbluser set  tbluser_kdlevel='8',tbluser_ketlevel='Head Quality Monitoring' ,tbluser_mgrcode='" + mngrcode + "',tbluser_mgrname='" + mngrname + "' "
            sStrsql = sStrsql + " where tbluser_userid ='" + sUserid + "'"
            M_OBJCONN.Execute (sStrsql)
      
          
            STRSQL = "insert into  tbluser_log (aksi,field_nama,field_nilai,user_by ,sts_form ) values ('New','Userid,Nama,Status','" + sUserid + "," + sName + ",Active " + " ','" + MDIForm1.TxtUsername.Text + "','8')"
            M_OBJCONN.Execute (STRSQL)
            
    
    Else
            
            sStrsql = "update tbluser set  tbluser_kdlevel='2',tbluser_ketlevel='Supervisor', tbluser_mgrcode='" + mngrcode + "',tbluser_mgrname='" + mngrname + "' "
            sStrsql = sStrsql + " where tbluser_userid ='" + sUserid + "'"
            M_OBJCONN.Execute (sStrsql)
            
            STRSQL = "insert into  tbluser_log (aksi,field_nama,field_nilai,user_by ,sts_form ) values ('Update','Userid,Nama,Status','" + sUserid + "," + sName + ",Active " + " ','" + MDIForm1.TxtUsername.Text + "','2')"
            M_OBJCONN.Execute (STRSQL)
    
    End If
   
    Exit Function
out:
    MsgBox err.Description
    updateSpv = False
End Function
Public Function deleteSpv(sUserid As String, sNminput As String, field_awal As String, field_nilai As String) As Boolean
On Error GoTo out
    deleteSpv = True
  
    
    
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Delete User','" + sNminput + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
'
        
    STRSQL = "insert into  tbluser_log (userid ,aksi,field_nama,field_nilai,user_by ,sts_form ) values ('" + sUserid + "','Delete','" + field_awal + "','" + field_nilai + "','" + MDIForm1.TxtUsername.Text + "','2')"
    M_OBJCONN.Execute (STRSQL)
      
    sStrsql = "delete  from  tbluser "
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
   
    Exit Function
out:
    MsgBox err.Description
    deleteSpv = False
End Function
Public Function finduseridSpv(sUserid As String) As ADODB.Recordset
On Error GoTo out
    sStrsql = "select * from  tbluser where tbluser_userid='" + sUserid + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set finduseridSpv = M_objrs
    Set M_objrs = Nothing
    Exit Function
out:
    MsgBox err.Description
End Function
Public Function finduseridSpvIcentra(sUserid As String, Optional connIcentra As ADODB.Connection) As ADODB.Recordset
On Error GoTo out
    sStrsql = "select * from  acd_agent where name='" + sUserid + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, connIcentra, adOpenDynamic, adLockOptimistic
    Set finduseridSpvIcentra = M_objrs
    Set M_objrs = Nothing
    Exit Function
out:
    MsgBox err.Description
End Function
Public Function getrecordallSpv() As ADODB.Recordset
On Error GoTo out
    sStrsql = "select * from  tbluser,tbllevel,tblstatus where tbluser.tbluser_kdlevel=tbllevel.tbllevel_kdlevel"
    sStrsql = sStrsql + " and  tbluser.tbluser_kdstatus= tblstatus.tblstatus_kdstatus  and  tbluser_lockkdstatus=tblstatus.tblstatus_kdstatus "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set getrecordallSpv = M_objrs
    Set M_objrs = Nothing
    Exit Function
out:
    MsgBox err.Description
End Function
Public Function update_image_user(sUserid As String, sPath As String) As Boolean
On Error GoTo out
Dim STRSQL As String
    update_image_user = True
    If sPath <> Empty Then
                        STRSQL = "update tbluser set  tbluser_foto=lo_import('"
                        STRSQL = STRSQL + Replace(sPath, "\", "/")
                        STRSQL = STRSQL + "') where tbluser_userid ='"
                        STRSQL = STRSQL + sUserid + "'"
                        M_OBJCONN.Execute STRSQL
    End If
    Exit Function
out:
    update_image_user = False
    MsgBox err.Description, App.Title
End Function
Public Function FindRecordUser(Optional strfield As String, Optional strValue As String, Optional sType As String, Optional sAktif As String, Optional sLevel As String, Optional sUsernik As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
STRSQL = " select * from tbluser "
MWHERE = ""
If strfield <> Empty Then
        If Len(MWHERE) = 0 Then
            MWHERE = " where " + strfield + "='" + strValue + "' "
        Else
            MWHERE = " and " + strfield + "='" + strValue + "' "
        End If
        
End If

If sType <> Empty Then
    If sType = "('2','5')" Then
            If Len(MWHERE) = 0 Then
                MWHERE = " where tbluser_kdlevel in ('2','5') "
            Else
                MWHERE = MWHERE + " and tbluser_kdlevel in ('2','5') "
            End If
    Else
            If Len(MWHERE) = 0 Then
                MWHERE = " where tbluser_kdlevel ='" + sType + "' "
            Else
                MWHERE = MWHERE + " and tbluser_kdlevel ='" + sType + "' "
            End If
    End If
        
End If



If sAktif <> Empty Then
        If Len(MWHERE) = 0 Then
            MWHERE = " where    tbluser_kdstatus   ='1'"
        Else
         MWHERE = MWHERE + " and    tbluser_kdstatus   ='1'"
        End If
End If



If sLevel <> Empty Then
     Select Case sLevel
        Case "Administrator"
            
        
        Case "Supervisor"
            If Len(MWHERE) = 0 Then
                MWHERE = " where      tbluser_groupspvcode ='" + sUsernik + "' "
            Else
                MWHERE = MWHERE + " and    tbluser_groupspvcode ='" + sUsernik + "'"
            End If
        Case "Assisten Manager"
             If Len(MWHERE) = 0 Then
                MWHERE = " where       tbluser_amcode ='" + sUsernik + "'"
            Else
                MWHERE = MWHERE + " and    tbluser_amcode ='" + sUsernik + "'"
            End If
            
     End Select
     
        
End If
Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE + " order by tbluser_kdlevel , tbluser_name asc ", M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindRecordUser = RSTEMP
   Set RSTEMP = Nothing
End Function



Public Function FindRecordUserIcentra(Optional strfield As String, Optional strValue As String, Optional connIcentra As ADODB.Connection) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset

STRSQL = " select * from acd_agent "
MWHERE = ""
If strfield <> Empty Then
        If Len(MWHERE) = 0 Then
            MWHERE = " where " + strfield + "='" + strValue + "' "
        Else
            MWHERE = " and " + strfield + "='" + strValue + "' "
        End If
        
End If

Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE, connIcentra, adOpenDynamic, adLockOptimistic
   Set FindRecordUserIcentra = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function FindRecordUserHST(Optional sAction As String, Optional sUserby As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
STRSQL = " select * from  tbluser_log where sts_form ='1'"
MWHERE = ""

If sAction <> Empty Then
            MWHERE = MWHERE + " and  lower(aksi)  like '%" + StrConv(sAction, vbLowerCase, vbTextCompare) + "%'"
End If


If sUserby <> Empty Then
            MWHERE = MWHERE + " and  lower(user_by)  like '%" + StrConv(sUserby, vbLowerCase, vbTextCompare) + "%'"
End If


Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindRecordUserHST = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function FindRecordUserhstQA(Optional sAction As String, Optional sUserby As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
STRSQL = " select * from  tbluser_log where sts_form ='7'"
MWHERE = ""

If sAction <> Empty Then
            MWHERE = MWHERE + " and  lower(aksi)  like '%" + StrConv(sAction, vbLowerCase, vbTextCompare) + "%'"
End If


If sUserby <> Empty Then
            MWHERE = MWHERE + " and  lower(user_by)  like '%" + StrConv(sUserby, vbLowerCase, vbTextCompare) + "%'"
End If


Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindRecordUserhstQA = RSTEMP
   Set RSTEMP = Nothing
End Function

Public Function saveAdministrator(sUserid As String, sName As String, sAktif As String, sPwd As String, sNminput As String) As Boolean
Dim M_objrs  As New ADODB.Recordset
Dim dExpired As Date
On Error GoTo out
    saveAdministrator = True
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        sStrsql = " select current_date + cast('30 day' as interval) as tglexpired "
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        If Not M_objrs.EOF Then
          dExpired = IIf(IsNull(M_objrs!tglexpired), Null, M_objrs!tglexpired)
        End If
    Set M_objrs = Nothing
        
    sStrsql = "insert into   tbluser (tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired) "
    sStrsql = sStrsql + " values    ('" + sUserid + "','" + sName + "','6','Administrator','1','9','" + sPwd + "','" + CStr(Format(dExpired, "yyyy-mm-dd")) + "')"
    M_OBJCONN.Execute (sStrsql)
    
    STRSQL = " UPDATE TBLUSER SET web_password=MD5('PASS!@#$%') WHERE tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (STRSQL)
    
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'New User','" + sNminput + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
      STRSQL = "insert into  tbluser_log (aksi,field_nama,field_nilai,user_by ,sts_form ) values ('New','Userid,Nama,Status','" + sUserid + "," + sName + ",Active " + " ','" + MDIForm1.TxtUsername.Text + "','6')"
      M_OBJCONN.Execute (STRSQL)


    Exit Function
out:
    MsgBox err.Description, vbInformation + vbOKOnly, App.Title
    saveAdministrator = False
End Function
Public Function updateAdministrator(sUserid As String, sName As String, sAktif As String, sPwd As String, sNminput As String, field_awal As String, field_nilai As String) As Boolean
On Error GoTo out
    updateAdministrator = True
    
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_tglupdate,tbluser_hst_tglentry ,tbluser_hst_tglregister ,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Update User',tbluser_tglupdate,tbluser_tglentry ,tbluser_tglregister,'" + sNminput + "' as nminput "
'    sstrsql = sstrsql + " from  tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
    
    
    sStrsql = "update tbluser set  tbluser_name='" + sName + "', tbluser_tglupdate =now(),tbluser_kdstatus ='" + sAktif + "'"
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    sStrsql = "update tbluser set tbluser_mgrname='" + sName + "'  where tbluser_userid='" + sUserid + "' "
    M_OBJCONN.Execute (sStrsql)
    
          STRSQL = "insert into  tbluser_log (userid,aksi,field_nama,field_nilai,user_by ,sts_form ) values ('" + sUserid + "','Edit','" + field_awal + "','" + field_nilai + "','" + MDIForm1.TxtUsername.Text + "','6')"
      M_OBJCONN.Execute (STRSQL)


    Exit Function
out:
    MsgBox err.Description
    updateAdministrator = False
End Function
Public Function deleteAdministrator(sUserid As String, sNminput As String, field_awal As String, field_nilai As String) As Boolean
On Error GoTo out
    deleteAdministrator = True
  
    
    
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_tglupdate,tbluser_hst_tglentry ,tbluser_hst_tglregister,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Delete User',tbluser_tglupdate,tbluser_tglentry ,tbluser_tglregister,'" + sNminput + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
'
     STRSQL = "insert into  tbluser_log (userid ,aksi,field_nama,field_nilai,user_by ,sts_form ) values ('" + sUserid + "','Delete','" + field_awal + "','" + field_nilai + "','" + MDIForm1.TxtUsername.Text + "','6')"
    M_OBJCONN.Execute (STRSQL)
    
    
    sStrsql = "delete  from  tbluser "
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    Exit Function
out:
    MsgBox err.Description
    deleteAdministrator = False
End Function

Public Function finduseridAdministrator(sUserid As String) As ADODB.Recordset
    On Error GoTo out
    sStrsql = "select * from  tbluser where tbluser_userid='" + sUserid + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set finduseridAdministrator = M_objrs
    Set M_objrs = Nothing
    Exit Function
out:
    MsgBox err.Description
End Function

'--- level agent
Public Function saveAgent(sUserid As String, sName As String, sAktif As String, sPwd As String, SPVCODE As String, SPVNAME As String, sNminput As String) As Boolean
    Dim M_objrs  As New ADODB.Recordset
    Dim dExpired As Date
    On Error GoTo out
    saveAgent = True
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    sStrsql = " select current_date + cast('30 day' as interval) as tglexpired "
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
      dExpired = IIf(IsNull(M_objrs!tglexpired), Null, M_objrs!tglexpired)
    End If
    Set M_objrs = Nothing
        
    sStrsql = "insert into   tbluser (tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired,tbluser_groupspvcode, tbluser_ketgroupspv ,tbluser_lastuser) "
    sStrsql = sStrsql + " values    ('" + sUserid + "','" + sName + "','1','Agent','1','5','" + sPwd + "','" + CStr(Format(dExpired, "yyyy-mm-dd")) + "','" + SPVCODE + "','" + SPVNAME + "','" + sNminput + "')"
    M_OBJCONN.Execute (sStrsql)
    
    Call save_acd_agent("SAVE", sUserid, sName, "Agent", sPwd, SPVCODE)

    Set M_objrs = New ADODB.Recordset
    sStrsql = "select * from tbluser where  tbluser_userid ='" + SPVCODE + "'"
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
      mngrcode = IIf(IsNull(M_objrs!tbluser_mgrcode), "", M_objrs!tbluser_mgrcode)
      mngrname = IIf(IsNull(M_objrs!tbluser_mgrname), "", M_objrs!tbluser_mgrname)
      amcode = IIf(IsNull(M_objrs!tbluser_amcode), "", M_objrs!tbluser_amcode)
      amname = IIf(IsNull(M_objrs!tbluser_amname), "", M_objrs!tbluser_amname)
    End If
   
    Set M_objrs = Nothing
    sStrsql = "update tbluser set  web_password=MD5('PASS!@#$%'),tbluser_amcode='" + amcode + "',tbluser_amname='" + amname + "' , tbluser_mgrcode='" + mngrcode + "',tbluser_mgrname='" + mngrname + "' "
    sStrsql = sStrsql + " where tbluser_userid ='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    STRSQL = "insert into  tbluser_log (aksi,field_nama,field_nilai,user_by ,sts_form ) values ('New','Userid,Nama,Supervisor,Status','" + sUserid + "," + sName + "," + SPVCODE + "-" + SPVNAME + ",Active " + " ','" + MDIForm1.TxtUsername.Text + "','1')"
    M_OBJCONN.Execute (STRSQL)
      
'      sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_amcode,tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'      sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password,tbluser_amcode ,tbluser_tglexpired ,'New User','" + sNminput + "' as nminput "
'      sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'      M_OBJCONN.Execute (sstrsql)
    
    Exit Function
out:
    MsgBox err.Description
    saveAgent = False
End Function

Public Function saveQA(sUserid As String, sName As String, sAktif As String, sPwd As String, SPVCODE As String, SPVNAME As String, sNminput As String) As Boolean
    Dim M_objrs  As New ADODB.Recordset
    Dim dExpired As Date
    On Error GoTo out
    saveQA = True
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    sStrsql = " select current_date + cast('30 day' as interval) as tglexpired "
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
      dExpired = IIf(IsNull(M_objrs!tglexpired), Null, M_objrs!tglexpired)
    End If
    Set M_objrs = Nothing
        
    sStrsql = "insert into   tbluser (tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired,tbluser_groupspvcode, tbluser_ketgroupspv ,tbluser_lastuser) "
    sStrsql = sStrsql + " values    ('" + sUserid + "','" + sName + "','7','Quality Monitoring','1','5','" + sPwd + "','" + CStr(Format(dExpired, "yyyy-mm-dd")) + "','" + SPVCODE + "','" + SPVNAME + "','" + sNminput + "')"
    M_OBJCONN.Execute (sStrsql)
    
    Set M_objrs = New ADODB.Recordset
    sStrsql = "select * from tbluser where  tbluser_userid ='" + SPVCODE + "'"
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
        mngrcode = IIf(IsNull(M_objrs!tbluser_mgrcode), "", M_objrs!tbluser_mgrcode)
        mngrname = IIf(IsNull(M_objrs!tbluser_mgrname), "", M_objrs!tbluser_mgrname)
        amcode = IIf(IsNull(M_objrs!tbluser_amcode), "", M_objrs!tbluser_amcode)
        amname = IIf(IsNull(M_objrs!tbluser_amname), "", M_objrs!tbluser_amname)
    End If
    
    Set M_objrs = Nothing
    sStrsql = "update tbluser set web_password=MD5('PASS!@#$%'), tbluser_amcode='" + amcode + "',tbluser_amname='" + amname + "' , tbluser_mgrcode='" + mngrcode + "',tbluser_mgrname='" + mngrname + "' "
    sStrsql = sStrsql + " where tbluser_userid ='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    STRSQL = "insert into  tbluser_log (aksi,field_nama,field_nilai,user_by ,sts_form ) values ('New','Userid,Nama,Supervisor,Status','" + sUserid + "," + sName + "," + SPVCODE + "-" + SPVNAME + ",Active " + " ','" + MDIForm1.TxtUsername.Text + "','7')"
    M_OBJCONN.Execute (STRSQL)
    
    '      sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_amcode,tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
    '      sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password,tbluser_amcode ,tbluser_tglexpired ,'New User','" + sNminput + "' as nminput "
    '      sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
    '      M_OBJCONN.Execute (sstrsql)
    
    Exit Function
out:
    MsgBox err.Description
    saveQA = False
End Function

Public Function updateAgent(sUserid As String, sName As String, sAktif As String, sPwd As String, SPVCODE As String, SPVNAME As String, snameInput As String, field_awal As String, field_nilai As String) As Boolean
On Error GoTo out
    updateAgent = True
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Update User','" + snameInput + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
    
    
    sStrsql = "update tbluser set  tbluser_name='" + sName + "', tbluser_tglupdate =now(),tbluser_kdstatus ='" + sAktif + "',tbluser_groupspvcode ='" + SPVCODE + "',tbluser_ketgroupspv ='" + SPVNAME + "' ,tbluser_lastuser ='" + snameInput + "'"
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
   Set M_objrs = New ADODB.Recordset
    sStrsql = "select * from tbluser where  tbluser_userid ='" + SPVCODE + "'"
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
        mngrcode = IIf(IsNull(M_objrs!tbluser_mgrcode), "", M_objrs!tbluser_mgrcode)
        mngrname = IIf(IsNull(M_objrs!tbluser_mgrname), "", M_objrs!tbluser_mgrname)
        amcode = IIf(IsNull(M_objrs!tbluser_amcode), "", M_objrs!tbluser_amcode)
        amname = IIf(IsNull(M_objrs!tbluser_amname), "", M_objrs!tbluser_amname)
    End If
    
    Set M_objrs = Nothing
    sStrsql = "update tbluser set  tbluser_amcode='" + amcode + "', tbluser_amname='" + amname + "' , tbluser_mgrcode='" + mngrcode + "',tbluser_mgrname='" + mngrname + "' "
    sStrsql = sStrsql + " where tbluser_userid ='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    STRSQL = "insert into  tbluser_log (userid,aksi,field_nama,field_nilai,user_by ,sts_form ) values ('" + sUserid + "','Edit','" + field_awal + "','" + field_nilai + "','" + MDIForm1.TxtUsername.Text + "','1')"
    M_OBJCONN.Execute (STRSQL)


    Exit Function
out:
    MsgBox err.Description
    updateAgent = False
End Function
Public Function updateQA(sUserid As String, sName As String, sAktif As String, sPwd As String, SPVCODE As String, SPVNAME As String, snameInput As String, field_awal As String, field_nilai As String) As Boolean
On Error GoTo out
    updateQA = True
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Update User','" + snameInput + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
    
    
    sStrsql = "update tbluser set  tbluser_name='" + sName + "', tbluser_tglupdate =now(),tbluser_kdstatus ='" + sAktif + "',tbluser_groupspvcode ='" + SPVCODE + "',tbluser_ketgroupspv ='" + SPVNAME + "' ,tbluser_lastuser ='" + snameInput + "'"
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
   Set M_objrs = New ADODB.Recordset
    sStrsql = "select * from tbluser where  tbluser_userid ='" + SPVCODE + "'"
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
        mngrcode = IIf(IsNull(M_objrs!tbluser_mgrcode), "", M_objrs!tbluser_mgrcode)
        mngrname = IIf(IsNull(M_objrs!tbluser_mgrname), "", M_objrs!tbluser_mgrname)
        amcode = IIf(IsNull(M_objrs!tbluser_amcode), "", M_objrs!tbluser_amcode)
        amname = IIf(IsNull(M_objrs!tbluser_amname), "", M_objrs!tbluser_amname)
    End If
    
    Set M_objrs = Nothing
    sStrsql = "update tbluser set  tbluser_amcode='" + amcode + "', tbluser_amname='" + amname + "' , tbluser_mgrcode='" + mngrcode + "',tbluser_mgrname='" + mngrname + "' "
    sStrsql = sStrsql + " where tbluser_userid ='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    STRSQL = "insert into  tbluser_log (userid,aksi,field_nama,field_nilai,user_by ,sts_form ) values ('" + sUserid + "','Edit','" + field_awal + "','" + field_nilai + "','" + MDIForm1.TxtUsername.Text + "','7')"
    M_OBJCONN.Execute (STRSQL)
    Exit Function
out:
    MsgBox err.Description
    updateQA = False
End Function

Public Function deleteAgent(sUserid As String, NMINPUT As String, field_awal As String, field_nilai As String) As Boolean
On Error GoTo out
    deleteAgent = True
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Delete User','" + NMINPUT + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
    
    sStrsql = "delete  from  tbluser "
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    STRSQL = "insert into  tbluser_log (userid ,aksi,field_nama,field_nilai,user_by ,sts_form ) values ('" + sUserid + "','Delete','" + field_awal + "','" + field_nilai + "','" + MDIForm1.TxtUsername.Text + "','1')"
      M_OBJCONN.Execute (STRSQL)
   
    Exit Function
out:
    MsgBox err.Description
    deleteAgent = False
End Function

Public Function deleteQA(sUserid As String, NMINPUT As String, field_awal As String, field_nilai As String) As Boolean
    On Error GoTo out
    deleteQA = True
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Delete User','" + NMINPUT + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
    
    sStrsql = "delete  from  tbluser "
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    STRSQL = "insert into  tbluser_log (userid ,aksi,field_nama,field_nilai,user_by ,sts_form ) values ('" + sUserid + "','Delete','" + field_awal + "','" + field_nilai + "','" + MDIForm1.TxtUsername.Text + "','7')"
      M_OBJCONN.Execute (STRSQL)
   
    Exit Function
out:
    MsgBox err.Description
    deleteQA = False
End Function


Public Function finduseridAgent(sUserid As String) As ADODB.Recordset
    On Error GoTo out
    sStrsql = "select * from  tbluser where tbluser_userid='" + sUserid + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set finduseridAgent = M_objrs
    Set M_objrs = Nothing
    Exit Function
out:
    MsgBox err.Description
End Function
'--- level Admin
Public Function saveAdmin(sUserid As String, sName As String, sAktif As String, sPwd As String, SPVCODE As String, SPVNAME As String, sNminput As String) As Boolean
    
    Dim M_objrs  As New ADODB.Recordset
    Dim dExpired As Date
    On Error GoTo out
    saveAdmin = True
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        sStrsql = " select current_date + cast('30 day' as interval) as tglexpired "
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        If Not M_objrs.EOF Then
          dExpired = IIf(IsNull(M_objrs!tglexpired), Null, M_objrs!tglexpired)
        End If
    Set M_objrs = Nothing
        
    sStrsql = "insert into   tbluser (tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired,tbluser_groupspvcode, tbluser_ketgroupspv ,tbluser_lastuser) "
    sStrsql = sStrsql + " values    ('" + sUserid + "','" + sName + "','3','Admin','1','9','" + sPwd + "','" + CStr(Format(dExpired, "yyyy-mm-dd")) + "','" + SPVCODE + "','" + SPVNAME + "','" + sNminput + "')"
    M_OBJCONN.Execute (sStrsql)
    
    
    
    Set M_objrs = New ADODB.Recordset
    sStrsql = "select * from tbluser where  tbluser_userid ='" + SPVCODE + "'"
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
        mngrcode = IIf(IsNull(M_objrs!tbluser_mgrcode), "", M_objrs!tbluser_mgrcode)
        mngrname = IIf(IsNull(M_objrs!tbluser_mgrname), "", M_objrs!tbluser_mgrname)
        amcode = IIf(IsNull(M_objrs!tbluser_amcode), "", M_objrs!tbluser_amcode)
        amname = IIf(IsNull(M_objrs!tbluser_amname), "", M_objrs!tbluser_amname)
    End If
    
    Set M_objrs = Nothing
    sStrsql = "update tbluser set  web_password=MD5('PASS!@#$%'), tbluser_amcode='" + amcode + "',tbluser_amname='" + amname + "' , tbluser_mgrcode='" + mngrcode + "',tbluser_mgrname='" + mngrname + "' "
    sStrsql = sStrsql + " where tbluser_userid ='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    STRSQL = "insert into  tbluser_log (aksi,field_nama,field_nilai,user_by ,sts_form ) values ('New','Userid,Nama,Status','" + sUserid + "," + sName + ",Active " + " ','" + MDIForm1.TxtUsername.Text + "','3')"
    M_OBJCONN.Execute (STRSQL)

    
    Exit Function
out:
    MsgBox err.Description
    saveAdmin = False
End Function
Public Function updateAdmin(sUserid As String, sName As String, sAktif As String, sPwd As String, SPVCODE As String, SPVNAME As String, sNminput As String, field_awal As String, field_nilai As String) As Boolean
On Error GoTo out
    updateAdmin = True
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Update User','" + sNminput + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)
    
    
    sStrsql = "update tbluser set  tbluser_name='" + sName + "', tbluser_tglupdate =now(),tbluser_kdstatus ='" + sAktif + "',tbluser_groupspvcode ='" + SPVCODE + "',tbluser_ketgroupspv ='" + SPVNAME + "' ,tbluser_lastuser ='" + sNminput + "'"
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    Set M_objrs = New ADODB.Recordset
    sStrsql = "select * from tbluser where  tbluser_userid ='" + SPVCODE + "'"
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
        mngrcode = IIf(IsNull(M_objrs!tbluser_mgrcode), "", M_objrs!tbluser_mgrcode)
        mngrname = IIf(IsNull(M_objrs!tbluser_mgrname), "", M_objrs!tbluser_mgrname)
        amcode = IIf(IsNull(M_objrs!tbluser_amcode), "", M_objrs!tbluser_amcode)
        amname = IIf(IsNull(M_objrs!tbluser_amname), "", M_objrs!tbluser_amname)
    End If
    
    Set M_objrs = Nothing
    sStrsql = "update tbluser set  tbluser_amcode='" + amcode + "', tbluser_amname='" + amname + "' , tbluser_mgrcode='" + mngrcode + "',tbluser_mgrname='" + mngrname + "',admin_divisi='" + admin_divisi + "' "
    sStrsql = sStrsql + " where tbluser_userid ='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    STRSQL = "insert into  tbluser_log (userid,aksi,field_nama,field_nilai,user_by ,sts_form ) values ('" + sUserid + "','Edit','" + field_awal + "','" + field_nilai + "','" + MDIForm1.TxtUsername.Text + "','3')"
    M_OBJCONN.Execute (STRSQL)
    
    Exit Function
out:
    MsgBox err.Description
    updateAdmin = False
End Function
Public Function deleteAdmin(sUserid As String, sNminput As String, field_awal As String, field_nilai As String) As Boolean
On Error GoTo out
    deleteAdmin = True
'    sstrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
'    sstrsql = sstrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Delete User','" + sNminput + "' as nminput "
'    sstrsql = sstrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
'    M_OBJCONN.Execute (sstrsql)

    STRSQL = "insert into  tbluser_log (userid ,aksi,field_nama,field_nilai,user_by ,sts_form ) values ('" + sUserid + "','Delete','" + field_awal + "','" + field_nilai + "','" + MDIForm1.TxtUsername.Text + "','3')"
    M_OBJCONN.Execute (STRSQL)
      
      
    
    sStrsql = "delete  from  tbluser "
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
   
    Exit Function
out:
    MsgBox err.Description
    deleteAdmin = False
End Function

Public Function finduseridAdmin(sUserid As String) As ADODB.Recordset
On Error GoTo out
    sStrsql = "select * from  tbluser where tbluser_userid='" + sUserid + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set finduseridAdmin = M_objrs
    Set M_objrs = Nothing
    Exit Function
out:
    MsgBox err.Description
End Function

'--- level Checker
Public Function saveChecker(sUserid As String, sName As String, sAktif As String, sPwd As String, SPVCODE As String, SPVNAME As String) As Boolean
    Dim M_objrs  As New ADODB.Recordset
    Dim dExpired As Date
    On Error GoTo out
    saveChecker = True
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    sStrsql = " select current_date + cast('30 day' as interval) as tglexpired "
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
      dExpired = IIf(IsNull(M_objrs!tglexpired), Null, M_objrs!tglexpired)
    End If
    Set M_objrs = Nothing
        
    sStrsql = "insert into   tbluser (tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired,tbluser_groupspvcode, tbluser_ketgroupspv ,tbluser_lastuser) "
    sStrsql = sStrsql + " values    ('" + sUserid + "','" + sName + "','7','Agent','1','9','" + sPwd + "','" + CStr(Format(dExpired, "yyyy-mm-dd")) + "','" + SPVCODE + "','" + SPVNAME + "','anto')"
    M_OBJCONN.Execute (sStrsql)
    
    Set M_objrs = New ADODB.Recordset
    sStrsql = "select * from tbluser where  tbluser_userid ='" + SPVCODE + "'"
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
        mngrcode = IIf(IsNull(M_objrs!tbluser_mgrcode), "", M_objrs!tbluser_mgrcode)
        mngrname = IIf(IsNull(M_objrs!tbluser_mgrname), "", M_objrs!tbluser_mgrname)
        amcode = IIf(IsNull(M_objrs!tbluser_amcode), "", M_objrs!tbluser_amcode)
        amname = IIf(IsNull(M_objrs!tbluser_amname), "", M_objrs!tbluser_amname)
    End If
       
    Set M_objrs = Nothing
    sStrsql = "update tbluser set  tbluser_amcode='" + amcode + "',tbluser_amname='" + amname + "' , tbluser_mgrcode='" + mngrcode + "',tbluser_mgrname='" + mngrname + "' "
    sStrsql = sStrsql + " where tbluser_userid ='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)

    sStrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_amcode,tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
    sStrsql = sStrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password,tbluser_amcode ,tbluser_tglexpired ,'New User','anto' as nminput "
    sStrsql = sStrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    Exit Function
out:
    MsgBox err.Description
    saveChecker = False
End Function

Public Function updateChecker(sUserid As String, sName As String, sAktif As String, sPwd As String, SPVCODE As String, SPVNAME As String) As Boolean
    On Error GoTo out
    updateChecker = True
    sStrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
    sStrsql = sStrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Update User','anto' as nminput "
    sStrsql = sStrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    
    sStrsql = "update tbluser set  tbluser_name='" + sName + "', tbluser_tglupdate =now(),tbluser_kdstatus ='" + sAktif + "',tbluser_groupspvcode ='" + SPVCODE + "',tbluser_ketgroupspv ='" + SPVNAME + "' ,tbluser_lastuser ='anto'"
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    Set M_objrs = New ADODB.Recordset
    sStrsql = "select * from tbluser where  tbluser_userid ='" + SPVCODE + "'"
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
        mngrcode = IIf(IsNull(M_objrs!tbluser_mgrcode), "", M_objrs!tbluser_mgrcode)
        mngrname = IIf(IsNull(M_objrs!tbluser_mgrname), "", M_objrs!tbluser_mgrname)
        amcode = IIf(IsNull(M_objrs!tbluser_amcode), "", M_objrs!tbluser_amcode)
        amname = IIf(IsNull(M_objrs!tbluser_amname), "", M_objrs!tbluser_amname)
    End If
    
    Set M_objrs = Nothing
    sStrsql = "update tbluser set  tbluser_amcode='" + amcode + "', tbluser_amname='" + amname + "' , tbluser_mgrcode='" + mngrcode + "',tbluser_mgrname='" + mngrname + "' "
    sStrsql = sStrsql + " where tbluser_userid ='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    Exit Function
out:
    MsgBox err.Description
     updateChecker = False
End Function

Public Function deleteChecker(sUserid As String) As Boolean
    On Error GoTo out
    deleteChecker = True
    sStrsql = "insert Into  tbluser_hst_hst (tbluser_hst_userid,tbluser_hst_name,tbluser_hst_kdlevel,tbluser_hst_ketlevel ,tbluser_hst_kdstatus,tbluser_hst_lockkdstatus ,tbluser_hst_password, tbluser_hst_tglexpired,tbluser_hst_action,tbluser_hst_nama_user  ) "
    sStrsql = sStrsql + " select tbluser_userid,tbluser_name,tbluser_kdlevel,tbluser_ketlevel ,tbluser_kdstatus,tbluser_lockkdstatus ,tbluser_password, tbluser_tglexpired ,'Update User','anto' as nminput "
    sStrsql = sStrsql + " from tbluser where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
    
    sStrsql = "delete  from  tbluser "
    sStrsql = sStrsql + " where tbluser_userid='" + sUserid + "'"
    M_OBJCONN.Execute (sStrsql)
   
    Exit Function
out:
    MsgBox err.Description
    deleteChecker = False
End Function

Public Function finduseridChecker(sUserid As String) As ADODB.Recordset
    On Error GoTo out
    sStrsql = "select * from  tbluser where tbluser_userid='" + sUserid + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set finduseridChecker = M_objrs
    Set M_objrs = Nothing
    Exit Function
out:
    MsgBox err.Description
End Function

Public Function SetUserMenuBar(sUserid As String, MDIForm1x As MDIForm, Level As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    Dim STRSQL As String
    Dim ctl As Control
    STRSQL = "SELECT tblmenu_role_key_menu as tbluser_role_key_menu  FROM tblmenu_role where  tblmenu_role_ketlevel  ='" + Level + "'"
 
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not RSTEMP.EOF
        ' Update By Izuddin
        For Each ctl In MDIForm1x
            If ctl.Name = Trim(RSTEMP!tbluser_role_key_menu) Then
                ctl.Enabled = True
            End If
        Next
        RSTEMP.MoveNext
    Wend
    Set SetUserMenuBar = RSTEMP
    Set RSTEMP = Nothing
    Exit Function
End Function

Public Function FindRecordGrid(Optional USERID As String, Optional nama As String, Optional SPVCODE As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    STRSQL = " select * from tbluser where tbluser_kdlevel ='1' "
    MWHERE = ""
    If USERID <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_userid) like '%" + StrConv(USERID, vbLowerCase, vbTextCompare) + "%'"
    End If
    
    If nama <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_name) like '%" + StrConv(nama, vbLowerCase, vbTextCompare) + "%'"
    End If

    If SPVCODE <> Empty Then
        MWHERE = MWHERE + " and lower( tbluser_groupspvcode) like '%" + StrConv(SPVCODE, vbLowerCase, vbTextCompare) + "%'"
    End If

    If MDIForm1.txtlevel.Text = "Supervisor" Then
        MWHERE = MWHERE + " and tbluser_groupspvcode in ('" + MDIForm1.TxtUsername.Text + "')"
    End If

    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE + " order by   tbluser_name asc , tbluser_userid ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set FindRecordGrid = RSTEMP
    Set RSTEMP = Nothing
End Function

Public Function FindRecordGridQA(Optional USERID As String, Optional nama As String, Optional SPVCODE As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    STRSQL = " select * from tbluser where tbluser_kdlevel ='7' "
    MWHERE = ""
    If USERID <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_userid) like '%" + StrConv(USERID, vbLowerCase, vbTextCompare) + "%'"
    End If

    If nama <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_name) like '%" + StrConv(nama, vbLowerCase, vbTextCompare) + "%'"
    End If

    If SPVCODE <> Empty Then
        MWHERE = MWHERE + " and lower( tbluser_groupspvcode) like '%" + StrConv(SPVCODE, vbLowerCase, vbTextCompare) + "%'"
    End If

    If MDIForm1.txtlevel.Text = "Supervisor" Then
        MWHERE = MWHERE + " and tbluser_groupspvcode in ('" + MDIForm1.TxtUsername.Text + "')"
    End If

    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE + " order by   tbluser_name asc , tbluser_userid ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set FindRecordGridQA = RSTEMP
    Set RSTEMP = Nothing
End Function

Public Function FindRecordGridspv(Optional USERID As String, Optional nama As String, Optional SPVCODE As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    STRSQL = " select * from tbluser where tbluser_kdlevel  in ('2','8') "
    MWHERE = ""
    If USERID <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_userid) like '%" + StrConv(USERID, vbLowerCase, vbTextCompare) + "%'"
    End If
    
    If nama <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_name) like '%" + StrConv(nama, vbLowerCase, vbTextCompare) + "%'"
    End If

    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE + " order by   tbluser_name asc , tbluser_userid ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set FindRecordGridspv = RSTEMP
    Set RSTEMP = Nothing
End Function

Public Function FindRecordUserHSTspv(Optional sAction As String, Optional sUserby As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    STRSQL = " select * from  tbluser_log where sts_form in ('2','8')"
    MWHERE = ""

    If sAction <> Empty Then
        MWHERE = MWHERE + " and  lower(aksi)  like '%" + StrConv(sAction, vbLowerCase, vbTextCompare) + "%'"
    End If

    If sUserby <> Empty Then
        MWHERE = MWHERE + " and  lower(user_by)  like '%" + StrConv(sUserby, vbLowerCase, vbTextCompare) + "%'"
    End If

    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set FindRecordUserHSTspv = RSTEMP
    Set RSTEMP = Nothing
End Function

Public Function FindRecordUserHSTadmin(Optional sAction As String, Optional sUserby As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    STRSQL = " select * from  tbluser_log where sts_form ='3'"
    MWHERE = ""

    If sAction <> Empty Then
        MWHERE = MWHERE + " and  lower(aksi)  like '%" + StrConv(sAction, vbLowerCase, vbTextCompare) + "%'"
    End If

    If sUserby <> Empty Then
        MWHERE = MWHERE + " and  lower(user_by)  like '%" + StrConv(sUserby, vbLowerCase, vbTextCompare) + "%'"
    End If

    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set FindRecordUserHSTadmin = RSTEMP
    Set RSTEMP = Nothing
End Function

Public Function FindRecordGridadmin(Optional USERID As String, Optional nama As String, Optional SPVCODE As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    STRSQL = " select * from tbluser where tbluser_kdlevel ='3' "
    MWHERE = ""
    If USERID <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_userid) like '%" + StrConv(USERID, vbLowerCase, vbTextCompare) + "%'"
    End If

    If nama <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_name) like '%" + StrConv(nama, vbLowerCase, vbTextCompare) + "%'"
    End If

    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE + " order by   tbluser_name asc , tbluser_userid ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set FindRecordGridadmin = RSTEMP
    Set RSTEMP = Nothing
End Function

Public Function FindRecordGridam(Optional USERID As String, Optional nama As String, Optional SPVCODE As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    STRSQL = " select * from tbluser where tbluser_kdlevel ='4' "
    MWHERE = ""
    If USERID <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_userid) like '%" + StrConv(USERID, vbLowerCase, vbTextCompare) + "%'"
    End If

    If nama <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_name) like '%" + StrConv(nama, vbLowerCase, vbTextCompare) + "%'"
    End If

    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE + " order by   tbluser_name asc , tbluser_userid ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set FindRecordGridam = RSTEMP
    Set RSTEMP = Nothing
End Function

Public Function FindRecordUserHSTam(Optional sAction As String, Optional sUserby As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    STRSQL = " select * from  tbluser_log where sts_form ='4'"
    MWHERE = ""

    If sAction <> Empty Then
        MWHERE = MWHERE + " and  lower(aksi)  like '%" + StrConv(sAction, vbLowerCase, vbTextCompare) + "%'"
    End If

    If sUserby <> Empty Then
        MWHERE = MWHERE + " and  lower(user_by)  like '%" + StrConv(sUserby, vbLowerCase, vbTextCompare) + "%'"
    End If

    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set FindRecordUserHSTam = RSTEMP
    Set RSTEMP = Nothing
End Function

Public Function FindRecordUserHSTManager(Optional sAction As String, Optional sUserby As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    STRSQL = " select * from  tbluser_log where sts_form ='5'"
    MWHERE = ""
    
    If sAction <> Empty Then
        MWHERE = MWHERE + " and  lower(aksi)  like '%" + StrConv(sAction, vbLowerCase, vbTextCompare) + "%'"
    End If

    If sUserby <> Empty Then
        MWHERE = MWHERE + " and  lower(user_by)  like '%" + StrConv(sUserby, vbLowerCase, vbTextCompare) + "%'"
    End If

    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set FindRecordUserHSTManager = RSTEMP
    Set RSTEMP = Nothing
End Function

Public Function FindRecordGridmanager(Optional USERID As String, Optional nama As String, Optional SPVCODE As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    STRSQL = " select * from tbluser where tbluser_kdlevel ='5' "
    MWHERE = ""
    If USERID <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_userid) like '%" + StrConv(USERID, vbLowerCase, vbTextCompare) + "%'"
    End If

    If nama <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_name) like '%" + StrConv(nama, vbLowerCase, vbTextCompare) + "%'"
    End If

    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE + " order by   tbluser_name asc , tbluser_userid ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set FindRecordGridmanager = RSTEMP
    Set RSTEMP = Nothing
End Function

Public Function FindRecordGridadministrator(Optional USERID As String, Optional nama As String, Optional SPVCODE As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    STRSQL = " select * from tbluser where tbluser_kdlevel ='6' "
    MWHERE = ""
    If USERID <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_userid) like '%" + StrConv(USERID, vbLowerCase, vbTextCompare) + "%'"
    End If
    
    If nama <> Empty Then
        MWHERE = MWHERE + " and lower(tbluser_name) like '%" + StrConv(nama, vbLowerCase, vbTextCompare) + "%'"
    End If
    
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE + " order by   tbluser_name asc , tbluser_userid ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set FindRecordGridadministrator = RSTEMP
    Set RSTEMP = Nothing
End Function

Public Function FindRecordUserHSTadministrator(Optional sAction As String, Optional sUserby As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    STRSQL = " select * from  tbluser_log where sts_form ='6'"
    MWHERE = ""
    
    If sAction <> Empty Then
        MWHERE = MWHERE + " and  lower(aksi)  like '%" + StrConv(sAction, vbLowerCase, vbTextCompare) + "%'"
    End If
    
    If sUserby <> Empty Then
        MWHERE = MWHERE + " and  lower(user_by)  like '%" + StrConv(sUserby, vbLowerCase, vbTextCompare) + "%'"
    End If

    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open STRSQL + MWHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set FindRecordUserHSTadministrator = RSTEMP
    Set RSTEMP = Nothing
End Function


