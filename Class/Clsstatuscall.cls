VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clsstatuscall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class untuk semua modul user
' 1) Modul Manager
Public Function saveStatuscall(sKdstatuscall As String, sKeterangan As String, sAktif, suSERINUT As String, sNminput As String, sGroupcall As String, sGroupcall2 As String) As Boolean
Dim M_objrs  As New ADODB.Recordset
Dim dExpired As Date
On Error GoTo out
    saveStatuscall = True
    
        
    sStrsql = "insert into  tblstatuscall (tblstatuscall_kdstscall,tblstatuscall_keterangan,tblstatuscall_kdstatus,tblstatuscall_kdusrwrite ,tblstatuscall_ketuserwrite,grp_call,grp_call2) "
    sStrsql = sStrsql + " values    ('" + sKdstatuscall + "','" + sKeterangan + "','1','" + suSERINUT + "','" + sNminput + "','" + sGroupcall + "','" + sGroupcall2 + "')"
    M_OBJCONN.Execute (sStrsql)
    
    
    sStrsql = "insert Into tblstatuscall_hst (tblstatuscall_hst_kdstscall,tblstatuscall_hst_keterangan, tblstatuscall_hst_kdstatus,tblstatuscall_hst_action,tblstatuscall_hst_nama_user,grp_call,grp_call2) "
    sStrsql = sStrsql + " select tblstatuscall_kdstscall,tblstatuscall_keterangan, tblstatuscall_kdstatus, 'New Status Call','" + suSERINUT + "',grp_call,grp_call2 "
    sStrsql = sStrsql + " from  tblstatuscall  where tblstatuscall_kdstscall='" + sKdstatuscall + "'"
    M_OBJCONN.Execute (sStrsql)
    
    Exit Function
out:
    MsgBox err.Description, vbInformation + vbOKOnly, App.Title
    saveStatuscall = False
End Function

Public Function updateStatusCall(sKdstatuscall As String, sKeterangan As String, sAktif As String, sNminput As String, sGroupcall As String, sGroupcall2 As String) As Boolean
On Error GoTo out
    updateStatusCall = True
    
    
    
   
    sStrsql = "insert Into tblstatuscall_hst (tblstatuscall_hst_kdstscall,tblstatuscall_hst_keterangan, tblstatuscall_hst_kdstatus,tblstatuscall_hst_action,tblstatuscall_hst_nama_user,grp_call,grp_call2) "
    sStrsql = sStrsql + " select tblstatuscall_kdstscall,tblstatuscall_keterangan, tblstatuscall_kdstatus, 'Update Status Call','" + sNminput + "',grp_call,grp_call2 "
    sStrsql = sStrsql + " from  tblstatuscall  where tblstatuscall_kdstscall='" + sKdstatuscall + "'"
    M_OBJCONN.Execute (sStrsql)
    
    
    sStrsql = "update tblstatuscall set  tblstatuscall_keterangan='" + sKeterangan + "',tblstatuscall_tglupdate    =now(),tblstatuscall_kdstatus  ='" + sAktif + "',grp_call='" + sGroupcall + "',grp_call2='" + sGroupcall2 + "'"
    sStrsql = sStrsql + " where  tblstatuscall_kdstscall='" + sKdstatuscall + "'"
    M_OBJCONN.Execute (sStrsql)
    
    Exit Function
out:
    MsgBox err.Description
    updateStatusCall = False
End Function
Public Function deleteStatuscall(sKdstatuscall As String, sNminput As String) As Boolean
On Error GoTo out
    deleteStatuscall = True
  
    
    sStrsql = "insert Into tblstatuscall_hst (tblstatuscall_hst_kdstscall,tblstatuscall_hst_keterangan, tblstatuscall_hst_kdstatus,tblstatuscall_hst_action,tblstatuscall_hst_nama_user,grp_call,grp_call2) "
    sStrsql = sStrsql + " select tblstatuscall_kdstscall,tblstatuscall_keterangan, tblstatuscall_kdstatus, 'Delete Status Call','" + sNminput + "',grp_call,grp_call2 "
    sStrsql = sStrsql + " from  tblstatuscall  where tblstatuscall_kdstscall='" + sKdstatuscall + "'"
    M_OBJCONN.Execute (sStrsql)
    
    
    sStrsql = "delete  from tblstatuscall "
    sStrsql = sStrsql + " where  tblstatuscall_kdstscall ='" + sKdstatuscall + "'"
    M_OBJCONN.Execute (sStrsql)
   
    Exit Function
out:
    MsgBox err.Description
    deleteStatuscall = False
End Function
Public Function findKdstatus(sKdstatus As String) As ADODB.Recordset
On Error GoTo out
    sStrsql = "select * from  tblstatuscall where  tblstatuscall_kdstscall ='" + sKdstatus + "' order by tblstatuscall_id"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set findKdstatus = M_objrs
    Set M_objrs = Nothing
    Exit Function
out:
    MsgBox err.Description
End Function
Public Function FindRecordStatuscall(Optional strfield As String, Optional strValue As String, Optional sType As String, Optional aktif As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select * from  tblstatuscall  "
mwhere = ""
If strfield <> Empty Then
            mwhere = " where " + strfield + " like '%" + strValue + "%' "
End If

If aktif <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = " where   tblstatuscall_kdstatus  ='1'"
        Else
         mwhere = mwhere + " and   tblstatuscall_kdstatus  ='1'"
        End If
        
End If

Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere + "  order by tblstatuscall_id asc ", M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindRecordStatuscall = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function FindRecordStatuscallHST(Optional strfield As String, Optional strValue As String, Optional sType As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select * from  tblstatuscall_hst "
If strfield <> Empty Then
            mwhere = " where " + strfield + "='" + strValue + "'"
End If

Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindRecordStatuscallHST = RSTEMP
   Set RSTEMP = Nothing
End Function





