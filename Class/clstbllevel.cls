VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clstbllevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class untuk semua modul user
' 1) Modul Manager
Public Function saveLevel(sKdlevel As String, sKeterangan As String, sAktif) As Boolean
    Dim M_objrs  As New ADODB.Recordset
    Dim dExpired As Date
    On Error GoTo out
    saveLevel = True
    
        
'    sstrsql = "insert into   tbllevel ( tbllevel_kdlevel,tbllevel_keterangan, tbllevel_kdstatus) "
    sStrsql = "INSERT into tbllevel ( tbllevel_kdlevel,level_name, tbllevel_kdstatus) "
    sStrsql = sStrsql + " values    ('" + sKdlevel + "','" + sKeterangan + "','1')"
    M_OBJCONN.Execute (sStrsql)
    
    sStrsql = "insert Into  tbllevel_hst (tbllevel_hst_kdlevel,tbllevel_hst_keterangan, tbllevel_hst_kdstatus,tbllevel_hst_action,tbllevel_hst_nama_user  ) "
    sStrsql = sStrsql + " select tbllevel_kdlevel,tbllevel_keterangan, tbllevel_kdstatus,'New Level','anto' as nminput "
    sStrsql = sStrsql + " from  tbllevel where tbllevel_kdlevel='" + sKdlevel + "'"
    M_OBJCONN.Execute (sStrsql)
    Exit Function
out:
    MsgBox err.Description, vbInformation + vbOKOnly, App.Title
    saveLevel = False
End Function
Public Function updateLevel(sKdlevel As String, sKeterangan As String, sAktif) As Boolean
On Error GoTo out
    updateLevel = True
    
    
    
    sStrsql = "insert Into  tbllevel_hst (tbllevel_hst_kdlevel,tbllevel_hst_keterangan, tbllevel_hst_kdstatus,tbllevel_hst_action,tbllevel_hst_nama_user  ) "
    sStrsql = sStrsql + " select tbllevel_kdlevel,tbllevel_keterangan, tbllevel_kdstatus,'Update Level','anto' as nminput "
    sStrsql = sStrsql + " from  tbllevel where tbllevel_kdlevel='" + sKdlevel + "'"
    M_OBJCONN.Execute (sStrsql)
    
    
    
'    sstrsql = "update tbllevel set  tbllevel_keterangan='" + sKeterangan + "', tbllevel_tglupdate  =now(),tbllevel_kdstatus ='" + sAktif + "'"
    sStrsql = "update tbllevel set level_name='" + sKeterangan + "', tbllevel_tglupdate  =now(),tbllevel_kdstatus ='" + sAktif + "'"
    sStrsql = sStrsql + " where tbllevel_kdlevel='" + sKdlevel + "'"
    M_OBJCONN.Execute (sStrsql)
    
    Exit Function
out:
    MsgBox err.Description
    updateLevel = False
End Function
Public Function deleteLevel(sKdlevel As String) As Boolean
On Error GoTo out
    deleteLevel = True
  
    
    sStrsql = "insert Into  tbllevel_hst (tbllevel_hst_kdlevel,tbllevel_hst_keterangan, tbllevel_hst_kdstatus,tbluser_hst_action,tbluser_hst_nama_user  ) "
    sStrsql = sStrsql + " select tbllevel_kdlevel,tbllevel_keterangan, tbllevel_kdstatus,'Update Level','anto' as nminput "
    sStrsql = sStrsql + " from  tbllevel where tbllevel_kdlevel='" + sKdlevel + "'"
    M_OBJCONN.Execute (sStrsql)
    
    
    
    
    
    sStrsql = "delete  from  tbllevel "
    sStrsql = sStrsql + " where tbllevel_kdlevel='" + sKdlevel + "'"
    M_OBJCONN.Execute (sStrsql)
   
    Exit Function
out:
    MsgBox err.Description
    deleteLevel = False
End Function
Public Function finduseridLevel(sKdlevel As String) As ADODB.Recordset
On Error GoTo out
    sStrsql = "select * from  tbllevel where tbllevel_kdlevel='" + sKdlevel + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Set finduseridLevel = M_objrs
    Set M_objrs = Nothing
    Exit Function
out:
    MsgBox err.Description
End Function
Public Function FindRecordLevel(Optional strfield As String, Optional strValue As String, Optional sType As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select * from  tbllevel "
If strfield <> Empty Then
            mwhere = " where " + strfield + "='" + strValue + "' "
End If

Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindRecordLevel = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function FindRecordLevelHST(Optional strfield As String, Optional strValue As String, Optional sType As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select * from  tbllevel_hst "
If strfield <> Empty Then
            mwhere = " where " + strfield + "='" + strValue + "'"
End If

Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindRecordLevelHST = RSTEMP
   Set RSTEMP = Nothing
End Function



