VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function FindMenu() As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select * from tblmenu "
Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindMenu = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function FindMenuSource(Optional sLevel As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select * from  tblmenu where tblmenu_key_menu not in ( "
strsql = strsql + " select tblmenu_role_key_menu from tblmenu_role where tblmenu_role_kdlevel='" + sLevel + "' )"
Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindMenuSource = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function FindMenuDestination(Optional sLevel As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select * from  tblmenu_role where tblmenu_role_kdlevel='" + sLevel + "'"
Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindMenuDestination = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function FindMenuRoleHst(Optional sLevel As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select * from  tblmenuhst_role "
Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindMenuRoleHst = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function FindMenulvlDist() As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = "SELECT a.*,b.level_name FROM ( select tblmenu_role_kdlevel,  tblmenu_role_ketlevel,date(tblmenu_role_tglentry)"
strsql = strsql + "   from  tblmenu_role  group by tblmenu_role_kdlevel,  tblmenu_role_ketlevel,date(tblmenu_role_tglentry)) a,tbllevel b WHERE a.tblmenu_role_kdlevel=b.tbllevel_kdlevel "

Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindMenulvlDist = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function FindMenuLevelsource(Optional sLevel As String, Optional sUserid As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select * from tblmenu_role where tblmenu_role_ketlevel ='" + sLevel + "' and tblmenu_role_key_menu not in (select   tbluser_role_key_menu from tbluser_role where  tbluser_role_userid ='" + sUserid + "')"
Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindMenuLevelsource = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function FindMenuuserDestination(Optional sUserid As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select * from tbluser_role where  tbluser_role_userid='" + sUserid + "'"
Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindMenuuserDestination = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function FindMenuListuserDist() As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select tbluser_role_userid , tbluser_role_username "
strsql = strsql + "   from   tbluser_role group by tbluser_role_userid , tbluser_role_username "

Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindMenuListuserDist = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function FindMenuuserRoleHst(Optional sLevel As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select * from  tbluserhst_role"
Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindMenuuserRoleHst = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function SetUserMenuBar(sUserid As String, MDIForm1x As MDIForm, Level As String) As ADODB.Recordset
    Dim RSTEMP As New ADODB.Recordset
    Dim strsql As String
    Dim sMenu As String
    Dim Ctl As Control
    strsql = "SELECT tblmenu_role_key_menu as tbluser_role_key_menu  FROM tblmenu_role where  tblmenu_role_ketlevel  ='" + Level + "' order by 1"
 
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
    '------------------------------------------------------------------------------------------------------
    While Not RSTEMP.EOF
       Select Case Trim(RSTEMP!tbluser_role_key_menu)
          
          Case "mnbar(0)"
                MDIForm1.mnbar(0).Enabled = True
          Case "mnfile(3)"
                MDIForm1.MnFile(3).Enabled = True
          Case "mnfile(5)"
               MDIForm1.MnFile(5).Enabled = True
          Case "mnfile(7)"
                MDIForm1.MnFile(7).Enabled = True
          Case "mnbar(1)"
                MDIForm1.mnbar(1).Enabled = True
          Case "mnoffice"
                MDIForm1.mnoffice.Enabled = True
          Case "mnagent"
                MDIForm1.mnagent.Enabled = True
          Case "mntl"
                MDIForm1.mntl.Enabled = True
          Case "mnmgr"
                MDIForm1.mnmgr.Enabled = True
          Case "mnrole"
                MDIForm1.mnrole.Enabled = True
          Case "mnNact"
                MDIForm1.mnNact.Enabled = True
          Case "mnblack"
                MDIForm1.mnblack.Enabled = True
          Case "mnbar(12)"
                MDIForm1.mnbar(12).Enabled = True
          Case "mnrdistribut"
                MDIForm1.mnrdistribut.Enabled = True
          Case "mn_monhly_bp"
                MDIForm1.mn_monhly_bp.Enabled = True
          Case "mnmonthcpa"
                MDIForm1.mnmonthcpa.Enabled = True
          Case "mnptppayment"
                MDIForm1.mnptppayment.Enabled = True
          Case "nmconfidenceanalisysagent"
                MDIForm1.nmconfidenceanalisysagent.Enabled = True
          Case "mn_confidence_list"
                MDIForm1.mn_confidence_list.Enabled = True
          Case "mn_performance"
                MDIForm1.mn_performance.Enabled = True
          Case "mntools"
                MDIForm1.mntools.Enabled = True
          Case "mndistribut"
                MDIForm1.mndistribut.Enabled = True
          Case "mnrecycle"
                MDIForm1.mnrecycle.Enabled = True
          Case "nmupload"
                MDIForm1.nmupload.Enabled = True
          Case "nmuploadcustomer"
                MDIForm1.nmuploadcustomer.Enabled = True
          Case "nmuploadpayment"
                MDIForm1.nmuploadpayment.Enabled = True
          Case "list_phone_review"
                MDIForm1.list_phone_review.Enabled = True
          Case "mnucallmonitor"
                MDIForm1.mnuCallmonitor.Enabled = True
          Case "mnrresult"
                MDIForm1.mnrresult.Enabled = True
             
             
        End Select
        RSTEMP.MoveNext
    
    Wend
   Set SetUserMenuBar = RSTEMP
   Set RSTEMP = Nothing
    Exit Function
End Function
Public Function IsControlArray(Ctl As Object) As Boolean
    Dim Ctrl As Control
    On Error Resume Next
    For Each Ctrl In Ctl
        IsControlArray = (err.number = 0)
        Exit Function
    Next Ctrl
End Function
Public Function InControlArray(Ctl As Object) As Boolean
    InControlArray = Not Ctl.Parent.Controls(Ctl.Name) Is Ctl
End Function

