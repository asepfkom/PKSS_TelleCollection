VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_FRMCUST_CC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ADD_OK As Boolean

Public Function QUERY_COMBO_PRODUCT(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset
CMDSQL = "SELECT * FROM PRODUCTTBL"
CMDSQL = CMDSQL + " WHERE UNIT ='Credit Card'"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " AND " + M_WHERE
 End If
CMDSQL = CMDSQL + " ORDER BY CODE"
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_PRODUCT = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_COMBO_CLOSSING(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset
CMDSQL = "SELECT * FROM CLSREASONTBL"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
CMDSQL = CMDSQL + " ORDER BY KDCLS"
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_CLOSSING = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_COMBO_DATASOURCE(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset
CMDSQL = "SELECT * FROM DATASOURCETBL"
'CMDSQL = CMDSQL + " WHERE STATUS ='A'"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
CMDSQL = CMDSQL + " ORDER BY KODEDS"
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_DATASOURCE = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_COMBO_DATASOURCE_ISI(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset
CMDSQL = "SELECT * FROM DATASOURCETBL"
CMDSQL = CMDSQL + " WHERE STATUS ='I'"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " AND " + M_WHERE
 End If
CMDSQL = CMDSQL + " ORDER BY KODEDS"
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_DATASOURCE_ISI = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_CUST(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset
CMDSQL = "SELECT * FROM mgm"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
CMDSQL = CMDSQL + " ORDER BY NAME"
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_CUST = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_NAME_CUST(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset
CMDSQL = "SELECT NAME FROM cc_custtbl"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
CMDSQL = CMDSQL + " ORDER BY NAME"
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_NAME_CUST = M_objrs
Set M_objrs = Nothing
End Function

'Public Function QUERY_HIST_CUST(M_OBJCONN As ADODB.Connection, M_WHERE As String, LAgent As String) As Object
Public Function QUERY_HIST_CUST(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String

On Error Resume Next
' Tambahan jika error invalid sequence byte Pake coalesce aja - Update By Izuddin
Dim M_objrs As ADODB.Recordset
    CMDSQL = "SELECT id,CUSTID,TGL,AGENT,coalesce(HST,'') as HST,KodeDS,KdComplaint,f_cek ,statuscall ,ststelpwith,user_log,stop_time,f_special,unique_id"
    'Cmdsql = "SELECT * "
    CMDSQL = CMDSQL + " From mgm_hst "
    'CMDSQL = "SELECT custid,agent,hst,tgl,kodeds from mgm_hst "
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
' If LAgent <> "Agent" Then
' Else
' CMDSQL = CMDSQL + " AND AGENT= '" + mdiform1.txtusername.text + "'"
' End If
 
CMDSQL = CMDSQL + " ORDER BY tgl DESC "
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.CursorType = adOpenDynamic
M_objrs.LockType = adLockOptimistic
M_objrs.ActiveConnection = M_OBJCONN
M_objrs.Open CMDSQL

Set QUERY_HIST_CUST = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_HIST_PAID(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
On Error Resume Next
Dim M_objrs As ADODB.Recordset
    'CMDSQL = "SELECT CUSTID,PayDate,Payment,Agent,FieldName,Id"
    'CMDSQL = CMDSQL + " From tbllunas"
    CMDSQL = "SELECT b.custid as custid1, a.CUSTID,a.PayDate,a.Payment,a.Agent,a.FieldName,a.Id from tbllunas a"
    CMDSQL = CMDSQL + " inner join mgm b on a.custid=b.custid "
    
    
 If Len(M_WHERE) <> 0 Then
    'CMDSQL = CMDSQL + " WHERE " + M_WHERE + " and dateadd(day,1,a.Paydate) > b.tglsource "
    CMDSQL = CMDSQL + " WHERE " + M_WHERE + ""
 End If
CMDSQL = CMDSQL + " order by a.PayDate desc "
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
Set QUERY_HIST_PAID = M_objrs
Set M_objrs = Nothing

End Function


Public Function ADD_HISTORY(M_OBJCONN As ADODB.Connection, CustId As String, dATETIME1 As String, JAM As String, AGENT As String, _
                            PRODUCTS As String, hst As String, KODEDS As String, KdComplaint As String, RemarkComplaint As String, kode_urut As Integer)
Dim CMDSQL As String
Dim DATETIME2 As String
JAM = Format(Time, "hh:mm")
JAM = Left(JAM, 2) & Right(JAM, 2)
CMDSQL = " INSERT INTO mgm_hst"

If PRODUCTS <> Empty Then
    CMDSQL = CMDSQL + " PRODUCTS,"
End If
CMDSQL = CMDSQL + " HST,kode_urut)"
CMDSQL = CMDSQL + " VALUES"
CMDSQL = CMDSQL + " ('" + CustId + "',"
CMDSQL = CMDSQL + " '" + Format(dATETIME1, "yyyymmdd") & JAM + "',"
CMDSQL = CMDSQL + " '" + AGENT + "',"
CMDSQL = CMDSQL + " '" + KODEDS + "',"
CMDSQL = CMDSQL + " '" + KdComplaint + "',"
CMDSQL = CMDSQL + " '" + RemarkComplaint + "',"
If PRODUCTS <> Empty Then
    CMDSQL = CMDSQL + " '" + UBAH_QUOTE(PRODUCTS) + "',"
End If
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(hst) + "','"
CMDSQL = CMDSQL + CStr(kode_urut) + "')"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox err.Description
End Function


Public Function ADD_Complaint(M_OBJCONN As ADODB.Connection, CustId As String, dATETIME1 As String, JAM As String, AGENT As String, _
                            PRODUCTS As String, hst As String, KODEDS As String, KdComplaint As String, RemarkComplaint As String)
Dim CMDSQL As String
JAM = Format(Time, "hh:mm")
JAM = Left(JAM, 2) & Right(JAM, 2)
CMDSQL = " INSERT INTO CC_CUSTTBL_COMPLAINT"
CMDSQL = CMDSQL + " (CUSTID,"
CMDSQL = CMDSQL + " DATETIME,"
CMDSQL = CMDSQL + " AGENT,"
CMDSQL = CMDSQL + " KODEDS,"
CMDSQL = CMDSQL + " KdComplaint,"
CMDSQL = CMDSQL + " RemarkComplaint,"
If PRODUCTS <> Empty Then
    CMDSQL = CMDSQL + " PRODUCTS,"
End If
CMDSQL = CMDSQL + " HST)"
CMDSQL = CMDSQL + " VALUES"
CMDSQL = CMDSQL + " ('" + CustId + "',"
CMDSQL = CMDSQL + " '" + Format(dATETIME1, "yyyymmdd") & JAM + "',"
CMDSQL = CMDSQL + " '" + KdComplaint + "',"
CMDSQL = CMDSQL + " '" + RemarkComplaint + "',"
CMDSQL = CMDSQL + " '" + AGENT + "',"
CMDSQL = CMDSQL + " '" + KODEDS + "',"
If PRODUCTS <> Empty Then
    CMDSQL = CMDSQL + " '" + UBAH_QUOTE(PRODUCTS) + "',"
End If
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(hst) + "')"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox err.Description
End Function




Public Function QUERY_CEK_ADDCUST(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset
CMDSQL = "SELECT * FROM CC_CUSTTBL"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
CMDSQL = CMDSQL + " ORDER BY NAME"
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_CEK_ADDCUST = M_objrs
Set M_objrs = Nothing
End Function
