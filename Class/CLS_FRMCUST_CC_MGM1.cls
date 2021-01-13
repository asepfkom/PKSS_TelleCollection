VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_FRMCUST_CC_MGM"
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
CMDSQL = "SELECT NAME FROM mgm"
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

Public Function QUERY_HIST_CUST(M_OBJCONN As ADODB.Connection, M_WHERE As String) As ADODB.Recordset
Dim CMDSQL As String
On Error GoTo eRR_QUERY_HIST_CUST
Dim M_objrs As ADODB.Recordset
    CMDSQL = "SELECT CUSTID,DATETIME,AGENT,HST, KDCOMPLAINT, REMARKCOMPLAINT "
    CMDSQL = CMDSQL + " From mgm_hst"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
CMDSQL = CMDSQL + " ORDER BY DATETIME ASC "
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_HIST_CUST = M_objrs
Set M_objrs = Nothing
EXIT_eRR_QUERY_HIST_CUST:
    Exit Function
eRR_QUERY_HIST_CUST:
    MsgBox err.Description
    Resume EXIT_eRR_QUERY_HIST_CUST
End Function
Public Function ADD_HISTORY(CustId As String, dATETIME1 As String, JAM As String, AGENT As String, _
                            PRODUCTS As String, hst As String, KODEDS As String, KdComplaint As String, RemarkComplaint As String, F_CEK As String, statuscall As String, _
                            lastsatatus As String, f_cek_new As String, stsaccount As String, _
                            ststelpwith As String, userlog As String, kode_unik As String, kode_urut As Integer, Optional wktstartcall As String, Optional wktstopcall As String, Optional nocard As String, Optional s_unique_id As String, Optional s_kat_aktif_telp As String)
Dim CMDSQL As String
Dim DATETIME2 As String


If wktstopcall = "" Then
    wktstopcall = Format(Date, "yyyy-mm-dd 00:00:00")
End If
's_unique_id = FrmCC_Colection.Text7.Text
JAM = Format(Time, "hh:mm")
JAM = Left(JAM, 2) & Right(JAM, 2)
CMDSQL = " INSERT INTO mgm_hst"
CMDSQL = CMDSQL + " (CUSTID,"
CMDSQL = CMDSQL + " DATETIME,"
CMDSQL = CMDSQL + " AGENT,"
CMDSQL = CMDSQL + " KODEDS,"
CMDSQL = CMDSQL + " KdComplaint,"
CMDSQL = CMDSQL + " F_CEK,"
CMDSQL = CMDSQL + " PHONENO,"
CMDSQL = CMDSQL + " RemarkComplaint,"
If PRODUCTS <> Empty Then
    CMDSQL = CMDSQL + " PRODUCTS,"
End If
CMDSQL = CMDSQL + " HST,STATUSCALL,lastcall,f_cek_new,kethslkerjadesc,ststelpwith,user_log,kode_unik,kode_urut,stop_time,unique_id,callwith )"
CMDSQL = CMDSQL + " VALUES"
CMDSQL = CMDSQL + " ('" + CustId + "',"
CMDSQL = CMDSQL + " '" + Format(dATETIME1, "yyyymmdd") & JAM + "',"
CMDSQL = CMDSQL + " '" + AGENT + "',"
CMDSQL = CMDSQL + " '" + KODEDS + "',"
CMDSQL = CMDSQL + " '" + KdComplaint + "',"
CMDSQL = CMDSQL + " '" + F_CEK + "',"
CMDSQL = CMDSQL + " '" + FrmCC_Colection.txtPhone.text + "',"
CMDSQL = CMDSQL + " '" + RemarkComplaint + "',"
If PRODUCTS <> Empty Then
    CMDSQL = CMDSQL + " '" + UBAH_QUOTE(PRODUCTS) + "',"
End If
CMDSQL = CMDSQL + "'" + UBAH_QUOTE(hst) + "',"
CMDSQL = CMDSQL + "'" + statuscall + "',"
CMDSQL = CMDSQL + "'" + lastsatatus + "',"
CMDSQL = CMDSQL + "'" + f_cek_new + "',"
CMDSQL = CMDSQL + "'" + stsaccount + "',"
CMDSQL = CMDSQL + "'" + ststelpwith + "',"
CMDSQL = CMDSQL + "'" + userlog + "'"
CMDSQL = CMDSQL + ",'" + CStr(kode_unik) + "','"
CMDSQL = CMDSQL + CStr(kode_urut) + "','"
CMDSQL = CMDSQL + wktstopcall + "',"
'CMDSQL = CMDSQL + nocard + "',"
'CMDSQL = CMDSQL + custno + "',"
CMDSQL = CMDSQL + "'" + MDIForm1.txt_unique_id.text + "','" & s_kat_aktif_telp & "')"
    On Error GoTo add_error
     M_OBJCONN.Execute CMDSQL
     ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox err.Description
End Function
Public Function ADD_tbllunas(M_OBJCONN As ADODB.Connection, CustId As String, dATETIME1 As Date, Payment As String, _
                             AGENT As String, fieldName As String, datafrom As String)
Dim CMDSQL As String
Dim DATETIME2 As String
'JAM = Format(Time, "hh:mm")
'JAM = Left(JAM, 2) & Right(JAM, 2)
CMDSQL = " INSERT INTO tbllunas"
CMDSQL = CMDSQL + " (CUSTID,"
CMDSQL = CMDSQL + " PayDate,"
CMDSQL = CMDSQL + " Payment,"
CMDSQL = CMDSQL + " AGENT,"
CMDSQL = CMDSQL + " FIeldName,"
CMDSQL = CMDSQL + "Datafrom)"

CMDSQL = CMDSQL + " VALUES"
CMDSQL = CMDSQL + " ('" + CustId + "',"
CMDSQL = CMDSQL + " '" + Format(dATETIME1, "yyyy/mm/dd") + "',"
CMDSQL = CMDSQL + " " + Payment + ","
CMDSQL = CMDSQL + " '" + AGENT + "',"
CMDSQL = CMDSQL + " '" + fieldName + "',"
CMDSQL = CMDSQL + "'Data')"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox err.Description
End Function
Public Function QUERY_HIST_PAID(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
On Error Resume Next
Dim M_objrs As ADODB.Recordset
    CMDSQL = "SELECT CUSTID,PayDate,Payment,Agent,FieldName "
    CMDSQL = CMDSQL + " From tbllunas"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
CMDSQL = CMDSQL + " ORDER BY PayDate ASC "
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_HIST_PAID = M_objrs
Set M_objrs = Nothing
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

Public Function ADD_HISTORY_OTO(M_OBJCONN As ADODB.Connection, CustId As String, dATETIME1 As String, JAM As String, AGENT As String, _
                            PRODUCTS As String, hst As String, KODEDS As String, KdComplaint As String, RemarkComplaint As String, F_CEK As String, statuscall As String, TELP As String)
Dim CMDSQL As String
Dim DATETIME2 As String
JAM = Format(Time, "hh:mm")
JAM = Left(JAM, 2) & Right(JAM, 2)
CMDSQL = " INSERT INTO mgm_hst"
CMDSQL = CMDSQL + " (CUSTID,"
CMDSQL = CMDSQL + " DATETIME,"
CMDSQL = CMDSQL + " AGENT,"
CMDSQL = CMDSQL + " KODEDS,"
CMDSQL = CMDSQL + " KdComplaint,"
CMDSQL = CMDSQL + " F_CEK,"
CMDSQL = CMDSQL + " PHONENO,"
CMDSQL = CMDSQL + " RemarkComplaint,"
If PRODUCTS <> Empty Then
    CMDSQL = CMDSQL + " PRODUCTS,"
End If
CMDSQL = CMDSQL + " HST,STATUSCALL)"
CMDSQL = CMDSQL + " VALUES"
CMDSQL = CMDSQL + " ('" + CustId + "',"
CMDSQL = CMDSQL + " '" + Format(dATETIME1, "yyyymmdd") & JAM + "',"
CMDSQL = CMDSQL + " '" + AGENT + "',"
CMDSQL = CMDSQL + " '" + KODEDS + "',"
CMDSQL = CMDSQL + " '" + KdComplaint + "',"
CMDSQL = CMDSQL + " '" + F_CEK + "',"
CMDSQL = CMDSQL + " '" + TELP + "',"
CMDSQL = CMDSQL + " '" + RemarkComplaint + "',"
If PRODUCTS <> Empty Then
    CMDSQL = CMDSQL + " '" + UBAH_QUOTE(PRODUCTS) + "',"
End If
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(hst) + "','" + statuscall + "')"
    On Error GoTo add_error
     M_OBJCONN.Execute CMDSQL
     ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox err.Description
End Function

