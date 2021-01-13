VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_DISTRIBUSISTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ADD_OK As Boolean



Private Function DELETE_TEMPCUSTTBL(M_OBJCONN As ADODB.Connection, USERID As String, TIPE_PRODUK As String)
Dim CMDSQL As String
    Select Case UCase(TIPE_PRODUK)
    Case "KTA"
        CMDSQL = "DELETE tempKTA_CUSTTBL WHERE AGENT = '" + USERID + "'"
    Case "KTA - CROSS SELL"
        CMDSQL = "DELETE tempCS_CUSTTBL WHERE AGENT = '" + USERID + "'"
    Case "CREDIT CARD"
        CMDSQL = "DELETE tempCC_CUSTTBL WHERE AGENT = '" + USERID + "'"
    Case "CC - CROSS SELL"
        CMDSQL = "DELETE tempCCCS_CUSTTBL WHERE AGENT = '" + USERID + "'"
    Case Else
        Exit Function
    End Select
    M_OBJCONN.Execute CMDSQL
End Function


Public Function QUERY_COMBO_DATASOURCE(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM DATASOURCETBL"
'CMDSQL = CMDSQL + " WHERE STATUS ='A'"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 Else
    CMDSQL = CMDSQL + " where left(kodeds,3)<>'mgm'"
    CMDSQL = CMDSQL + " and left(kodeds,3)<>'pre'"
 End If
CMDSQL = CMDSQL + " ORDER BY KODEDS"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_DATASOURCE = M_objrs
Set M_objrs = Nothing
End Function

'Public Function INSERT_DISTRIBUSI(M_RPTCONN As ADODB.Connection, M_OBJCONN As ADODB.Connection, SPVCODE As String, Tanggal As String)
'Dim CMDSQL As String
'Dim USERID As String
'Dim nama As String
'Dim TGLJAM1 As String
'Dim JAM As String
'Dim TGLJAM2 As String
'Dim i As Integer
'Dim M_objrs As ADODB.Recordset
'
'Call DELETE_DISTRIBUSI(M_RPTCONN)
'
'Set M_objrs = QUERY_USER(M_OBJCONN, SPVCODE)
'If M_objrs.RecordCount = 0 Then
'    FRM_SETUSER.ProgressBar1.Max = 100
'Else
'    FRM_SETUSER.ProgressBar1.Max = 100 * (M_objrs.RecordCount + 1)
'
'End If
'    FRM_SETUSER.ProgressBar1.Visible = True
'    FRM_SETUSER.ProgressBar1.Value = 100
'i = 100
'
'TGLJAM2 = Format(Tanggal, "mm/dd/yy")
'JAM = Format(TGLJAM2, "mm/dd/yy") + " " + Format(Now, "hh:mm")
'TGLJAM1 = Format(TGLJAM2, "yyyymmdd") + Format(Now, "hhmm")
'While Not M_objrs.EOF
'    FRM_SETUSER.ProgressBar1.Value = i
'    USERID = IIf(IsNull(M_objrs("USERID")), "", M_objrs("USERID"))
'    nama = IIf(IsNull(M_objrs("AGENT")), "", M_objrs("AGENT"))
'    CMDSQL = "INSERT INTO DISTRIBUSI"
'    CMDSQL = CMDSQL + " (USERID,"
'    CMDSQL = CMDSQL + " TGLJAM,"
'    CMDSQL = CMDSQL + " NAMA)"
'    CMDSQL = CMDSQL + " VALUES"
'    CMDSQL = CMDSQL + " ('" + Trim(USERID) + "',"
'    CMDSQL = CMDSQL + " '" + LTrim(TGLJAM1) + "',"
'    CMDSQL = CMDSQL + " '" + Trim(nama) + "')"
'    M_RPTCONN.Execute CMDSQL
'    M_objrs.MoveNext
'    i = i + 100
'Wend
'    FRM_SETUSER.ProgressBar1.Value = FRM_SETUSER.ProgressBar1.Max
'    FRM_SETUSER.ProgressBar1.Visible = False
'End Function

Private Function DELETE_DISTRIBUSI(M_RPTCONN As ADODB.Connection)
Dim CMDSQL As String
    CMDSQL = "DELETE * FROM DISTRIBUSI"
    M_RPTCONN.Execute CMDSQL
End Function

Public Function QUERY_USER(M_OBJCONN As ADODB.Connection, SPVCODE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM usertbl"
CMDSQL = CMDSQL + " WHERE USERTYPE ='1'"
 If Len(SPVCODE) <> 0 Then
    CMDSQL = CMDSQL + " AND SPVCODE = '" + SPVCODE + "'"
 End If
CMDSQL = CMDSQL + " AND AKTIF = '0'"
CMDSQL = CMDSQL + " ORDER BY USERID"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_USER = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_USER_ACC(M_RPTCONN As ADODB.Connection, SPVCODE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM DISTRIBUSI"
CMDSQL = CMDSQL + " ORDER BY USERID"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_USER_ACC = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_SPV(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM SPVTBL"
CMDSQL = CMDSQL + " WHERE UNIT <> 'Admin'"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " AND " + M_WHERE
 End If
CMDSQL = CMDSQL + " ORDER BY SPVCODE"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SPV = M_objrs
Set M_objrs = Nothing
End Function

Public Function HITUNG_TEMPCUST_KTA(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Currency
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT COUNT(CUSTID) AS JML FROM tempKTA_CUSTTBL"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If M_objrs.RecordCount <> 0 Then
    HITUNG_TEMPCUST_KTA = M_objrs("JML")
End If
Set M_objrs = Nothing
End Function

Public Function HITUNG_TEMPCUST_CC(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Currency
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT COUNT(CUSTID) AS JML FROM tempCC_CUSTTBL"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If M_objrs.RecordCount <> 0 Then
    HITUNG_TEMPCUST_CC = M_objrs("JML")
End If
Set M_objrs = Nothing
End Function

Public Function HITUNG_TEMPCUST_CS(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Currency
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT COUNT(CUSTID) AS JML FROM tempCS_CUSTTBL"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If M_objrs.RecordCount <> 0 Then
    HITUNG_TEMPCUST_CS = M_objrs("JML")
End If
Set M_objrs = Nothing
End Function

Public Function HITUNG_TEMPCUST_CCCS(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Currency
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT COUNT(CUSTID) AS JML FROM tempCCCS_CUSTTBL"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If M_objrs.RecordCount <> 0 Then
    HITUNG_TEMPCUST_CCCS = M_objrs("JML")
End If
Set M_objrs = Nothing
End Function
