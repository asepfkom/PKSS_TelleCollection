VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_FRMCUST_KTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ADD_OK As Boolean

Public Function QUERY_COMBO_PRODUCT(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim m_objrs As ADODB.Recordset
   
CMDSQL = "SELECT * FROM PRODUCTTBL"
CMDSQL = CMDSQL + " WHERE UNIT ='KTA'"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " AND " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY CODE"
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_PRODUCT = m_objrs
Set m_objrs = Nothing
End Function

Public Function QUERY_COMBO_CLOSSING(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim m_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM CLSREASONTBL"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY KDCLS"
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_CLOSSING = m_objrs
Set m_objrs = Nothing
End Function


Public Function QUERY_COMBO_DATASOURCE(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim m_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM DATASOURCETBL"
'CMDSQL = CMDSQL + " WHERE STATUS ='A'"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY KODEDS"
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_DATASOURCE = m_objrs
Set m_objrs = Nothing
End Function

Public Function QUERY_COMBO_DATASOURCE_ISI(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim m_objrs As ADODB.Recordset
CMDSQL = "SELECT * FROM DATASOURCETBL"
CMDSQL = CMDSQL + " WHERE STATUS ='I'"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " AND " + M_WHERE
 End If
CMDSQL = CMDSQL + " ORDER BY KODEDS"
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_DATASOURCE_ISI = m_objrs
Set m_objrs = Nothing
End Function

Public Function QUERY_CUST(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim m_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM KTA_CUSTTBL"

 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY NAME"
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_CUST = m_objrs
Set m_objrs = Nothing
End Function

Public Function QUERY_NAME_CUST(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim m_objrs As ADODB.Recordset

CMDSQL = "SELECT NAME FROM KTA_CUSTTBL"
    
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY NAME"
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_NAME_CUST = m_objrs
Set m_objrs = Nothing
End Function


Public Function QUERY_HIST_CUST(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim m_objrs As ADODB.Recordset
On Error Resume Next
    CMDSQL = "SELECT CUSTID,DATETIME,AGENT,HST"
    CMDSQL = CMDSQL + " From KTA_CUSTHSTTBL"
    
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY DATETIME ASC "
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_HIST_CUST = m_objrs
Set m_objrs = Nothing
End Function

Public Function ADD_CUSTOMER_BARU(M_OBJCONN As ADODB.Connection, KETHSLKERJA As String, STATUS_FIELD_LAMA As String, DOK1 As String)

Dim m_objrs As New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
'm_objrs.Open "SELECT * FROM KTA_CUSTTBL WHERE AGENT ='N/A'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
m_objrs.Open "SELECT * FROM KTA_CUSTTBL WHERE CUSTID ='" + FRMCUST_CC.Text1(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_objrs.RecordCount <> 0 Then
    FRMCUST_CC.Text1(1).Text = "KTA-I-" & CUSTNOMOR(M_OBJCONN, UCase(FRMCUST_CC))
End If
With FRMCUST_CC
    m_objrs.AddNew
    m_objrs!CUSTID = .Text1(1).Text
    m_objrs!Name = .Text1(0).Text
    m_objrs!Title = .Combo1(2).Text
    m_objrs!PRIOR = .Combo5.Text
    If .TDBDate1(0).ValueIsNull Then
        m_objrs!BIRTHD = Null
    Else
        m_objrs!BIRTHD = Format(.TDBDate1(0).Value, "mm/dd/yyyy")
    End If
    If .Option7(0).Value Then
        m_objrs!QUALIFIED = 1
    Else
        m_objrs!QUALIFIED = 0
    End If
    m_objrs!ADDRNOW = .RichTextBox1(0).Text
    m_objrs!ZIPNOW = .Text1(6).Text
    m_objrs!CITYNOW = .Text1(7).Text
    If Len(.TDBMask1(0).Value) > 5 Then
        m_objrs!HOMENO = .TDBMask1(0).Value
    End If
    If Len(.TDBMask1(1).Value) > 5 Then
        m_objrs!HOMENO2 = .TDBMask1(1).Value
    End If
    If Len(.TDBMask1(2).Value) > 5 Then
        m_objrs!MOBILENO = .TDBMask1(2).Value
    End If
    If Len(.TDBMask1(3).Value) > 5 Then
        m_objrs!MOBILENO2 = .TDBMask1(3).Value
    End If
    If .Option4(0).Value Then
        m_objrs!CAT = "0"
        m_objrs!NAMAPT = .Text1(20).Text
        m_objrs!ADDRPT = .RichTextBox1(2).Text
    Else
        m_objrs!CAT = "1"
        m_objrs!JENISUSAHA = .Text1(34).Text
        m_objrs!NAMAPT = .Text1(18).Text
        m_objrs!ADDRPT = .RichTextBox1(4).Text
    End If
    If Len(.TDBMask1(6).Value) > 5 Then
        m_objrs!FAXNO = .TDBMask1(6).Value
    End If
    If Len(.TDBMask1(7).Value) > 5 Then
        m_objrs!FAXNO2 = .TDBMask1(7).Value
    End If
    If Len(.TDBMask1(4).Value) > 5 Then
        m_objrs!OFFICENO = .TDBMask1(4).Value
        m_objrs!extoffice = .Text1(21).Text
    End If
    If Len(.TDBMask1(5).Value) > 5 Then
        m_objrs!OFFICENO2 = .TDBMask1(5).Value
        m_objrs!EXTOFFICE2 = .Text1(32).Text
    End If
    m_objrs!agent = MDIForm1.Text1.Text
    m_objrs!NEXTACT = .Text2.Text & " " & .RichTextBox1(3).Text
    m_objrs!NEXTACTDATE = .TDBDate1(3).Value & " " & .TDBTime1.Value
    m_objrs!PRODUCTOFFERED = .Combo2(0).Text
    m_objrs!VOLOFFERED = CCur(.TDBNumber1(2).Value)
    m_objrs!PRODUCTAPPROVED = .Combo2(2).Text
    m_objrs!VOLAPPROVED = CCur(.TDBNumber1(1).Value)
    m_objrs!RECSOURCE = .Combo1(0).Text
    m_objrs!TGLSOURCE = MDIForm1.TDBDate1.Value
    m_objrs!RECSTATUS = STATUS_FIELD_LAMA
    m_objrs!TGLSTATUS = MDIForm1.TDBDate1.Value & " " & Time
    m_objrs!KETHSLKERJA = KETHSLKERJA
    m_objrs!OTHERS = .Text3.Text
    m_objrs!NOLAP = .Text1(4).Text
    m_objrs!AHOMENO = .TDBMask1(8).Value
    m_objrs!AHOMENO2 = .TDBMask1(9).Value
    m_objrs!AOFFICENO = .TDBMask1(10).Value
    m_objrs!AOFFICENO2 = .TDBMask1(11).Value
    m_objrs!AFAXNO = .TDBMask1(12).Value
    m_objrs!AFAXNO2 = .TDBMask1(13).Value
    m_objrs!DOK1 = DOK1
    m_objrs!KD_CLS = .Combo3(0).Text
'    m_objrs!LIMIT_CR = .TDBNumber1(2).Value
 '   m_objrs!LIMIT_CR_APP = .TDBNumber1(3).Value
    m_objrs.UPDATE
    m_objrs.Requery
End With
m_objrs.Close
Set m_objrs = Nothing
End Function

Public Function ADD_HISTORY(M_OBJCONN As ADODB.Connection, CUSTID As String, dATETIME1 As String, JAM As String, agent As String, _
                            PRODUCTS As String, HST As String)
Dim CMDSQL As String
JAM = Format(Time, "hh:mm")
JAM = Left(JAM, 2) & Right(JAM, 2)
CMDSQL = " INSERT INTO KTA_CUSTHSTTBL"
CMDSQL = CMDSQL + " (CUSTID,"
CMDSQL = CMDSQL + " DATETIME,"
CMDSQL = CMDSQL + " AGENT,"
CMDSQL = CMDSQL + " PRODUCTS,"
CMDSQL = CMDSQL + " HST)"
CMDSQL = CMDSQL + " VALUES"
CMDSQL = CMDSQL + " ('" + CUSTID + "',"
CMDSQL = CMDSQL + " '" + Format(dATETIME1, "yyyymmdd") & JAM + "',"
CMDSQL = CMDSQL + " '" + agent + "',"
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(PRODUCTS) + "',"
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(HST) + "')"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox Err.Description
End Function

Public Function UPDATE_CUSTOMER_BARU(M_OBJCONN As ADODB.Connection, KETHSLKERJA As String, STATUS_FIELD_LAMA As String, M_CALL As String, M_STATUS As String, DOK1 As String)
Dim m_objrs As New ADODB.Recordset
Dim m_nama As String
With FRMCUST_CC
m_objrs.CursorLocation = adUseClient
m_objrs.Open "SELECT * FROM CC_CUSTTBL WHERE CUSTID = '" + .Text1(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_objrs.RecordCount <> 0 Then
    m_nama = UCase(IIf(IsNull(m_objrs("name")), "", m_objrs("name")))
    If m_nama <> "" Then
        If m_nama <> UCase(.Text1(0).Text) Then
            Call ADD_HISTORY(M_OBJCONN, .Text1(1).Text, MDIForm1.TDBDate1.Text, CStr(Time), MDIForm1.Text1.Text, "", "UBAH NAMA OLEH " & MDIForm1.Text1.Text)
        End If
    End If
End If
If m_objrs.RecordCount <> 0 Then
    m_objrs!Name = .Text1(0).Text
    m_objrs!PRIOR = .Combo5.Text
    m_objrs!Title = .Combo1(2).Text
    If .TDBDate1(0).ValueIsNull Then
    Else
        m_objrs!BIRTHD = Format(.TDBDate1(0).Value, "mm/dd/yyyy")
    End If
    If .Option7(0).Value Then
        m_objrs!QUALIFIED = 1
    Else
        If .Option7(1).Value Then
            m_objrs!QUALIFIED = 0
        End If
    End If
    m_objrs!ADDRNOW = .RichTextBox1(0).Text
    m_objrs!ZIPNOW = .Text1(6).Text
    m_objrs!CITYNOW = .Text1(7).Text
    If Len(.TDBMask1(0).Value) > 1 Then
        m_objrs!HOMENO = .TDBMask1(0).Value
    End If
    If Len(.TDBMask1(1).Text) > 1 Then
        m_objrs!HOMENO2 = .TDBMask1(1).Value
    End If
    If Len(.TDBMask1(2).Text) > 1 Then
        m_objrs!MOBILENO = .TDBMask1(2).Value
    End If
    If Len(.TDBMask1(3).Text) > 1 Then
        m_objrs!MOBILENO2 = .TDBMask1(3).Value
    End If
    If .Option4(0).Value Then
        m_objrs!CAT = "0"
        m_objrs!NAMAPT = .Text1(20).Text
        m_objrs!ADDRPT = .RichTextBox1(2).Text
    Else
        m_objrs!CAT = "1"
        m_objrs!JENISUSAHA = .Text1(34).Text
        m_objrs!NAMAPT = .Text1(18).Text
        m_objrs!ADDRPT = .RichTextBox1(4).Text
    End If
    If Len(.TDBMask1(6).Value) > 1 Then
        m_objrs!FAXNO = .TDBMask1(6).Value
    End If
    If Len(.TDBMask1(7).Value) > 1 Then
        m_objrs!FAXNO2 = .TDBMask1(7).Value
    End If
    If Len(.TDBMask1(4).Value) > 1 Then
        m_objrs!OFFICENO = .TDBMask1(4).Value
        m_objrs!extoffice = .Text1(21).Text
    End If
    If Len(.TDBMask1(5).Value) > 1 Then
        m_objrs!OFFICENO2 = .TDBMask1(5).Value
        m_objrs!EXTOFFICE2 = .Text1(32).Text
    End If
    m_objrs!NEXTACT = .Text2.Text & " " & .RichTextBox1(3).Text
    m_objrs!NEXTACTDATE = .TDBDate1(3).Value & " " & .TDBTime1.Value
    m_objrs!PRODUCTOFFERED = .Combo2(0).Text
    m_objrs!VOLOFFERED = CCur(.TDBNumber1(2).Value)
    m_objrs!PRODUCTAPPROVED = .Combo2(2).Text
    m_objrs!VOLAPPROVED = CCur(.TDBNumber1(1).Value)
    m_objrs!RECSOURCE = .Combo1(0).Text
'    If M_STATUS = "1" Then
        m_objrs!RECSTATUS = STATUS_FIELD_LAMA
        If .Check2(3).Value = 1 Then
            m_objrs!TGLSTATUS = .TDBDate1(1).Value & " " & Time
        Else
            m_objrs!TGLSTATUS = MDIForm1.TDBDate1.Value & " " & Time
        End If
        m_objrs!KETHSLKERJA = KETHSLKERJA
 '   End If
    If CALL_OK Then
        m_objrs!TGLCALL = MDIForm1.TDBDate1.Value
    End If
    m_objrs!OTHERS = .Text3.Text
    m_objrs!NOLAP = .Text1(4).Text
    m_objrs!AHOMENO = .TDBMask1(8).Value
    m_objrs!AHOMENO2 = .TDBMask1(9).Value
    m_objrs!AFAXNO = .TDBMask1(12).Value
    m_objrs!AFAXNO2 = .TDBMask1(13).Value
    m_objrs!AOFFICENO = .TDBMask1(10).Value
    m_objrs!AOFFICENO2 = .TDBMask1(11).Value
    m_objrs!DOK1 = DOK1
    m_objrs!KD_CLS = .Combo3(0).Text
    m_objrs.UPDATE
    m_objrs.Requery
End If
m_objrs.Close
Set m_objrs = Nothing
End With

End Function

Public Function QUERY_CEK_ADDCUST(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim m_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM KTA_CUSTTBL"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
CMDSQL = CMDSQL + " ORDER BY NAME"
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_CEK_ADDCUST = m_objrs
Set m_objrs = Nothing
End Function
