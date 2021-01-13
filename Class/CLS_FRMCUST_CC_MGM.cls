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
Dim cmdsql As String
Dim m_objrs As ADODB.Recordset
cmdsql = "SELECT * FROM PRODUCTTBL"
cmdsql = cmdsql + " WHERE UNIT ='Credit Card'"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " AND " + M_WHERE
 End If
cmdsql = cmdsql + " ORDER BY CODE"
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_PRODUCT = m_objrs
Set m_objrs = Nothing
End Function

Public Function QUERY_COMBO_CLOSSING(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim m_objrs As ADODB.Recordset
cmdsql = "SELECT * FROM CLSREASONTBL"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " WHERE " + M_WHERE
 End If
cmdsql = cmdsql + " ORDER BY KDCLS"
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_CLOSSING = m_objrs
Set m_objrs = Nothing
End Function

Public Function QUERY_COMBO_DATASOURCE(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim m_objrs As ADODB.Recordset
cmdsql = "SELECT * FROM DATASOURCETBL"
'CMDSQL = CMDSQL + " WHERE STATUS ='A'"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " WHERE " + M_WHERE
 End If
cmdsql = cmdsql + " ORDER BY KODEDS"
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_DATASOURCE = m_objrs
Set m_objrs = Nothing
End Function

Public Function QUERY_COMBO_DATASOURCE_ISI(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim m_objrs As ADODB.Recordset
cmdsql = "SELECT * FROM DATASOURCETBL"
cmdsql = cmdsql + " WHERE STATUS ='I'"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " AND " + M_WHERE
 End If
cmdsql = cmdsql + " ORDER BY KODEDS"
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_DATASOURCE_ISI = m_objrs
Set m_objrs = Nothing
End Function

Public Function QUERY_CUST(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim m_objrs As ADODB.Recordset
cmdsql = "SELECT * FROM MGM"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " WHERE " + M_WHERE
 End If
cmdsql = cmdsql + " ORDER BY NAME"
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_CUST = m_objrs
Set m_objrs = Nothing
End Function

Public Function QUERY_NAME_CUST(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim m_objrs As ADODB.Recordset
cmdsql = "SELECT NAME FROM MGM"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " WHERE " + M_WHERE
 End If
cmdsql = cmdsql + " ORDER BY NAME"
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_NAME_CUST = m_objrs
Set m_objrs = Nothing
End Function

Public Function QUERY_HIST_CUST(M_OBJCONN As ADODB.Connection, M_WHERE As String) As ADODB.Recordset
Dim cmdsql As String
On Error GoTo eRR_QUERY_HIST_CUST
Dim m_objrs As ADODB.Recordset
    cmdsql = "SELECT CUSTID,DATETIME,AGENT,HST, KDCOMPLAINT, REMARKCOMPLAINT "
    cmdsql = cmdsql + " From MGM_HST"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " WHERE " + M_WHERE
 End If
cmdsql = cmdsql + " ORDER BY DATETIME ASC "
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_HIST_CUST = m_objrs
Set m_objrs = Nothing
EXIT_eRR_QUERY_HIST_CUST:
    Exit Function
eRR_QUERY_HIST_CUST:
    MsgBox Err.Description
    Resume EXIT_eRR_QUERY_HIST_CUST
End Function


Public Function ADD_HISTORY(M_OBJCONN As ADODB.Connection, CUSTID As String, dATETIME1 As String, JAM As String, agent As String, _
                            PRODUCTS As String, HST As String, KODEDS As String, KdComplaint As String, RemarkComplaint As String)
Dim cmdsql As String
Dim DATETIME2 As String
JAM = Format(Time, "hh:mm")
JAM = Left(JAM, 2) & Right(JAM, 2)
cmdsql = " INSERT INTO MGM_HST"
cmdsql = cmdsql + " (CUSTID,"
cmdsql = cmdsql + " DATETIME,"
cmdsql = cmdsql + " AGENT,"
cmdsql = cmdsql + " KODEDS,"
cmdsql = cmdsql + " KdComplaint,"
cmdsql = cmdsql + " RemarkComplaint,"
If PRODUCTS <> Empty Then
    cmdsql = cmdsql + " PRODUCTS,"
End If
cmdsql = cmdsql + " HST)"
cmdsql = cmdsql + " VALUES"
cmdsql = cmdsql + " ('" + CUSTID + "',"
cmdsql = cmdsql + " '" + Format(dATETIME1, "yyyymmdd") & JAM + "',"
cmdsql = cmdsql + " '" + agent + "',"
cmdsql = cmdsql + " '" + KODEDS + "',"
cmdsql = cmdsql + " '" + KdComplaint + "',"
cmdsql = cmdsql + " '" + RemarkComplaint + "',"
If PRODUCTS <> Empty Then
    cmdsql = cmdsql + " '" + UBAH_QUOTE(PRODUCTS) + "',"
End If
cmdsql = cmdsql + " '" + UBAH_QUOTE(HST) + "')"
    On Error GoTo add_error
    M_OBJCONN.Execute cmdsql
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox Err.Description
End Function

Public Function UPDATE_CUSTOMER_BARU(M_OBJCONN As ADODB.Connection, KETHSLKERJA As String, STATUS_FIELD_LAMA As String, M_CALL As String, M_STATUS As String, DOK1 As String)
Dim m_objrs As New ADODB.Recordset
Dim cmdsql As String
Dim m_nama As String
On Error GoTo ADDERR
With FRMCUST_CC_MGM
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "SELECT * FROM MGM WHERE CUSTID = '" + .Text1(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_objrs.RecordCount <> 0 Then
        m_objrs!Name = .Text1(0).Text
        m_objrs!Title = .Combo1(2).Text
        If .TDBDate1(0).ValueIsNull Then
        Else
            m_objrs!BIRTHD = Format(.TDBDate1(0).Value, "mm/dd/yyyy")
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
        MsgBox "16"
        m_objrs!NEXTACT = .Text2.Text & " " & .RichTextBox1(3).Text
        MsgBox "17"
        m_objrs!NEXTACTDATE = .TDBDate1(3).Value
        MsgBox "18"
        m_objrs!RECSOURCE = .Combo1(0).Text
        MsgBox "19"
        m_objrs!TGLSTATUS = MDIForm1.TDBDate1.Value
        MsgBox "20"
        If .Check2(1).Value = 1 Then
            m_objrs!RECSTATUS = "N"
            m_objrs!KETHSLKERJA = .Combo3(0).Text
            m_objrs!StsLastCall = .Combo3(0).Text
            MsgBox "21"
        Else
            If .Check2(0).Value = 1 Then
                m_objrs!RECSTATUS = "C"
                m_objrs!KETHSLKERJA = .Combo2(0).Text
                m_objrs!StsLastCall = .Combo2(0).Text
                MsgBox "22"
            Else
                If .Check2(2).Value = 1 Then
                MsgBox "23"
                Else
                    m_objrs!RECSTATUS = "1A"
                    m_objrs!KETHSLKERJA = "1A"
                    MsgBox "24"
                End If
            End If
        End If
        MsgBox "25"
       ' m_objrs!OTHERS = .Text3.Text
        MsgBox "26"
        m_objrs!AHOMENO = .TDBMask1(8).Value
        MsgBox "27"
        m_objrs!AHOMENO2 = .TDBMask1(9).Value
        MsgBox "28"
        m_objrs!AFAXNO = .TDBMask1(12).Value
        MsgBox "29"
        m_objrs!AFAXNO2 = .TDBMask1(13).Value
        MsgBox "30"
        m_objrs!AOFFICENO = .TDBMask1(10).Value
        MsgBox "31"
        m_objrs!AOFFICENO2 = .TDBMask1(11).Value
        MsgBox "32"
        m_objrs!NAMAAGENT = MDIForm1.Text7.Text
        MsgBox "33"
        m_objrs!KD_CLS = .Combo4(0).Text
        MsgBox "34"
        m_objrs![PRIOR] = .Combo5.Text
        MsgBox "35"
        m_objrs!KdComplaint = .Combo6.Text
        MsgBox "36"
        m_objrs!RemarkComplaint = .RichTextBox1(1).Text
        MsgBox "37"
        m_objrs.UPDATE
        MsgBox "38"
        m_objrs.Requery
        MsgBox "39"
    End If
    MsgBox "40"
    m_objrs.Close
    MsgBox "41"
    Set m_objrs = Nothing
    MsgBox "42"
    End With
    MsgBox "43"
    Exit Function
ADDERR:
MsgBox Err.Description
End Function

Public Function QUERY_CEK_ADDCUST(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim m_objrs As ADODB.Recordset
cmdsql = "SELECT * FROM CC_CUSTTBL"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " WHERE " + M_WHERE
 End If
cmdsql = cmdsql + " ORDER BY NAME"
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_CEK_ADDCUST = m_objrs
Set m_objrs = Nothing
End Function
