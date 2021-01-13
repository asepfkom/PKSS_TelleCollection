VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_DISTRIBUSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ADD_OK As Boolean

'Private Function ADD_CUSTTBL(M_OBJCONN As ADODB.Connection, CustId As String, NAME1 As String, TITLE1 As String, _
'                            BIRTHD As String, AddrNow As String, ZIPNOW As String, CITYNOW As String, AHOMENO As String, _
'                            HOMENO As String, AHOMENO2 As String, HOMENO2 As String, MOBILENO As String, MOBILENO2 As String, _
'                            CAT As String, JENISUSAHA As String, NAMAPT As String, ADDRPT As String, AFAXNO As String, FAXNO As String, _
'                            AFAXNO2 As String, FAXNO2 As String, AOFFICENO As String, OFFICENO As String, EXTOFFICENO As String, _
'                            AOFFICENO2 As String, OFFICENO2 As String, EXTOFFICENO2 As String, AGENT As String, NEXTACT As String, _
'                            NEXTACTDATE As String, RECSOURCE As String, TGLSOURCE As String, RECSTATUS As String, KETHSLKERJA As String, _
'                            TGLSTATUS As String, OTHERS As String, TIPE_PRODUK As String, NOLAP As String, NamaAgent As String)
'Dim CMDSQL As String
'Dim m_objcek As New ADODB.Recordset
'm_objcek.CursorLocation = adUseClient
''If FRM_DISTRIBUTE.Option1(0).Value Then
'    CMDSQL = " INSERT INTO mgm"
''Else
'    CMDSQL = " INSERT INTO XSELLBANK"
''End If
'     CMDSQL = CMDSQL + " (CUSTID,"
'     CMDSQL = CMDSQL + " NAME,"
''     cmdsql = cmdsql + " Title,"
'     If BIRTHD <> Empty Then
'        CMDSQL = CMDSQL + " BIRTHD,"
'     End If
'     CMDSQL = CMDSQL + " ADDRNOW,"
'     CMDSQL = CMDSQL + " ZIPNOW,"
'     CMDSQL = CMDSQL + " CITYNOW,"
'     CMDSQL = CMDSQL + " AHOMENO,"
'     CMDSQL = CMDSQL + " HOMENO,"
'     CMDSQL = CMDSQL + " AHOMENO2,"
'     CMDSQL = CMDSQL + " HOMENO2,"
'     CMDSQL = CMDSQL + " MOBILENO,"
'     CMDSQL = CMDSQL + " MOBILENO2,"
' '    cmdsql = cmdsql + " CAT,"
' '    cmdsql = cmdsql + " JENISUSAHA,"
'     CMDSQL = CMDSQL + " NAMAPT,"
'     CMDSQL = CMDSQL + " ADDRPT,"
'     CMDSQL = CMDSQL + " AFAXNO,"
'     CMDSQL = CMDSQL + " FAXNO,"
'     CMDSQL = CMDSQL + " AFAXNO2,"
'     CMDSQL = CMDSQL + " FAXNO2,"
'     CMDSQL = CMDSQL + " AOFFICENO,"
'     CMDSQL = CMDSQL + " OFFICENO,"
'     CMDSQL = CMDSQL + " EXTOFFICE,"
'     CMDSQL = CMDSQL + " AOFFICENO2,"
'     CMDSQL = CMDSQL + " OFFICENO2,"
'     CMDSQL = CMDSQL + " EXTOFFICE2,"
'     CMDSQL = CMDSQL + " agent,"
' '    cmdsql = cmdsql + " NEXTACT,"
' '    cmdsql = cmdsql + " NEXTACTDATE,"
'     CMDSQL = CMDSQL + " RECSOURCE,"
''     cmdsql = cmdsql + " TGLSOURCE,"
'     CMDSQL = CMDSQL + " RECSTATUS,"
'     CMDSQL = CMDSQL + " TGLSTATUS,"
' '    cmdsql = cmdsql + " TGLDISTRIBUSI,"
'     CMDSQL = CMDSQL + " KETHSLKERJA,"
''    cmdsql = cmdsql + " NOLAP,"
'    CMDSQL = CMDSQL + " NamaAgent,"
'     CMDSQL = CMDSQL + " OTHERS)"
'    CMDSQL = CMDSQL + " VALUES"
'    CMDSQL = CMDSQL + " ('" + CustId + "',"
'    CMDSQL = CMDSQL + " '" + NAME1 + "',"
''    cmdsql = cmdsql + " '" + TITLE1 + "',"
'    If BIRTHD <> Empty Then
'        CMDSQL = CMDSQL + " '" + Format(BIRTHD, "mm/dd/yy") + "',"
'    End If
'    CMDSQL = CMDSQL + " '" + AddrNow + "',"
'    CMDSQL = CMDSQL + " '" + ZIPNOW + "',"
'    CMDSQL = CMDSQL + " '" + CITYNOW + "',"
'    CMDSQL = CMDSQL + " '" + AHOMENO + "',"
'    CMDSQL = CMDSQL + " '" + HOMENO + "',"
'    CMDSQL = CMDSQL + " '" + AHOMENO2 + "',"
'    CMDSQL = CMDSQL + " '" + HOMENO2 + "',"
'    CMDSQL = CMDSQL + " '" + MOBILENO + "',"
'    CMDSQL = CMDSQL + " '" + MOBILENO2 + "',"
' '   cmdsql = cmdsql + " '" + CAT + "',"
' '   cmdsql = cmdsql + " '" + JENISUSAHA + "',"
'    CMDSQL = CMDSQL + " '" + NAMAPT + "',"
'    CMDSQL = CMDSQL + " '" + ADDRPT + "',"
'    CMDSQL = CMDSQL + " '" + AFAXNO + "',"
'    CMDSQL = CMDSQL + " '" + FAXNO + "',"
'    CMDSQL = CMDSQL + " '" + AFAXNO2 + "',"
'    CMDSQL = CMDSQL + " '" + FAXNO2 + "',"
'    CMDSQL = CMDSQL + " '" + AOFFICENO + "',"
'    CMDSQL = CMDSQL + " '" + OFFICENO + "',"
'    CMDSQL = CMDSQL + " '" + EXTOFFICENO + "',"
'    CMDSQL = CMDSQL + " '" + AOFFICENO2 + "',"
'    CMDSQL = CMDSQL + " '" + OFFICENO2 + "',"
'    CMDSQL = CMDSQL + " '" + EXTOFFICENO2 + "',"
'    CMDSQL = CMDSQL + " '" + AGENT + "',"
' '   cmdsql = cmdsql + " '" + NEXTACT + "',"
' '   cmdsql = cmdsql + " '" + Format(NEXTACTDATE, "mm/dd/yy hh:mm") + "',"
'    CMDSQL = CMDSQL + " '" + FRM_DISTRIBUTE.Combo1(0).Text + "',"
' '   cmdsql = cmdsql + " '" + Format(TGLSOURCE, "mm/dd/yy hh:mm") + "',"
'    CMDSQL = CMDSQL + " '" + RECSTATUS + "',"
'    CMDSQL = CMDSQL + " '" + Format(MDIForm1.TDBDate1.Text, "mm/dd/yyyy") + "',"
'    'cmdsql = cmdsql + " '" + Format(MDIForm1.TDBDate1.Text, "mm/dd/yyyy") + "',"
'    CMDSQL = CMDSQL + " '" + KETHSLKERJA + "',"
''    cmdsql = cmdsql + " '" + NOLAP + "',"
'    CMDSQL = CMDSQL + " '" + NamaAgent + "',"
'    CMDSQL = CMDSQL + " '" + OTHERS + "')"
'
'    On Error GoTo add_error
'  '  WaitSecs (1)
'    M_OBJCONN.Execute CMDSQL
'    ADD_OK = True
'    Exit Function
'add_error:
'    ADD_OK = False
'  ''  MsgBox Err.Description
'End Function


'Private Function QUERY_TEMPCUSTTBL_KTA(M_OBJCONN As ADODB.Connection, TIPE_PRODUK As String, USERID As String, NamaAgent As String)
'Dim CMDSQL As String
'Dim i As Currency
'Dim CustId As String, NAME1 As String, TITLE1 As String, _
'                            BIRTHD As String, AddrNow As String, ZIPNOW As String, CITYNOW As String, AHOMENO As String, _
'                            HOMENO As String, AHOMENO2 As String, HOMENO2 As String, MOBILENO As String, MOBILENO2 As String, _
'                            CAT As String, JENISUSAHA As String, NAMAPT As String, ADDRPT As String, AFAXNO As String, FAXNO As String, _
'                            AFAXNO2 As String, FAXNO2 As String, AOFFICENO As String, OFFICENO As String, EXTOFFICENO As String, _
'                            AOFFICENO2 As String, OFFICENO2 As String, EXTOFFICENO2 As String, AGENT As String, NEXTACT As String, _
'                            NEXTACTDATE As String, PRODUCTOFFERED As String, VOLOFFERED As String, PRODUCTAPPROVED As String, _
'                            VOLAPPROVED As String, RECSOURCE As String, TGLSOURCE As String, RECSTATUS As String, KETHSLKERJA As String, _
'                            TGLSTATUS As String, OTHERS As String, NOLAP As String
'Dim M_objrs As ADODB.Recordset
'
'CMDSQL = "SELECT * FROM tempCC_CUSTTBL"
'CMDSQL = CMDSQL + " WHERE AGENT = '" + USERID + "'"
'CMDSQL = CMDSQL + " ORDER BY NAME"
'
'Set M_objrs = New ADODB.Recordset
'M_objrs.CursorLocation = adUseClient
'M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'If M_objrs.RecordCount = 0 Then
'    'FRM_DISTRIBUTE.ProgressBar1.Max = 100
'Else
'    'FRM_DISTRIBUTE.ProgressBar1.Max = (M_objrs.RecordCount + 1)
'
'End If
'    'FRM_DISTRIBUTE.SSFrame1.Visible = True
'    'FRM_DISTRIBUTE.ProgressBar1.Visible = True
'i = 1
'While Not M_objrs.EOF
'    FRM_DISTRIBUTE.ProgressBar1.Value = M_objrs.Bookmark
'    CustId = IIf(IsNull(M_objrs("CUSTID")), "", M_objrs("CUSTID"))
'    NOLAP = IIf(IsNull(M_objrs("NOLAP")), "", M_objrs("NOLAP"))
'    NamaAgent = IIf(IsNull(M_objrs("NamaAgent")), "", M_objrs("NamaAgent"))
'    NAME1 = IIf(IsNull(M_objrs("NAME")), "", M_objrs("NAME"))
'    TITLE1 = IIf(IsNull(M_objrs("TITLE")), "", M_objrs("TITLE"))
'    BIRTHD = IIf(IsNull(M_objrs("BIRTHD")), "", M_objrs("BIRTHD"))
'    AddrNow = IIf(IsNull(M_objrs("ADDRNOW")), "", M_objrs("ADDRNOW"))
'    ZIPNOW = IIf(IsNull(M_objrs("ZIPNOW")), "", M_objrs("ZIPNOW"))
'    CITYNOW = IIf(IsNull(M_objrs("CITYNOW")), "", M_objrs("CITYNOW"))
'    AHOMENO = IIf(IsNull(M_objrs("AHOMENO")), "", M_objrs("AHOMENO"))
'    HOMENO = IIf(IsNull(M_objrs("HOMENO")), "", M_objrs("HOMENO"))
'    AHOMENO2 = IIf(IsNull(M_objrs("AHOMENO2")), "", M_objrs("AHOMENO2"))
'    HOMENO2 = IIf(IsNull(M_objrs("HOMENO2")), "", M_objrs("HOMENO2"))
'    MOBILENO = IIf(IsNull(M_objrs("MOBILENO")), "", M_objrs("MOBILENO"))
'    MOBILENO2 = IIf(IsNull(M_objrs("MOBILENO2")), "", M_objrs("MOBILENO2"))
'    CAT = IIf(IsNull(M_objrs("CAT")), "", M_objrs("CAT"))
'    JENISUSAHA = IIf(IsNull(M_objrs("JENISUSAHA")), "", M_objrs("JENISUSAHA"))
'    NAMAPT = IIf(IsNull(M_objrs("NAMAPT")), "", M_objrs("NAMAPT"))
'    ADDRPT = IIf(IsNull(M_objrs("ADDRPT")), "", M_objrs("ADDRPT"))
'    AFAXNO = IIf(IsNull(M_objrs("AFAXNO")), "", M_objrs("AFAXNO"))
'    FAXNO = IIf(IsNull(M_objrs("FAXNO")), "", M_objrs("FAXNO"))
'    AFAXNO2 = IIf(IsNull(M_objrs("AFAXNO2")), "", M_objrs("AFAXNO2"))
'    FAXNO2 = IIf(IsNull(M_objrs("FAXNO2")), "", M_objrs("FAXNO2"))
'    AOFFICENO = IIf(IsNull(M_objrs("AOFFICENO")), "", M_objrs("AOFFICENO"))
'    OFFICENO = IIf(IsNull(M_objrs("OFFICENO")), "", M_objrs("OFFICENO"))
'    EXTOFFICENO = IIf(IsNull(M_objrs("EXTOFFICE")), "", M_objrs("EXTOFFICE"))
'    AOFFICENO2 = IIf(IsNull(M_objrs("AOFFICENO2")), "", M_objrs("AOFFICENO2"))
'    OFFICENO2 = IIf(IsNull(M_objrs("OFFICENO2")), "", M_objrs("OFFICENO2"))
'    EXTOFFICENO2 = IIf(IsNull(M_objrs("EXTOFFICE2")), "", M_objrs("EXTOFFICE2"))
'    AGENT = IIf(IsNull(M_objrs("AGENT")), "", M_objrs("AGENT"))
'    NEXTACT = IIf(IsNull(M_objrs("NEXTACT")), "", M_objrs("NEXTACT"))
'    NEXTACTDATE = IIf(IsNull(M_objrs("NEXTACTDATE")), "", M_objrs("NEXTACTDATE"))
'    RECSOURCE = IIf(IsNull(M_objrs("RECSOURCE")), "", M_objrs("RECSOURCE"))
'    TGLSOURCE = IIf(IsNull(M_objrs("TGLSOURCE")), "", M_objrs("TGLSOURCE"))
'    RECSTATUS = IIf(IsNull(M_objrs("RECSTATUS")), "", M_objrs("RECSTATUS"))
'    KETHSLKERJA = IIf(IsNull(M_objrs("KETHSLKERJA")), "", M_objrs("KETHSLKERJA"))
'    TGLSTATUS = IIf(IsNull(M_objrs("TGLSTATUS")), "", M_objrs("TGLSTATUS"))
'    OTHERS = IIf(IsNull(M_objrs("OTHERS")), "", M_objrs("OTHERS"))
'    Call ADD_CUSTTBL(M_OBJCONN, CustId, NAME1, TITLE1, _
'                            BIRTHD, AddrNow, ZIPNOW, CITYNOW, AHOMENO, _
'                            HOMENO, AHOMENO2, HOMENO2, MOBILENO, MOBILENO2, _
'                            CAT, JENISUSAHA, NAMAPT, ADDRPT, AFAXNO, FAXNO, _
'                            AFAXNO2, FAXNO2, AOFFICENO, OFFICENO, EXTOFFICENO, _
'                            AOFFICENO2, OFFICENO2, EXTOFFICENO2, AGENT, NEXTACT, _
'                            NEXTACTDATE, RECSOURCE, TGLSOURCE, RECSTATUS, KETHSLKERJA, _
'                            TGLSTATUS, OTHERS, TIPE_PRODUK, NOLAP, NamaAgent)
'    M_objrs.MoveNext
'    'i = i + 100
'Wend
'    FRM_DISTRIBUTE.ProgressBar1.Value = FRM_DISTRIBUTE.ProgressBar1.Max
'    WaitSecs (2)
'    Call DELETE_TEMPCUSTTBL(M_OBJCONN, USERID, TIPE_PRODUK)
'Set M_objrs = Nothing
'End Function


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

Public Function PROSES(M_OBJCONN As ADODB.Connection, M_RPTCONN As ADODB.Connection, DATASOURCE As String, USERID As String, jumlah As String, tgljam As String, TIPE_PRODUK As String, NamaAgent As String)
Dim TGL As String
Dim JAM As String
Dim Tgl1 As String
Dim M_objrs As ADODB.Recordset
On Error GoTo hell

M_OBJCONN.BeginTrans
    If Len(tgljam) < 11 Then
        TGL = Format(MDIForm1.TDBDate1.Text, "mm/dd/yy")
        JAM = Format(Now, "hh:mm")
        Tgl1 = TGL + " " + JAM
    Else
        TGL = Mid(tgljam, 5, 2) + "/" + Mid(tgljam, 7, 2) + "/" + Left(tgljam, 4)
        JAM = Mid(tgljam, 9, 2) + ":" + Right(tgljam, 2)
        Tgl1 = TGL + " " + JAM
    End If
    If TIPE_PRODUK = Empty Then
    Exit Function
    Else
'        WaitSecs (1)
        'Call UPDATE_TEMPCUSTTBL_KTA(M_OBJCONN, USERID, JUMLAH, Tgl1, TIPE_PRODUK, NamaAgent)
        WaitSecs (1)
        'Call QUERY_TEMPCUSTTBL_KTA(M_OBJCONN, TIPE_PRODUK, USERID, NamaAgent)
    End If
    'FRM_DISTRIBUTE.ProgressBar1.Value = FRM_DISTRIBUTE.ProgressBar1.Max
    'FRM_DISTRIBUTE.ProgressBar1.Visible = False
    'FRM_DISTRIBUTE.ProgressBar1.Value = 0
    
Set M_objrs = Nothing
M_OBJCONN.CommitTrans
ADD_OK = True
Exit Function
hell:
    ADD_OK = False
    MsgBox err.Description
  '     Resume
    M_OBJCONN.RollbackTrans
    

End Function

'Private Function UPDATE_TEMPCUSTTBL_KTA(M_OBJCONN As ADODB.Connection, USERID As String, JUMLAH As String, tgljam As String, TIPE_PRODUK As String, NamaAgent As String)
'Dim CMDSQL As String
'Dim CustId As String
'Dim M_objrs As ADODB.Recordset
'Set M_objrs = New ADODB.Recordset
'M_objrs.CursorLocation = adUseClient
'
'    CMDSQL = "SELECT TOP " + JUMLAH
'        Select Case UCase(TIPE_PRODUK)
'        Case "KTA"
'            CMDSQL = CMDSQL + " CUSTID FROM tempKTA_CUSTTBL"
'        Case "KTA - CROSS SELL"
'            CMDSQL = CMDSQL + " CUSTID FROM tempCS_CUSTTBL"
'        Case "CREDIT CARD"
'            CMDSQL = CMDSQL + " CUSTID FROM tempCC_CUSTTBL"
'        Case "CC - CROSS SELL"
'            CMDSQL = CMDSQL + " CUSTID FROM tempCCCS_CUSTTBL"
'        End Select
'    CMDSQL = CMDSQL + " WHERE RECSOURCE ='" + FRM_DISTRIBUTE.Combo1(0).Text + "' ORDER BY CUSTID"
'
'M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'CMDSQL = Empty
'While Not M_objrs.EOF
'    CustId = IIf(IsNull(M_objrs("CUSTID")), " ", M_objrs("CUSTID"))
'    If CustId <> " " Then
'            Select Case UCase(TIPE_PRODUK)
'                Case "KTA"
'                    CMDSQL = "UPDATE tempKTA_CUSTTBL"
'                Case "KTA - CROSS SELL"
'                    CMDSQL = "UPDATE tempCS_CUSTTBL"
'                Case "CREDIT CARD"
'                    CMDSQL = "UPDATE tempCC_CUSTTBL"
'                Case "CC - CROSS SELL"
'                    CMDSQL = "UPDATE tempCCCS_CUSTTBL"
'            End Select
'        CMDSQL = CMDSQL + " SET AGENT = '" + Trim(USERID) + "',"
'        CMDSQL = CMDSQL + " NamaAgent = '" + NamaAgent + "',"
'        CMDSQL = CMDSQL + " NEXTACTDATE = '" + tgljam + "'"
'        CMDSQL = CMDSQL + " WHERE CUSTID = '" + CustId + "'"
'        M_OBJCONN.Execute CMDSQL
'    End If
'    M_objrs.MoveNext
'Wend
'Set M_objrs = Nothing
'End Function

Public Function QUERY_COMBO_DATASOURCE(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM DATASOURCETBL"
'CMDSQL = CMDSQL + " WHERE STATUS ='A'"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
    Else
    CMDSQL = CMDSQL + " where substring(kodeds,1,3)<>'inf'"
 End If
    
CMDSQL = CMDSQL + " ORDER BY KODEDS"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_DATASOURCE = M_objrs
Set M_objrs = Nothing
End Function

Public Function INSERT_DISTRIBUSI(M_RPTCONN As ADODB.Connection, M_OBJCONN As ADODB.Connection, SPVCODE As String, Tanggal As String)
'Dim CMDSQL As String
'Dim USERID As String
'Dim nama As String
'Dim TGLJAM1 As String
'Dim JAM As String
'Dim TGLJAM2 As String
'Dim i As Currency
'Dim M_objrs As ADODB.Recordset
'
'Call DELETE_DISTRIBUSI(M_RPTCONN)
''SPVCODE = Empty
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
End Function

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

CMDSQL = "SELECT sum(AmountWo) AS JML FROM mgm"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If M_objrs.RecordCount <> 0 Then
    HITUNG_TEMPCUST_CC = IIf(IsNull(M_objrs("JML")), 0, M_objrs("JML"))
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
