VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FormReport_sms 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report SMS"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtPath 
      Enabled         =   0   'False
      Height          =   480
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   160
   End
   Begin MSComctlLib.ListView List_ShowReport 
      Height          =   4845
      Left            =   270
      TabIndex        =   6
      Top             =   1605
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   8546
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdProses 
      Caption         =   "Proses"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7290
      TabIndex        =   5
      Top             =   855
      Width           =   975
   End
   Begin VB.CommandButton CmdExport 
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8280
      TabIndex        =   4
      Top             =   855
      Width           =   975
   End
   Begin VB.ComboBox txt_flag_export 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "FormReport_sms.frx":0000
      Left            =   1500
      List            =   "FormReport_sms.frx":000A
      TabIndex        =   3
      Top             =   1200
      Width           =   1995
   End
   Begin MSComDlg.CommonDialog CD_Save 
      Left            =   9495
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TDBDate6Ctl.TDBDate TdTglCall1 
      Height          =   315
      Left            =   1500
      TabIndex        =   7
      Top             =   780
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   556
      Calendar        =   "FormReport_sms.frx":001D
      Caption         =   "FormReport_sms.frx":0135
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FormReport_sms.frx":01A1
      Keys            =   "FormReport_sms.frx":01BF
      Spin            =   "FormReport_sms.frx":021D
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mm-yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__-__-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate TdTglCall2 
      Height          =   315
      Left            =   3285
      TabIndex        =   8
      Top             =   780
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   556
      Calendar        =   "FormReport_sms.frx":0245
      Caption         =   "FormReport_sms.frx":035D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FormReport_sms.frx":03C9
      Keys            =   "FormReport_sms.frx":03E7
      Spin            =   "FormReport_sms.frx":0445
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mm-yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__-__-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2955
      TabIndex        =   9
      Top             =   765
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kriteria"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   300
      TabIndex        =   2
      Top             =   1230
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   810
      Width           =   780
   End
   Begin VB.Line Line1 
      X1              =   300
      X2              =   9480
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report SMS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   300
      TabIndex        =   0
      Top             =   210
      Width           =   1530
   End
End
Attribute VB_Name = "FormReport_sms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim sQuerySelectTemp As String
Dim sQuerySelectTempID As String
Private Sub CmdProses_Click()
     If txt_flag_export.text = "INBOX" Then
        If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
            If Len(sWhere) = 0 Then
                sWhere = " where  date(received_sms_Date) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' "
                sWhere = sWhere + " and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
            Else
                sWhere = sWhere + "  and date(received_sms_Date) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' "
                sWhere = sWhere + " and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
            End If
        End If
    ElseIf txt_flag_export.text = "OUTBOX" Then
        If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
            If Len(sWhere) = 0 Then
                sWhere = " where  date(tgl_kirim) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' "
                sWhere = sWhere + " and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
            Else
                sWhere = sWhere + "  and date(tgl_kirim) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' "
                sWhere = sWhere + " and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
            End If
        End If
    End If
   
   If txt_flag_export.text = "" Then
    MsgBox "Harus diisi kriteria!"
    Exit Sub
   End If
   If txt_flag_export.text = "INBOX" Then
    
    sQuerySelect = " SELECT agent as ""Agent"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,custid as ""Customer ID"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,sender_number as ""Sender Number"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,text_sms as ""Message"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,received_sms_Date as ""Date Time"""
    sQuerySelect = sQuerySelect + vbCrLf + " FROM tbl_notif_sms"
    sQuerySelect = sQuerySelect + vbCrLf + sWhere
    
    ElseIf txt_flag_export.text = "OUTBOX" Then
     
    sQuerySelect = " SELECT agent as ""Agent"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,custid as ""Customer ID"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,NAME as ""Customer Name"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,notelp as ""Handphone Number"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,pesan as ""Message"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,tgl_kirim as ""Send Date"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,tgl_approve as ""Approval Date"""
    sQuerySelect = sQuerySelect + vbCrLf + " FROM request_sms"
    sQuerySelect = sQuerySelect + vbCrLf + sWhere
    
  End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sQuerySelect, M_OBJCONN, adOpenKeyset, adLockOptimistic
    
    sQuerySelectTemp = sQuerySelect
    Call ShowListView(rs, List_ShowReport)
    sQuerySelectTempID = "SELECT id FROM tbl_notif_sms " & sWhere
   If txt_flag_export.text = "INBOX" Then
    Call CmdInbox
   End If
    
    
End Sub
Private Sub CmdExport_Click()
   Call isi_data_ex(sQuerySelectTemp)
End Sub
Public Sub ShowListView(ByRef rsS As ADODB.Recordset, list As ListView, Optional no As Boolean = True)
    Dim j As Double
    j = 1
    list.ColumnHeaders.CLEAR
    If no = True Then
        list.ColumnHeaders.ADD 1, , "NO"
    End If
    If no = True Then
        For i = 0 To rsS.fields.Count - 1
            list.ColumnHeaders.ADD i + 2, , rsS.fields(i).Name
        Next i
    Else
        For i = 0 To rsS.fields.Count - 1
            list.ColumnHeaders.ADD i + 1, , rsS.fields(i).Name
        Next i
    End If
    list.ListItems.CLEAR
    While Not rsS.EOF
        If no = True Then
            Set listv = list.ListItems.ADD(, , j)
            For i = 0 To rsS.fields.Count - 1
                listv.SubItems(i + 1) = cnull(rsS(i))
            Next i
        Else
            Set listv = list.ListItems.ADD(, , cnull(rsS(0)))
            For i = 1 To rsS.fields.Count - 1
                listv.SubItems(i) = cnull(rsS(i))
            Next i
        End If
        j = j + 1
        rsS.MoveNext
    Wend
End Sub
Private Sub isi_data_ex(strsql As String)
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim cmdsql_update As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i As Integer
    Dim m_msgbox As String
     i = 1
     'STRSQL = txtSintaks.Text
 Set M_objrs = New ADODB.Recordset
 M_objrs.CursorLocation = adUseClient
 If txt_flag_export.text = "INBOX" Then
    M_objrs.Open strsql, M_OBJCONN1, adOpenDynamic, adLockOptimistic
 ElseIf txt_flag_export.text = "OUTBOX" Then
   M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
 Else
    MsgBox "isi kriteria dulu!", vbOKOnly + vbInformation, "Informasi"
    Exit Sub
End If
Form_Save:
    CD_save.ShowSave
    TxtPath.text = CD_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Download dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo Form_Save        '-> maka goto form_save
        End If
    End If
    'Set excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
    
    On Error GoTo Salah
    'Proses pengsisian nama field ke excel
    Dim X, Y    As Integer
        If M_objrs.State = 1 Then
            X = 0
            Y = M_objrs.fields().Count - 1
            Do Until X > Y
                DoEvents
                objSheet.Cells(1, i).Value = CStr(M_objrs.fields(X).Name)
                i = i + 1
                X = X + 1
            Loop
        End If
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs TxtPath.text, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
 
Salah:
    Exit Sub
End Sub
Private Sub CmdInbox()
    
    Dim satu As String
    Dim dua As String
    Dim tiga As String
    Dim empat As String
    Dim M_objrs As New ADODB.Recordset
    Dim cmdsql34 As String
    Dim TELPo As String
    Dim lst As ListItem
    Dim JmlBelumBaca As Integer
    Dim JmlSudahBaca As Integer

    'On Error Resume Next
    

    

    TELPo = "Select `ReceivingDateTime`, `SenderNumber`, `TextDecoded`,`ID`,`Processed` FROM inbox WHERE `SenderNumber` in ('a',"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
     
        'MsgBox "Silahkan tunggu! Program akan mencari inbox dari semua team anda!", vbOKOnly + vbInformation, "Informasi"
        'cmdsql34 = "SELECT contact1,contact2,mobileno FROM tbl_address WHERE custid in (SELECT custno FROM mgm WHERE agent in ("
        cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm "
     
    
    
    M_objrs.Open cmdsql34, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.EOF = False Then
        If M_objrs.RecordCount <> 0 Then
            'Pb1.Max = M_objrs.RecordCount
        Else
            MsgBox "Tidak ada data customer!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
    End If
    
    While Not M_objrs.EOF
        'Pb1.Value = M_objrs.Bookmark
        
        If M_objrs("mobileno2") <> "" Then
            TELPo = TELPo + "'" + Replace(Trim(M_objrs("mobileno2")), " ", "") + "',"
        End If
        If M_objrs("mobilenoadd1") <> "" Then
            TELPo = TELPo + "'" + Replace(Trim(M_objrs("mobilenoadd1")), " ", "") + "',"
        End If
        If M_objrs("mobileno") <> "" Then
            TELPo = TELPo + "'" + Replace(Trim(M_objrs("mobileno")), " ", "") + "',"
        End If
        If M_objrs("mobilenoadd2") <> "" Then
            TELPo = TELPo + "'" + Replace(Trim(M_objrs("mobilenoadd2")), " ", "") + "',"
        End If
    
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
    
    TELPo = Left(TELPo, Len(TELPo) - 1)
    Dim TELPo1
    Dim TELPo2
    If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
            If Len(sWhere) = 0 Then
                sWhere = " where  date(`ReceivingDateTime`) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' "
                sWhere = sWhere + " and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
            Else
                sWhere = sWhere + "  and date(`ReceivingDateTime`) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' "
                sWhere = sWhere + " and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
            End If
    End If
    If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
        TELPo1 = TELPo + ") AND date(`ReceivingDateTime`) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "' order by `ReceivingDateTime` desc " 'Ini yang belum pernah di baca
    Else
        TELPo1 = TELPo + ")  order by `ReceivingDateTime` desc " 'Ini yang belum pernah di baca
    End If
    'TELPo2 = TELPo + ") and `Processed`='true' order by `ReceivingDateTime` desc " 'Ini yang udah pernah di baca
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open TELPo1, M_OBJCONN1, adOpenDynamic, adLockOptimistic
    
    'Ini buat data inbox yang belum dibaca
    JmlBelumBaca = M_objrs.RecordCount
    If M_objrs.RecordCount <> 0 Then
        'Pb1.Max = JmlBelumBaca
    Else
        Dim Update_Status As String
        'MsgBox "Tidak ada sms baru!", vbOKOnly + vbInformation, "Informasi"
        'Update status sms di usertbl jadi null, supaya ga blink
        Update_Status = "update usertbl set status_sms=null where userid='"
        Update_Status = Update_Status + Trim(MDIForm1.txtusername.text) + "'"
        M_OBJCONN.Execute Update_Status
        'MDIForm1.TimerBlink.Enabled = False
        MDIForm1.Label9.ForeColor = vbBlack
    End If
    While Not M_objrs.EOF
        'Pb1.Value = M_objrs.Bookmark
        
        S = Format(M_objrs!receivingdatetime, "DD-MM-YYYY hh:mm:ss")
        t = Trim(M_objrs!sendernumber)
        u = M_objrs!textdecoded
        v = FindReplace(t, "+62", "0")
    
        If (Left(v, 3) = "021") Then
            v = Mid(v, 4, 20)
        End If
    
        Dim showlist As New ADODB.Recordset
        Dim TOTPTP As Currency
        Dim ssql As String
        
        If showlist.State = 1 Then showlist.Close
        ssql = "SELECT custid,agent, name FROM mgm WHERE mobileno='" & v & "'  or mobileno2='" & v & "'  or mobilenoadd1='" & v & "'  or mobilenoadd2='" & v & "'"
        showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        If showlist.EOF = False Then
            isicustid = showlist!CustId
            isiname = showlist!Name
            ISIAGENT = showlist!AGENT
            Set showlist = Nothing
        End If
        
        Set lst = List_ShowReport.ListItems.ADD(, , M_objrs.Bookmark) 'custid
            lst.SubItems(1) = Trim(ISIAGENT)  'agent
            lst.SubItems(2) = Trim(isicustid)  'Isi custid
            'lst.SubItems(1) = Trim(isiname)  'Isi nama
            lst.SubItems(3) = Trim(v) 'Telepon
            lst.SubItems(5) = S 'Receivingdatetime
            lst.SubItems(4) = Trim(IIf(IsNull(M_objrs("TextDecoded")), "", M_objrs("TextDecoded"))) 'Textsms
            'lst.SubItems(6) = M_objrs("id")
            'lst.SubItems(6) = M_objrs("Processed")
            lst.Bold = True
'            List_ShowReport.SelectedItem.ForeColor = vbRed
'
'            lst.ListSubItems(1).ForeColor = vbRed
'            lst.ListSubItems(2).ForeColor = vbRed
'            lst.ListSubItems(3).ForeColor = vbRed
'            lst.ListSubItems(4).ForeColor = vbRed
'            lst.ListSubItems(5).ForeColor = vbRed
'            lst.ListSubItems(6).ForeColor = vbRed
            M_objrs.MoveNext
    Wend
    sQuerySelectTemp = TELPo1
    Set M_objrs = Nothing
    
    
End Sub

Function FindReplace(SourceString, Searchstring, Replacestring) As String
  Dim tmpString1
  Dim tmpString2
  tmpString1 = SourceString
 
      tmpString2 = tmpString1
      tmpString1 = ReplaceFirstInstance(tmpString1, _
                   Searchstring, Replacestring)
      
      FindReplace = tmpString1
End Function
Function ReplaceFirstInstance(SourceString, _
Searchstring, Replacestring)
Dim StartLoc
Dim FoundLoc
  If StartLoc = 0 Then StartLoc = 1
  FoundLoc = InStr(StartLoc, SourceString, Searchstring) '*
  If FoundLoc <> 0 And FoundLoc < 2 Then
     ReplaceFirstInstance = Left(SourceString, FoundLoc - 1) & Replacestring & Right(SourceString, Len(SourceString) - (FoundLoc - 1) - Len(Searchstring))
     StartLoc = FoundLoc + Len(Replacestring)
  ElseIf FoundLoc > 1 Then
  
      ReplaceFirstInstance = Replacestring & "21" & SourceString

  Else
     StartLoc = 1

    ReplaceFirstInstance = SourceString
  End If
End Function

Private Sub HeaderList()
    List_ShowReport.ColumnHeaders.ADD , , "No", 1000
    List_ShowReport.ColumnHeaders.ADD , , "Agent", 1500
    List_ShowReport.ColumnHeaders.ADD , , "Customer ID", 1500
    List_ShowReport.ColumnHeaders.ADD , , "Sender Number", 1600
    List_ShowReport.ColumnHeaders.ADD , , "Message", 3000
    List_ShowReport.ColumnHeaders.ADD , , "Date Time", 1500
    List_ShowReport.ColumnHeaders.ADD , , "id", 0
    List_ShowReport.ColumnHeaders.ADD , , "processed", 0
End Sub

Private Sub Form_Load()
    Call HeaderList
End Sub
