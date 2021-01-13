VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form_Tarik_Data 
   Caption         =   "Tarik Data"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13185
   LinkTopic       =   "Form2"
   ScaleHeight     =   8925
   ScaleWidth      =   13185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "Export dan Tarik"
      Height          =   466
      Left            =   10800
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8295
      UseMaskColor    =   -1  'True
      Width           =   2205
   End
   Begin VB.PictureBox Form_Report_Submit 
      BackColor       =   &H80000005&
      Height          =   8850
      Left            =   0
      ScaleHeight     =   8790
      ScaleWidth      =   13095
      TabIndex        =   0
      Top             =   0
      Width           =   13155
      Begin VB.TextBox TxtPath 
         Height          =   285
         Left            =   10440
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cmb_kdagent 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Form_Open_Data.frx":0000
         Left            =   5280
         List            =   "Form_Open_Data.frx":0002
         TabIndex        =   11
         Top             =   8880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmb_nmagent 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Form_Open_Data.frx":0004
         Left            =   6570
         List            =   "Form_Open_Data.frx":0006
         TabIndex        =   10
         Top             =   8880
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "Parkir Data"
         Height          =   465
         Left            =   10665
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8820
         UseMaskColor    =   -1  'True
         Width           =   1125
      End
      Begin VB.ComboBox Combo7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Form_Open_Data.frx":0008
         Left            =   1680
         List            =   "Form_Open_Data.frx":000A
         TabIndex        =   4
         Top             =   8880
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Search"
         Height          =   465
         Left            =   10800
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7800
         UseMaskColor    =   -1  'True
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000005&
         Caption         =   "Check All"
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   7770
         Width           =   1350
      End
      Begin VB.TextBox txt_jmldata 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   9885
         TabIndex        =   1
         Top             =   8130
         Width           =   705
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6780
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   11959
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComDlg.CommonDialog Cd_save 
         Left            =   12360
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Ms. Excel 97/2000/XP|*.xls"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Campaign Code"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00333333&
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   8880
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00333333&
         Height          =   330
         Index           =   1
         Left            =   4680
         TabIndex        =   12
         Top             =   8880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tools Tarik Data"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   1
         Left            =   1950
         TabIndex        =   9
         Top             =   255
         Width           =   3585
      End
      Begin VB.Image Image2 
         Height          =   825
         Left            =   -15
         Picture         =   "Form_Open_Data.frx":000C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   19245
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Verifikasi && Hardcopy"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   0
         Left            =   2175
         TabIndex        =   7
         Top             =   270
         Width           =   3585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00333333&
         Height          =   345
         Index           =   2
         Left            =   8775
         TabIndex        =   6
         Top             =   8175
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form_Tarik_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim get_noapp As String
Dim where_get_noapp As String
Dim where_get_noapp_vw As String

Private Sub cmb_kdagent_Click()
    cmb_nmagent.ListIndex = cmb_kdagent.ListIndex
End Sub

Private Sub cmb_kdagent_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmb_nmagent_Click()
    cmb_kdagent.ListIndex = cmb_nmagent.ListIndex
End Sub

Private Sub cmb_nmagent_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Check1_Click()
    If ListView2.ListItems.Count > 0 Then
        i = ListView2.ListItems.Count
        If Check1.Value = vbChecked Then
            For a = 1 To i
                ListView2.ListItems(a).Checked = True
            Next a
        Else
            For a = 1 To i
                ListView2.ListItems(a).Checked = False
            Next a
        End If
    Else
        If Check1.Value = 1 Then
            MsgBox "Tidak Ada Data", vbOKOnly + vbInformation, "Informasi"
            Check1.Value = vbUnchecked
            Exit Sub
        End If
    End If
End Sub

Private Sub Command1_Click()
    Dim M_M_objrs As New ADODB.Recordset
    
    Dim c1 As Integer
    Dim buka As Integer

    sStrsql = "select column_name from information_schema.columns  where table_name = 'mgm' order by ordinal_position"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

    ListView2.ListItems.CLEAR
    ListView2.ColumnHeaders.CLEAR
    If M_objrs.RecordCount > 0 Then
        c1 = M_objrs.RecordCount
        With ListView2.ColumnHeaders
            For i = 1 To c1
                .ADD i, , cnull(M_objrs!column_name)
                M_objrs.MoveNext
            Next i
        End With
    End If
    Set M_objrs = Nothing

    wherebuka = " and f_parkir = 0 and agent = 'TARIK' "
    'whereagent = " and agent = 'PARKIR' "
    'wherebuka = ""
    
'    If tgl_upload.Value <> "" Then
'        datewhere = " and tanggal_upload = '" & Format(tgl_upload.Value, "yyyy-mm-dd") & "'"
'    End If
    
    mwhere = ""
     If Combo7.text <> "" Then
        mwhere = mwhere & " and recsource = '" & Combo7.text & "' "
     End If
     
'     If cmb_kdagent.text <> "" Then
'        mwhere = mwhere & " and agent = '" & cmb_kdagent.text & "' "
'     End If
    
    sStrsql = "select * from mgm where 1=1" & mwhere & wherebuka
    '& mwhere & wherebuka & datewhere
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

    While Not M_objrs.EOF

        Set ListItem = ListView2.ListItems.ADD(, , cnull(M_objrs("custid")))
            For i = 1 To c1 - 1
                ListItem.SubItems(i) = cnull(M_objrs(i))
            Next i
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
    txt_jmldata.text = ListView2.ListItems.Count
End Sub
Private Sub isi_combo_agent()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    'Menegecek jenis User yang login, jika dia agent
    'maka combo nama agent di kunci
    
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        cmb_kdagent.text = MDIForm1.TxtUsername.text
        cmb_nmagent.text = MDIForm1.txtnama.text
        cmb_kdagent.Enabled = False
        
        cmb_nmagent.Enabled = False
    End If
    
    'Jika yang login=administrator
    If MDIForm1.txtlevel.text = "Administrator" Or MDIForm1.txtlevel.text = "Admin" Then
            CMDSQL = "SELECT * FROM USERTBL where   AKTIF='1'  order by  KDLEVEL='1' DESC,agent "
    End If
    'Jika yang login Supervisor
    If MDIForm1.txtlevel.text = "Supervisor" Then
        CMDSQL = "select  userid,agent from usertbl where   AKTIF='1' and ( spvcode='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "' or userid ='" + MDIForm1.TxtUsername.text + "')"
        CMDSQL = CMDSQL + " order by  KDLEVEL='1' DESC,agent "
    ElseIf MDIForm1.txtlevel.text = "Agent" Then
        CMDSQL = "select userid,agent from usertbl where userid='" + MDIForm1.TxtUsername.text + "'"
    End If
    'Jika yang login TeamLeader
    
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    'Jika tidak ada data agent/TL maka tutup form viewmgmdata
  
    
    While Not M_objrs.EOF
        
        cmb_kdagent.AddItem IIf(IsNull(M_objrs("userid")), "", M_objrs("userid"))
        cmb_nmagent.AddItem IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub
Private Sub campaign()
    sStrsql = "select distinct recsource from mgm "
    
    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        sStrsql = "select distinct recsource from mgm where agent in (select userid from usertbl where spvcode = '" & MDIForm1.TxtUsername.text & "' or userid = '" & MDIForm1.TxtUsername.text & "')"
        '============='
    End If
    
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql & " order by  recsource ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        '============='
        Combo7.CLEAR
        While Not M_objrs.EOF
            Combo7.AddItem IIf(IsNull(M_objrs!recsource), "", M_objrs!recsource)
            M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub

'Private Sub Command2_Click()
Private Sub tarik_data()
    Call getapp
    
    If ListView2.ListItems.Count = 0 Then
            MsgBox "Tidak Ada Data", vbOKOnly + vbCritical, "Informasi"
        Exit Sub
    End If
        
    If where_get_noapp = "" Then
        MsgBox "Pilih Terlebih Dahulu Data Yang Akan Dibuka!!", vbOKOnly + vbCritical, "Informasi"
        Exit Sub
    End If
    
    'M_OBJCONN.Execute "update mgm set f_parkir=1, date_parkir=now() where 1=1 " & where_get_noapp
    M_OBJCONN.Execute "update mgm set date_tarik=now() where 1=1 " & where_get_noapp
    'M_OBJCONN.Execut = "select column_name from information_schema.columns  where table_name = 'mgm' order by ordinal_position"
    
    MsgBox "Data Berhasil Ditarik", vbOKOnly + vbInformation, "Informasi"
    Call Command1_Click
End Sub

Private Sub getapp()
    get_noapp = ""
    where_get_noapp = ""
    For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).Checked Then
            get_noapp = get_noapp & "'" & ListView2.ListItems(i) & "',"
        End If
    Next i
    
    If get_noapp <> "" Then
        get_noapp = Left(get_noapp, Len(get_noapp) - 1)
        where_get_noapp = " and custid in (" & get_noapp & ")"
        where_get_noapp_vw = "and ""NOMER_KARTU"" in (" & get_noapp & ")"
    End If
End Sub

Private Sub Command3_Click()
    Dim strQuery As String
    strQuery = createQuery
    isi_dataSTATUS strQuery
End Sub

Private Sub Form_Load()
    Call isi_combo_agent
    'Call export_data
    Call campaign
End Sub

Public Function createQuery()
    Call getapp
    
    strsql = "select * from tarik_data where 1=1 " & where_get_noapp_vw
    createQuery = strsql
End Function


Private Sub isi_dataSTATUS(strsql As String)
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
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

    'Jika data tidak ada, maka keluar dari fungsi ini!
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data Blank!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
   
Form_Save:
    Cd_save.ShowSave
    TxtPath.text = Cd_save.FileName
    
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
    Call tarik_data
 
Salah:
    Exit Sub
    
End Sub
