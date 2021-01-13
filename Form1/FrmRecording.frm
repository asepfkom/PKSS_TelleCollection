VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRecording 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recording"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10755
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Pb1 
      Height          =   315
      Left            =   2820
      TabIndex        =   19
      Top             =   6840
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdHeringRecording 
      Caption         =   "Hearing selected recording..."
      Height          =   375
      Left            =   6180
      TabIndex        =   18
      Top             =   1740
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton CmdDownloadRecording 
      Caption         =   "Download selected recording..."
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   1740
      Width           =   3255
   End
   Begin VB.CommandButton CmdUnCekAll 
      Caption         =   "&Uncek All"
      Height          =   375
      Left            =   1380
      TabIndex        =   16
      Top             =   1740
      Width           =   1275
   End
   Begin VB.CommandButton CmdCekAll 
      Caption         =   "&Cek All"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1740
      Width           =   1275
   End
   Begin VB.TextBox TxtJmlhRecording 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   14
      Text            =   "0"
      Top             =   6900
      Width           =   975
   End
   Begin MSComctlLib.ListView LvRecording 
      Height          =   4635
      Left            =   60
      TabIndex        =   12
      Top             =   2160
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   8176
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton CmdCariRecording 
      Caption         =   "&Cari recording..."
      Height          =   915
      Left            =   8520
      TabIndex        =   11
      Top             =   540
      Width           =   2115
   End
   Begin VB.ComboBox CmbServer 
      Height          =   315
      ItemData        =   "FrmRecording.frx":0000
      Left            =   2340
      List            =   "FrmRecording.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox TxtCustid 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2340
      TabIndex        =   3
      Top             =   480
      Width           =   3435
   End
   Begin TDBDate6Ctl.TDBDate TxtTgl1 
      Height          =   315
      Left            =   2340
      TabIndex        =   6
      Top             =   1260
      Width           =   1380
      _Version        =   65536
      _ExtentX        =   2434
      _ExtentY        =   556
      Calendar        =   "FrmRecording.frx":0022
      Caption         =   "FrmRecording.frx":013A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmRecording.frx":01A6
      Keys            =   "FrmRecording.frx":01C4
      Spin            =   "FrmRecording.frx":0222
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "yyyy-mm-dd"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy-mm-dd"
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
      Text            =   "____-__-__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.12794198814265E-317
      CenturyMode     =   0
   End
   Begin TDBTime6Ctl.TDBTime TxtWaktu1 
      Height          =   315
      Left            =   3060
      TabIndex        =   7
      Top             =   1260
      Width           =   675
      _Version        =   65536
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   "FrmRecording.frx":024A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmRecording.frx":02B6
      Spin            =   "FrmRecording.frx":0306
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.99999
      MidnightMode    =   0
      MinTime         =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.02960316199441E-317
   End
   Begin TDBDate6Ctl.TDBDate txtTgl2 
      Height          =   315
      Left            =   5040
      TabIndex        =   9
      Top             =   1320
      Width           =   1380
      _Version        =   65536
      _ExtentX        =   2434
      _ExtentY        =   556
      Calendar        =   "FrmRecording.frx":032E
      Caption         =   "FrmRecording.frx":0446
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmRecording.frx":04B2
      Keys            =   "FrmRecording.frx":04D0
      Spin            =   "FrmRecording.frx":052E
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "yyyy-mm-dd"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy-mm-dd"
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
      Text            =   "____-__-__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.12794198814265E-317
      CenturyMode     =   0
   End
   Begin TDBTime6Ctl.TDBTime TxtWaktu2 
      Height          =   315
      Left            =   5595
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   795
      _Version        =   65536
      _ExtentX        =   1402
      _ExtentY        =   556
      Caption         =   "FrmRecording.frx":0556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmRecording.frx":05C2
      Spin            =   "FrmRecording.frx":0612
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.99999
      MidnightMode    =   0
      MinTime         =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.02960316199441E-317
   End
   Begin VB.Label Label3 
      Caption         =   "Jumlah Recording:"
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   6900
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "Sampai dengan"
      Height          =   255
      Index           =   3
      Left            =   3780
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Waktu telepon:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1260
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Custid Recording:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Lihat recording di server:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   900
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Pada form ini anda dapat mendownload/mendengarkan recording ..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10755
   End
End
Attribute VB_Name = "FrmRecording"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrKoneksi As String
Dim ConIcentra As ADODB.Connection
Dim NamaServer As String
Dim LinkDownload As String
Dim IPServerLink As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal IpOperation As String, ByVal IpFile As String, ByVal IpParameters As String, ByVal IpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub BukaLink(x As String)
    'Dim link As Long
    link = ShellExecute(0, vbNullString, x, "", "", vbNormalFocus)
End Sub

Private Sub CmbServer_Click()
    If Trim(UCase(CmbServer.Text)) = "SERVER 4" Then
        LvRecording.ListItems.CLEAR
        NamaServer = "SERVER 4"
        IPServerLink = "192.168.10.4"
        'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=localhost; PORT=5432; Database=icentra_4; UID=admin; PWD=admin321"
        'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.10.4; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
        'STRKONEKSI UNTUK OBELISK
        StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.10.4; PORT=5432; Database=obelisk; UID=crm; PWD=CRM!@#$%"
    ElseIf Trim(UCase(CmbServer.Text)) = "SERVER 5" Then
        LvRecording.ListItems.CLEAR
        NamaServer = "SERVER 5"
        IPServerLink = "192.168.10.5"
        'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=localhost; PORT=5432; Database=icentra_4; UID=admin; PWD=admin321"
        StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.10.5; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    End If
   
End Sub

Private Sub header()
    LvRecording.ColumnHeaders.ADD 1, , "Record ID", 900
    LvRecording.ColumnHeaders.ADD 2, , "Calldate", 1500
    LvRecording.ColumnHeaders.ADD 3, , "Destination", 1500
    LvRecording.ColumnHeaders.ADD 4, , "Duration (Sec)", 2000
    LvRecording.ColumnHeaders.ADD 5, , "SERVER", 2000
End Sub

Private Sub CmdCariRecording_Click()
    Dim cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim listItem As listItem
    
'    If TxtTgl1.ValueIsNull = True Or _
'       'TxtWaktu1.ValueIsNull = True Or _
'       txtTgl2.ValueIsNull = True Or _
'       'TxtWaktu2.ValueIsNull = True Then

    If TxtTgl1.ValueIsNull = True Or _
       txtTgl2.ValueIsNull = True Then
        
        MsgBox "Tanggal tidak boleh kosong!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    Set ConIcentra = New ADODB.Connection
    ConIcentra.Open StrKoneksi
    
    If Trim(UCase(CmbServer.Text)) = "SERVER 5" Then
        cmdsql = "select * from ota_bill_record as obr,ota_log_record as olr "
        cmdsql = cmdsql + " where date(obr.calldate) between '"
    '    Cmdsql = Cmdsql & Format(TxtTgl1.Value, "yyyy-mm-dd") & " " & Format(TxtWaktu1.Value, "hh:nn:ss") & "' and '"
    '    Cmdsql = Cmdsql & Format(txtTgl2.Value, "yyyy-mm-dd") & " " & Format(TxtWaktu2.Value, "hh:nn:ss") & "' and "
        cmdsql = cmdsql & Format(TxtTgl1.Value, "yyyy-mm-dd") & "' and '"
        cmdsql = cmdsql & Format(txtTgl2.Value, "yyyy-mm-dd") & "' and "
        cmdsql = cmdsql + " obr.uniqueid=olr.unique_id and obr.duration<>'0' "
        cmdsql = cmdsql + " and obr.dst<>'108' and olr.nocase='"
        cmdsql = cmdsql & TxtCustid.Text + "'"
        
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        DoEvents
        'vbHourglass
        M_Objrs.Open cmdsql, ConIcentra, adOpenDynamic, adLockOptimistic, adCmdText
    
        TxtJmlhRecording.Text = M_Objrs.RecordCount
        LvRecording.ListItems.CLEAR
    
        If M_Objrs.RecordCount = 0 Then
            MsgBox "Mohon maaf, recording yang anda cari tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
            Set ConIcentra = Nothing
            Exit Sub
        End If
    
        PB1.Max = M_Objrs.RecordCount
    
        While Not M_Objrs.EOF
            PB1.Value = M_Objrs.Bookmark
            Set listItem = LvRecording.ListItems.ADD(, , M_Objrs("uniqueid"))
                listItem.SubItems(1) = Format(M_Objrs("calldate"), "yyyy-mm-dd hh:nn:ss")
                listItem.SubItems(2) = IIf(IsNull(M_Objrs("dst")), "", M_Objrs("dst"))
                listItem.SubItems(3) = IIf(IsNull(M_Objrs("duration")), "", M_Objrs("duration"))
                listItem.SubItems(4) = NamaServer
            M_Objrs.MoveNext
        Wend
    ElseIf Trim(UCase(CmbServer.Text)) = "SERVER 4" Then
        cmdsql = "SELECT * FROM call_log"
        cmdsql = cmdsql + " WHERE date(start_time) between '"
        cmdsql = cmdsql & Format(TxtTgl1.Value, "yyyy-mm-dd") & "' AND '"
        cmdsql = cmdsql & Format(txtTgl2.Value, "yyyy-mm-dd") & "' AND "
        cmdsql = cmdsql + " customer_id = '" & TxtCustid.Text & "' "
        cmdsql = cmdsql + " and destination <> '108'"
        
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        DoEvents
        'vbHourglass
        M_Objrs.Open cmdsql, ConIcentra, adOpenDynamic, adLockOptimistic, adCmdText
        
        TxtJmlhRecording.Text = M_Objrs.RecordCount
        LvRecording.ListItems.CLEAR
        
        
        If M_Objrs.RecordCount = 0 Then
            MsgBox "Mohon maaf, recording yang anda cari tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
            Set ConIcentra = Nothing
            Exit Sub
        End If
        
        PB1.Max = M_Objrs.RecordCount
        
        While Not M_Objrs.EOF
            PB1.Value = M_Objrs.Bookmark
            Set listItem = LvRecording.ListItems.ADD(, , M_Objrs("unique_id"))
                listItem.SubItems(1) = Format(M_Objrs("start_time"), "yyyy-mm-dd hh:nn:ss")
                listItem.SubItems(2) = IIf(IsNull(M_Objrs("destination")), "", M_Objrs("destination"))
                listItem.SubItems(3) = IIf(IsNull(M_Objrs("duration")), "", M_Objrs("duration"))
                listItem.SubItems(4) = NamaServer
            M_Objrs.MoveNext
        Wend
        End If
        
'        Set M_Objrs = New ADODB.Recordset
'        M_Objrs.CursorLocation = adUseClient
'        DoEvents
'        'vbHourglass
'        M_Objrs.Open cmdsql, ConIcentra, adOpenDynamic, adLockOptimistic, adCmdText
'
'        TxtJmlhRecording.Text = M_Objrs.RecordCount
'        LvRecording.ListItems.CLEAR
'
'
'        If M_Objrs.RecordCount = 0 Then
'            MsgBox "Mohon maaf, recording yang anda cari tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
'            Set ConIcentra = Nothing
'            Exit Sub
'        End If
'
'        PB1.Max = M_Objrs.RecordCount
'
'        While Not M_Objrs.EOF
'            PB1.Value = M_Objrs.Bookmark
'            Set listItem = LvRecording.ListItems.ADD(, , M_Objrs("uniqueid"))
'                listItem.SubItems(1) = Format(M_Objrs("calldate"), "yyyy-mm-dd hh:nn:ss")
'                listItem.SubItems(2) = IIf(IsNull(M_Objrs("dst")), "", M_Objrs("dst"))
'                listItem.SubItems(3) = IIf(IsNull(M_Objrs("duration")), "", M_Objrs("duration"))
'                listItem.SubItems(4) = NamaServer
'            M_Objrs.MoveNext
'        Wend
        
        Set M_Objrs = Nothing
        Set ConIcentra = Nothing
End Sub

Private Sub CmdCekAll_Click()
    Dim w As Integer
    
    
    If LvRecording.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For w = 1 To LvRecording.ListItems.Count
        LvRecording.ListItems(w).Checked = True
    Next w
End Sub

Private Sub CmdDownloadRecording_Click()
    Dim w As Integer
    Dim K As Integer
    Dim S As Integer
    
    If LvRecording.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    K = 0
    
    For S = 1 To LvRecording.ListItems.Count
        If LvRecording.ListItems(S).Checked = True Then
            K = K + 1
        End If
    Next S
    
    If K = 0 Then
        MsgBox "Untuk mendownload recording, anda harus mencentang salah satu data!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For w = 1 To LvRecording.ListItems.Count
        If LvRecording.ListItems(w).Checked = True Then
            If Obelisk = False Then
                LinkDownload = "http://" & IPServerLink & "/tnis_new/webapp/recording_tnis.php?record_id=" & LvRecording.ListItems(w).Text & "&type=wav"
            Else
                LinkDownload = "http://" & IPServerLink & "/admin/recording/" & LvRecording.ListItems(w).Text & ""
            End If
            Call BukaLink(LinkDownload)
        End If
    Next w
    
End Sub

Private Sub CmdHeringRecording_Click()
    Dim w As Integer
    Dim K As Integer
    Dim S As Integer
    
    If LvRecording.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    K = 0
    
    For S = 1 To LvRecording.ListItems.Count
        If LvRecording.ListItems(S).Checked = True Then
            K = K + 1
        End If
    Next S
    
    If K = 0 Then
        MsgBox "Untuk mendengarkan recording, anda harus mencentang salah satu data!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If K > 1 Then
        MsgBox "Untuk mendengarkan recording, anda hanya boleh mencentang salah satu data!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    
    For w = 1 To LvRecording.ListItems.Count
        If LvRecording.ListItems(w).Checked = True Then
            LinkDownload = "http://" & IPServerLink & "/tnis_new/webapp/recording1_tnis.php?record_id=" & LvRecording.ListItems(w).Text & "&play=true"
            Call BukaLink(LinkDownload)
        End If
    Next w
End Sub

Private Sub CmdUnCekAll_Click()
    Dim w As Integer
    
    
    If LvRecording.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For w = 1 To LvRecording.ListItems.Count
        LvRecording.ListItems(w).Checked = False
    Next w
End Sub

Private Sub Form_Load()
    Call header
    CmbServer.Text = "SERVER 4"
    'LinkDownload = "http://" & IPServerLink & "/tnis_new/webapp/recording.php?record_id=" & LvRecording.SelectedItem.Text & "&type=wav"
    
End Sub

Private Sub LvRecording_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvRecording.SortKey = ColumnHeader.Index - 1
    LvRecording.Sorted = True
End Sub
