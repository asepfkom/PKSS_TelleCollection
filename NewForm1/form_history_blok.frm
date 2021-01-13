VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form form_history_blok 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "History Blok"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11115
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdexcel 
      Caption         =   "Export to Excel"
      Height          =   375
      Left            =   8880
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.ListView listopenblok 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin TDBDate6Ctl.TDBDate date1 
      Height          =   285
      Left            =   7680
      TabIndex        =   3
      Top             =   480
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   503
      Calendar        =   "form_history_blok.frx":0000
      Caption         =   "form_history_blok.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "form_history_blok.frx":0184
      Keys            =   "form_history_blok.frx":01A2
      Spin            =   "form_history_blok.frx":0200
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   12648384
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
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
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   6815745
      Value           =   39876
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate date2 
      Height          =   285
      Left            =   9600
      TabIndex        =   4
      Top             =   480
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   494
      Calendar        =   "form_history_blok.frx":0228
      Caption         =   "form_history_blok.frx":0340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "form_history_blok.frx":03AC
      Keys            =   "form_history_blok.frx":03CA
      Spin            =   "form_history_blok.frx":0428
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   12648384
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
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
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   6815745
      Value           =   39876
      CenturyMode     =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Search by Tanggal"
      Height          =   255
      Left            =   7680
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "to"
      Height          =   255
      Left            =   9240
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "form_history_blok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub search()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    CMDSQL = "select * from mandiri.hstopenblock where date(tanggalopen) between '" + Format(date1.Value, "yyyy-mm-dd") + "' and '" + Format(date2.Value, "yyyy-mm-dd") + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    'TxtJmlAgentLogin.Text = M_Objrs.RecordCount
    listopenblok.ListItems.CLEAR
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = listopenblok.ListItems.ADD(, , M_objrs("agentterblock"))
                ListItem.SubItems(1) = M_objrs("alasan")
                ListItem.SubItems(2) = M_objrs("tanggalopen")
                ListItem.SubItems(3) = M_objrs("pembukablock")
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub

Private Sub HeaderAgentBlok()
    listopenblok.ColumnHeaders.ADD 1, , "AGENT", 2500
    listopenblok.ColumnHeaders.ADD 2, , "ALASAN", 4000
    listopenblok.ColumnHeaders.ADD 3, , "TANGGAL OPEN", 5000
    listopenblok.ColumnHeaders.ADD 4, , "OPEN By", 5000
End Sub

Private Sub isiAgentBlok()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    CMDSQL = "select * from mandiri.hstopenblock"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    'TxtJmlAgentLogin.Text = M_Objrs.RecordCount
    listopenblok.ListItems.CLEAR
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = listopenblok.ListItems.ADD(, , M_objrs("agentterblock"))
                ListItem.SubItems(1) = M_objrs("alasan")
                ListItem.SubItems(2) = M_objrs("tanggalopen")
                ListItem.SubItems(3) = M_objrs("pembukablock")
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub

Private Sub cmdsearch_Click()
    If date1.Text = "__/__/____" Or date2.Text = "__/__/____" Then
        MsgBox "Pilih Tanggal yang ingin dicari"
        Exit Sub
    End If
    Call search
End Sub

Private Sub Form_Load()
    Call HeaderAgentBlok
    Call isiAgentBlok
End Sub
