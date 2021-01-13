VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmPilihanAtmLainnya 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pilihan ATM Lainnya"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   5820
      Width           =   1215
   End
   Begin MSComctlLib.ListView LstATM 
      Height          =   5385
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   9499
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   10147522
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
   Begin VB.Label Label1 
      Caption         =   "Pilih Salah Satu ATM"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2355
   End
End
Attribute VB_Name = "FrmPilihanAtmLainnya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
    LstATM.ColumnHeaders.ADD 1, , "No.", 1000
    LstATM.ColumnHeaders.ADD 2, , "Nama ATM", 3000
End Sub

Private Sub IsiData()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    Dim W As Integer
    
    CMDSQL = "select * from tbl_atm where aktif='1' order by nama_atm asc"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    W = 0
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            W = W + 1
            Set ListItem = LstATM.ListItems.ADD(, , CStr(W))
            ListItem.SubItems(1) = IIf(IsNull(M_objrs("nama_atm")), "", M_objrs("nama_atm"))
            M_objrs.MoveNext
        Wend
    End If
    
    Set M_objrs = Nothing
End Sub

Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Call header
    Call IsiData
End Sub


Private Sub LstATM_Click()
    If LstATM.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia"
        Exit Sub
    End If
    
    FrmCC_Colection.CmbViaPtp.Text = Trim(LstATM.SelectedItem.SubItems(1))
End Sub

Private Sub LstATM_DblClick()
    If LstATM.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia"
        Exit Sub
    End If

    FrmCC_Colection.CmbViaPtp.AddItem Trim(LstATM.SelectedItem.SubItems(1))
    FrmCC_Colection.CmbViaPtp.Text = Trim(LstATM.SelectedItem.SubItems(1))
    Me.Hide
End Sub
