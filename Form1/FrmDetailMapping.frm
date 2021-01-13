VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmDetailMapping 
   BackColor       =   &H009AD6C2&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2115
   ClientLeft      =   750
   ClientTop       =   4035
   ClientWidth     =   5310
   LinkTopic       =   "Form2"
   ScaleHeight     =   2115
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3330
      TabIndex        =   0
      Top             =   1740
      Width           =   1965
   End
   Begin MSComctlLib.ListView LvDetail 
      Height          =   1350
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   2381
      View            =   3
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Info Kartu Detail Mapping"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   2115
   End
End
Attribute VB_Name = "FrmDetailMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
    LvDetail.ColumnHeaders.ADD , , "Card No", 2000
    LvDetail.ColumnHeaders.ADD , , "Class", 1800
    LvDetail.ColumnHeaders.ADD , , "Balance", 2000
End Sub



Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call header
    Call isi
End Sub


Private Sub isi()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    
    CMDSQL = "select * from tbldetailmapping where custid='"
    CMDSQL = CMDSQL + FrmCC_Colection.lblCustId.Text + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvDetail.ListItems.ADD(, , M_objrs("cardno"))
                ListItem.SubItems(1) = M_objrs("class")
                ListItem.SubItems(2) = cnull(M_objrs("balance"))
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub
