VERSION 5.00
Begin VB.Form FRM_VER_REJECT 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1215
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1560
      TabIndex        =   6
      Top             =   1065
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   570
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Tutup"
      Height          =   375
      Index           =   0
      Left            =   3150
      TabIndex        =   1
      Top             =   705
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reject"
      Height          =   375
      Index           =   1
      Left            =   1635
      TabIndex        =   0
      Top             =   705
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "CustId Database  Upload :"
      Height          =   480
      Left            =   135
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Agent :"
      Height          =   255
      Left            =   195
      TabIndex        =   4
      Top             =   585
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   405
      Left            =   60
      TabIndex        =   2
      Top             =   75
      Width           =   4575
   End
End
Attribute VB_Name = "FRM_VER_REJECT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click(Index As Integer)
Dim cmdsql As String
Dim reason As String
Select Case Index
    Case 0
        Unload Me
        
    Case 1
        M_OBJCONN.Execute "Update RequestInbound set StatusRequest =2 where custid ='" + Label1.Caption + "'"
  '      M_OBJCONN.Execute "Delete From tempCC_CUSTTBL where custid ='" + Text1.Text + "'"
        reason = "Data Inbound Yang Anda Masukan Telah di Reject dan Dihapus Oleh " & MDIForm1.Text1.Text
        cmdsql = "Insert Into RequestInboundRst "
        cmdsql = cmdsql + " (custid, "
        cmdsql = cmdsql + " AgentLama, "
        cmdsql = cmdsql + " AgentBaru, "
        cmdsql = cmdsql + " Reason) "
        cmdsql = cmdsql + " Values "
        cmdsql = cmdsql + " ('" + Label1.Caption + "' , "
        cmdsql = cmdsql + " '" + FRM_VER_INBOUND.ListView1.SelectedItem.SubItems(14) + "', "
        cmdsql = cmdsql + " '" + Combo1.Text + "', "
        cmdsql = cmdsql + " '" + reason + "') "
        M_OBJCONN.Execute cmdsql
        M_OBJCONN.Execute "Update cc_custtbl set agent ='" + Combo1.Text + "' where custid ='" + Label1.Caption + "'"
        MsgBox "Proses Selesai"
        Unload Me
End Select
End Sub

Private Sub Form_Load()
Dim m_objrs As New ADODB.Recordset
    Me.Top = 1000
    Me.Left = 2000
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from usertbl", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not m_objrs.EOF
        Combo1.AddItem m_objrs!USERID
        m_objrs.MoveNext
    Wend
    Combo1.Text = FRM_VER_INBOUND.ListView1.SelectedItem.SubItems(14)
Set m_objrs = Nothing
End Sub
