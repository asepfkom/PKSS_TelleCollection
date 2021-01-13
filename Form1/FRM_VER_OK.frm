VERSION 5.00
Begin VB.Form FRM_VER_OK 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3090
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   1560
      TabIndex        =   7
      Top             =   1905
      Width           =   2100
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1560
      TabIndex        =   6
      Top             =   1455
      Width           =   2100
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   990
      Width           =   2115
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Tutup"
      Height          =   375
      Index           =   0
      Left            =   3015
      TabIndex        =   1
      Top             =   2550
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Verifikasi"
      Height          =   375
      Index           =   1
      Left            =   1500
      TabIndex        =   0
      Top             =   2550
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Agent Lama :"
      Height          =   255
      Left            =   1590
      TabIndex        =   10
      Top             =   585
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Agent Lama :"
      Height          =   255
      Left            =   195
      TabIndex        =   9
      Top             =   600
      Width           =   1290
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Data  Agent Yg Tidak Valid :"
      Height          =   480
      Left            =   135
      TabIndex        =   8
      Top             =   1830
      Width           =   1350
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Data  Upload Yg Dipindahkan :"
      Height          =   480
      Left            =   135
      TabIndex        =   5
      Top             =   1350
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Agent :"
      Height          =   255
      Left            =   195
      TabIndex        =   4
      Top             =   1005
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
Attribute VB_Name = "FRM_VER_OK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click(Index As Integer)
Dim CMDSQL As String
Dim reason As String
Select Case Index
    Case 0
        Unload Me
        
    Case 1
        If Combo1.Text = Empty Then
            MsgBox "Agent Harus DiIsi", vbInformation + vbOKOnly, "Telegrandi"
            Combo1.SetFocus
            Exit Sub
        End If
        M_OBJCONN.Execute "Update RequestInbound set StatusRequest =1 where custid ='" + Label1.Caption + "'"
        M_OBJCONN.Execute "Delete From tempCC_CUSTTBL where custid ='" + Text1.Text + "'"
        M_OBJCONN.Execute "Delete From CC_CUSTTBL where custid ='" + Text2.Text + "'"
        If Trim(UCase(Combo1.Text)) = Trim(UCase(FRM_VER_INBOUND.ListView1.SelectedItem.SubItems(2))) Then
            reason = "Data Inbound Yang Anda Masukan Telah di Cek Oleh " & MDIForm1.Text1.Text
        Else
            reason = "Data Telah di Pindahkan dari " & FRM_VER_INBOUND.ListView1.SelectedItem.SubItems(2) & " Ke " & Combo1.Text & " Oleh " & MDIForm1.Text1.Text
        End If
        CMDSQL = "Insert Into RequestInboundRst "
        CMDSQL = CMDSQL + " (custid, "
        CMDSQL = CMDSQL + " AgentLama, "
        CMDSQL = CMDSQL + " AgentBaru, "
        CMDSQL = CMDSQL + " Reason) "
        CMDSQL = CMDSQL + " Values "
        CMDSQL = CMDSQL + " ('" + Label1.Caption + "' , "
        CMDSQL = CMDSQL + " '" + FRM_VER_INBOUND.ListView1.SelectedItem.SubItems(14) + "', "
        CMDSQL = CMDSQL + " '" + Combo1.Text + "', "
        CMDSQL = CMDSQL + " '" + reason + "') "
        M_OBJCONN.Execute CMDSQL
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
Set m_objrs = Nothing
End Sub
