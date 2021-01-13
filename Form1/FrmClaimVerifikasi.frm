VERSION 5.00
Begin VB.Form FrmClaimVerifikasi 
   Caption         =   "Verifikasi Claim"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "FrmClaimVerifikasi.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Verifikasi Void"
      Height          =   510
      Left            =   1845
      TabIndex        =   5
      Top             =   1755
      Width           =   945
   End
   Begin VB.TextBox TxtAgentClaim 
      Height          =   390
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Search Data &Telah Distribusi"
      Height          =   570
      Left            =   2595
      TabIndex        =   3
      Top             =   1065
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search Data &Belum Di Distribusi"
      Height          =   570
      Left            =   720
      TabIndex        =   2
      Top             =   1065
      Width           =   1485
   End
   Begin VB.TextBox TxtId 
      Height          =   390
      Left            =   1635
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox TxtNamaDiKartu 
      Height          =   330
      Left            =   810
      TabIndex        =   0
      Top             =   510
      Width           =   3120
   End
End
Attribute VB_Name = "FrmClaimVerifikasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    FrmClaimSBlmDis.Show
End Sub

Private Sub Command2_Click()
 FrmClaimSTlhDis.Show
End Sub

Private Sub Command3_Click()
Dim cmdsql As String
    cmdsql = "UPDATE ClaimSheet SET "
    cmdsql = cmdsql + " AgentLama ='Data Belum Didistribusi', "
    cmdsql = cmdsql + " KodeStatus ='1', "
    cmdsql = cmdsql + " Keterangan ='Telah Diverifikasi oleh " + MDIForm1.Text1.Text + " ' "
    cmdsql = cmdsql + " where id = " + FrmClaimVerifikasi.TxtId.Text + ""
    M_OBJCONN.Execute cmdsql
    Unload Me
    MsgBox "Proses Selesai", vbInformation + vbOKOnly, "Telegrandi"
    FrmClaimList.ListView1.SelectedItem.SubItems(2) = "Data Milik Orang Lain"
    FrmClaimList.ListView1.SelectedItem.SubItems(5) = "2"
    FrmClaimList.ListView1.SelectedItem.SubItems(6) = "Telah Di Batalkan oleh " & MDIForm1.Text1.Text
End Sub
