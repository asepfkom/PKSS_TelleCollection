VERSION 5.00
Begin VB.Form frmClaim 
   Caption         =   "Claim Sheet"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5205
   Icon            =   "frmClaim.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   1665
   ScaleWidth      =   5205
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   3780
      TabIndex        =   5
      Top             =   1095
      Width           =   1110
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   360
      Left            =   2595
      TabIndex        =   4
      Top             =   1095
      Width           =   1110
   End
   Begin VB.TextBox TxtNoTelp 
      Height          =   315
      Left            =   1935
      TabIndex        =   3
      Top             =   570
      Width           =   2910
   End
   Begin VB.TextBox TxtNamaKk 
      Height          =   315
      Left            =   1935
      TabIndex        =   2
      Top             =   255
      Width           =   2910
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "No Telephone :"
      Height          =   270
      Left            =   210
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nama Di Kartu Kredit :"
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   1695
   End
End
Attribute VB_Name = "frmClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OK As Boolean

Private Sub Command1_Click()
    OK = True
    If TxtNamaKk.Text = Empty Or TxtNoTelp.Text = Empty Then
        MsgBox "Data Tidak Lengkap", vbCritical + vbOKOnly, "Telegrandi"
        TxtNamaKk.SetFocus
        Exit Sub
    End If
    Me.Hide
    M_OBJCONN.Execute "Insert Into ClaimSheet (NamaDiKartu, Telp, AgentClaim, KodeStatus, Keterangan) values ( '" + TxtNamaKk.Text + "','" + TxtNoTelp.Text + "', '" + MDIForm1.Text1.Text + "' , '0', 'Belum di Verifikasi') "
End Sub

Private Sub Command2_Click()
    OK = False
    Unload Me
End Sub
