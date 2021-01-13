VERSION 5.00
Begin VB.Form frm_blacklist 
   Caption         =   "Tambah No.Telp BlackList"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   ControlBox      =   0   'False
   Icon            =   "frm_blacklist.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2295
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNotelp 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   4935
   End
   Begin VB.TextBox TxtId 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   1620
      Width           =   1455
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   1620
      Width           =   1455
   End
   Begin VB.TextBox TxtKeterangan 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   6375
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Input Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   540
      TabIndex        =   8
      Top             =   60
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   5
      Left            =   30
      Picture         =   "frm_blacklist.frx":058A
      Stretch         =   -1  'True
      Top             =   30
      Width           =   420
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Id:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Telepon:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   435
      Index           =   8
      Left            =   0
      Picture         =   "frm_blacklist.frx":1094
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20700
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000A&
      Height          =   1800
      Index           =   0
      Left            =   30
      Top             =   450
      Width           =   8340
   End
End
Attribute VB_Name = "frm_blacklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ok As Boolean

Private Sub Cmdbatal_Click()
    ok = False
    Unload Me
End Sub

Private Sub CmdOk_Click()
    Dim VSAVE As Boolean
    
    VSAVE = True
    VSAVE = VSAVE And txtNotelp.Text <> Empty
    
    If VSAVE Then
     If Len(txtNotelp.Text) > 20 Then
      MsgBox "Maksimal jumlah digit no telp:20!", vbInformation + vbOKOnly, "Informasi"
      Exit Sub
     End If
     ok = True
     Me.Hide
     frm_BlackListNo_List.LVBlackList.SetFocus
    Else
      MsgBox "Data Yang Anda Masukan Tidak Lengkap", vbInformation, "Informasi"
    End If
    
End Sub



Private Sub TxtNoTelp_KeyPress(KeyAscii As Integer)
 'Hanya numeric yang dapat diinput
 If KeyAscii < 47 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

