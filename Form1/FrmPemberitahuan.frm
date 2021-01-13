VERSION 5.00
Begin VB.Form FrmPemberitahuan 
   Caption         =   "Pesan "
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBukaForm 
      Caption         =   "&Buka Form Request"
      Height          =   435
      Left            =   1080
      TabIndex        =   4
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CheckBox ChkJngnTampil 
      Caption         =   "Jangan Tampilkan lagi saat ini..."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2100
      Width           =   3015
   End
   Begin VB.Timer TimerPemberitahuan 
      Interval        =   15000
      Left            =   4020
      Top             =   660
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Tutup"
      Height          =   435
      Left            =   3240
      TabIndex        =   1
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "(Untuk meng-approve penambahan nomor telepon (additional phone) ada di menu master-> List Request Number Phone."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   60
      TabIndex        =   3
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label LblPemberitahuan 
      Caption         =   "Ada request number telephone!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   4455
   End
End
Attribute VB_Name = "FrmPemberitahuan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdBukaForm_Click()
    FrmListReqTlp.Show vbModal
End Sub

Private Sub cmdOK_Click()
'   If ChkJngnTampil.Value Then
'        MDIForm1.TimerRequest.Enabled = False
'   Else
'        MDIForm1.TimerRequest.Enabled = True
'   End If
    Unload Me
End Sub

Private Sub TimerPemberitahuan_Timer()
    Unload Me
End Sub
