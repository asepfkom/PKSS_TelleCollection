VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FRMTERIMAPOPUP 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5550
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   660
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   3525
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   4620
      Left            =   0
      TabIndex        =   2
      Top             =   405
      Width           =   4500
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   4440
         Left            =   30
         TabIndex        =   1
         Top             =   135
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   7832
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"FRMTERIMAPOPUP.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Dari :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   195
      TabIndex        =   3
      Top             =   135
      Width           =   465
   End
End
Attribute VB_Name = "FRMTERIMAPOPUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  '  SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub
