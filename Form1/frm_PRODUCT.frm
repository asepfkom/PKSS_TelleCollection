VERSION 5.00
Begin VB.Form frm_PRODUCT 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1605
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000004&
      Caption         =   "Non Aktif"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Top             =   885
      Width           =   1290
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000004&
      Caption         =   "Aktif"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   885
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   6315
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1065
      UseMaskColor    =   -1  'True
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1065
      UseMaskColor    =   -1  'True
      Width           =   825
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1545
      MaxLength       =   50
      TabIndex        =   1
      Top             =   540
      Width           =   5925
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1545
      MaxLength       =   12
      TabIndex        =   0
      Top             =   225
      Width           =   2370
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1545
      TabIndex        =   10
      Top             =   855
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   4
      Left            =   75
      TabIndex        =   8
      Top             =   900
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   45
      TabIndex        =   7
      Top             =   570
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Kode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   90
      TabIndex        =   6
      Top             =   240
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Unit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   1
      Left            =   165
      TabIndex        =   9
      Top             =   885
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "frm_PRODUCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ok As Boolean

Private Sub Combo1_LostFocus()
Select Case UCase(Combo1.Text)
    Case "KTA"
    Case "CREDIT CARD"
    Case "KTA - CROSS SELL"
    Case "CC - CROSS SELL"
    Case Else
        Combo1.Text = Empty
End Select
End Sub

Private Sub Command1_Click(Index As Integer)
Dim VSAVE As Boolean
VSAVE = True
Select Case Index
    Case 0
        VSAVE = VSAVE And Text1.Text <> Empty
        VSAVE = VSAVE And Text2.Text <> Empty
        If VSAVE Then
            ok = True
            Me.Hide
            FRM_PRODUCT_LIST.ListView1.SetFocus
        Else
            MsgBox "Data Yang Anda Masukan Tidak Lengkap", vbInformation, "Informasi"
        End If
    Case 1
        ok = False
        Unload Me
        FRM_PRODUCT_LIST.ListView1.SetFocus
End Select
End Sub

Private Sub Form_Load()
   Combo1.AddItem "KTA", 0
   Combo1.AddItem "CREDIT CARD", 1
   Combo1.AddItem "KTA - CROSS SELL", 2
   Combo1.AddItem "CC - CROSS SELL", 3
   Combo1.Text = "CREDIT CARD"
End Sub

