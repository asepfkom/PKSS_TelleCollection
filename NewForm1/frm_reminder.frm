VERSION 5.00
Begin VB.Form frm_reminder 
   BackColor       =   &H00FFFFFF&
   Caption         =   "::: R E M I N D E R"
   ClientHeight    =   2640
   ClientLeft      =   8310
   ClientTop       =   4890
   ClientWidth     =   5025
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5025
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   1815
      Left            =   225
      TabIndex        =   2
      Top             =   165
      Width           =   4575
      Begin VB.TextBox txtcustid 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "XXXXXXXXXXX"
         Top             =   240
         Width           =   2010
      End
      Begin VB.TextBox txtnama 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "XXXXXXXXXXX"
         Top             =   600
         Width           =   2010
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Id"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date / Time "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   7
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Automatic close in"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "FOLLOW UP"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2025
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "LATER"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4995
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "frm_reminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private xx As Integer

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    sReminder_CUST_ID = txtcustid.Text
    Unload Me
    If bAktif_form_customer Then
        Unload FrmCC_Colection
    End If
    bReminder_agent = True
    FrmCC_Colection.Show vbModal
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.ScaleWidth) - 500
    Me.Top = (Screen.Height - Me.ScaleHeight) - 700
    'xx = 30
End Sub

'Private Sub Timer1_Timer()
'    xx = xx - 1
'    Label1(8).Caption = xx
'    If xx <= 0 Then Unload Me
'End Sub
