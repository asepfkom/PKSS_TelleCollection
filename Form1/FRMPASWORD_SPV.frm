VERSION 5.00
Begin VB.Form FRMPASWORD_SPV 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Masukan  Password Supervisor"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FRMPASWORD_SPV.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPassword 
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1515
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   555
      Width           =   2340
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   2445
      TabIndex        =   3
      Top             =   1065
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ok"
      Default         =   -1  'True
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
      Left            =   765
      TabIndex        =   2
      Top             =   1065
      Width           =   1140
   End
   Begin VB.TextBox txtUserName 
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
      Height          =   345
      Left            =   1530
      TabIndex        =   0
      Top             =   150
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Password:"
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
      Height          =   270
      Index           =   1
      Left            =   390
      TabIndex        =   5
      Top             =   570
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "&User Name:"
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
      Height          =   270
      Index           =   0
      Left            =   405
      TabIndex        =   4
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "FRMPASWORD_SPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TOLAK_OK As Boolean

Private Sub cmdCancel_Click()
    TOLAK_OK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim m_objrs As ADODB.Recordset
    If txtUserName = Empty Then
        MsgBox "Username Belum Diisi", vbCritical + vbOKOnly, "Peringatan"
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    Else
        If txtPassword = Empty Then
            MsgBox "Password Belum Diisi", vbCritical + vbOKOnly, "Peringatan"
            txtPassword.SetFocus
            SendKeys "{Home}+{End}"
            Exit Sub
        End If
    End If
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "SELECT USERID, ACCREC FROM USERTBL WHERE USERID = '" + txtUserName + "' AND USERTYPE <> 1 ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_objrs.RecordCount <> 0 Then
    If txtPassword <> m_objrs("ACCREC") Then
        MsgBox "Password Yang Anda Masukan Salah", vbCritical + vbOKOnly, "Peringatan"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        TOLAK_OK = False
    Else
        TOLAK_OK = True
        Unload Me
    End If
Else
    MsgBox "User Name Yang Anda Masukan Tidak Terdaftar", vbCritical + vbOKOnly, "Peringatan"
    txtUserName.SetFocus
    SendKeys "{Home}+{End}"
End If
    Set m_objrs = Nothing
End Sub

