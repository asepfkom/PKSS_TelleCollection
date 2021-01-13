VERSION 5.00
Begin VB.Form FRM_CAPTURE 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Capture"
   ClientHeight    =   10395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13965
   Icon            =   "FRM_CAPTURE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10395
   ScaleWidth      =   13965
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Esc --> Tutup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   30
      Width           =   1650
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   10020
      Left            =   60
      ScaleHeight     =   9960
      ScaleWidth      =   13800
      TabIndex        =   0
      Top             =   330
      Width           =   13860
   End
End
Attribute VB_Name = "FRM_CAPTURE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Dim tempat As String
On Error Resume Next
Picture1.Height = Me.Height
Picture1.Width = Me.Width
Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "SELECT WSTATION FROM USERLOG WHERE AGENT = '" + FRM_LOGIN_LOGOUT.ListView1.SelectedItem.Text + "'AND DATETIME >= '" + Format(MDIForm1.TDBDate1.Text, "mm/dd/yyyy") & "-00:00" + "' and DATETIME <= '" + Format(MDIForm1.TDBDate1.Text, "mm/dd/yyyy") & "-23:59" + "' AND ACTIVITY ='LOGIN'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    FRM_LOGIN_LOGOUT.MousePointer = 11
    FRM_LOGIN_LOGOUT.ListView1.MousePointer = 11
'    WaitSecs (10)
    If m_objrs.RecordCount <> 0 Then
        tempat = "\\" + IIf(IsNull(m_objrs("WSTATION")), "", m_objrs("WSTATION")) + "\C\CAPTURE.BMP"
  '      WaitSecs (10)
        Picture1.Picture = LoadPicture(tempat)
    End If
    FRM_LOGIN_LOGOUT.MousePointer = 0
    FRM_LOGIN_LOGOUT.ListView1.MousePointer = 0
    Set m_objrs = Nothing
End Sub


Function DELETE(filespec)
On Error Resume Next
  Dim fso, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.GetFile(filespec)
  'f.DELETE
End Function

Private Sub Form_Unload(Cancel As Integer)
'Dim m_objrs As ADODB.Recordset
'Dim tempat As String
'Set m_objrs = New ADODB.Recordset
'    m_objrs.CursorLocation = adUseClient
'    m_objrs.Open "SELECT VARVALUE FROM COMMONCFG WHERE VARNAME = '1'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If m_objrs.RecordCount <> 0 Then
'        tempat = IIf(IsNull(m_objrs("VARVALUE")), "", m_objrs("VARVALUE"))
'    End If
'    Set m_objrs = Nothing
'  '  Call DELETE(tempat)
End Sub

