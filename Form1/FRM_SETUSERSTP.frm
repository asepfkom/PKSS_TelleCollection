VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRM_SETUSERSTP 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "FRM_SETUSERSTP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   1335
      Visible         =   0   'False
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   3810
      TabIndex        =   3
      Top             =   825
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Tampilkan Agen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2370
      TabIndex        =   2
      Top             =   825
      Width           =   1395
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   1
      Left            =   2175
      TabIndex        =   1
      Top             =   285
      Width           =   3060
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   0
      Left            =   1290
      TabIndex        =   0
      Top             =   285
      Width           =   900
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "Supervisor :"
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
      Left            =   150
      TabIndex        =   4
      Top             =   315
      Width           =   1260
   End
End
Attribute VB_Name = "FRM_SETUSERSTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim sSearchText As String
Dim lReturn As Long
Select Case Index
Case 0, 1
If KeyAscii = 13 Then
   Combo1_Click (Index)
   KeyAscii = 0
Else
   sSearchText = Left$(Combo1(Index).Text, Combo1(Index).SelStart) & Chr$(KeyAscii)
   lReturn = SendMessage(Combo1(Index).hWnd, CB_FINDSTRING, -1, ByVal sSearchText)
   If lReturn <> CB_ERR Then
      mbIgnoreListClick = True
      Combo1(Index).ListIndex = lReturn
      mbIgnoreListClick = False
      Combo1(Index).Text = Combo1(Index).List(lReturn)
      Combo1(Index).SelStart = Len(sSearchText)
      Combo1(Index).SelLength = Len(Combo1(Index).Text)
      KeyAscii = 0
   End If
End If
End Select
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim M_DATA As New CLS_DISTRIBUSISTP
Dim m_objrs As ADODB.Recordset
Select Case Index
    Case 0
        Set m_objrs = M_DATA.QUERY_SPV(M_OBJCONN, "SPVCODE = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("SPVCODE")
            Combo1(1).Text = m_objrs("SPVNAME")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 1
        Set m_objrs = M_DATA.QUERY_SPV(M_OBJCONN, "SPVNAME = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("SPVCODE")
            Combo1(1).Text = m_objrs("SPVNAME")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    End Select
Set m_objrs = Nothing
Set M_DATA = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim M_DATA As New CLS_DISTRIBUSISTP
Select Case Index
    Case 0
        If Combo1(0).Text = Empty Then
            MsgBox "Pilih TeamLeader..!!!", vbInformation + vbOKOnly, "Informasi"
        Else
            M_DATA.INSERT_DISTRIBUSI M_RPTCONN, M_OBJCONN, Combo1(0).Text, MDIForm1.TDBDate1.Text
            FRM_DISTRIBUTESTP.Show vbModal
        End If
    Case 1
        Unload Me
End Select
Set M_DATA = Nothing
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New CLS_DISTRIBUSISTP
    Set m_objrs = M_DATA.QUERY_SPV(M_OBJCONN, "")
        While Not m_objrs.EOF
            Combo1(0).AddItem m_objrs("SPVCODE")
            Combo1(0).DataField = m_objrs("SPVCODE")
            Combo1(1).AddItem m_objrs("SPVNAME")
            Combo1(1).DataField = m_objrs("SPVNAME")
            m_objrs.MoveNext
        Wend
    Set m_objrs = Nothing
    Set M_DATA = Nothing
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim M_DATA As New CLS_DISTRIBUSISTP
Dim m_objrs As ADODB.Recordset
Select Case Index
    Case 0
        Set m_objrs = M_DATA.QUERY_SPV(M_OBJCONN, "SPVCODE = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("SPVCODE")
            Combo1(1).Text = m_objrs("SPVNAME")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 1
        Set m_objrs = M_DATA.QUERY_SPV(M_OBJCONN, "SPVNAME = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("SPVCODE")
            Combo1(1).Text = m_objrs("SPVNAME")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    End Select
Set m_objrs = Nothing
Set M_DATA = Nothing
End Sub
