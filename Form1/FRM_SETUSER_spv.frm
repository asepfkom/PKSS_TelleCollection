VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRM_SETUSER_spv 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
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
Attribute VB_Name = "FRM_SETUSER_spv"
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
'Dim M_DATA As New CLS_DISTRIBUSI
Dim m_objrs As ADODB.Recordset
Select Case Index
    Case 0
        Set m_objrs = QUERY_SPV(M_OBJCONN, "SPVCODE = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("SPVCODE")
            Combo1(1).Text = m_objrs("SPVNAME")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 1
        Set m_objrs = QUERY_SPV(M_OBJCONN, "SPVNAME = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("SPVCODE")
            Combo1(1).Text = m_objrs("SPVNAME")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    End Select
Set m_objrs = Nothing
'Set M_DATA = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
'Dim M_DATA As New CLS_DISTRIBUSI
Select Case Index
    Case 0
        If Combo1(0).Text = Empty Then
            MsgBox "Pilih TeamLeader..!!!", vbInformation + vbOKOnly, "Informasi"
        Else
            INSERT_DISTRIBUSI M_RPTCONN, M_OBJCONN, Combo1(0).Text, MDIForm1.TDBDate1.Text
            FRM_distribute_spv1.Show vbModal
        End If
    Case 1
        Unload Me
End Select
'Set M_DATA = Nothing
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
'Dim M_DATA As New CLS_DISTRIBUSI
    Set m_objrs = QUERY_SPV(M_OBJCONN, "")
        While Not m_objrs.EOF
            Combo1(0).AddItem m_objrs("spvcode")
            Combo1(0).DataField = m_objrs("spvcode")
            Combo1(1).AddItem m_objrs("spvname")
            Combo1(1).DataField = m_objrs("spvname")
            m_objrs.MoveNext
        Wend
    Set m_objrs = Nothing
    'Set M_DATA = Nothing
End Sub

Private Sub Combo1_Click(Index As Integer)
'Dim M_DATA As New CLS_DISTRIBUSI
Dim m_objrs As ADODB.Recordset
Select Case Index
    Case 0
        Set m_objrs = QUERY_SPV(M_OBJCONN, "SPVCODE = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("spvcode")
            Combo1(1).Text = m_objrs("spvname")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 1
        Set m_objrs = QUERY_SPV(M_OBJCONN, "agent = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("spvcode")
            Combo1(1).Text = m_objrs("spvname")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    End Select
Set m_objrs = Nothing
'Set M_DATA = Nothing
End Sub

Public Function QUERY_SPV(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim cmdsql As String
Dim m_objrs As ADODB.Recordset

cmdsql = "SELECT * FROM spvtbl"
cmdsql = cmdsql + " WHERE SPVCODE = '" + MDIForm1.Text1.Text + "'"
 If Len(M_WHERE) <> 0 Then
    cmdsql = cmdsql + " AND " + M_WHERE
 End If
cmdsql = cmdsql + " ORDER BY spvcode"
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SPV = m_objrs
Set m_objrs = Nothing
End Function

Public Function INSERT_DISTRIBUSI(M_RPTCONN As ADODB.Connection, M_OBJCONN As ADODB.Connection, SPVCODE As String, TANGGAL As String)
Dim cmdsql As String
Dim USERID As String
Dim Nama As String
Dim TGLJAM1 As String
Dim JAM As String
Dim TGLJAM2 As String
Dim i As Integer
Dim m_objrs As ADODB.Recordset

Call DELETE_DISTRIBUSI(M_RPTCONN)

Set m_objrs = QUERY_USER(M_OBJCONN, SPVCODE)
If m_objrs.RecordCount = 0 Then
    FRM_SETUSER_spv.ProgressBar1.Max = 100
Else
    FRM_SETUSER_spv.ProgressBar1.Max = 100 * (m_objrs.RecordCount + 1)
    
End If
    FRM_SETUSER_spv.ProgressBar1.Visible = True
    FRM_SETUSER_spv.ProgressBar1.Value = 100
i = 100

TGLJAM2 = Format(TANGGAL, "mm/dd/yy")
JAM = Format(TGLJAM2, "mm/dd/yy") + " " + Format(Now, "hh:mm")
TGLJAM1 = Format(TGLJAM2, "yyyymmdd") + Format(Now, "hhmm")
While Not m_objrs.EOF
    FRM_SETUSER_spv.ProgressBar1.Value = i
    USERID = IIf(IsNull(m_objrs("USERID")), "", m_objrs("USERID"))
    Nama = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
    cmdsql = "INSERT INTO DISTRIBUSI"
    cmdsql = cmdsql + " (USERID,"
    cmdsql = cmdsql + " TGLJAM,"
    cmdsql = cmdsql + " NAMA)"
    cmdsql = cmdsql + " VALUES"
    cmdsql = cmdsql + " ('" + Trim(USERID) + "',"
    cmdsql = cmdsql + " '" + LTrim(TGLJAM1) + "',"
    cmdsql = cmdsql + " '" + Trim(Nama) + "')"
    M_RPTCONN.Execute cmdsql
    m_objrs.MoveNext
    i = i + 100
Wend
    FRM_SETUSER_spv.ProgressBar1.Value = FRM_SETUSER_spv.ProgressBar1.Max
    FRM_SETUSER_spv.ProgressBar1.Visible = False
End Function

Private Function DELETE_DISTRIBUSI(M_RPTCONN As ADODB.Connection)
Dim cmdsql As String
    cmdsql = "DELETE * FROM DISTRIBUSI"
    M_RPTCONN.Execute cmdsql
End Function

Public Function QUERY_USER(M_OBJCONN As ADODB.Connection, SPVCODE As String) As Object
Dim cmdsql As String
Dim m_objrs As ADODB.Recordset

cmdsql = "SELECT * FROM USERTBL"
cmdsql = cmdsql + " WHERE USERTYPE ='1'"
 If Len(SPVCODE) <> 0 Then
    cmdsql = cmdsql + " AND SPVCODE = '" + SPVCODE + "'"
 End If
cmdsql = cmdsql + " AND AKTIF = '0'"
cmdsql = cmdsql + " ORDER BY USERID"
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_USER = m_objrs
Set m_objrs = Nothing
End Function
