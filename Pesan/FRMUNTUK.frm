VERSION 5.00
Begin VB.Form FRMUNTUK 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5715
      Left            =   15
      TabIndex        =   4
      Top             =   -90
      Width           =   5760
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   4350
         TabIndex        =   6
         Top             =   2340
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Per Team"
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
         Index           =   2
         Left            =   4395
         TabIndex        =   5
         Top             =   2715
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Cancel          =   -1  'True
         Caption         =   "Kel&uar"
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
         Left            =   4395
         TabIndex        =   3
         Top             =   1740
         Width           =   1215
      End
      Begin VB.ListBox List1 
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
         Height          =   5520
         ItemData        =   "FRMUNTUK.frx":0000
         Left            =   45
         List            =   "FRMUNTUK.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   0
         Top             =   135
         Width           =   4230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Ambil"
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
         Left            =   4410
         TabIndex        =   1
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Semua"
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
         Left            =   4410
         TabIndex        =   2
         Top             =   615
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FRMUNTUK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim agen As String * 10

Private Sub Command1_Click(Index As Integer)
Dim m_objrs As ADODB.Recordset
Dim i As Integer
Select Case Index
Case 0
    If FRMSENDMSG.Text1.Text = Empty Then
         For i = 0 To List1.ListCount - 1
            If List1.Selected(i) Then
                FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text & List1.List(i) & ";"
            End If
         Next i
    Else
         For i = 0 To List1.ListCount - 1
            If List1.Selected(i) Then
                FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text & List1.List(i) & ";"
            End If
         Next i
    End If
    Set m_objrs = Nothing
    Unload Me
Case 1
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        m_objrs.Open "SELECT USERID FROM USERTBL WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERTYPE =1 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            For i = 1 To m_objrs.RecordCount
                agen = m_objrs("USERID")
                FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text + agen & ";"
                m_objrs.MoveNext
            Next i
        Set m_objrs = Nothing
    Else
        m_objrs.Open "SELECT USERID FROM USERTBL order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            For i = 1 To m_objrs.RecordCount
                agen = m_objrs("USERID")
                FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text + agen & ";"
                m_objrs.MoveNext
            Next i
        Set m_objrs = Nothing
    End If
    FRMSENDMSG.Command2.Enabled = False
Unload Me
Case 2
    List1.Clear
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from usertbl where spvcode ='" + Combo2(0).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    For i = 1 To m_objrs.RecordCount
        agen = m_objrs("USERID")
        List1.AddItem agen & "!" & IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
        m_objrs.MoveNext
    Next i
    Set m_objrs = Nothing
End Select
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Dim i As Integer
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
    If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        m_objrs.Open "SELECT USERID,AGENT FROM USERTBL order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        For i = 1 To m_objrs.RecordCount
           agen = IIf(IsNull(m_objrs("USERID")), "", m_objrs("USERID"))
           List1.AddItem agen & "!" & IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
            m_objrs.MoveNext
        Next i
    Else
        If UCase(MDIForm1.Text2.Text) = "AGENT" Then
            Command1(1).Visible = False
            m_objrs.Open "SELECT USERID,AGENT FROM USERTBL order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                For i = 1 To m_objrs.RecordCount
                   agen = IIf(IsNull(m_objrs("USERID")), "", m_objrs("USERID"))
                   List1.AddItem agen & "!" & IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
                    m_objrs.MoveNext
                Next i
        Else
            m_objrs.Open "SELECT USERID,AGENT FROM USERTBL order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                For i = 1 To m_objrs.RecordCount
                    agen = IIf(IsNull(m_objrs("USERID")), "", m_objrs("USERID"))
                    List1.AddItem agen & "!" & IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
                    m_objrs.MoveNext
                Next i
        End If
    End If
Set m_objrs = Nothing
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from spvtbl", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    Combo2(0).AddItem m_objrs!SPVCODE
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
End Sub

