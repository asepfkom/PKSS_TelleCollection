VERSION 5.00
Begin VB.Form FRMUNTUK 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5715
      Left            =   15
      TabIndex        =   4
      Top             =   -90
      Width           =   5760
      Begin VB.CommandButton CmdServer5 
         Caption         =   "Server 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4665
         TabIndex        =   10
         Top             =   6165
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton CmdServer4 
         Caption         =   "Server 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4665
         TabIndex        =   9
         Top             =   5745
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Team"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5955
         TabIndex        =   8
         Top             =   3900
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Group 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Spv"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5955
         TabIndex        =   7
         Top             =   4380
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   4350
         TabIndex        =   6
         Top             =   2340
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   5
         Top             =   2715
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Kel&uar"
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
         Left            =   4440
         TabIndex        =   3
         Top             =   1020
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   5430
         ItemData        =   "FRMUNTUK.frx":0000
         Left            =   45
         List            =   "FRMUNTUK.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   0
         Top             =   135
         Width           =   4230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Ambil"
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
         Index           =   0
         Left            =   4440
         TabIndex        =   1
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Semua"
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
         Index           =   1
         Left            =   4440
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

Private Sub CmdServer4_Click()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim i As Integer
    
    List1.CLEAR
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        ' UPDATE 29 MEI 2013 BY IZUDDIN
        CMDSQL = "SELECT TEAM as USERID,SPVNAME as agent FROM SPVTBL WHERE TEAM IN ("
        CMDSQL = CMDSQL + " SELECT TEAM FROM usertbl WHERE USERID='" & MDIForm1.TxtUsername.Text & "')"
    Else
        'CMDSQL = "Select * from usertbl where spvcode ='" + Combo2(0).Text + "'"
        CMDSQL = "select usertbl.agent,usertbl.userid from tbl_ip,usertbl where tbl_ip.ip_addr in "
        CMDSQL = CMDSQL + " (select ip from tbl_ip_icentra where ip_icentra='192.168.10.4') "
        CMDSQL = CMDSQL + " and usertbl.userid=tbl_ip.agent "
    End If
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    
    For i = 1 To M_objrs.RecordCount
        agen = M_objrs("USERID")
        List1.AddItem agen & "!" & IIf(IsNull(M_objrs("AGENT")), "", M_objrs("AGENT"))
        M_objrs.MoveNext
    Next i
    Set M_objrs = Nothing
End Sub

Private Sub CmdServer5_Click()
        Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim i As Integer
    
    List1.CLEAR
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        ' UPDATE 29 MEI 2013 BY IZUDDIN
        CMDSQL = "SELECT TEAM as USERID,SPVNAME as agent FROM SPVTBL WHERE TEAM IN ("
        CMDSQL = CMDSQL + " SELECT TEAM FROM usertbl WHERE USERID='" & MDIForm1.TxtUsername.Text & "')"
    Else
        'CMDSQL = "Select * from usertbl where spvcode ='" + Combo2(0).Text + "'"
        CMDSQL = "select usertbl.agent,usertbl.userid from tbl_ip,usertbl where tbl_ip.ip_addr in "
        CMDSQL = CMDSQL + " (select ip from tbl_ip_icentra where ip_icentra='192.168.10.5') "
        CMDSQL = CMDSQL + " and usertbl.userid=tbl_ip.agent "
    End If
    
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    
    For i = 1 To M_objrs.RecordCount
        agen = M_objrs("USERID")
        List1.AddItem agen & "!" & IIf(IsNull(M_objrs("AGENT")), "", M_objrs("AGENT"))
        M_objrs.MoveNext
    Next i
    Set M_objrs = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim M_objrs As ADODB.Recordset
Dim i As Integer
Select Case Index
Case 0
    If FRMSENDMSG.Text1.Text = Empty Then
         For i = 0 To List1.ListCount - 1
            If List1.Selected(i) Then
                FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text & List1.list(i) & ";"
            End If
         Next i
    Else
         For i = 0 To List1.ListCount - 1
            If List1.Selected(i) Then
                FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text & List1.list(i) & ";"
            End If
         Next i
    End If
    Set M_objrs = Nothing
    Unload Me
Case 1
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        M_objrs.Open "SELECT USERID FROM usertbl WHERE AKTIF = 1 AND SPVCODE ='" + MDIForm1.TxtUsername.Text + "' AND kdlevel =1 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            For i = 1 To M_objrs.RecordCount
                agen = M_objrs("USERID")
                FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text + agen & ";"
                M_objrs.MoveNext
            Next i
        Set M_objrs = Nothing
    Else
        M_objrs.Open "SELECT USERID FROM usertbl WHERE AKTIF = 1 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            For i = 1 To M_objrs.RecordCount
                agen = M_objrs("USERID")
                FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text + agen & ";"
                M_objrs.MoveNext
            Next i
        Set M_objrs = Nothing
    End If
    FRMSENDMSG.Command2.Enabled = False
Unload Me
Case 2
    Dim CMDSQL As String
    List1.CLEAR
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = "Select * from usertbl where spvcode ='" + Combo2(0).Text + "'"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    For i = 1 To M_objrs.RecordCount
        agen = M_objrs("USERID")
        List1.AddItem agen & "!" & IIf(IsNull(M_objrs("AGENT")), "", M_objrs("AGENT"))
        M_objrs.MoveNext
    Next i
    Set M_objrs = Nothing
End Select
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
Dim M_objrs As ADODB.Recordset
Dim i As Integer

    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    'If UCase(mdiform1.txtlevel.text) = "TEAMLEADER" Then
        M_objrs.Open "SELECT USERID FROM usertbl WHERE AKTIF = 0 AND  USERTYPE =20 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            For i = 1 To M_objrs.RecordCount
                agen = M_objrs("USERID")
                FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text + agen & ";"
                M_objrs.MoveNext
            Next i
        Set M_objrs = Nothing
    'Else
     '   m_objrs.Open "SELECT USERID FROM usertbl WHERE AKTIF = 0 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
      '      For i = 1 To m_objrs.RecordCount
       '         agen = m_objrs("USERID")
        '        FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text + agen & ";"
         '       m_objrs.MoveNext
          '  Next i
        'Set m_objrs = Nothing
   ' End If
    FRMSENDMSG.Command2.Enabled = False
Unload Me
End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim i As Integer
    Dim ssql As String
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient

    If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Then
    ssql = " SELECT USERID,KDLEVEL,AGENT FROM usertbl WHERE AKTIF = 1 AND spvcode ='" + MDIForm1.TxtUsername.Text + "'"
    ssql = ssql + " union all (select userid,KDLEVEL,agent from  usertbl where  userid in (select userid from usertbl where kdlevel in('2','5') and aktif='1')order by kdlevel,userid )"
        M_objrs.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

        For i = 1 To M_objrs.RecordCount
           agen = IIf(IsNull(M_objrs("USERID")), "", M_objrs("USERID"))
           List1.AddItem agen & "!" & IIf(IsNull(M_objrs("AGENT")), "", M_objrs("AGENT"))
            M_objrs.MoveNext
        Next i
    Else
        If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then

'            ssql = "SELECT TEAM,SPVNAME FROM SPVTBL WHERE TEAM IN ("
'            ssql = ssql + " SELECT TEAM FROM usertbl WHERE USERID='" & MDIForm1.TxtUsername.Text & "')"
            ssql = "select userid,KDLEVEL,agent from  usertbl where  aktif='1' and userid in (select spvcode from  usertbl where  userid ='" + MDIForm1.TxtUsername.Text + "') "

            'Command1(1).Visible = False
            'M_OBJRS.Open "SELECT SPVCODE,SPVNAME FROM SPVTBL order by SPVCODE", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            M_objrs.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                For i = 1 To M_objrs.RecordCount
                   agen = IIf(IsNull(M_objrs("userid")), "", M_objrs("userid"))
                   List1.AddItem agen & "!" & IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
                    M_objrs.MoveNext
                Next i
                'Command1(1).Visible = False
                Command1(2).Visible = False
                Combo2(0).Visible = False
                Command3.Visible = False
                Group.Visible = False
                Command1(1).Visible = False
        Else

            M_objrs.Open "SELECT USERID,AGENT FROM usertbl WHERE AKTIF = 1 order by KDLEVEL,userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

                For i = 1 To M_objrs.RecordCount
                    agen = IIf(IsNull(M_objrs("USERID")), "", M_objrs("USERID"))
                    List1.AddItem agen & "!" & IIf(IsNull(M_objrs("AGENT")), "", M_objrs("AGENT"))
                    M_objrs.MoveNext
                Next i
        End If
    End If
Set M_objrs = Nothing
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient

M_objrs.Open "Select userid from usertbl where kdlevel in ('2','5') and aktif='1' order by kdlevel='2'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

While Not M_objrs.EOF
    Combo2(0).AddItem M_objrs!USERID
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing
End Sub


Private Sub Group_Click()
Dim M_objrs As ADODB.Recordset
Dim i As Integer

    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    'If UCase(mdiform1.txtlevel.text) = "TEAMLEADER" Then
        M_objrs.Open "SELECT USERID FROM usertbl WHERE AKTIF = 1 AND  USERTYPE =2 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            For i = 1 To M_objrs.RecordCount
                agen = M_objrs("USERID")
                FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text + agen & ";"
                M_objrs.MoveNext
            Next i
        Set M_objrs = Nothing
    'Else
     '   m_objrs.Open "SELECT USERID FROM usertbl WHERE AKTIF = 0 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
      '      For i = 1 To m_objrs.RecordCount
       '         agen = m_objrs("USERID")
        '        FRMSENDMSG.Text1.Text = FRMSENDMSG.Text1.Text + agen & ";"
         '       m_objrs.MoveNext
          '  Next i
        'Set m_objrs = Nothing
   ' End If
    FRMSENDMSG.Command2.Enabled = False
Unload Me
End Sub



