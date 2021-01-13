VERSION 5.00
Begin VB.Form View_AlihData 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pindah Data"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5580
   ControlBox      =   0   'False
   Icon            =   "View_AlihData.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNote 
      Height          =   1155
      Left            =   900
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   495
      Width           =   4410
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   345
      Left            =   4395
      TabIndex        =   4
      Top             =   1755
      Width           =   915
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process"
      Height          =   345
      Left            =   3375
      TabIndex        =   3
      Top             =   1755
      Width           =   915
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   1
      Left            =   2310
      TabIndex        =   1
      Top             =   165
      Width           =   3000
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   0
      Left            =   915
      TabIndex        =   0
      Top             =   165
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Note :"
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
      Height          =   240
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   525
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name :"
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
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   210
      Width           =   675
   End
End
Attribute VB_Name = "View_AlihData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ok As Boolean

Private Sub cmdExit_Click()
    ok = False
    Me.Hide
End Sub

Private Sub Alih_SdhDistribusi()
On Error GoTo aerr
   ' M_OBJCONN.BeginTrans
    M_OBJCONN.Execute "Update mandiri.mgm set Agent ='" + Combo2(0).Text + "', NamaAgent='" + Combo2(1).Text + "', nextact = 'Data Pindahan' where CUSTID = '" + VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) + "'"
    TxtNote.Text = "Data Baru - " & TxtNote.Text
   ' M_OBJCONN.Execute "Insert into TrMutasiData (Custid, Keterangan, AgentLama, AgentBaru, AlihOleh) values ('" + VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(1) + "', '" + TxtNote.Text + "','" + VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(7) + "','" + Combo2(0).Text + "', '" + mdiform1.txtusername.text + "') "
    M_OBJCONN.CommitTrans
    ok = True
    MsgBox "Done"
    Me.Hide
    Exit Sub
aerr:
   ' M_OBJCONN.RollbackTrans
    MsgBox err.Description
    ok = False
End Sub

Private Sub cmdProcess_Click()
Dim m_msgbox As Variant
m_msgbox = MsgBox("Yakin Akan dilakukan ??", vbYesNo + vbExclamation, "Aplikasi")
If m_msgbox = vbNo Then
    Exit Sub
End If
'If VIEW_mgmDATA.F_DISTRIBUSI = "SDHDISTRIBUSI" Then
    Call Alih_SdhDistribusi
'Else
  '  Call Alih_BlmDistribusi
'End If

End Sub

Private Sub Alih_BlmDistribusi()
Dim m_objDt As New ADODB.Recordset
Dim m_objDt1 As ADODB.Recordset
Dim Lcustid As String
Dim LName  As String
Dim LAHOMENO  As String
Dim LHOMENO  As String
Dim LMOBILENO  As String
Dim LAFAXNO  As String
Dim LFAXNO  As String
Dim LAOFFICENO  As String
Dim LOFFICENO  As String
Dim LEXTOFFICE  As String
Dim LAgent  As String
Dim LRECSOURCE  As String
Dim LOTHERS  As String
Dim LNAMAAGENT  As String
Dim LNEXTACT  As String
Dim CMDSQL As String
On Error GoTo aerr
    M_OBJCONN.BeginTrans
    m_objDt.CursorLocation = adUseClient
    m_objDt.Open "Select * from mandiri.tempCC_CUSTTBL where custid = '" + VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_objDt.RecordCount <> 0 Then
        Lcustid = IIf(IsNull(m_objDt!CustId), "", m_objDt!CustId)
        LName = IIf(IsNull(m_objDt!Name), "", m_objDt!Name)
        LAHOMENO = IIf(IsNull(m_objDt!AHOMENO), "", m_objDt!AHOMENO)
        LHOMENO = IIf(IsNull(m_objDt!HOMENO), "", m_objDt!HOMENO)
        LMOBILENO = IIf(IsNull(m_objDt!MOBILENO), "", m_objDt!MOBILENO)
        LAFAXNO = IIf(IsNull(m_objDt!AFAXNO), "", m_objDt!AFAXNO)
        LFAXNO = IIf(IsNull(m_objDt!FAXNO), "", m_objDt!FAXNO)
        LAOFFICENO = IIf(IsNull(m_objDt!AOFFICENO), "", m_objDt!AOFFICENO)
        LOFFICENO = IIf(IsNull(m_objDt!OFFICENO), "", m_objDt!OFFICENO)
        LEXTOFFICE = IIf(IsNull(m_objDt!extoffice), "", m_objDt!extoffice)
        LAgent = Combo2(0).Text
        LRECSOURCE = IIf(IsNull(m_objDt!RECSOURCE), "", m_objDt!RECSOURCE)
        LOTHERS = IIf(IsNull(m_objDt!OTHERS), "", m_objDt!OTHERS)
        LNAMAAGENT = Combo2(1).Text
        LNEXTACT = "DATA BARU CLAIM INBOUND"
        CMDSQL = " Insert Into mandiri.XSELLBANK "
        CMDSQL = CMDSQL + " (CUSTID, NAME, AHOMENO, HOMENO, MOBILENO, AFAXNO, FAXNO, AOFFICENO, OFFICENO, EXTOFFICE, Agent, RECSOURCE, OTHERS, NAMAAGENT, NEXTACT)"
        CMDSQL = CMDSQL + " VALUES"
        CMDSQL = CMDSQL + " ('" + Lcustid + "',"
        CMDSQL = CMDSQL + " '" + LName + "',"
        CMDSQL = CMDSQL + " '" + LAHOMENO + "',"
        CMDSQL = CMDSQL + " '" + LHOMENO + "',"
        CMDSQL = CMDSQL + " '" + LMOBILENO + "',"
        CMDSQL = CMDSQL + " '" + LAFAXNO + "',"
        CMDSQL = CMDSQL + " '" + LFAXNO + "',"
        CMDSQL = CMDSQL + " '" + LAOFFICENO + "',"
        CMDSQL = CMDSQL + " '" + LOFFICENO + "',"
        CMDSQL = CMDSQL + " '" + LEXTOFFICE + "',"
        CMDSQL = CMDSQL + " '" + LAgent + "',"
        CMDSQL = CMDSQL + " '" + LRECSOURCE + "',"
        CMDSQL = CMDSQL + " '" + LOTHERS + "',"
        CMDSQL = CMDSQL + " '" + LNAMAAGENT + "',"
        CMDSQL = CMDSQL + " '" + LNEXTACT + "')"
    M_OBJCONN.Execute CMDSQL
    TxtNote.Text = "Data Baru - " & TxtNote.Text
    M_OBJCONN.Execute "Insert into mandiri.TrMutasiData (Custid, Keterangan, AgentLama, AgentBaru, AlihOleh) values ('" + VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.Text + "', '" + TxtNote.Text + "','" + VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(7) + "','" + Combo2(0).Text + "', '" + MDIForm1.TxtUsername.Text + "') "
    M_OBJCONN.Execute " tempCC_CUSTTBL where custid = '" + VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) + "'"
    End If
    Set m_objDt = Nothing
    M_OBJCONN.CommitTrans
    ok = True
    MsgBox "Done"
    Me.Hide
    Exit Sub
aerr:
    M_OBJCONN.RollbackTrans
    MsgBox err.Description
    ok = False
End Sub


Private Sub Combo2_Click(Index As Integer)
    Call Combo2_LostFocus(Index)
End Sub

Private Sub Combo2_LostFocus(Index As Integer)
Dim m_combo As New ADODB.Recordset
Select Case Index
Case 0
    m_combo.CursorLocation = adUseClient
    m_combo.Open "Select USERID, agent from mandiri.usertbl where USERID ='" + Combo2(0).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_combo.RecordCount = 1 Then
        Combo2(0).Text = m_combo!USERID
        Combo2(1).Text = m_combo!AGENT
    Else
        Combo2(0).Text = ""
        Combo2(1).Text = ""
    End If
    Set m_combo = Nothing
Case 1
    m_combo.CursorLocation = adUseClient
    m_combo.Open "Select USERID, agent from mandiri.usertbl where AGENT ='" + Combo2(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_combo.RecordCount = 1 Then
        Combo2(0).Text = m_combo!USERID
        Combo2(1).Text = m_combo!AGENT
    Else
        Combo2(0).Text = ""
        Combo2(1).Text = ""
    End If
    Set m_combo = Nothing
End Select
End Sub

Private Sub Form_Load()
Dim m_combo As New ADODB.Recordset
m_combo.CursorLocation = adUseClient
'm_combo.Open "Select * from usertbl where usertype =1 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Then
    m_combo.Open "Select * from mandiri.usertbl where usertype =1 AND TEAM ='" + MDIForm1.TxtUsername.Text + "' order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ElseIf Left(UCase(MDIForm1.txtlevel.Text), 5) = "ADMIN" Then
    m_combo.Open "Select * from mandiri.usertbl where USERTYPE =1 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
End If
While Not m_combo.EOF
    Combo2(0).AddItem m_combo!USERID
    Combo2(1).AddItem m_combo!AGENT
    m_combo.MoveNext
Wend
Set m_combo = Nothing
End Sub
