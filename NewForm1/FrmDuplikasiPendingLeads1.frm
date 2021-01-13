VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmDuplikasiPendingLeads1 
   Caption         =   "Pending Closing Duplicate Database"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13140
   LinkTopic       =   "Form2"
   ScaleHeight     =   7095
   ScaleWidth      =   13140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   15
      TabIndex        =   2
      Top             =   600
      Width           =   6855
      Begin MSComctlLib.ListView LstExLeads 
         Height          =   5325
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   9393
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.ComboBox cmbAgentBaru 
      Height          =   315
      Index           =   1
      Left            =   3375
      TabIndex        =   1
      Top             =   6420
      Width           =   2130
   End
   Begin VB.ComboBox cmbAgentBaru 
      Height          =   315
      Index           =   0
      Left            =   930
      TabIndex        =   0
      Top             =   6405
      Width           =   2085
   End
   Begin Threed.SSCommand CmdSearch 
      Height          =   390
      Left            =   5580
      TabIndex        =   4
      Top             =   6375
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   688
      _Version        =   196610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Script MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Searching.."
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Existing Leads....."
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Agent :"
      BeginProperty Font 
         Name            =   "Vladimir Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   6390
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "to"
      Height          =   195
      Index           =   1
      Left            =   3090
      TabIndex        =   5
      Top             =   6450
      Width           =   135
   End
End
Attribute VB_Name = "FrmDuplikasiPendingLeads1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub isi_data(kriteria1 As String, kriteria2 As String)
Dim listitem As listitem
Dim m_rs As New ADODB.Recordset
LstExLeads.ListItems.Clear
m_rs.CursorLocation = adUseClient
m_rs.Open "Select * from Vw_PendingCloseDuplikasi where agent between '" + kriteria1 + "' and  '" + kriteria2 + "' ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_rs.EOF
    Set listitem = LstExLeads.ListItems.ADD(, , CStr(i))
        listitem.SubItems(1) = IIf(IsNull(m_rs("CUSTID")), "", m_rs("CUSTID"))
        listitem.SubItems(2) = IIf(IsNull(m_rs("NAME")), "", m_rs("NAME"))
        listitem.SubItems(3) = IIf(IsNull(m_rs("HOMENO")), "", m_rs("HOMENO")) & "-* " & IIf(IsNull(m_rs("HOMENO2")), "", m_rs("HOMENO2"))
        listitem.SubItems(4) = IIf(IsNull(m_rs("OFFICENO")), "", m_rs("OFFICENO")) & "-* " & IIf(IsNull(m_rs("OFFICENO2")), "", m_rs("OFFICENO2"))
        listitem.SubItems(5) = IIf(IsNull(m_rs("MOBILENO")), "", m_rs("MOBILENO")) & "-* " & IIf(IsNull(m_rs("MOBILENO2")), "", m_rs("MOBILENO2"))
        listitem.SubItems(6) = Format(IIf(IsNull(m_rs("BIRTHD")), "", m_rs("BIRTHD")), "yyyy/mm/dd")
        listitem.SubItems(7) = IIf(IsNull(m_rs("AGENT")), "", m_rs("AGENT"))
        listitem.SubItems(8) = IIf(IsNull(m_rs("SPVCODE")), "", m_rs("SPVCODE"))
        listitem.SubItems(9) = IIf(IsNull(m_rs("KETHSLKERJA")), "", m_rs("KETHSLKERJA"))
        listitem.SubItems(10) = Format(IIf(IsNull(m_rs("TGLSTATUS")), "", m_rs("TGLSTATUS")), "yyyy/mm/dd")
        listitem.SubItems(11) = IIf(IsNull(m_rs("CUSTIDBARU")), "", m_rs("CUSTIDBARU"))
    m_rs.MoveNext
Wend
Set m_rs = Nothing
End Sub

Private Sub Form_Load()
Dim M_OBJCMB As New ADODB.Recordset
    Call header
M_OBJCMB.CursorLocation = adUseClient
M_OBJCMB.Open "Select Userid from usertbl where SPVCODE ='" + MDIForm1.Text1.Text + "' Order by Userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_OBJCMB.EOF
    cmbAgentBaru(0).AddItem M_OBJCMB!USERID
    cmbAgentBaru(1).AddItem M_OBJCMB!USERID
    M_OBJCMB.MoveNext
Wend
Set M_OBJCMB = Nothing
End Sub

Private Sub header()
    LstExLeads.ColumnHeaders.ADD 1, , "No", 3 * TXT
    LstExLeads.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
    LstExLeads.ColumnHeaders.ADD 3, , "Nama Leads", 10 * TXT
    LstExLeads.ColumnHeaders.ADD 4, , "Home Phone", 10 * TXT
    LstExLeads.ColumnHeaders.ADD 5, , "Bussiness Phone", 10 * TXT
    LstExLeads.ColumnHeaders.ADD 6, , "Handphone", 10 * TXT
    LstExLeads.ColumnHeaders.ADD 7, , "DOB", 10 * TXT
    LstExLeads.ColumnHeaders.ADD 8, , "AgentCode", 8 * TXT
    LstExLeads.ColumnHeaders.ADD 9, , "Team", 10 * TXT
    LstExLeads.ColumnHeaders.ADD 10, , "Status LastCall", 10 * TXT
    LstExLeads.ColumnHeaders.ADD 11, , "LastCall Date", 10 * TXT
    LstExLeads.ColumnHeaders.ADD 12, , "CustIDBaru", 10 * TXT
End Sub

Private Sub LstExLeads_BeforeLabelEdit(Cancel As Integer)
M_OBJCONN.BeginTrans
    reff_Duplikasi1 = True
    FRMCUST_CC.Show vbModal
    If FRMCUST_CC.closeOk = True Then
        M_OBJCONN.Execute "update Vw_PendingCloseDuplikasi set sts =1 where custidbaru ='" + LstExLeads.SelectedItem.SubItems(11) + "'"
        MsgBox "Done"
        Call isi_data(cmbAgentBaru(0), cmbAgentBaru(1))
    End If
    Unload FRMCUST_CC
    reff_Duplikasi1 = False
M_OBJCONN.CommitTrans
Exit Sub
errrdes:
    MsgBox Err.Description
    M_OBJCONN.RollbackTrans
End Sub
