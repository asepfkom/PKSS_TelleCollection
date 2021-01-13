VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmDuplikasiPendingLeads 
   Caption         =   "Pending Closing Duplicate Database"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13845
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   5535
      Left            =   6915
      TabIndex        =   8
      Top             =   615
      Width           =   6855
      Begin MSComctlLib.ListView LstNewLeads 
         Height          =   5325
         Left            =   0
         TabIndex        =   9
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "New Leads....."
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
      Left            =   7035
      TabIndex        =   10
      Top             =   15
      Width           =   6735
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
Attribute VB_Name = "FrmDuplikasiPendingLeads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdSearch_Click()
    Call isi_data(cmbAgentBaru(0), cmbAgentBaru(1))
End Sub

Private Sub Form_Load()
Dim M_OBJCMB As New ADODB.Recordset

Call ExistingLeads
Call newLeads
M_OBJCMB.CursorLocation = adUseClient
M_OBJCMB.Open "Select Userid from usertbl where SPVCODE ='" + MDIForm1.Text1.Text + "' Order by Userid", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_OBJCMB.EOF
    cmbAgentBaru(0).AddItem M_OBJCMB!USERID
    cmbAgentBaru(1).AddItem M_OBJCMB!USERID
    M_OBJCMB.MoveNext
Wend
Set M_OBJCMB = Nothing
End Sub

Private Sub ExistingLeads()
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
    LstExLeads.ColumnHeaders.ADD 13, , "Reason Closing", 11 * TXT
    LstExLeads.ColumnHeaders.ADD 14, , "Agent Baru", 11 * TXT
End Sub

Private Sub newLeads()
    LstNewLeads.ColumnHeaders.ADD 1, , "No", 3 * TXT
    LstNewLeads.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
    LstNewLeads.ColumnHeaders.ADD 3, , "Nama Leads", 10 * TXT
    LstNewLeads.ColumnHeaders.ADD 4, , "Home Phone", 10 * TXT
    LstNewLeads.ColumnHeaders.ADD 5, , "Bussiness Phone", 10 * TXT
    LstNewLeads.ColumnHeaders.ADD 6, , "HandPhone", 10 * TXT
    LstNewLeads.ColumnHeaders.ADD 7, , "DOB", 8 * TXT
    LstNewLeads.ColumnHeaders.ADD 8, , "Agent Code", 10 * TXT
    LstNewLeads.ColumnHeaders.ADD 9, , "Team", 10 * TXT
End Sub

Private Sub isi_data(kriteria1 As String, kriteria2 As String)
Dim listitem As listitem
Dim listitem1 As listitem
Dim m_cek1 As New ADODB.Recordset
Dim cekidNewLeads As String
Dim i As Integer

LstExLeads.ListItems.Clear
LstNewLeads.ListItems.Clear

m_cek1.CursorLocation = adUseClient
m_cek1.Open "Select * from Vw_PendingCloseDuplikasi where AGENTBARU BETWEEN '" + kriteria1 + "' AND '" + kriteria2 + "' AND sts = 0", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
i = 0
While Not m_cek1.EOF
    If m_cek1!CUSTIDBARU <> cekidNewLeads Then
        i = i + 1
    End If
        Set listitem = LstExLeads.ListItems.ADD(, , CStr(i))
                listitem.SubItems(1) = IIf(IsNull(m_cek1("CUSTID")), "", m_cek1("CUSTID"))
                listitem.SubItems(2) = IIf(IsNull(m_cek1("NAME")), "", m_cek1("NAME"))
                listitem.SubItems(3) = IIf(IsNull(m_cek1("HOMENO")), "", m_cek1("HOMENO")) & "-* " & IIf(IsNull(m_cek1("HOMENO2")), "", m_cek1("HOMENO2"))
                listitem.SubItems(4) = IIf(IsNull(m_cek1("OFFICENO")), "", m_cek1("OFFICENO")) & "-* " & IIf(IsNull(m_cek1("OFFICENO2")), "", m_cek1("OFFICENO2"))
                listitem.SubItems(5) = IIf(IsNull(m_cek1("MOBILENO")), "", m_cek1("MOBILENO")) & "-* " & IIf(IsNull(m_cek1("MOBILENO2")), "", m_cek1("MOBILENO2"))
                listitem.SubItems(6) = Format(IIf(IsNull(m_cek1("BIRTHD")), "", m_cek1("BIRTHD")), "yyyy/mm/dd")
                listitem.SubItems(7) = IIf(IsNull(m_cek1("AGENT")), "", m_cek1("AGENT"))
                listitem.SubItems(8) = IIf(IsNull(m_cek1("SPVCODE")), "", m_cek1("SPVCODE"))
                listitem.SubItems(9) = IIf(IsNull(m_cek1("KETHSLKERJA")), "", m_cek1("KETHSLKERJA"))
                listitem.SubItems(10) = Format(IIf(IsNull(m_cek1("TGLSTATUS")), "", m_cek1("TGLSTATUS")), "yyyy/mm/dd")
                listitem.SubItems(11) = IIf(IsNull(m_cek1("CUSTIDBARU")), "", m_cek1("CUSTIDBARU"))
                listitem.SubItems(12) = IIf(IsNull(m_cek1("ReasonClosing")), "", m_cek1("ReasonClosing"))
                listitem.SubItems(13) = IIf(IsNull(m_cek1("AGENTBARU")), "", m_cek1("AGENTBARU"))
                
        If m_cek1!CUSTIDBARU <> cekidNewLeads Then
            cekidNewLeads = m_cek1!CUSTIDBARU
            Set listitem1 = LstNewLeads.ListItems.ADD(, , CStr(i))
                listitem1.SubItems(1) = IIf(IsNull(m_cek1("CUSTIDBARU")), "", m_cek1("CUSTIDBARU"))
                listitem1.SubItems(2) = IIf(IsNull(m_cek1("NAMABARU")), "", m_cek1("NAMABARU"))
                listitem1.SubItems(3) = IIf(IsNull(m_cek1("HOMENOBARU")), "", m_cek1("HOMENOBARU"))
                listitem1.SubItems(4) = IIf(IsNull(m_cek1("OFFICENOBARU")), "", m_cek1("OFFICENOBARU"))
                listitem1.SubItems(5) = IIf(IsNull(m_cek1("MOBILENOBARU")), "", m_cek1("MOBILENOBARU"))
                listitem1.SubItems(6) = Format(IIf(IsNull(m_cek1("DOBBARU")), "", m_cek1("DOBBARU")), "yyyy/mm/dd")
                listitem1.SubItems(7) = IIf(IsNull(m_cek1("AGENTBARU")), "", m_cek1("AGENTBARU"))
                listitem1.SubItems(8) = IIf(IsNull(m_cek1("SPVCODEBARU")), "", m_cek1("SPVCODEBARU"))
        Else
            Set listitem1 = LstNewLeads.ListItems.ADD(, , CStr(i))
                listitem1.SubItems(1) = "------"
                listitem1.SubItems(2) = "------"
                listitem1.SubItems(3) = "------"
                listitem1.SubItems(4) = "------"
                listitem1.SubItems(5) = "------"
                listitem1.SubItems(6) = "------"
                listitem1.SubItems(7) = "------"
                listitem1.SubItems(8) = "------"
        End If
    
    m_cek1.MoveNext
Wend
Set m_cek1 = Nothing
End Sub


Private Sub LstExLeads_DblClick()
M_OBJCONN.BeginTrans
    reff_Duplikasi1 = True
    FRMCUST_CC.TxtReasonClosing = LstExLeads.SelectedItem.SubItems(12)
    FRMCUST_CC.Show vbModal
    If FRMCUST_CC.closeOk = True Then
   '     M_OBJCONN.Execute "update TBL_DUPLIKASI set sts =1 where custidbaru ='" + LstExLeads.SelectedItem.SubItems(11) + "'"
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


