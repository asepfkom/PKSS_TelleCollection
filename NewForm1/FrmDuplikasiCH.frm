VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmDuplikasiCh 
   Caption         =   "Duplikasi Data"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14190
   Icon            =   "FrmDuplikasiCH.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7080
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxNewLeadsId 
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
      Height          =   315
      Left            =   8175
      MaxLength       =   20
      TabIndex        =   13
      Top             =   6525
      Width           =   2325
   End
   Begin VB.ComboBox cmbAgentBaru 
      Height          =   315
      Index           =   0
      Left            =   1035
      TabIndex        =   9
      Top             =   6405
      Width           =   2085
   End
   Begin VB.ComboBox cmbAgentBaru 
      Height          =   315
      Index           =   1
      Left            =   3480
      TabIndex        =   8
      Top             =   6420
      Width           =   2130
   End
   Begin VB.Frame Frame2 
      Height          =   5535
      Left            =   7200
      TabIndex        =   1
      Top             =   600
      Width           =   6855
      Begin MSComctlLib.ListView LstNewLeads 
         Height          =   5325
         Left            =   0
         TabIndex        =   5
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
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6855
      Begin MSComctlLib.ListView LstExLeads 
         Height          =   5325
         Left            =   0
         TabIndex        =   4
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
   Begin Threed.SSCommand CmdReleaseData 
      Height          =   495
      Left            =   12585
      TabIndex        =   6
      Top             =   6375
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
      Caption         =   "Release Data"
   End
   Begin Threed.SSCommand CmdSearch 
      Height          =   390
      Left            =   5685
      TabIndex        =   7
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
   Begin Threed.SSCommand CmdCloseData 
      Height          =   495
      Left            =   10665
      TabIndex        =   12
      Top             =   6435
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
      Caption         =   "Close Data"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "to"
      Height          =   195
      Index           =   1
      Left            =   3195
      TabIndex        =   11
      Top             =   6450
      Width           =   135
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
      Left            =   105
      TabIndex        =   10
      Top             =   6390
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "New CH....."
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
      Left            =   7320
      TabIndex        =   3
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Existing CH....."
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
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "FrmDuplikasiCh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCloseData_Click()
    Dim m_msgbox As Variant
    Dim m_rs As ADODB.Recordset
    Dim CMDSQL As String
    If Len(TxNewLeadsId.Text) < 3 Then
        MsgBox "Pilih New CH Yang Akan Di Close", vbCritical + vbOKOnly, "Telegrandi"
        Exit Sub
    End If
    m_msgbox = MsgBox("Close Data ???..", vbInformation + vbOKCancel, "Telegrandi")
    If m_msgbox = vbCancel Then
        Exit Sub
    End If
    'masukin datanya neh
    Set m_rs = New ADODB.Recordset
    m_rs.CursorLocation = adUseClient
    CMDSQL = "Select * from TBL_DUPLIKASICH WHERE CUSTIDBARU ='" + TxNewLeadsId.Text + "'"
    m_rs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_rs.RecordCount <> 0 Then
        m_rs!sts = 1
        m_rs.UPDATE
        m_rs.Requery
    End If
    MsgBox "Done"
    Set m_rs = Nothing
End Sub

Private Sub CmdReleaseData_Click()
    Dim m_msgbox As Variant
    Dim m_rs As ADODB.Recordset
    Dim CMDSQL As String
    Dim CUSTID1 As String
    If Len(TxNewLeadsId.Text) < 10 Then
        MsgBox "Pilih New CH Yang Akan Di Release", vbCritical + vbOKOnly, "Telegrandi"
        Exit Sub
    End If
    m_msgbox = MsgBox("Release Data ???..", vbInformation + vbOKCancel, "Telegrandi")
    If m_msgbox = vbCancel Then
        Exit Sub
    End If
    'masukin datanya neh
    Set m_rs = New ADODB.Recordset
    m_rs.CursorLocation = adUseClient
    CMDSQL = "Select * from TBL_DUPLIKASICH where CUSTIDBARU ='" + TxNewLeadsId.Text + "'"
    m_rs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_rs.RecordCount <> 0 Then
        CMDSQL = "Insert into CC_CUSTTBL(CUSTID, NAME, HOMENO, MOBILENO, OFFICENO, AGENT, RECSOURCE,CustIdRef,"
        If IsNull(m_rs!BIRTHD) = False Then
            CMDSQL = CMDSQL + "BIRTHD,"
        End If
        CMDSQL = CMDSQL + " RecSourceRef) values"
        CMDSQL = CMDSQL + "('" + CUSTID1 + "',"
        CMDSQL = CMDSQL + "'" + m_rs!NAMABARU + "',"
        CMDSQL = CMDSQL + "'" + m_rs!HOMENOBARU + "',"
        CMDSQL = CMDSQL + "'" + m_rs!MOBILENOBARU + "',"
        CMDSQL = CMDSQL + "'" + m_rs!OFFICENOBARU + "',"
        CMDSQL = CMDSQL + "'" + m_rs!AGENTBARU + "',"
        CMDSQL = CMDSQL + "'" + m_rs!recsourcebaru + "',"
        CMDSQL = CMDSQL + "'" + m_rs!CUSTID + "',"
        If IsNull(m_rs!BIRTHD) = False Then
            CMDSQL = CMDSQL + "'" + Format(m_rs!BIRTHD, "yyyy/mm/dd") + "',"
        End If
        CMDSQL = CMDSQL + "'" + m_rs!recsourcebaru + "')"
        M_OBJCONN.Execute CMDSQL
        m_rs!sts = 1
        m_rs.UPDATE
        m_rs.Requery
    End If
    Set m_rs = Nothing
        MsgBox "Data sudah tersimpan", vbInformation + vbOKOnly, "Telegrandi"
        TxNewLeadsId.Text = Empty
End Sub

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
Dim CMDSQL As String
CMDSQL = "Select * from VWDUPLIKASICH where AGENTBARU BETWEEN '" + kriteria1 + "' AND '" + kriteria2 + "' AND sts = 0"
m_cek1.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
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
    reff_Duplikasi = True
    FRMCUST_CC_MGM.Show vbModal
    If FRMCUST_CC_MGM.closeOk = True Then
'        M_OBJCONN.Execute "update TBL_DUPLIKASICH set sts =1 where custidbaru ='" + LstExLeads.SelectedItem.SubItems(11) + "'"
        MsgBox "Done"
        Call isi_data(cmbAgentBaru(0), cmbAgentBaru(1))
    End If
    Unload FRMCUST_CC
    reff_Duplikasi = False
M_OBJCONN.CommitTrans
Exit Sub
errrdes:
    MsgBox Err.Description
    M_OBJCONN.RollbackTrans
End Sub

Private Sub LstNewLeads_DblClick()
TxNewLeadsId.Text = LstNewLeads.SelectedItem.SubItems(1)
End Sub
