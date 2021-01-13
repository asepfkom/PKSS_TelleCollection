VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmClaimList 
   Caption         =   "Claim Sheet"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12675
   LinkTopic       =   "Form2"
   ScaleHeight     =   6945
   ScaleWidth      =   12675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Verifikasi"
      Height          =   375
      Left            =   10995
      TabIndex        =   2
      Top             =   585
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   375
      Left            =   10995
      TabIndex        =   1
      Top             =   180
      Width           =   1320
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6900
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   12171
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
      MousePointer    =   1
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "2= ClaimTolak "
      Height          =   285
      Left            =   11055
      TabIndex        =   5
      Top             =   1605
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1=Claim Di Setujui "
      Height          =   285
      Left            =   11055
      TabIndex        =   4
      Top             =   1320
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "0=Belum Di Verifikasi "
      Height          =   285
      Left            =   11055
      TabIndex        =   3
      Top             =   1020
      Width           =   1545
   End
End
Attribute VB_Name = "FrmClaimList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Id", 10 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Agent Claim", 10 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Agent Lama", 10 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Nama Sesuai DiKartu", 20 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Telp", 10 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Kode Status", 10 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Keterangan", 15 * TXT
End Sub

Private Sub Command1_Click()
    frmClaim.Show vbModal
    If frmClaim.OK Then
        Call showdata
    End If
    Unload frmClaim
End Sub

Private Sub Command2_Click()
If ListView1.ListItems.Count = 0 Then
    Exit Sub
End If
If Trim(ListView1.SelectedItem.SubItems(5)) <> "0" Then
    Exit Sub
End If
    Unload FrmClaimSBlmDis
    
    FrmClaimVerifikasi.TxtId = ListView1.SelectedItem.Text
    FrmClaimVerifikasi.TxtAgentClaim.Text = ListView1.SelectedItem.SubItems(1)
    FrmClaimVerifikasi.TxtNamaDiKartu.Text = ListView1.SelectedItem.SubItems(3)
    FrmClaimVerifikasi.Show
End Sub

Private Sub form_load()
    Call header
    Call showdata
If MDIForm1.Text2.Text <> "Agent" Then
    Command2.Visible = True
Else
    Command2.Visible = False
End If
End Sub


Private Sub showdata()
Dim listitem As listitem
Dim m_claimlist As New ADODB.Recordset
ListView1.ListItems.Clear
m_claimlist.CursorLocation = adUseClient
If MDIForm1.Text2.Text <> "Agent" Then
    m_claimlist.Open "Select * from ClaimSheet where KodeStatus ='0'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    m_claimlist.Open "Select * from ClaimSheet where agentCLAIM ='" + MDIForm1.Text1.Text + "' OR AgentLama ='" + MDIForm1.Text1.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
End If
While Not m_claimlist.EOF
    Set listitem = ListView1.ListItems.ADD(, , m_claimlist("Id"))
        listitem.SubItems(1) = IIf(IsNull(m_claimlist("AgentClaim")), "", m_claimlist("AgentClaim"))
        listitem.SubItems(2) = IIf(IsNull(m_claimlist("AgentLama")), "", m_claimlist("AgentLama"))
        listitem.SubItems(3) = IIf(IsNull(m_claimlist("NamaDiKartu")), "", m_claimlist("NamaDiKartu"))
        listitem.SubItems(4) = IIf(IsNull(m_claimlist("Telp")), "", m_claimlist("Telp"))
        listitem.SubItems(5) = IIf(IsNull(m_claimlist("KodeStatus")), "", m_claimlist("KodeStatus"))
        listitem.SubItems(6) = IIf(IsNull(m_claimlist("KETERANGAN")), "", m_claimlist("KETERANGAN"))
    m_claimlist.MoveNext
Wend
m_claimlist.Close
Set m_claimlist = Nothing
End Sub


