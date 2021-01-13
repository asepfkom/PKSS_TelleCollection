VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmKurirSpvApp 
   Caption         =   "Kurir..."
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   Icon            =   "FrmKurirSpvApp.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1290
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport RPT 
      Left            =   3420
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.OptionButton OptionStatus 
      Caption         =   "Kirim"
      Height          =   225
      Left            =   1035
      TabIndex        =   4
      Top             =   30
      Width           =   2130
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   960
      TabIndex        =   1
      Top             =   315
      Width           =   3720
   End
   Begin Threed.SSCommand CmdShow 
      Height          =   390
      Left            =   2985
      TabIndex        =   2
      Top             =   780
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   688
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Show"
   End
   Begin Threed.SSCommand CmdCancel 
      Height          =   390
      Left            =   3735
      TabIndex        =   3
      Top             =   780
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   688
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Cancel"
   End
   Begin Threed.SSCommand CmdApproved 
      Height          =   390
      Left            =   285
      TabIndex        =   5
      Top             =   765
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   688
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Approved"
   End
   Begin Threed.SSCommand CmdPrint 
      Height          =   390
      Left            =   1575
      TabIndex        =   6
      Top             =   780
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   688
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Print"
   End
   Begin VB.Label Label1 
      Caption         =   "No. POD :"
      Height          =   360
      Left            =   105
      TabIndex        =   0
      Top             =   375
      Width           =   855
   End
End
Attribute VB_Name = "FrmKurirSpvApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IDCH As String

Private Sub CmdApproved_Click()
    M_OBJCONN.Execute "Insert TblPODSend (NoPod) values ('" + Text1.Text + "')"
    MsgBox "Done"
    Call CmdPrint_Click
    Text1.Text = Empty
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
RPT.Reset
RPT.Formulas(1) = "@NoPod = totext('" + CStr(Text1.Text) + "')"
RPT.ReportFileName = App.Path + "\RpKurir.rpt"
Call SHOW_PRN
End Sub

Private Sub CmdShow_Click()
Dim m_show As New ADODB.Recordset
m_show.CursorLocation = adUseClient
If OptionStatus.Value = True Then
    m_show.Open "Select CUSTID from cc_custtbl where PODSEND = '" + Text1.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    m_show.Open "Select CUSTID from cc_custtbl where PODSEND = '" + Text1.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
End If
If m_show.RecordCount <> 0 Then
    IDCH = m_show!CUSTID
    POD = True
    FRMCUST_CC.Show vbModal
    Me.MousePointer = vbNormal
Else
    MsgBox "No POD Tidak Ada", vbInformation + vbOKOnly, "Telegrandi"
End If
Set m_show = Nothing
End Sub


Private Sub SHOW_PRN()
    RPT.RetrieveDataFiles
    RPT.WindowLeft = 0
    RPT.WindowTop = 0
    RPT.WindowState = crptMaximized
    RPT.WindowShowPrintBtn = True
    RPT.WindowShowRefreshBtn = True
    RPT.WindowShowSearchBtn = True
    RPT.WindowShowPrintSetupBtn = True
    RPT.WindowControls = True
    RPT.PrintReport
    RPT.Reset
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdShow_Click
End If
End Sub
