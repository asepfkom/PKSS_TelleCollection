VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRMclaimResult 
   Caption         =   "Verifikasi Inbound"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   Icon            =   "FRMclaimResult.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8625
   ScaleWidth      =   12615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Tutup"
      Height          =   390
      Left            =   11415
      TabIndex        =   1
      Top             =   8205
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8235
      Left            =   15
      TabIndex        =   0
      Top             =   390
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   14526
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0 = Belum Diproses  1 = Cek Valid   2 = Di Reject / Di Hapus"
      Height          =   270
      Left            =   7860
      TabIndex        =   2
      Top             =   60
      Width           =   4680
   End
End
Attribute VB_Name = "FRMclaimResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "AOC-Agent Lama", 15 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Nama-Agent Lama", 15 * TXT
    ListView1.ColumnHeaders.ADD 3, , "AOC-Agent Baru", 10 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Nama-Agent Baru", 8 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Refference No", 10 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Refference Name", 10 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Leads", 10 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Status", 10 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Reason", 50 * TXT
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim m_objrs As New ADODB.Recordset
Dim LISTITEM As LISTITEM
Dim cmdsql As String
Call header
m_objrs.CursorLocation = adUseClient
m_objrs.Open "SELECT * FROM RequestInbound where AGENTlama ='" + MDIForm1.Text1.Text + "' or  AGENTBaru ='" + MDIForm1.Text1.Text + "' ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
        Set LISTITEM = ListView1.ListItems.ADD(, , IIf(IsNull(m_objrs("AGENTlama")), "", m_objrs("AGENTlama")))
        LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("NamaAGENTlama")), "", m_objrs("NamaAGENTlama"))
        LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("AGENTBaru")), "", m_objrs("AGENTBaru"))
        LISTITEM.SubItems(3) = IIf(IsNull(m_objrs("NamaAGENTBaru")), "", m_objrs("NamaAGENTBaru"))
        LISTITEM.SubItems(4) = IIf(IsNull(m_objrs("NOREF")), "", m_objrs("NOREF"))
        LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
        LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("LEADSNAME")), "", m_objrs("LEADSNAME"))
        LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("Status")), "", m_objrs("Status"))
        LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("Reason")), "", m_objrs("Reason"))
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
End Sub
