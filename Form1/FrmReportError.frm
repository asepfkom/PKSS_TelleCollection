VERSION 5.00
Begin VB.Form FrmReportError 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Error Reporting"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2775
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdReportTelepon 
      Caption         =   "Report Masalah Telepon"
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   2655
   End
   Begin VB.CommandButton CmdReportHeadset 
      Caption         =   "Report Masalah Headset"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2655
   End
End
Attribute VB_Name = "FrmReportError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdReportHeadset_Click()
    FrmReportHeadset.Show vbModal
End Sub

'Private Sub IsiJenisMasalah()
'    Dim M_Objrs As ADODB.Recordset
'    Dim Cmdsql As String
'
'    Cmdsql = "select * from tbl_jenis_masalah where jenis_masalah is not null and status='1' "
'    Cmdsql = Cmdsql + "order by jenis_problem asc "
'    Set M_Objrs = New ADODB.Recordset
'    M_Objrs.CursorLocation = adUseClient
'    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    CmbJenisMasalah.CLEAR
'
'    If M_Objrs.RecordCount > 0 Then
'        While Not M_Objrs.EOF
'            CmbJenisMasalah.AddItem UCase(M_Objrs("jenis_problem"))
'            M_Objrs.MoveNext
'        Wend
'    End If
'
'    Set M_Objrs = Nothing
'End Sub
Private Sub CmdReportTelepon_Click()
    FrmReportTelepon.Show vbModal
End Sub
