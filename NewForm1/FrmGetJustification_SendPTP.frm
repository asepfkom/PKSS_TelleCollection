VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmGetJustification_SendPTP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Get Justification From Remarks"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12240
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   11100
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
   End
   Begin MSComctlLib.ListView listview1 
      Height          =   4080
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7197
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   10147522
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmGetJustification_SendPTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdkeluar_Click()
    Unload Me
End Sub

Private Sub HEADER_HISTORY()
    ListView1(1).ColumnHeaders.ADD 1, , "Tanggal(mm-dd-yyyy)", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 2, , "History", 70 * TXT
    ListView1(1).ColumnHeaders.ADD 3, , "User Log", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 4, , "Handle By", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 5, , "Sts Account", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 6, , "Sts Call", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 7, , "Sts Telp With", 25 * TXT
    ListView1(1).ColumnHeaders.ADD 8, , "Id", 25 * TXT
End Sub

Private Sub AmbilHistory()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    CMDSQL = "select * from mandiri.mgm_hst where custid='"
    CMDSQL = CMDSQL + Trim(FrmViewPTP.txtcardno.Text) + "' order by tgl desc"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = ListView1(1).ListItems.ADD(, , Format(IIf(IsNull(M_objrs("TGL")), "", M_objrs!TGL), "mm-dd-yyyy hh:mm:ss"))
                ListItem.SubItems(1) = IIf(IsNull(M_objrs("HST")), "", M_objrs("HST"))
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("user_log")), "", M_objrs("user_log"))
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("AGENT")), "", M_objrs("AGENT"))
                ListItem.SubItems(4) = IIf(IsNull(M_objrs("KodeDs")), "", M_objrs("KodeDs"))
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("statuscall")), "", M_objrs("statuscall"))
                ListItem.SubItems(6) = IIf(IsNull(M_objrs("ststelpwith")), "", M_objrs("ststelpwith"))
                ListItem.SubItems(7) = IIf(IsNull(M_objrs("id")), "", M_objrs("id"))
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub

Private Sub cmdOK_Click()
     If ListView1(1).ListItems.Count = 0 Then
        MsgBox "Data Remarks belum tersedia!", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If
    
    If FrmViewPTP.txtjust.Text = "" Then
        FrmViewPTP.txtjust.Text = ListView1(1).SelectedItem.SubItems(1)
    Else
        FrmViewPTP.txtjust.Text = FrmViewPTP.txtjust.Text + vbCrLf + vbCrLf + "-" + ListView1(1).SelectedItem.SubItems(1)
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call HEADER_HISTORY
    Call AmbilHistory
End Sub



Private Sub ListView1_DblClick(Index As Integer)
    If ListView1(1).ListItems.Count = 0 Then
        MsgBox "Data Remarks belum tersedia!", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If
    
    If FrmViewPTP.txtjust.Text = "" Then
        FrmViewPTP.txtjust.Text = ListView1(1).SelectedItem.SubItems(1)
    Else
        FrmViewPTP.txtjust.Text = FrmViewPTP.txtjust.Text + vbCrLf + vbCrLf + "-" + ListView1(1).SelectedItem.SubItems(1)
    End If
    
    
    Unload Me
End Sub

