VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmListHotProspect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List Hot Prospect"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10185
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus status hot prospect"
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   60
      Width           =   2295
   End
   Begin VB.CommandButton CmdCekAll 
      Caption         =   "&Cek All"
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&UnCek All"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox TxtJml 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "0"
      Top             =   4020
      Width           =   1035
   End
   Begin MSComctlLib.ListView LvHotPr 
      Height          =   4020
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   7091
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   1
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
   Begin VB.Label Label1 
      Caption         =   "Jumlah Data:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
End
Attribute VB_Name = "FrmListHotProspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
    LvHotPr.ColumnHeaders.ADD 1, , "Custid", 2500
    LvHotPr.ColumnHeaders.ADD 2, , "Nama", 3000
    LvHotPr.ColumnHeaders.ADD 3, , "Status Kept", 1500
End Sub


Private Sub CmdCekAll_Click()
    Dim W As Integer
    If LvHotPr.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    For W = 1 To LvHotPr.ListItems.Count
        LvHotPr.ListItems(W).Checked = True
    Next W
End Sub

Private Sub cmdHapus_Click()
    Dim W As Integer
    Dim CMDSQL As String
    Dim K As String
    
    If LvHotPr.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    K = MsgBox("Anda yakin akan menghapus data hot prospect yang dicentang?", vbQuestion + vbYesNo, "Konfirmasi")
    If K = vbNo Then
        Exit Sub
    End If
    
    For W = 1 To LvHotPr.ListItems.Count
        If LvHotPr.ListItems(W).Checked = True Then
            CMDSQL = "update mgm set status_htc=null where custid='"
            CMDSQL = CMDSQL + LvHotPr.ListItems(W).Text + "'"
            M_OBJCONN.Execute CMDSQL
        End If
    Next W
    Call IsiData
    MsgBox "Status Hot Prospect berhasil dihapus!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub Command1_Click()
    Dim W As Integer
    If LvHotPr.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    For W = 1 To LvHotPr.ListItems.Count
        LvHotPr.ListItems(W).Checked = False
    Next W
End Sub

Private Sub Form_Load()
    Call header
    Call IsiData
End Sub

Private Sub IsiData()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    If UCase(MDIForm1.txtlevel.Text) = "ADMIN" Then
        CMDSQL = "select custid,name,status_keep from mgm where status_htc='1' order by name asc"
    End If
    If UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Then
        CMDSQL = "select custid,name,status_keep from mgm where status_htc='1' order by name asc"
    End If
    If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Then
        CMDSQL = "select custid,name,status_keep from mgm where status_htc='1' order by name asc"
    End If
    If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        CMDSQL = "select custid,name,status_keep from mgm where status_htc='1' and agent in ("
        CMDSQL = CMDSQL + "select userid from  usertbl where team='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "' and usertype='1')  order by name asc"
    End If
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        CMDSQL = "select custid,name,status_keep from mgm where status_htc='1' and agent='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "' "
        CMDSQL = CMDSQL + "order by name asc"
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvHotPr.ListItems.CLEAR
    
    txtjml.Text = M_objrs.RecordCount
    
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data Hot Prospect tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    While Not M_objrs.EOF
        Set ListItem = LvHotPr.ListItems.ADD(, , M_objrs("custid"))
            ListItem.SubItems(1) = M_objrs("name")
            If M_objrs("status_keep") = "1" Then
                ListItem.SubItems(2) = "KEPT"
            End If
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub
