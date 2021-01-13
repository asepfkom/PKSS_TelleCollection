VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmListKeepAcc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Keep Account"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtJml 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1260
      TabIndex        =   5
      Text            =   "0"
      Top             =   4080
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&UnCek All"
      Height          =   375
      Left            =   7860
      TabIndex        =   3
      Top             =   1020
      Width           =   1035
   End
   Begin VB.CommandButton CmdCekAll 
      Caption         =   "&Cek All"
      Height          =   375
      Left            =   7860
      TabIndex        =   2
      Top             =   660
      Width           =   1035
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   7860
      TabIndex        =   1
      Top             =   120
      Width           =   1035
   End
   Begin MSComctlLib.ListView LvKeepAcc 
      Height          =   4020
      Left            =   60
      TabIndex        =   0
      Top             =   60
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
      Left            =   120
      TabIndex        =   4
      Top             =   4140
      Width           =   1095
   End
End
Attribute VB_Name = "FrmListKeepAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
    LvKeepAcc.ColumnHeaders.ADD 1, , "ID", 1000
    LvKeepAcc.ColumnHeaders.ADD 2, , "Custid", 2000
    LvKeepAcc.ColumnHeaders.ADD 3, , "Nama CH", 2000
    LvKeepAcc.ColumnHeaders.ADD 4, , "Tgl.Keep", 2000
    LvKeepAcc.ColumnHeaders.ADD 5, , "Keterangan", 5000
End Sub

Private Sub IsiKeepAcc()
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    Dim CMDSQL As String
    
    CMDSQL = "select * from tbl_keep_acc where date_part('year',tglkeep)="
    CMDSQL = CMDSQL + "date_part('year',now()) and date_part('month',tglkeep)="
    CMDSQL = CMDSQL + "date_part('month',now()) and agent='"
    CMDSQL = CMDSQL + FrmCC_Colection.lblaoc.Caption + "' order by idkeepacc desc"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvKeepAcc.ListItems.CLEAR
    txtjml.Text = M_objrs.RecordCount
    
    If M_objrs.RecordCount = 0 Then
        MsgBox "Anda tidak memiliki account keep!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    While Not M_objrs.EOF
        Set ListItem = LvKeepAcc.ListItems.ADD(, , M_objrs("idkeepacc"))
            ListItem.SubItems(1) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("nama")), "", M_objrs("nama"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("tglkeep")), "", Format(M_objrs("tglkeep"), "yyyy-mm-dd"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("keterangan")), "", M_objrs("keterangan"))
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End Sub

Private Sub CmdCekAll_Click()
    Dim W As Integer
    
    If LvKeepAcc.ListItems.Count = 0 Then
        MsgBox "Tidak ada data keep account!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvKeepAcc.ListItems.Count
        LvKeepAcc.ListItems(W).Checked = True
    Next W
End Sub

Private Sub cmdHapus_Click()
    Dim CMDSQL As String
    Dim a As String
    Dim W As Integer
    
    If LvKeepAcc.ListItems.Count = 0 Then
        MsgBox "Tidak ada data keep account!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin akan menghapus data yang dicentang?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbYes Then
        For W = 1 To LvKeepAcc.ListItems.Count
            If LvKeepAcc.ListItems(W).Checked = True Then
                CMDSQL = "delete from tbl_keep_acc where idkeepacc='"
                CMDSQL = CMDSQL + CStr(LvKeepAcc.ListItems(W).Text) + "'"
                M_OBJCONN.Execute CMDSQL
                
                CMDSQL = "update mgm set status_keep=null where custid='"
                CMDSQL = CMDSQL + CStr(LvKeepAcc.ListItems(W).SubItems(1)) + "'"
                M_OBJCONN.Execute CMDSQL
            End If
        Next W
        MsgBox "Penghapusan data yang dicentang berhasil!", vbOKOnly + vbInformation, "Informasi"
        Call IsiKeepAcc
    End If
End Sub

Private Sub Command1_Click()
        Dim W As Integer
    
    If LvKeepAcc.ListItems.Count = 0 Then
        MsgBox "Tidak ada data keep account!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvKeepAcc.ListItems.Count
        LvKeepAcc.ListItems(W).Checked = False
    Next W
End Sub

Private Sub Form_Load()
    Call header
    Call IsiKeepAcc
End Sub
