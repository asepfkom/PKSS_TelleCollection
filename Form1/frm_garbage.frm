VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_garbage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Garbage"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8040
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5820
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   10266
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
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frm_garbage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddelete_Click()
    Dim query As String
        
        query = "DELETE FROM tampungtransferdata where custid not in (select custid from mgm)"
        M_OBJCONN.Execute query
    
    MsgBox "Berhasil didelete"
    Call isilv
End Sub

Private Sub Form_Load()
    Call HeaderLv
    Call isilv
End Sub
Private Sub isilv()
    Dim CustId, sQuery, where, tgl_telfon As String
    Dim RS_Lv As ADODB.Recordset
    Dim num As Integer
    
    sQuery = "SELECT * FROM tampungtransferdata where custid not in (select custid from mgm)"

    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    ListView1.ListItems.CLEAR
    If RS_Lv.RecordCount > 0 Then
        num = 0
        Do Until RS_Lv.EOF
            num = num + 1
            tanggalupload = Format(RS_Lv("tanggalupload"), "yyyy-mm-dd hh:mm:ss")
            Set listItem = ListView1.ListItems.ADD(, , num)
            listItem.SubItems(1) = Trim(cnull(RS_Lv("custid")))
            listItem.SubItems(2) = tanggalupload
            listItem.SubItems(3) = Trim(cnull(RS_Lv("pengupload")))
            listItem.SubItems(4) = Trim(cnull(RS_Lv("tujapproval")))
            RS_Lv.MoveNext
        Loop
    Else
        MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
    End If
End Sub

Private Sub HeaderLv()
    ListView1.ColumnHeaders.ADD , , "No", 600
    ListView1.ColumnHeaders.ADD , , "Custid", 1100
    ListView1.ColumnHeaders.ADD , , "Tanggal Upload", 2000
    ListView1.ColumnHeaders.ADD , , "PengUpload", 2000
    ListView1.ColumnHeaders.ADD , , "Pengaprove", 2000
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   IndexColumnHEader = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub
