VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Formhsttelecolection 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "History Call"
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   LinkTopic       =   "Form3"
   ScaleHeight     =   5445
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&E&x&port"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   8387
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog CD_save 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Formhsttelecolection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub Command2_Click()
Dim objExcel As New Excel.Application
Dim objExcelSheet As Excel.Worksheet
Dim col, Row As Integer
Dim a As String
If ListView1.ListItems.Count > 0 Then
    objExcel.Workbooks.ADD
    Set objExcelSheet = objExcel.Worksheets.ADD
 

    For col = 1 To ListView1.ColumnHeaders.Count
        objExcelSheet.Cells(1, col).Value = ListView1.ColumnHeaders(col)
    Next
 
    For Row = 2 To ListView1.ListItems.Count + 1
        For col = 1 To ListView1.ColumnHeaders.Count
        If col = 1 Then
                objExcelSheet.Cells(Row, col).Value = ListView1.ListItems(Row - 1).Text
        Else
            '" 'cararandy 29032016 "
            Dim hasil1 As String
                hasil1 = "'" + ListView1.ListItems(Row - 1).SubItems(col - 1)
                objExcelSheet.Cells(Row, col).Value = hasil1
            End If
        Next
    Next
 
    objExcelSheet.Columns.AutoFit
    CD_save.ShowOpen
    a = CD_save.FileName
 
    objExcelSheet.SaveAs a & ".xls"
    MsgBox "Export Completed", vbInformation, Me.Caption
 
    objExcel.Workbooks.Open a & ".xls"
    objExcel.Visible = True
Else
    MsgBox "No data to export", vbInformation, Me.Caption
End If
End Sub

Private Sub Form_Load()
    Call headerhst
    Call Isilv
End Sub

Public Sub headerhst()
    ListView1.ColumnHeaders.ADD , , "Tanggal", 2000
    ListView1.ColumnHeaders.ADD , , "Agent Lama", 2000
    ListView1.ColumnHeaders.ADD , , "Agent Baru", 2000
    ListView1.ColumnHeaders.ADD , , "Create By", 2000
    ListView1.ColumnHeaders.ADD , , "List Do", 2000
End Sub

Public Sub Isilv()
    Dim CustId, sQuery, tgl_telfon As String
    Dim RS_Lv As ADODB.Recordset
    
    sQuery = "SELECT * FROM hst_telecollection"
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    ListView1.ListItems.CLEAR
    If RS_Lv.RecordCount > 0 Then
        Do Until RS_Lv.EOF
            Tanggal = Format(RS_Lv("tanggal"), "yyyy-mm-dd hh:mm:ss")
            Set listItem = ListView1.ListItems.ADD(, , Tanggal)
            listItem.SubItems(1) = Trim(cnull(RS_Lv("agent_lama")))
            listItem.SubItems(2) = Trim(cnull(RS_Lv("agent_batu")))
            listItem.SubItems(3) = Trim(cnull(RS_Lv("createby")))
            listItem.SubItems(4) = Trim(cnull(RS_Lv("listdo")))
            RS_Lv.MoveNext
        Loop
    Else
        MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
    End If
End Sub

