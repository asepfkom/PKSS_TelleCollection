VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormHistoryLog 
   Caption         =   "History Log"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form3"
   ScaleHeight     =   7275
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexcel 
      Caption         =   "Export to Excel"
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin MSComctlLib.ListView listhistorylog 
      Height          =   7035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   12409
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
   Begin TDBDate6Ctl.TDBDate date1 
      Height          =   285
      Left            =   7920
      TabIndex        =   2
      Top             =   480
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   503
      Calendar        =   "FormHistoryLog.frx":0000
      Caption         =   "FormHistoryLog.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FormHistoryLog.frx":0184
      Keys            =   "FormHistoryLog.frx":01A2
      Spin            =   "FormHistoryLog.frx":0200
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   12648384
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   6815745
      Value           =   39876
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate date2 
      Height          =   285
      Left            =   9840
      TabIndex        =   3
      Top             =   480
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   494
      Calendar        =   "FormHistoryLog.frx":0228
      Caption         =   "FormHistoryLog.frx":0340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FormHistoryLog.frx":03AC
      Keys            =   "FormHistoryLog.frx":03CA
      Spin            =   "FormHistoryLog.frx":0428
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   12648384
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   6815745
      Value           =   39876
      CenturyMode     =   0
   End
   Begin MSComDlg.CommonDialog CD_save 
      Left            =   10560
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "to"
      Height          =   255
      Left            =   9480
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Search by Tanggal"
      Height          =   255
      Left            =   7920
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FormHistoryLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdexcel_Click()
 Dim objExcel As New Excel.Application
Dim objExcelSheet As Excel.Worksheet

If listhistorylog.ListItems.Count > 0 Then
    objExcel.Workbooks.ADD
    Set objExcelSheet = objExcel.Worksheets.ADD
 

    For col = 1 To listhistorylog.ColumnHeaders.Count
        objExcelSheet.Cells(1, col).Value = listhistorylog.ColumnHeaders(col)
    Next
 
    For Row = 2 To listhistorylog.ListItems.Count + 1
        For col = 1 To listhistorylog.ColumnHeaders.Count
        If col = 1 Then
                objExcelSheet.Cells(Row, col).Value = listhistorylog.ListItems(Row - 1).Text
        Else
            '" 'cararandy 29032016 "
            'Dim hasil As Double
            Dim hasil1 As String
            'Dim hasil2 As Long
            'If Col = 3 Then
                hasil1 = "'" + listhistorylog.ListItems(Row - 1).SubItems(col - 1)
'                objExcelSheet.Cells.NumberFormat = "@"
                objExcelSheet.Cells(Row, col).Value = hasil1
'            Else
'                hasil1 = listhistorylog.ListItems(Row - 1).SubItems(Col - 1)
'                objExcelSheet.Cells(Row, Col).Value = hasil1
            End If
'        End If
        Next
    Next
 
    objExcelSheet.Columns.AutoFit
    CD_save.ShowOpen
    a = CD_save.FileName
 
    objExcelSheet.SaveAs a & ".xls"
    MsgBox "Export Completed", vbInformation, Me.Caption
 
    objExcel.Workbooks.Open a & ".xls"
    objExcel.Visible = True
    'objExcel.Quit
Else
    MsgBox "No data to export", vbInformation, Me.Caption
End If

End Sub

Private Sub cmdsearch_Click()
     Dim CustId, sQuery, where, tgl_telfon As String
    Dim RS_Lv As ADODB.Recordset
    Dim num As Integer
    
    sQuery = "SELECT * FROM tblloglistreview where date(tanggal_telfon) between '" + Format(date1.Value, "yyyy-mm-dd") + "'  and '" + Format(date2.Value, "yyyy-mm-dd") + "'  "
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    listhistorylog.ListItems.CLEAR
    If RS_Lv.RecordCount > 0 Then
        num = 0
        Do Until RS_Lv.EOF
            num = num + 1
            tanggal_telfon = Format(RS_Lv("tanggal_telfon"), "yyyy-mm-dd hh:mm:ss")
            Set listItem = listhistorylog.ListItems.ADD(, , num)
            listItem.SubItems(1) = Trim(cnull(RS_Lv("agent")))
            listItem.SubItems(2) = Trim(cnull(RS_Lv("custid")))
            listItem.SubItems(3) = Trim(cnull(RS_Lv("no_telfon")))
            listItem.SubItems(4) = tanggal_telfon
            listItem.SubItems(5) = Trim(cnull(RS_Lv("user_release")))
            RS_Lv.MoveNext
        Loop
    Else
        MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
    End If
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_Load()
    Call HeaderLv
    Call Isilv
End Sub
Private Sub HeaderLv()
    listhistorylog.ColumnHeaders.ADD , , "ID", 500
    listhistorylog.ColumnHeaders.ADD , , "Agent", 1100
    listhistorylog.ColumnHeaders.ADD , , "Customer ID", 3300
    listhistorylog.ColumnHeaders.ADD , , "Phone Number", 2400
    listhistorylog.ColumnHeaders.ADD , , "Call Date", 2000
    listhistorylog.ColumnHeaders.ADD , , "Pe-Release", 2000
End Sub

Private Sub Isilv()
    Dim CustId, sQuery, where, tgl_telfon As String
    Dim RS_Lv As ADODB.Recordset
    Dim num As Integer
    
    sQuery = "SELECT * FROM tblloglistreview"
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    listhistorylog.ListItems.CLEAR
    If RS_Lv.RecordCount > 0 Then
        num = 0
        Do Until RS_Lv.EOF
            num = num + 1
            tanggal_telfon = Format(RS_Lv("tanggal_telfon"), "yyyy-mm-dd hh:mm:ss")
            Set listItem = listhistorylog.ListItems.ADD(, , num)
            listItem.SubItems(1) = Trim(cnull(RS_Lv("agent")))
            listItem.SubItems(2) = Trim(cnull(RS_Lv("custid")))
            listItem.SubItems(3) = Trim(cnull(RS_Lv("no_telfon")))
            listItem.SubItems(4) = tanggal_telfon
            listItem.SubItems(5) = Trim(cnull(RS_Lv("user_release")))
            RS_Lv.MoveNext
        Loop
    Else
        MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
    End If
End Sub
