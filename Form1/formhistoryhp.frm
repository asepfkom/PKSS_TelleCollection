VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form formhistoryhp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History Call"
   ClientHeight    =   8145
   ClientLeft      =   1935
   ClientTop       =   1395
   ClientWidth     =   6480
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   6480
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "List Phone Number"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   6135
      Begin MSComctlLib.ListView lvhistoryhp 
         Height          =   5430
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   9578
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Criteria Report"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cbhistory 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "formhistoryhp.frx":0000
         Left            =   1320
         List            =   "formhistoryhp.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton btnsearch 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboteam 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   660
         Width           =   1575
      End
      Begin TDBDate6Ctl.TDBDate StartDate 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1560
         _Version        =   65536
         _ExtentX        =   2752
         _ExtentY        =   556
         Calendar        =   "formhistoryhp.frx":0030
         Caption         =   "formhistoryhp.frx":0148
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "formhistoryhp.frx":01B4
         Keys            =   "formhistoryhp.frx":01D2
         Spin            =   "formhistoryhp.frx":0230
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   1.12794198814265E-317
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate EndDate 
         Height          =   315
         Left            =   3430
         TabIndex        =   3
         Top             =   240
         Width           =   1560
         _Version        =   65536
         _ExtentX        =   2752
         _ExtentY        =   556
         Calendar        =   "formhistoryhp.frx":0258
         Caption         =   "formhistoryhp.frx":0370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "formhistoryhp.frx":03DC
         Keys            =   "formhistoryhp.frx":03FA
         Spin            =   "formhistoryhp.frx":0458
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   1.12794198814265E-317
         CenturyMode     =   0
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Search By     :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Team             :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3040
         TabIndex        =   5
         Top             =   310
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Call :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   310
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CD_save 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   8505
      Left            =   0
      Picture         =   "formhistoryhp.frx":0480
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11640
   End
End
Attribute VB_Name = "formhistoryhp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CustId, sQuery, where, tgl_telfon, tanggalcall As String
Private num As Long
Private listItem As listItem
Private RS_Lv As ADODB.Recordset

Private Sub btnexport_Click()
    Me.MousePointer = vbArrowHourglass
    'cmd_download.Enabled = False
    'Call My_Export_Excel
    Me.MousePointer = vbArrow
End Sub

Private Sub btnsearch_Click()
    Call Isilvdb
End Sub

Private Sub Command1_Click()
Dim objExcel As New Excel.Application
Dim objExcelSheet As Excel.Worksheet
Dim col, Row As Integer
Dim a As String
If lvhistoryhp.ListItems.Count > 0 Then
    objExcel.Workbooks.ADD
    Set objExcelSheet = objExcel.Worksheets.ADD
 

    For col = 1 To lvhistoryhp.ColumnHeaders.Count
        objExcelSheet.Cells(1, col).Value = lvhistoryhp.ColumnHeaders(col)
    Next
 
    For Row = 2 To lvhistoryhp.ListItems.Count + 1
        For col = 1 To lvhistoryhp.ColumnHeaders.Count
        If col = 1 Then
                objExcelSheet.Cells(Row, col).Value = lvhistoryhp.ListItems(Row - 1).Text
        Else
            '" 'cararandy 29032016 "
            Dim hasil1 As String
                hasil1 = "'" + lvhistoryhp.ListItems(Row - 1).SubItems(col - 1)
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
    Call HeaderLv
    Call Isi_Team
End Sub

Private Sub Isi_Team()
    sQuery = "SELECT DISTINCT agent FROM tbl_manual_dial order by agent"
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If RS_Lv.RecordCount > 0 Then
        While Not RS_Lv.EOF
            cboteam.AddItem RS_Lv!agent
            RS_Lv.MoveNext
        Wend
    End If

End Sub

Private Sub HeaderLv()
    lvhistoryhp.ColumnHeaders.ADD , , "No", 500
    lvhistoryhp.ColumnHeaders.ADD , , "Agent", 1500
    lvhistoryhp.ColumnHeaders.ADD , , "Phone Number", 1800
    lvhistoryhp.ColumnHeaders.ADD , , "Tanggal Call", 2000
End Sub

Private Sub Isilvdb()
        sQuery = "select * from tbl_manual_dial where phone_number is not null"

    If StartDate.Value <> "" Or EndDate.Value <> "" Then
        sQuery = sQuery + " and date(tgl_call) between '" + Format(StartDate.Value, "yyyy-mm-dd") + "'  and '" + Format(EndDate.Value, "yyyy-mm-dd") + "'"
    End If
    If cboteam.Text <> "" Then
        sQuery = sQuery + " and agent =  '" + cboteam.Text + "'"
    End If
    If cbhistory.Text <> "" Then
        If cbhistory.Text = "DATABASE CALL" Then
            sQuery = sQuery + " and EXISTS_NUMBER = 1"
        ElseIf cbhistory.Text = "FREE CALL" Then
            sQuery = sQuery + " and EXISTS_NUMBER = 0"
        End If
    End If
        
        Set RS_Lv = New ADODB.Recordset
        RS_Lv.CursorLocation = adUseClient
        RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        lvhistoryhp.ListItems.CLEAR
        If RS_Lv.RecordCount > 0 Then
            num = 0
            Do Until RS_Lv.EOF
                num = num + 1
                tanggalcall = Format(RS_Lv("tgl_call"), "yyyy-mm-dd hh:mm:ss")
                Set listItem = lvhistoryhp.ListItems.ADD(, , num)
                listItem.SubItems(1) = Trim(cnull(RS_Lv("agent")))
                listItem.SubItems(2) = Trim(cnull(RS_Lv("phone_number")))
                listItem.SubItems(3) = tanggalcall
                RS_Lv.MoveNext
            Loop
        Else
            MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
        End If
        Label2.Caption = RS_Lv.RecordCount
End Sub

