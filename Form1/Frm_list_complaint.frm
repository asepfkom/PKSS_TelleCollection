VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_list_complaint 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Complaint"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12345
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   12135
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   6720
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Export"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10080
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "Frm_list_complaint.frx":0000
         Left            =   1560
         List            =   "Frm_list_complaint.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "Frm_list_complaint.frx":0036
         Left            =   1560
         List            =   "Frm_list_complaint.frx":0043
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin TDBDate6Ctl.TDBDate tglcomplaintfrom 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   720
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   503
         Calendar        =   "Frm_list_complaint.frx":0056
         Caption         =   "Frm_list_complaint.frx":016E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_list_complaint.frx":01DA
         Keys            =   "Frm_list_complaint.frx":01F8
         Spin            =   "Frm_list_complaint.frx":0256
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
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
      Begin TDBDate6Ctl.TDBDate tglcomplaintto 
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Top             =   720
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   494
         Calendar        =   "Frm_list_complaint.frx":027E
         Caption         =   "Frm_list_complaint.frx":0396
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_list_complaint.frx":0402
         Keys            =   "Frm_list_complaint.frx":0420
         Spin            =   "Frm_list_complaint.frx":047E
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Complaint"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Kriteria"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.TextBox txtjml 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   10440
      TabIndex        =   1
      Top             =   6720
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4725
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   8334
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Data (s)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   14
      Top             =   6800
      Width           =   1215
   End
End
Attribute VB_Name = "Frm_list_complaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs_complaint As ADODB.Recordset
Private rs_view As ADODB.Recordset

Private strQuery As String
Private i As Integer

Private Sub Command1_Click()
    Dim list As listItem
    
    strQuery = "SELECT * FROM tbl_complaint WHERE id IS NOT NULL"
    
    If Combo1.Text <> "" And Text1.Text <> "" Then
        Select Case Combo1.Text
        Case "CUST ID"
            strQuery = strQuery & " AND custid='" & Trim(Text1.Text) & "'"
        Case "CUST NAME"
            strQuery = strQuery & " AND lower(cust_name) like '%" & Trim(LCase(Text1.Text)) & "%'"
        Case "COMPLAINER"
            strQuery = strQuery & " AND lower(complainer) like '%" & Trim(LCase(Text1.Text)) & "%'"
        End Select
    End If
    
    If Not tglcomplaintfrom.ValueIsNull And Not tglcomplaintto.ValueIsNull Then
        strQuery = strQuery & " AND (date(tgl_complaint) BETWEEN '" & Format(tglcomplaintfrom.Value, "yyyy-mm-dd") & "' AND '" & Format(tglcomplaintto.Value, "yyyy-mm-dd") & "')"
    End If
    
    If Combo2.Text <> "" Then
        strQuery = strQuery & " AND (lower(status)='" & LCase(Combo2.Text) & "')"
    End If
    
    ListView1.ListItems.CLEAR
    i = 0
    
    If rs_complaint.state = 1 Then rs_complaint.Close
    rs_complaint.Open strQuery
    If rs_complaint.RecordCount > 0 Then
        Do Until rs_complaint.EOF
            i = i + 1
            Set list = ListView1.ListItems.ADD(, , i)
            list.SubItems(1) = IIf(IsNull(rs_complaint!tgl_complaint), "", Format(rs_complaint!tgl_complaint, "dd-mm-yyyy"))
            list.SubItems(2) = IIf(IsNull(rs_complaint!CustId), "", rs_complaint!CustId)
            list.SubItems(3) = IIf(IsNull(rs_complaint!cust_name), "", rs_complaint!cust_name)
            list.SubItems(4) = IIf(IsNull(rs_complaint!complainer), "", rs_complaint!complainer)
            list.SubItems(5) = IIf(IsNull(rs_complaint!problem), "", rs_complaint!problem)
            list.SubItems(6) = IIf(IsNull(rs_complaint!STATUS), "", rs_complaint!STATUS)
            list.SubItems(7) = IIf(IsNull(rs_complaint!action_taken), "", rs_complaint!action_taken)
            list.SubItems(8) = IIf(IsNull(rs_complaint!date_solved), "", Format(rs_complaint!date_solved, "dd-mm-yyyy"))
            list.SubItems(9) = IIf(IsNull(rs_complaint!agent), "", rs_complaint!agent)
            list.SubItems(10) = IIf(IsNull(rs_complaint!ID), "", rs_complaint!ID)
            rs_complaint.MoveNext
        Loop
        Command2.Enabled = True
    Else
        MsgBox "Data tidak ditemukan!!", vbOKOnly + vbInformation, "INFO"
        Command2.Enabled = False
    End If
    txtjml.Text = ListView1.ListItems.Count
End Sub

Private Sub Command2_Click()
    'On Error GoTo SALAH
    Dim Txtpath         As String
    Dim cmdsql          As String
    Dim listItem        As listItem
    Dim cmdsql_update   As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim m_msgbox        As String
    Dim i               As Integer
   
form_save:
    CD_save.ShowSave
    CD_save.Filter = "Excel File |*.xls"
    Txtpath = CD_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If Txtpath = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Download dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
            MsgBox "Download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
            GoTo form_save        '-> maka goto form_save
        End If
    End If

    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
    
    strQuery = "SELECT tgl_complaint,custid,cust_name,complainer,problem,status,action_taken,date_solved,agent FROM tbl_complaint WHERE id IS NOT NULL"
    
    If Combo1.Text <> "" And Text1.Text <> "" Then
        Select Case Combo1.Text
        Case "CUST ID"
            strQuery = strQuery & " AND custid='" & Trim(Text1.Text) & "'"
        Case "CUST NAME"
            strQuery = strQuery & " AND lower(cust_name) like '%" & Trim(LCase(Text1.Text)) & "%'"
        Case "COMPLAINER"
            strQuery = strQuery & " AND lower(complainer) like '%" & Trim(LCase(Text1.Text)) & "%'"
        End Select
    End If
    
    If Not tglcomplaintfrom.ValueIsNull And Not tglcomplaintto.ValueIsNull Then
        strQuery = strQuery & " AND (date(tgl_complaint) BETWEEN '" & Format(tglcomplaintfrom.Value, "yyyy-mm-dd") & "' AND '" & Format(tglcomplaintto.Value, "yyyy-mm-dd") & "')"
    End If
    
    If Combo2.Text <> "" Then
        strQuery = strQuery & " AND (lower(status)='" & LCase(Combo2.Text) & "')"
    End If
    
    If rs_complaint.state = 1 Then rs_complaint.Close
    rs_complaint.Open strQuery & " ORDER BY tgl_complaint"
        
    On Error GoTo SALAH
    'Proses pengsisian nama field ke excel
    Dim x, Y    As Integer
    If rs_complaint.state = 1 Then
        x = 0
        i = 1
        Y = rs_complaint.fields().Count - 1
        Do Until x > Y
            DoEvents
            objSheet.Cells(1, i).Value = CStr(rs_complaint.fields(x).Name)
            i = i + 1
            x = x + 1
        Loop
    End If
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset rs_complaint '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs Txtpath, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
 
SALAH:
    Exit Sub
End Sub

Private Sub Form_Load()
    Set rs_complaint = New ADODB.Recordset
    rs_complaint.ActiveConnection = M_OBJCONN
    rs_complaint.CursorLocation = adUseClient
    rs_complaint.CursorType = adOpenDynamic
    rs_complaint.LockType = adLockOptimistic
    
    Set rs_view = New ADODB.Recordset
    rs_view.ActiveConnection = M_OBJCONN
    rs_view.CursorLocation = adUseClient
    rs_view.CursorType = adOpenDynamic
    rs_view.LockType = adLockOptimistic
    
    With ListView1
        .ColumnHeaders.ADD , , "NO", 500
        .ColumnHeaders.ADD , , "Tgl Complaint"
        .ColumnHeaders.ADD , , "Cust ID"
        .ColumnHeaders.ADD , , "Cust Name"
        .ColumnHeaders.ADD , , "Complainer"
        .ColumnHeaders.ADD , , "Problem"
        .ColumnHeaders.ADD , , "STATUS"
        .ColumnHeaders.ADD , , "Action Taken"
        .ColumnHeaders.ADD , , "Date Solved"
        .ColumnHeaders.ADD , , "agent"
        .ColumnHeaders.ADD , , "ID"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_complaint = Nothing
    Set rs_view = Nothing
End Sub

Private Sub ListView1_DblClick()
    If rs_view.state = 1 Then rs_view.Close
    rs_view.Open "SELECT * FROM tbl_complaint WHERE id=" & ListView1.SelectedItem.SubItems(10) & ""
    If rs_view.RecordCount > 0 Then
        With Form_complaint
            .txt_custid.Text = IIf(IsNull(rs_view!CustId), "", rs_view!CustId)
            .txt_custname.Text = IIf(IsNull(rs_view!cust_name), "", rs_view!cust_name)
            .txt_agent.Text = IIf(IsNull(rs_view!agent), "", rs_view!agent)
            .cb_status.Text = IIf(IsNull(rs_view!STATUS), "", rs_view!STATUS)
            .txt_complainer.Text = IIf(IsNull(rs_view!complainer), "", rs_view!complainer)
            .txt_problem.Text = IIf(IsNull(rs_view!problem), "", rs_view!problem)
            .txt_tglcomplaint.Value = IIf(IsNull(rs_view!tgl_complaint), Null, rs_view!tgl_complaint)
            .txt_action.Text = IIf(IsNull(rs_view!action_taken), "", rs_view!action_taken)
            .txttglsolved = IIf(IsNull(rs_view!date_solved), Null, rs_view!date_solved)
            .Frame2.Enabled = True
            .Frame1.Enabled = False
            .lbl_ket.Caption = "U"
            .lbl_id.Caption = ListView1.SelectedItem.SubItems(10)
            .Show 1
        End With
    End If
End Sub
