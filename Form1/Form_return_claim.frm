VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form_return_claim 
   Caption         =   "Batal Claim Account"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12120
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   12120
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD_save 
      Left            =   3360
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   11040
      TabIndex        =   18
      Top             =   7560
      Width           =   855
   End
   Begin VB.PictureBox CD_save1 
      Height          =   480
      Left            =   6480
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   19
      Top             =   7440
      Width           =   1200
   End
   Begin VB.TextBox txtpath 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   7560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Export Claim Approve"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   16
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtfield_kriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   240
      Width           =   2535
   End
   Begin VB.ComboBox cbkriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Form_return_claim.frx":0000
      Left            =   960
      List            =   "Form_return_claim.frx":0010
      TabIndex        =   14
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      TabIndex        =   10
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Export Batal Claim"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   7
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check All"
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
      TabIndex        =   2
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Return Back To Agent Claim"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   7560
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   11040
      TabIndex        =   0
      Top             =   3840
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2865
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   5054
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
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
   Begin MSComctlLib.ListView ListView2 
      Height          =   2625
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   4630
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
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
   Begin TDBDate6Ctl.TDBDate TDBDate2 
      Height          =   285
      Left            =   8640
      TabIndex        =   8
      Top             =   240
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   494
      Calendar        =   "Form_return_claim.frx":0037
      Caption         =   "Form_return_claim.frx":014F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Form_return_claim.frx":01BB
      Keys            =   "Form_return_claim.frx":01D9
      Spin            =   "Form_return_claim.frx":0237
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
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   285
      Left            =   6720
      TabIndex        =   9
      Top             =   240
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   494
      Calendar        =   "Form_return_claim.frx":025F
      Caption         =   "Form_return_claim.frx":0377
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Form_return_claim.frx":03E3
      Keys            =   "Form_return_claim.frx":0401
      Spin            =   "Form_return_claim.frx":045F
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
   Begin VB.Line Line2 
      X1              =   1200
      X2              =   11880
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   10080
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label label1 
      Caption         =   "Criteria"
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
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   735
   End
   Begin VB.Label label1 
      Caption         =   "Claim Date"
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
      Index           =   2
      Left            =   5400
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label label1 
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
      Index           =   0
      Left            =   8280
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Batal Claim"
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
      TabIndex        =   6
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label label1 
      Caption         =   "Claim Approve"
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
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form_return_claim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs As ADODB.Recordset
Private sqlstr As String

Private Function sqlfilter() As String
    Dim sqlTeks As String
    
    sqlfilter = ""
    sqlTeks = ""
    If cbkriteria.Text <> "" And txtfield_kriteria.Text <> "" Then
        sqlTeks = sqlTeks & " AND lower(" & cbkriteria.Text & ") like '%" & Trim(LCase(txtfield_kriteria.Text)) & "%'"
    End If
    
    If Not TDBDate1.ValueIsNull And Not TDBDate2.ValueIsNull Then
        sqlTeks = sqlTeks & " AND date(tgl_claim) between '" & Format(TDBDate1.Value, "yyyy-mm-dd") & "' AND '" & Format(TDBDate2.Value, "yyyy-mm-dd") & "'"
    End If
    
    sqlfilter = sqlTeks
End Function

Private Sub Check1_Click()
    Dim xx As Integer
    If ListView1.ListItems.Count > 0 Then
        If Check1.Value = vbChecked Then
            For xx = 1 To ListView1.ListItems.Count
                ListView1.ListItems(xx).Checked = True
            Next xx
        Else
            For xx = 1 To ListView1.ListItems.Count
                ListView1.ListItems(xx).Checked = False
            Next xx
        End If
    End If
End Sub

Private Sub Command1_Click()
    Dim xx As Integer
    If ListView1.ListItems.Count > 0 Then
        For xx = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(xx).Checked = True Then
                M_OBJCONN.Execute "UPDATE mgm SET agent='" & ListView1.ListItems(xx).SubItems(3) & "' WHERE custid='" & ListView1.ListItems(xx).SubItems(2) & "'"
                
                M_OBJCONN.Execute "INSERT INTO log_claim_back_hst(custid,agent_claim,agent_asli,reason,tgl_claim) " & _
                                "VALUES('" & ListView1.ListItems(xx).SubItems(2) & "','" & ListView1.ListItems(xx).SubItems(3) & "','" & ListView1.ListItems(xx).SubItems(4) & "','Return To Agent Claim','" & Format(ListView1.ListItems(xx).SubItems(6), "yyyy-mm-dd") & "')"
                
                M_OBJCONN.Execute "DELETE FROM log_claim_back WHERE id='" & ListView1.ListItems(xx).SubItems(8) & "'"
            End If
        Next xx
        MsgBox "Data berhasil dikembalikan ke agent claim!!!", vbOKOnly + vbInformation, "INFO"
        Call Command4_Click
    Else
        MsgBox "Data Tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If
End Sub

Private Sub Command2_Click()
    If rs.State = 1 Then rs.Close
    rs.Open Q_show_data_appclaim
    Export_Excel rs, "Approve Claim"
End Sub

Private Sub Command3_Click()
    If rs.State = 1 Then rs.Close
    rs.Open Q_show_data_batalclaim
    Export_Excel rs, "Batal Claim"
End Sub

Private Sub Command4_Click()
    Call show_data_appclaim
    Call show_data_batalclaim
End Sub

Private Sub Form_Load()
    Call koneksi
    Call create_header
End Sub

Private Sub koneksi()
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    rs.ActiveConnection = M_OBJCONN
End Sub

Private Sub show_data_batalclaim()
    Dim lsitem As ListItem
    Dim X As Integer
    X = 0
    With ListView1
        If rs.State = 1 Then rs.Close
        rs.Open Q_show_data_batalclaim
        ListView1.ListItems.CLEAR
        Text2.Text = 0
        If rs.RecordCount > 0 Then
            Do Until rs.EOF
                X = X + 1
                Set lsitem = ListView1.ListItems.ADD(, , X)
                lsitem.SubItems(1) = IIf(IsNull(rs!tgl_log), "", Format(rs!tgl_log, "yyyy-mm-dd"))
                lsitem.SubItems(2) = IIf(IsNull(rs!CustId), "", rs!CustId)
                lsitem.SubItems(3) = IIf(IsNull(rs!agent_claim), "", rs!agent_claim)
                lsitem.SubItems(4) = IIf(IsNull(rs!agent_asli), "", rs!agent_asli)
                lsitem.SubItems(5) = IIf(IsNull(rs!Reason), "", rs!Reason)
                lsitem.SubItems(6) = IIf(IsNull(rs!tgl_claim), "", Format(rs!tgl_claim, "yyyy-mm-dd"))
                lsitem.SubItems(7) = IIf(IsNull(rs!tgl_janji), "", Format(rs!tgl_janji, "yyyy-mm-dd"))
                lsitem.SubItems(8) = IIf(IsNull(rs!ID), "", rs!ID)
                rs.MoveNext
            Loop
            Text2.Text = rs.RecordCount
        End If
    End With
End Sub

Private Sub show_data_appclaim()
    With ListView2
        If rs.State = 1 Then rs.Close
        rs.Open Q_show_data_appclaim
        ListView2.ListItems.CLEAR
        Text1.Text = 0
        If rs.RecordCount > 0 Then
            Do Until rs.EOF
                X = X + 1
                Set lsitem = ListView2.ListItems.ADD(, , X)
                lsitem.SubItems(1) = IIf(IsNull(rs!tgl_app), "", Format(rs!tgl_app, "yyyy-mm-dd"))
                lsitem.SubItems(2) = IIf(IsNull(rs!CustId), "", rs!CustId)
                lsitem.SubItems(3) = IIf(IsNull(rs!agent_claim), "", rs!agent_claim)
                lsitem.SubItems(4) = IIf(IsNull(rs!agent_asli), "", rs!agent_asli)
                lsitem.SubItems(5) = IIf(IsNull(rs!alasan), "", rs!alasan)
                lsitem.SubItems(6) = IIf(IsNull(rs!tgl_claim), "", Format(rs!tgl_claim, "yyyy-mm-dd"))
                lsitem.SubItems(7) = IIf(IsNull(rs!ID), "", rs!ID)
                rs.MoveNext
            Loop
            Text1.Text = rs.RecordCount
        End If
    End With
End Sub

Private Function Q_show_data_batalclaim() As String
    Q_show_data_batalclaim = "SELECT * FROM log_claim_back WHERE id is not null " & sqlfilter & " ORDER by agent_claim "
End Function

Private Function Q_show_data_appclaim() As String
    Q_show_data_appclaim = "SELECT a.* FROM tbl_approve_claim a,(SELECT custid,max(tgl_app) as Tgl_akhir FROM tbl_approve_claim WHERE id is not null " & sqlfilter & " GROUP BY custid) b WHERE a.custid=b.custid AND a.tgl_app=b.tgl_akhir ORDER by agent_claim "
End Function

Private Sub create_header()
    With ListView1
        .ColumnHeaders.ADD 1, , "No", 500
        .ColumnHeaders.ADD 2, , "Tgl Log"
        .ColumnHeaders.ADD 3, , "Cust ID"
        .ColumnHeaders.ADD 4, , "Agent Claim"
        .ColumnHeaders.ADD 5, , "Agent Asli"
        .ColumnHeaders.ADD 6, , "Reason"
        .ColumnHeaders.ADD 7, , "Tgl Claim"
        .ColumnHeaders.ADD 8, , "Tgl Janji"
        .ColumnHeaders.ADD 9, , "ID"
    End With
    
    With ListView2 'Approve
        .ColumnHeaders.ADD 1, , "No", 500
        .ColumnHeaders.ADD 2, , "Tgl Approve"
        .ColumnHeaders.ADD 3, , "Cust ID"
        .ColumnHeaders.ADD 4, , "Agent Claim"
        .ColumnHeaders.ADD 5, , "Agent Asli"
        .ColumnHeaders.ADD 6, , "Reason"
        .ColumnHeaders.ADD 7, , "Tgl Claim"
        .ColumnHeaders.ADD 8, , "ID"
    End With
End Sub

Private Sub Export_Excel(M_objrs As ADODB.Recordset, judul_excel As String)
    On Error GoTo Salah
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim cmdsql_update As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i               As Integer
    Dim m_msgbox        As String
    
    i = 1

Form_Save:
    Cd_save.ShowSave
    TxtPath.Text = Cd_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.Text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Download dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo Form_Save        '-> maka goto form_save
        End If
    End If

    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
    On Error GoTo Salah
    
    Dim X, Y    As Integer
    If M_objrs.State = 1 Then
        X = 0
        Y = M_objrs.fields().Count - 1
        Do Until X > Y
            DoEvents
            objSheet.Cells(1, i).Value = CStr(M_objrs.fields(X).Name)
            i = i + 1
            X = X + 1
        Loop
    End If
    objSheet.Cells(1, 1).Value = judul_excel
    objSheet.Cells(1, 1).Font.Name = "Verdana"
    objSheet.Cells(1, 1).Font.Bold = True
    
    objSheet.Cells(2, 1).Value = "Tanggal : " + Format(Now, "dd-mm-yyyy")
    objSheet.Cells(2, 1).Font.Name = "Verdana"
    objSheet.Cells(2, 1).Font.Bold = True:
    
    objSheet.Range("A4").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs TxtPath.Text, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
    
    Exit Sub
Salah:
    MsgBox err.Description
    Set M_objrs = Nothing
    Exit Sub
End Sub
