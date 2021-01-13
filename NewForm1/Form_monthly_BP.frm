VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form_monthly_BP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Monthly Data BP(Broken Promise)"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13530
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_download 
      BackColor       =   &H0000FF00&
      Caption         =   "Download Data To Excel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9000
      Width           =   3855
   End
   Begin VB.CheckBox chk_team 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3720
      TabIndex        =   9
      Top             =   1320
      Width           =   195
   End
   Begin VB.ComboBox cmb_team 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Form_monthly_BP.frx":0000
      Left            =   1200
      List            =   "Form_monthly_BP.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmd_showbp 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Show BP"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox cmb_agent 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Form_monthly_BP.frx":0004
      Left            =   1200
      List            =   "Form_monthly_BP.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "List Account BP"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   13215
      Begin VB.CheckBox cek_all_ptp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cek All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.ListView LvBP 
         Height          =   5190
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   9155
         View            =   3
         LabelEdit       =   1
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
      Begin VB.Label lbltotal 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL : IDR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   6000
         Width           =   8055
      End
      Begin VB.Label lbldata 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   6000
         Width           =   2655
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM-yyyy"
      Format          =   118816771
      CurrentDate     =   41610
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label lbl_team 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "Team   :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "Agent  :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "List Account BP (Broken Promise)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   630
      TabIndex        =   12
      Top             =   60
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   4
      Left            =   120
      Picture         =   "Form_monthly_BP.frx":0008
      Stretch         =   -1  'True
      Top             =   120
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "Form_monthly_BP.frx":0B12
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13560
   End
End
Attribute VB_Name = "Form_monthly_BP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs_list As ADODB.Recordset
Dim f_team As Boolean

Private Sub koneksi()
    Set Rs_list = New ADODB.Recordset
    Rs_list.CursorLocation = adUseClient
    Rs_list.ActiveConnection = M_OBJCONN
    Rs_list.CursorType = adOpenDynamic
    Rs_list.LockType = adLockOptimistic
End Sub

Private Sub chk_team_Click()
    If chk_team.Value = vbChecked Then
        Call Isi_TL
        cmb_agent.ListIndex = 0
        cmb_agent.Enabled = False
        cmb_team.Enabled = True
        f_team = True
    Else
        cmb_agent.Enabled = True
        cmb_team.Enabled = False
        cmb_team.ListIndex = 0
        f_team = False
    End If
End Sub

Private Sub Isi_TL()
    If Rs_list.state = 1 Then Rs_list.Close
    
    Rs_list.Open "SELECT DISTINCT team FROM mandiri.usertbl where team ilike  'TL%' "
    
    cmb_team.AddItem " "
    
    While Not Rs_list.EOF
        cmb_team.AddItem Rs_list("team")
        Rs_list.MoveNext
    Wend
End Sub

Private Sub IsiListBP()
    Dim ListItem As ListItem

    If Rs_list.state = 1 Then Rs_list.Close
        
    Rs_list.Open "SELECT * FROM mandiri.tbl_listbp WHERE f_bp = '1' ORDER BY agent"
        
    LvBP.ListItems.CLEAR
        
    If Rs_list.RecordCount > 0 Then
        nomor = 0
          Do Until Rs_list.EOF
              nomor = nomor + 1
              Set ListItem = LvBP.ListItems.ADD(, , nomor)
                              ListItem.SubItems(1) = IIf(IsNull(Rs_list!CustId), "", Rs_list!CustId)
                              ListItem.SubItems(2) = IIf(IsNull(Rs_list!AGENT), "", Rs_list!AGENT)
                              ListItem.SubItems(3) = Format(IIf(IsNull(Rs_list!PromisePay), "", Rs_list!PromisePay), "##,###")
                              ListItem.SubItems(4) = Format(IIf(IsNull(Rs_list!PromiseDate), "", Rs_list!PromiseDate), "DD-MM-YYYY")
                              ListItem.SubItems(5) = cnull(IIf(IsNull(Rs_list!custname), "", Rs_list!custname))
                              ListItem.SubItems(6) = cnull(IIf(IsNull(Rs_list!PRODUCT), "", Rs_list!PRODUCT))
                              
                              Total = Total + IIf(IsNull(Rs_list!PromisePay), "0", Rs_list!PromisePay)
                              
              Rs_list.MoveNext
          Loop
          lbldata.Caption = "Jumlah Data  : " & Rs_list.RecordCount & " Rows"
          lbltotal.Caption = "Total : IDR " & Format(Total, "##,###") & " "
          'txt_total_ptp.Text = Total
      Else
          MsgBox "Data Tidak Tersedia !", vbOKOnly + vbInformation, "Info"
          
          LvBP.ListItems.CLEAR
          lbldata.Caption = "Rows : 0"
          lbltotal.Caption = "Total : IDR 0 "
          'txt_total_ptp.Text = 0
      End If
      Me.MousePointer = vbNormal
End Sub



Private Sub cmd_download_Click()
    If LvBP.ListItems.Count = 0 Then
        MsgBox "Show Data Terlebih Dahulu!", vbOKOnly + vbInformation, "Informasi"
    Exit Sub
    Else
        Call My_Export_Excel_BP
    End If
End Sub

Private Sub My_Export_Excel_BP()
    Dim a           As Long
    Dim B           As Long
    Dim ExlObj      As Excel.Application
    Dim listcustid  As String
    Dim rs          As ADODB.Recordset
    Dim RS2         As ADODB.Recordset
    Dim iRow        As Integer
    Dim i           As Integer
    Dim sQuery      As String
    Dim AGENT As String

    
    sQuery = "SELECT * FROM mandiri.tbl_listbp WHERE f_bp = '1' ORDER BY agent"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    
    Set ExlObj = CreateObject("excel.application")
    ExlObj.Workbooks.ADD
    ExlObj.Visible = True
    
    ExlObj.Range("A1:N1").MergeCells = True
    'ExlObj.Range("A2:N2").MergeCells = True
    ExlObj.Range("A4:N4").Font.Bold = True
    
    
    With ExlObj.ActiveSheet
        .Cells(1, 1).Value = "List BP - Periode " & Format(DTPicker1.Value, "MMM-YYYY") & " "
        .Cells(1, 1).Font.Name = "Verdana"
        .Cells(1, 1).Font.Bold = True
        .Cells(4, 1).Value = "NO"
        .Cells(4, 2).Value = "CARD NUMBER"
        .Cells(4, 3).Value = "CH NAME"
        .Cells(4, 4).Value = "AGENT"
        .Cells(4, 5).Value = "PROMISEPAY"
        .Cells(4, 6).Value = "PROMISEDATE"
        .Cells(4, 7).Value = "PRODUCT"

        iRow = 4
        If rs.RecordCount > 0 Then
            ProgressBar1.Max = rs.RecordCount
            i = 0
            Do Until rs.EOF
                i = i + 1
                iRow = iRow + 1
                ProgressBar1.Value = rs.Bookmark
                .Cells(iRow, 1).Value = i
                .Cells(iRow, 2).Value = IIf(IsNull(rs!CustId), "", rs!CustId)
                .Cells(iRow, 3).Value = IIf(IsNull(rs!custname), "", rs!custname)
                .Cells(iRow, 4).Value = IIf(IsNull(rs!AGENT), "", rs!AGENT)
                .Cells(iRow, 5).Value = Format(IIf(IsNull(rs!PromisePay), "", rs!PromisePay), "##,###")
                .Cells(iRow, 6).Value = IIf(IsNull(rs("PromiseDate")), "", Format(rs("PromiseDate"), "DD-MM-YYYY"))
                .Cells(iRow, 7).Value = IIf(IsNull(rs!PRODUCT), "", rs!PRODUCT)
                rs.MoveNext
            Loop
        End If
    
        'OTOMATISASI CELL
        For iColom = 1 To 14
            ExlObj.Cells(4, iColom).EntireColumn.AutoFit
        Next
        
        MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
        ProgressBar1.Value = 0
        cmd_download.Enabled = True
    
        Set ExlObj = Nothing
        Set rs = Nothing

        'StartMeUp (Txtlocation.Text)
        'FILL COLOR CELL
        'ExlObj.Range(.Cells(NoUrut, 1), .Cells(NoUrut, 7)).Interior.Color = RGB(6, 207, 250)
    End With
End Sub


Private Sub cmd_showbp_Click()
    
    Call IsiTblBP
    Call UpdateTblBP
    Call IsiListBP
End Sub

Private Sub IsiTblBP()
    Dim AGENT As String
    Dim bulan_sekarang As String
    Dim tahun_sekarang As String
    
    tanggal_sekarang = Format(DTPicker1.Value, "yyyy-mm-dd")
    
    bulan_sekarang = Format(tanggal_sekarang, "MM")
    tahun_sekarang = Format(tanggal_sekarang, "YYYY")
    
    
    AGENT = cmb_agent.Text
    
    If cmb_agent = " " Then
        AGENT = ""
    End If
    
    M_OBJCONN.Execute "DELETE FROM mandiri.tbl_listbp"
    
    If cmb_agent = "ALL" Then
        M_OBJCONN.Execute "INSERT INTO mandiri.tbl_listbp(" & _
                          " agent, custid, promisedate, promisepay, inputdate, custname, product) " & _
                          " (SELECT b.agent, a.custid, promisedate, promisepay, inputdate, name, acc_type FROM ( " & _
                          " SELECT custid, promisedate, promisepay, inputdate, name, acc_type FROM mandiri.reportPTP where date_part('month', promisedate) = '" & bulan_sekarang & "' " & _
                          " AND date_part('year', promisedate) = '" & tahun_sekarang & "') as a " & _
                          " LEFT JOIN (SELECT custid, agent FROM mandiri.mgm) as b " & _
                          " ON a.custid = b.custid) "
    Else
        If f_team = False Then
            M_OBJCONN.Execute "INSERT INTO mandiri.tbl_listbp(" & _
                          " agent, custid, promisedate, promisepay, inputdate, custname, product) " & _
                          " (SELECT b.agent, a.custid, promisedate, promisepay, inputdate, name, acc_type FROM ( " & _
                          " SELECT custid, promisedate, promisepay, inputdate, name, acc_type FROM mandiri.reportPTP where date_part('month', promisedate) = '" & bulan_sekarang & "' " & _
                          " AND date_part('year', promisedate) = '" & tahun_sekarang & "' AND agent = '" & AGENT & "') as a " & _
                          " LEFT JOIN (SELECT custid, agent FROM mandiri.mgm) as b " & _
                          " ON a.custid = b.custid) "
        Else
            M_OBJCONN.Execute "INSERT INTO mandiri.tbl_listbp(" & _
                          " agent, custid, promisedate, promisepay, inputdate, custname, product) " & _
                          " (SELECT b.agent, a.custid, promisedate, promisepay, inputdate, name, acc_type FROM ( " & _
                          " SELECT custid, promisedate, promisepay, inputdate, name, acc_type FROM mandiri.reportPTP where date_part('month', promisedate) = '" & bulan_sekarang & "' " & _
                          " AND date_part('year', promisedate) = '" & tahun_sekarang & "' AND agent in (select userid from mandiri.usertbl where team = '" & cmb_team.Text & "' AND userid ilike  'D%')) as a " & _
                          " LEFT JOIN (SELECT custid, agent FROM mandiri.mgm) as b " & _
                          " ON a.custid = b.custid) "
        End If
    End If
End Sub

Private Function GetLPD() As String
    Dim sQuery As String
    Dim RsLPD As ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    
    sQuery = "SELECT MAX(paydate) FROM mandiri.tbllunas "
    Set RsLPD = New ADODB.Recordset
    RsLPD.CursorLocation = adUseClient
    RsLPD.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    GetLPD = Format(RsLPD!Max, "DD-MM-YYYY")
    
End Function

Private Sub UpdateTblBP()
    Dim bulan_sekarang As String
    Dim tahun_sekarang As String
    Dim custid_bayar As String
    Dim id_ptp As String
    
    tanggal_sekarang = Format(DTPicker1.Value, "yyyy-mm-dd")
    
    bulan_sekarang = Format(tanggal_sekarang, "MM")
    tahun_sekarang = Format(tanggal_sekarang, "YYYY")
    
    tanggal_max = GetLPD
    
    If Rs_list.state = 1 Then Rs_list.Close
    
    Rs_list.Open "SELECT a.id as id_ptp, a.custid as custid_ptp, coalesce(b.custid,'') as custid_bayar, agent, " & _
                 " custname, promisedate, promisepay, inputdate, product FROM ( " & _
                 " (SELECT id,agent, custid, promisedate, promisepay, inputdate, custname, product " & _
                 " FROM mandiri.tbl_listbp WHERE promisedate < '" & tanggal_max & "') as a LEFT JOIN " & _
                 " (SELECT distinct custid FROM mandiri.tbllunas " & _
                 " WHERE date_part('month',paydate) = '" & bulan_sekarang & "' AND date_part('year',paydate) = '" & tahun_sekarang & "' ) as b " & _
                 " ON a.custid = b.custid " & _
                 " ) order by custid_bayar asc "
    
    If Rs_list.RecordCount > 0 Then
        Do Until Rs_list.EOF
            custid_bayar = IIf(IsNull(Rs_list!custid_bayar), "", Rs_list!custid_bayar)
            id_ptp = IIf(IsNull(Rs_list!id_ptp), "", Rs_list!id_ptp)
                If custid_bayar = "" Then
                    M_OBJCONN.Execute "UPDATE mandiri.tbl_listbp SET f_bp = '1' WHERE id = '" & id_ptp & "' "
                End If
        Rs_list.MoveNext
        Loop
    Else
        'MsgBox "Data Tidak Tersedia !", vbOKOnly + vbInformation, "Info"
    Exit Sub
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call koneksi
    Call IsiAgent
    
    f_team = False
    
    
    DTPicker1.Value = Now
    
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        lbl_team.Visible = False
        cmb_team.Visible = False
        chk_team.Visible = False
        cmd_download.Visible = False
        'cmd_download_payment.Visible = False
    Else
        If Rs_list.state = 1 Then Rs_list.Close

        If Left(MDIForm1.TxtUsername.Text, 2) = "TL" Then
            Rs_list.Open "select userid from mandiri.usertbl where usertype = '1' AND userid ilike 'D%' AND  team = '" & MDIForm1.TxtUsername.Text & "' Order by userid"
        Else
            Rs_list.Open "SELECT DISTINCT team from mandiri.usertbl WHERE team ilike 'TL%'"
        End If

    End If
    Call HeaderLvBP
End Sub

Private Sub IsiAgent()
    If Rs_list.state = 1 Then Rs_list.Close
    
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        Rs_list.Open "select userid from mandiri.usertbl where usertype = '1' and userid = '" & MDIForm1.TxtUsername.Text & "' Order by userid"
    ElseIf UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        Rs_list.Open "select userid from mandiri.usertbl where usertype = '1' AND userid ilike 'D%' AND  team = '" & MDIForm1.TxtUsername.Text & "' Order by userid"
    Else
        Rs_list.Open "select userid from mandiri.usertbl where usertype = '1' and userid ilike 'D%' Order by userid"
    End If
    
    cmb_agent.AddItem ""
    
    If UCase(MDIForm1.txtlevel.Text) <> "AGENT" And UCase(MDIForm1.txtlevel.Text) <> "TEAMLEADER" Then
        cmb_agent.AddItem "ALL"
    End If
    
    While Not Rs_list.EOF
        cmb_agent.AddItem Rs_list("USERID")
        Rs_list.MoveNext
    Wend
    cmb_agent.ListIndex = 1
End Sub

Private Sub HeaderLvBP()
    LvBP.ColumnHeaders.ADD , , "No", 560
    LvBP.ColumnHeaders.ADD , , "Custid", 2100
    LvBP.ColumnHeaders.ADD , , "Agent", 1000
    LvBP.ColumnHeaders.ADD , , "PromisePay", 1300
    LvBP.ColumnHeaders.ADD , , "PromiseDate", 1300
    LvBP.ColumnHeaders.ADD , , "CH Name", 2350
    LvBP.ColumnHeaders.ADD , , "Product", 1500
End Sub
   
Private Sub LvBP_DblClick()
    If LvBP.ListItems.Count > 0 Then
        VIEW_MGMDATA.txtnocard.Text = LvBP.SelectedItem.SubItems(1)
        Form_monthly_BP.Hide
        VIEW_MGMDATA.Show
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If
End Sub


