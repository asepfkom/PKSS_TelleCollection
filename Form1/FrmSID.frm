VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmSID 
   Caption         =   "LIST SID"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11325
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11325
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   10440
      TabIndex        =   8
      Top             =   6720
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "SEARCH ACCOUNT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtlocation 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   4800
         Width           =   2175
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Import"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   4200
         Width           =   2535
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   840
         Top             =   5040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add To List"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Export To Excel"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   6360
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox TxtCustid 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   150
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   180
         TabIndex        =   3
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search by list"
         Height          =   375
         Left            =   180
         TabIndex        =   2
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   2760
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label5 
         Caption         =   "Custid:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6375
      Left            =   3120
      TabIndex        =   12
      Top             =   240
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   11245
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Has Been Export"
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
      Height          =   255
      Left            =   8040
      TabIndex        =   16
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   5640
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "New Data"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   15
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   7680
      Top             =   6720
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9840
      TabIndex        =   9
      Top             =   6795
      Width           =   855
   End
End
Attribute VB_Name = "FrmSID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs1 As ADODB.Recordset
Private sqlstr As String
Private sCari_list As Boolean

Private Sub Check1_Click()
    Dim xx As Integer
    
    For xx = 1 To ListView1.ListItems.Count
        If Check1.Value = vbChecked Then
            ListView1.ListItems(xx).Checked = True
        Else
            ListView1.ListItems(xx).Checked = False
        End If
    Next xx
End Sub

Private Sub Command1_Click()
    Dim ListItem As ListItem
    Dim cust_exist As Boolean
    Dim xx As Integer
    Dim list_cust_sel As String
    Dim flag_export As Integer
    

    sqlstr = "SELECT custid,acc_type,name,dob,ktpno,'" & MDIForm1.TxtUsername.text & "/ID/HBAP/HSBC',mother FROM mgm WHERE custid IS NOT NULL "
    
    If sCari_list = True Then
        For xx = 0 To list1.ListCount - 1
            list_cust_sel = list_cust_sel & "'" & list1.list(xx) & "',"
        Next xx
        
        list_cust_sel = Mid(list_cust_sel, 1, Len(list_cust_sel) - 1)

        sqlstr = sqlstr & " AND custid in (" & list_cust_sel & ")"
    Else
        If txtcustid.text <> "" Then
            sqlstr = sqlstr & " AND custid like '%" & txtcustid.text & "%'"
        End If
    End If
    
    M_OBJCONN.Execute "DELETE FROM tblcpa_sid;"
    M_OBJCONN.Execute "INSERT INTO tblcpa_sid(ref_num,prd_type,name_,dob,id_no,requestor_name,mother_maiden_name) " & sqlstr
    M_OBJCONN.Execute "UPDATE tblcpa_sid SET flag_export_sid = '1' WHERE ref_num in (SELECT custid FROM tbl_temp_export_sid)"
    ListView1.ListItems.CLEAR
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT * FROM tblcpa_sid;"
    
    If rs1.RecordCount > 0 Then
        cust_exist = False
        If txtcustid.text <> "" Then
            For xx = 0 To list1.ListCount - 1
                If txtcustid.text = list1.list(xx) Then
                    cust_exist = True
                End If
            Next xx
            ' Add list customer
            If cust_exist = False Then
                list1.AddItem txtcustid.text
            End If
        End If
        
        Do Until rs1.EOF
            'ListView1.ForeColor = vbGreen
            Set ListItem = ListView1.ListItems.ADD(, , cnull(rs1!ref_num))
            ListItem.SubItems(1) = cnull(rs1!prd_type)
            ListItem.SubItems(2) = clean_sid(cnull(rs1!name_))
            ListItem.SubItems(3) = Format(cnull(rs1!DOB), "DDMMYYYY")
            ListItem.SubItems(4) = cnull(rs1!id_no)
            ListItem.SubItems(5) = cnull(rs1!requestor_name)
            ListItem.SubItems(6) = cnull(rs1!mother_maiden_name)
            flag_export = cnull(rs1!flag_export_sid)
            If flag_export = 1 Then
                For randy = 1 To 6
                    ListItem.ForeColor = vbGreen
                    ListItem.ListSubItems(randy).ForeColor = vbGreen
                Next randy
            End If
            rs1.MoveNext
        Loop
    End If
    
    Text1.text = rs1.RecordCount
End Sub



Private Sub Command2_Click()
    Dim sLokasiExcel As String
    Dim xListCustid As String
    Dim sQuery, CustId, custid_sid As String

    CD_save.Filter = "Excel Files |*.xls"
    'CD_save.Filter = "Excel Files|*.xls;*.xlsx"
    CD_save.ShowSave
    
    sLokasiExcel = CD_save.FileName
    
    For i = 1 To ListView1.ListItems.Count
        xListCustid = xListCustid & "'" & ListView1.ListItems(i).text & "',"
    Next i
    
    xListCustid = Mid(xListCustid, 1, Len(xListCustid) - 1)
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT * FROM tblcpa_sid WHERE ref_num in (" & xListCustid & ") "
    
    'UPDATE tblcpa_sid
     M_OBJCONN.Execute "INSERT INTO tblcpa_sid(ref_num,prd_type,name_,dob,id_no,requestor_name,mother_maiden_name) " & sqlstr

    
    Call ConvertToExcel_me(rs1, sLokasiExcel)
    
    'RANDY
    If rs1.State = 1 Then rs1.Close
    
    rs1.Open "SELECT * FROM (" & _
             "(SELECT custid as custid_mgm FROM mgm WHERE custid in (" & xListCustid & ")) " & _
             "AS a LEFT JOIN " & _
             "(SELECT custid as custid_sid FROM tbl_temp_export_sid WHERE custid in(" & xListCustid & ")) " & _
             "As b " & _
             "On a.custid_mgm = b.custid_sid )"
    
    If rs1.RecordCount > 0 Then
        Do Until rs1.EOF
            CustId = cnull(rs1!custid_mgm)
            custid_sid = cnull(rs1!custid_sid)
            If custid_sid = "" Then
                M_OBJCONN.Execute "INSERT INTO tbl_temp_export_sid VALUES ('" & CustId & "')"
            End If
            rs1.MoveNext
        Loop
    End If
End Sub

Private Sub Command3_Click()
    If list1.ListCount > 0 Then
        sCari_list = True
        Call Command1_Click
        s_cari_list = False
    End If
End Sub

Private Sub Command4_Click()
    Dim cek, K, W, index_hapus As Integer
'    If List1.ListCount > 0 Then
'        List1.RemoveItem List1.ListIndex
'    End If
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    cek = 0
    
    For K = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
    If cek = 0 Then
        MsgBox "Pilih Data Yang Akan Di-Remove!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
a:
    index_hapus = ListView1.ListItems.Count
    

    For W = 1 To index_hapus
        If ListView1.ListItems(W).Checked = True Then
           ListView1.ListItems.Remove ListView1.ListItems(W).Index
            GoTo a
        End If
    Next W
 
End Sub

Private Sub Command5_Click()
    Dim cst_exist As Boolean

    cst_exist = False
    If txtcustid.text <> "" Then
        For xx = 0 To list1.ListCount - 1
            If txtcustid.text = list1.list(xx) Then
                cst_exist = True
            End If
        Next xx
        ' Add list customer
        If cst_exist = False Then
            list1.AddItem txtcustid.text
        End If
    End If
End Sub

Private Sub Command6_Click()
    Dim OBJRECORD As ADODB.Recordset
    Dim ssql As String
    Dim CustId As String
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim M_objrs As ADODB.Recordset
    Dim M_XLSCONN As New ADODB.Connection
    
    On Error GoTo Salah
    
    With CD_save
        .DialogTitle = "Pilih file excel"
        .Filter = "Excel Files|*.xls"
        '.Filter = "Excel Files|*.xls;*.xlsx"
        .ShowOpen
    End With
        
    txtlocation.text = CD_save.FileName
    
    If CD_save.FileName = "" Then Exit Sub
    If M_XLSCONN.State = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CD_save.FileName & ";Extended Properties=Excel 8.0;"
    'M_XLSCONN.OpenSchema (adSchemaTables)
    'Set rs1 = M_XLSCONN.OpenSchema(adSchemaTables)
    
    Set OBJRECORD = New ADODB.Recordset
    OBJRECORD.CursorLocation = adUseClient
    ssql = "SELECT * FROM [Sheet1$] where [custid] is not null"

    OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
    If OBJRECORD.RecordCount > 0 Then
        Do Until OBJRECORD.EOF
            list1.AddItem cnull(OBJRECORD!CustId)
            OBJRECORD.MoveNext
        Loop
    End If
    Set OBJRECORD = Nothing
    Exit Sub
Salah:
    MsgBox err.Description, vbCritical
End Sub

Private Sub Form_Load()
    Call koneksi
    
    ListView1.ColumnHeaders.ADD , , "Cust ID"
    ListView1.ColumnHeaders.ADD , , "Product"
    ListView1.ColumnHeaders.ADD , , "Name"
    ListView1.ColumnHeaders.ADD , , "Dob"
    ListView1.ColumnHeaders.ADD , , "ID No"
    ListView1.ColumnHeaders.ADD , , "Requestor"
    ListView1.ColumnHeaders.ADD , , "Mother"
End Sub

Private Sub koneksi()
    Set rs1 = New ADODB.Recordset
    rs1.CursorLocation = adUseClient
    rs1.CursorType = adOpenDynamic
    rs1.ActiveConnection = M_OBJCONN
    rs1.LockType = adLockOptimistic
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs1 = Nothing
End Sub

Private Function clean_sid(sKata As String) As String
    Dim stemp_kata As String
    
    stemp_kata = Trim(sKata)
    stemp_kata = Replace(stemp_kata, "MRS", "")
    stemp_kata = Replace(stemp_kata, "MR", "")
    stemp_kata = Replace(stemp_kata, "MISS", "")
    stemp_kata = Replace(stemp_kata, ".", "")
    stemp_kata = Replace(stemp_kata, ",", "")
    stemp_kata = Trim(stemp_kata)
    
    clean_sid = stemp_kata
End Function

Private Sub ConvertToExcel_me(M_objrs As ADODB.Recordset, TxtPath As String)
    Dim ListItem        As ListItem
    Dim cmdsql_update   As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i As Double
    Dim m_msgbox As String
    
    i = 1
  
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath = Empty Then
        MsgBox "Nama file tidak boleh kosong, download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Set excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
    
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
    
    On Error GoTo Salah
    'Proses pengsisian nama field ke excel
    Dim X, Y    As Double
    If M_objrs.State = 1 Then
        X = 0
        Y = M_objrs.fields().Count - 1
        Do Until X > Y
            DoEvents
            objSheet.Cells(1, i).Value = UCase(Replace(CStr(M_objrs.fields(X).Name), "_", " "))
            i = i + 1
            X = X + 1
        Loop
    End If
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    'objSheet.Range("A2").CopyFromRecordset M_Objrs '-> Proses pengisian data dimulai dari Cell A2
    
    i = 2
    With M_objrs
        .MoveFirst
        While Not .EOF
            objSheet.Cells(i, 1).Value = cnull(!capture_date)
            objSheet.Cells(i, 2).Value = "'" & cnull(!ref_num)
            objSheet.Cells(i, 3).Value = "COLECTION_" & Format(cnull(!input_date), "MMDDYYYY")
            objSheet.Cells(i, 4).Value = clean_sid(cnull(!name_))
            objSheet.Cells(i, 5).Value = "'" & Format((cnull(!DOB)), "MM/DD/YYYY")
            objSheet.Cells(i, 6).Value = "'" & cnull(!id_no)
            objSheet.Cells(i, 7).Value = cnull(!bank1)
            objSheet.Cells(i, 8).Value = cnull(!marketing_source)
            objSheet.Cells(i, 9).Value = cnull(!input_date)
            objSheet.Cells(i, 10).Value = cnull(!requestor_name)
            objSheet.Cells(i, 11).Value = cnull(!user_)
            objSheet.Cells(i, 12).Value = cnull(!result)
            objSheet.Cells(i, 13).Value = cnull(!idi)
            objSheet.Cells(i, 14).Value = cnull(!diff_id)
            objSheet.Cells(i, 15).Value = cnull(!gender)
            objSheet.Cells(i, 16).Value = cnull(!mother_maiden_name)
            objSheet.Cells(i, 17).Value = cnull(!info_tujuan_permintaan_data_bi)
            i = i + 1
            .MoveNext
        Wend
    End With
    
    objBook.SaveAs TxtPath, xlWorkbookNormal
    objExcel.Quit
    
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    
Salah:
    Exit Sub
End Sub

