VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form form_upload_customer 
   Caption         =   "Upload Data"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13065
   LinkTopic       =   "Form2"
   ScaleHeight     =   9375
   ScaleWidth      =   13065
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      Begin VB.CommandButton cmdbrowse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "...."
         Height          =   315
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox txtlocation 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   8355
      End
      Begin VB.ComboBox cbosheet 
         Height          =   315
         Left            =   1070
         TabIndex        =   1
         Top             =   660
         Width           =   2535
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   315
         Width           =   795
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sheet"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   555
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   14631
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Data All"
      TabPicture(0)   =   "form_upload_customer.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data New"
      TabPicture(1)   =   "form_upload_customer.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Data Duplicate"
      TabPicture(2)   =   "form_upload_customer.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         BackColor       =   &H80000003&
         Height          =   9495
         Left            =   -75000
         TabIndex        =   17
         Top             =   360
         Width           =   13095
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   18
            Top             =   7515
            Width           =   1095
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   7095
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   12515
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Upload :"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   7560
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000003&
         Height          =   9495
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   13095
         Begin VB.CommandButton Command1 
            Caption         =   "Upload"
            Height          =   375
            Left            =   11640
            TabIndex        =   14
            Top             =   7440
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   13
            Top             =   7515
            Width           =   1095
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   7095
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   12515
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Upload :"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   7560
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000003&
         Height          =   9495
         Left            =   -75000
         TabIndex        =   7
         Top             =   360
         Width           =   13095
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   9
            Top             =   7515
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Upload"
            Height          =   375
            Left            =   11640
            TabIndex        =   8
            Top             =   7440
            Width           =   1335
         End
         Begin MSComctlLib.ListView lstview 
            Height          =   7095
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   12515
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Upload :"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   7560
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "form_upload_customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public M_XLSCONN As New ADODB.Connection
'On Error Resume Next

Private Sub cbosheet_Click()
    If txtlocation.text <> "" Then
        Call sheet("tbl_upload_customer_temp", cbosheet.text)
        Call isilist("tbl_upload_customer_temp", lstview, "all")
        Call isilist("tbl_upload_customer_temp", ListView1, "new")
        Call isilist("tbl_upload_customer_temp", ListView2, "duplicate")
    End If
End Sub

Private Sub cmdbrowse_Click()
    With CommonDialog1
        .DialogTitle = "Import From File"
        '.Filter = "Xlsx Files (*.xlsx)|*.xlsx"
        .Filter = "Excel Files|*.xls;*.xlsx"
        .ShowOpen
    End With
    
        If CommonDialog1.FileName = "" Then Exit Sub
        If M_XLSCONN.State = adStateOpen Then M_XLSCONN.Close
        M_XLSCONN.Open "Provider = Microsoft.ACE.OLEDB.12.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 12.0;"
        Set M_objrs = M_XLSCONN.OpenSchema(adSchemaTables)
        
        txtlocation.text = ""
        txtlocation.text = CommonDialog1.FileName
        cbosheet.CLEAR
        
        If M_objrs.EOF And M_objrs.BOF Then Exit Sub
        While Not M_objrs.EOF
            cbosheet.AddItem IIf(IsNull(M_objrs!table_name), "", M_objrs!table_name)
            M_objrs.MoveNext
        Wend
    
    M_objrs.Close
    M_XLSCONN.Close
    Set M_objrs = Nothing
    Set M_XLSCONN = Nothing
End Sub

Private Sub sheet(a As String, B As String)
    If M_XLSCONN.State = adStateOpen Then M_XLSCONN.Close
    
    M_XLSCONN.Open "Provider = Microsoft.ACE.OLEDB.12.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 12.0;"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    ssql = "SELECT  *  FROM [" & B & "]"
    M_objrs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
    
    If M_objrs.EOF And M_objrs.BOF Then Exit Sub
    headers = ""
    For i = 0 To M_objrs.fields.Count - 1
        headers = headers & """" & M_objrs.fields(i).Name & """" & " varchar,"
    Next i
    headers = Left(headers, Len(headers) - 1)
    
    M_OBJCONN.Execute "Drop table if exists " & a & ";"
    Creates = "Create table " & a & " ( " & headers & " );"
    M_OBJCONN.Execute Creates
    
    For i = 0 To M_objrs.RecordCount - 1
        B = ""
        For c = 0 To M_objrs.fields.Count - 1
            B = B + "'" & M_objrs(c) & "',"
        Next c
            B = Left(B, Len(B) - 1)
        M_OBJCONN.Execute "insert into " & a & " values (" & B & ");"
        M_objrs.MoveNext
    Next i
    
    M_XLSCONN.Close
    Set M_objrs = Nothing
    Set rs = Nothing
    Set M_XLSCONN = Nothing
End Sub

Private Sub isilist(a As String, B As ListView, c As String)
    Dim c1 As Integer
    On Error GoTo lanjut

    sStrsql = "select column_name from information_schema.columns  where table_name = '" & a & "' order by ordinal_position"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    B.ListItems.CLEAR
    B.ColumnHeaders.CLEAR
    If M_objrs.RecordCount > 0 Then
        c1 = M_objrs.RecordCount
        With B.ColumnHeaders
            For i = 1 To c1
                .ADD i, , cnull(M_objrs!column_name)
                M_objrs.MoveNext
            Next i
        End With
    End If
    Set M_objrs = Nothing
    
    If c = "all" Then
        sStrsql = "select * from " & a
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        While Not M_objrs.EOF
            
            Set ListItem = B.ListItems.ADD(, , cnull(M_objrs(0)))
                For i = 1 To c1 - 1
                    ListItem.SubItems(i) = cnull(M_objrs(i))
                Next i
            M_objrs.MoveNext
        Wend
        Text1.text = lstview.ListItems.Count
        SSTab1.TabCaption(0) = SSTab1.TabCaption(0) & " (" & Text1.text & ")"
    ElseIf c = "new" Then
        sStrsql = "select * from " & a & " where ""NOMER_KARTU"" not in (select custid from mgm);"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        While Not M_objrs.EOF
            
            Set ListItem = B.ListItems.ADD(, , cnull(M_objrs(0)))
                For i = 1 To c1 - 1
                    ListItem.SubItems(i) = cnull(M_objrs(i))
                Next i
            M_objrs.MoveNext
        Wend
        Text2.text = ListView1.ListItems.Count
        SSTab1.TabCaption(1) = SSTab1.TabCaption(1) & " (" & Text2.text & ")"
    ElseIf c = "duplicate" Then
        sStrsql = "select * from " & a & " where ""NOMER_KARTU"" in (select custid from mgm);"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        While Not M_objrs.EOF
            
            Set ListItem = B.ListItems.ADD(, , cnull(M_objrs(0)))
                For i = 1 To c1 - 1
                    ListItem.SubItems(i) = cnull(M_objrs(i))
                Next i
            M_objrs.MoveNext
        Wend
        Text3.text = ListView2.ListItems.Count
        SSTab1.TabCaption(2) = SSTab1.TabCaption(2) & " (" & Text3.text & ")"
    End If
    
    Set M_objrs = Nothing
    Exit Sub
lanjut:
    MsgBox "Format Upload Salah"
End Sub

Private Sub upload(up As String)
    sStrsql = "select * from maping_upload order by 1"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    strdb = ""
    strexcel = "("
    a = ""
    B = 0
    c = 0
    If M_objrs.RecordCount > 0 Then
        For i = 1 To M_objrs.RecordCount
            If a <> M_objrs!Db Then
                If B = 1 Then
                    If M_objrs!Excel = "CO_Date" Then
                        strdb = strdb & ",to_date(""" & M_objrs!Excel & """,'yyyy-mm-dd'),"
                    Else
                        strdb = strdb & ",""" & M_objrs!Excel & ""","
                    End If
                    B = 0
                    GoTo bawah
                End If
                If M_objrs!Excel = "CM_TOT_BALANCE" Or M_objrs!Excel = "CM_CYCLE" Then
                    strdb = strdb & """" & M_objrs!Excel & ""","
                Else
                    strdb = strdb & """" & M_objrs!Excel & ""","
                End If
bawah:
                strexcel = strexcel & """" & M_objrs!Db & ""","
                c = 1
            Else
                If c = 1 Then
                    strdb = Left(strdb, Len(strdb) - 1)
                End If
                strdb = strdb & "||" & """" & M_objrs!Excel & """"
 
                B = 1
                c = 0
            End If
            a = M_objrs!Db
            M_objrs.MoveNext
        Next i
        strdb = Left(strdb, Len(strdb) - 1) & ""
        strexcel = Left(strexcel, Len(strexcel) - 1) & ")"
        
        If up = "all" Then
            'M_OBJCONN.Execute "DELETE from mgm where custid in (select custid from tbl_upload_customer_temp)"
            M_OBJCONN.Execute "insert into mgm " & strexcel & " select " & strdb & " from tbl_upload_customer_temp;"
        ElseIf up = "new" Then
            M_OBJCONN.Execute "insert into mgm " & strexcel & " select " & strdb & " from tbl_upload_customer_temp;"
        End If
        
        Call bikin_campaign
        Call log_upload
        MsgBox "Done"
    End If
End Sub

Private Sub Command1_Click()
    Call upload("new")
End Sub

Private Sub Command2_Click()
    Call upload("all")
End Sub

Private Sub log_upload()
    sStrsql = "select column_name from information_schema.columns  where table_name = 'upload_log' order by ordinal_position"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount = 0 Then
        screate = "create table upload_log (" & vbCrLf
        screate = screate & " id serial not null, " & vbCrLf
        screate = screate & " location varchar, " & vbCrLf
        screate = screate & " sheet varchar, " & vbCrLf
        screate = screate & " jml_data int, " & vbCrLf
        screate = screate & " jml_data_new int, " & vbCrLf
        screate = screate & " jml_data_dup int, " & vbCrLf
        screate = screate & " tgl_upload timestamp default now(), " & vbCrLf
        screate = screate & " users varchar )"
        M_OBJCONN.Execute screate
    End If
    
    sinsert = "insert into upload_log (location, sheet, jml_data, jml_data_new, jml_data_dup, users) values " & vbCrLf
    sinsert = sinsert & "('" & txtlocation.text & "', '" & cbosheet.text & "', " & Text1.text & ", " & Text2.text & ", " & Text3.text & ", '" & MDIForm1.TxtUsername.text & "') "
    M_OBJCONN.Execute sinsert
    
End Sub

Private Sub bikin_campaign()
    supdate = "update mgm set recsource = 'PKSS'||to_char(now(),'yyyymmddhhmiss') where coalesce(recsource,'') = '';"
    M_OBJCONN.Execute supdate
End Sub


