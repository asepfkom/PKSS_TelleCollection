VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form_setting_upload_address 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup Map Upload"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14820
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Setting Upload"
      Height          =   1845
      Left            =   0
      TabIndex        =   7
      Top             =   450
      Width           =   14775
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   12360
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "...."
         Height          =   315
         Left            =   9990
         TabIndex        =   12
         Top             =   1050
         Width           =   615
      End
      Begin VB.TextBox txtlocation 
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1050
         Width           =   8445
      End
      Begin VB.ComboBox cbosheet 
         Height          =   315
         Left            =   1500
         TabIndex        =   10
         Top             =   1410
         Width           =   2385
      End
      Begin VB.ComboBox cbomapsource 
         Height          =   315
         Left            =   1500
         TabIndex        =   9
         Tag             =   "0"
         Top             =   210
         Width           =   3615
      End
      Begin VB.TextBox txtmapdesc 
         Height          =   315
         Left            =   1500
         TabIndex        =   8
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "Location"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   1110
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Sheet"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label cbomapping 
         Caption         =   "Mapping Source"
         Height          =   285
         Left            =   150
         TabIndex        =   14
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Mapping Desc"
         Height          =   285
         Left            =   180
         TabIndex        =   13
         Top             =   690
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Setting"
      Height          =   6975
      Left            =   30
      TabIndex        =   0
      Top             =   2370
      Width           =   14745
      Begin VB.CommandButton cmdsavesetting 
         Caption         =   "Save Setting"
         Height          =   525
         Left            =   11220
         TabIndex        =   2
         Top             =   6390
         Width           =   1575
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "Exit"
         Height          =   525
         Left            =   12900
         TabIndex        =   1
         Top             =   6390
         Width           =   1575
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5955
         Left            =   90
         TabIndex        =   3
         Top             =   240
         Width           =   14625
         _ExtentX        =   25797
         _ExtentY        =   10504
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "View Data upload "
         TabPicture(0)   =   "frm_setting_upload_address.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fx_mapping"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Cboexecelmap"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "History user "
         TabPicture(1)   =   "frm_setting_upload_address.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lsthstuser"
         Tab(1).ControlCount=   1
         Begin VB.ComboBox Cboexecelmap 
            Height          =   315
            Left            =   2520
            TabIndex        =   4
            Top             =   1770
            Visible         =   0   'False
            Width           =   1605
         End
         Begin MSComctlLib.ListView lsthstuser 
            Height          =   5355
            Left            =   -74880
            TabIndex        =   5
            Top             =   450
            Width           =   14385
            _ExtentX        =   25374
            _ExtentY        =   9446
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
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSFlexGridLib.MSFlexGrid fx_mapping 
            Height          =   5415
            Left            =   120
            TabIndex        =   6
            Top             =   450
            Width           =   14175
            _ExtentX        =   25003
            _ExtentY        =   9551
            _Version        =   393216
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   60
         X2              =   14700
         Y1              =   6300
         Y2              =   6300
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup Map Upload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   540
      TabIndex        =   17
      Top             =   30
      Width           =   3585
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   0
      Left            =   0
      Picture         =   "frm_setting_upload_address.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   465
      Left            =   0
      Picture         =   "frm_setting_upload_address.frx":0B42
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "Form_setting_upload_address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public M_XLSCONN As New ADODB.Connection

Private Sub Cboexecelmap_Click()
  fx_mapping.TextMatrix(fx_mapping.row, 2) = Cboexecelmap.text
  '  fx_mapping.SetFocus
  fx_mapping_Click
End Sub

Private Sub cbomapsource_Click()
    cbomapsource_LostFocus
End Sub

Private Sub cbomapsource_DropDown()
    loadCboMap
End Sub

Private Sub cbomapsource_LostFocus()

ssql = "SELECT * FROM  tbl_setting_upload_address WHERE KODE_SOURCE='" + cbomapsource + "'"
Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
        M_objrs.MoveFirst
        MsgBox "Data Sudah Ada", vbInformation + vbOKOnly, "Information"
        txtmapdesc.text = IIf(IsNull(M_objrs("nama_source")), "", M_objrs("nama_source"))
        txtlocation.text = IIf(IsNull(M_objrs("location_source")), "", M_objrs("location_source"))
        cbosheet.text = IIf(IsNull(M_objrs("table_source")), "", M_objrs("table_source"))
        findFx cbomapsource.text, True
        Cboexecelmap.Tag = 1
    Else
     findFx "", False
     Cboexecelmap.Tag = 0
    End If
    Set M_objrs = Nothing
 End Sub



Private Sub cbosheet_Click()
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & cbosheet.text & "] "
        M_objrs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Cboexecelmap.CLEAR
        Cboexecelmap.AddItem ""
        If M_objrs.EOF And M_objrs.BOF Then Exit Sub
            For i = 0 To M_objrs.fields.Count - 1
                On Error Resume Next
                Cboexecelmap.AddItem M_objrs.fields(i).Name
            Next i
    Set M_objrs = Nothing
End Sub

Private Sub cmdbrowse_Click()
  Dim dir_listbulantem$
    With CommonDialog1
        .DialogTitle = "Import From File"
        '.Filter = "Excel Files|*.xls;*.xlsx"
        .Filter = "Excel Files|*.xls"
        .ShowOpen
    End With
    txtlocation.text = ""
    If CommonDialog1.FileName = "" Then Exit Sub
    txtlocation.text = CommonDialog1.FileName
    If M_XLSCONN.State = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 8.0;"
    Set M_objrs = M_XLSCONN.OpenSchema(adSchemaTables)
    cbosheet.CLEAR
    If M_objrs.EOF And M_objrs.BOF Then Exit Sub
    While Not M_objrs.EOF
        cbosheet.AddItem IIf(IsNull(M_objrs!table_name), "", M_objrs!table_name)
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    
    
End Sub
Public Sub findFx(ByVal xCodeMap As String, bBool As Boolean)
If bBool = False Then


    sStrsql = " select *,'' as field_destination from ("
    sStrsql = sStrsql + " SELECT column_name as nama_kolom From information_schema.Columns WHERE table_name='tbl_address' and table_schema='public'"
    sStrsql = sStrsql + "  ORDER BY ordinal_position) as tblbaru "

Else
   sStrsql = " select nama_kolom,field_destination from ( "
    sStrsql = sStrsql + " select * from ( "
    sStrsql = sStrsql + " SELECT column_name as nama_kolom From information_schema.Columns WHERE table_name='tbl_address' and table_schema='public'"
    sStrsql = sStrsql + "  ORDER BY ordinal_position) as tblbaru "
    sStrsql = sStrsql + " full join  ( "
    sStrsql = sStrsql + "  select field_source,field_destination from tbl_setting_upload_address where kode_source='" + cbomapsource.text + "' ) "
    sStrsql = sStrsql + " as tbldua on tblbaru.nama_kolom =tbldua.field_source) as tblbaru "
End If
  
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        fx_mapping.CLEAR
        CreateFx_Upload
        If M_objrs.EOF And M_objrs.BOF Then Exit Sub
            fx_mapping.Rows = 2
        For i = 1 To M_objrs.RecordCount
            fx_mapping.TextMatrix(i, 1) = IIf(IsNull(M_objrs("nama_kolom")), "", M_objrs("nama_kolom"))
            fx_mapping.TextMatrix(i, 2) = IIf(IsNull(M_objrs("field_destination")), "", M_objrs("field_destination"))
            fx_mapping.Rows = fx_mapping.Rows + 1
            M_objrs.MoveNext
        Next i
        
        fx_mapping.Rows = M_objrs.RecordCount + 1
        Set M_objrs = Nothing
End Sub
Public Sub CreateFx_Upload()
    With fx_mapping
        .Rows = 2
        .Cols = 3
        .TextMatrix(0, 1) = "Delta Net Field(s)"
        .ColWidth(1) = 3000
        .TextMatrix(0, 2) = "Excel Field(s)"
        .ColWidth(2) = 3000
        .RowHeightMin = Cboexecelmap.Height
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdsavesetting_Click()
Dim lst As ListItem
Dim RSTEMP As New ADODB.Recordset
            If cbomapsource.text = "" Then
                MsgBox "Kode map belum Di isi", vbOKOnly, "Information"
                Exit Sub
            End If
    
            If txtmapdesc.text = "" Then
                MsgBox "Description map belum Di isi", vbOKOnly, "Information"
                Exit Sub
            End If
    
            If cbosheet.text = "" Then
                MsgBox "Sheets Belum di isi", vbOKOnly, "Information"
                Exit Sub
            End If
    
            ssql = "DELETE FROM tbl_setting_upload_address WHERE kode_source ='" & cbomapsource.text & "'"
            M_OBJCONN.Execute (ssql)

            For brs = 1 To fx_mapping.Rows - 1
                If fx_mapping.TextMatrix(brs, 2) <> vbNullString Then
                    ssql = "INSERT INTO tbl_setting_upload_address (kode_source,nama_source,table_source ,location_source,field_source , field_destination) "
                    ssql = ssql & "VALUES "
                    With fx_mapping
                        ssql = ssql & "('" & cbomapsource.text & "', "
                        ssql = ssql & "'" & txtmapdesc.text & "', "
                        ssql = ssql & "'" & cbosheet.text & "', "
                        ssql = ssql & "'" & Replace(txtlocation.text, "\", "/") & "', "
                        ssql = ssql & "'" & .TextMatrix(brs, 1) & "', "
                        ssql = ssql & "'" & .TextMatrix(brs, 2) & "')"
                    End With
        M_OBJCONN.Execute (ssql)
        End If
    Next brs
    
 
    If Cboexecelmap.Tag = 0 Then
        MsgBox "Data Telah Di simpan ", vbOKOnly, "Information"
        sAction = "New Mapping"
    Else
        MsgBox "Data Telah Di Edit ", vbOKOnly, "Information"
        sAction = "Edit Mapping"
    End If
    findFx "", False
    
    strsql = "insert into tbl_hst_setting_upload (user_input,action_user) values ('" + MDIForm1.TxtUsername.text + "','" + sAction + "') "
    M_OBJCONN.Execute (strsql)
    Set lst = lsthstuser.ListItems.ADD(, , lsthstuser.ListItems.Count + 1)
        lst.SubItems(1) = Format(Date, "dd/mm/yyyy")
        lst.SubItems(2) = MDIForm1.TxtUsername.text
        lst.SubItems(3) = sAction
    
End Sub

Private Sub Form_Load()
    CreateFx_Upload
    create_header_hst_setting_upload
    load_hst_setting_upload
    findFx "", False
End Sub
Private Sub fx_mapping_Click()
 Select Case fx_mapping.col
    Case 2
        Cboexecelmap.Top = fx_mapping.CellTop + fx_mapping.Top
        Cboexecelmap.Left = fx_mapping.CellLeft + fx_mapping.Left
        Cboexecelmap.Width = fx_mapping.CellWidth
        Cboexecelmap.Visible = True
        Cboexecelmap.SetFocus
        If Not (fx_mapping.text = "") Then
            Cboexecelmap.text = fx_mapping.TextMatrix(fx_mapping.row, fx_mapping.col)
            Else
            Cboexecelmap.text = ""
            
        End If
    End Select
End Sub
Public Sub create_header_hst_setting_upload()
    lsthstuser.ColumnHeaders.ADD 1, , "No", 5 * TXT
    lsthstuser.ColumnHeaders.ADD 2, , "Tgl_Insert", 15 * TXT
    lsthstuser.ColumnHeaders.ADD 3, , "User", 15 * TXT
    lsthstuser.ColumnHeaders.ADD 4, , "Action", 7 * TXT
End Sub
Public Sub load_hst_setting_upload()
Dim list As ListItem
Dim no As Double
sStrsql = "select * from tbl_hst_setting_upload "
Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    no = 0
    lsthstuser.ListItems.CLEAR
    While Not M_objrs.EOF
        no = no + 1
        Set list = lsthstuser.ListItems.ADD(, , no)
            list.SubItems(1) = Format(IIf(IsNull(M_objrs!tgl_insert), "", M_objrs!tgl_insert), "dd/mm/yyyy")
            list.SubItems(2) = IIf(IsNull(M_objrs!user_input), "", M_objrs!user_input)
            list.SubItems(3) = IIf(IsNull(M_objrs!action_user), "", M_objrs!action_user)
        M_objrs.MoveNext
    Wend
   
Set M_objrs = Nothing
End Sub
Public Sub loadCboMap()
cbomapsource.CLEAR
 ssql = "select DISTINCT(kode_source) from tbl_setting_upload_address "
 Set M_objrs = New ADODB.Recordset
 M_objrs.CursorLocation = adUseClient
 M_objrs.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
 While Not M_objrs.EOF
    cbomapsource.AddItem IIf(IsNull(M_objrs("kode_source")), "", M_objrs("kode_source"))
    M_objrs.MoveNext
 Wend
 
 Set M_objrs = Nothing

End Sub


