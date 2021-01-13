VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_upload_payment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Upload Payment"
   ClientHeight    =   10005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17265
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   17265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Upload Data"
      Height          =   2025
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   17235
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   10320
         Top             =   1800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox TxtPath 
         Height          =   315
         Left            =   9840
         TabIndex        =   16
         Top             =   330
         Visible         =   0   'False
         Width           =   3555
      End
      Begin VB.CommandButton cmdproses 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Verify"
         Height          =   285
         Left            =   3990
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   990
         Width           =   1095
      End
      Begin VB.ComboBox cbosheet 
         Height          =   315
         Left            =   1380
         TabIndex        =   14
         Top             =   990
         Width           =   2565
      End
      Begin VB.CommandButton cmdbrowse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "...."
         Height          =   315
         Left            =   9870
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   555
      End
      Begin VB.CommandButton cmdcreatemap 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Create Map"
         Height          =   285
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtlocation 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   8445
      End
      Begin VB.ComboBox cbomap 
         Height          =   315
         ItemData        =   "frm_upload_payment.frx":0000
         Left            =   1380
         List            =   "frm_upload_payment.frx":0002
         TabIndex        =   10
         Top             =   270
         Width           =   2595
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   360
         Left            =   5220
         TabIndex        =   17
         Top             =   990
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label5 
         Height          =   345
         Left            =   7590
         TabIndex        =   22
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label lblstatus 
         Height          =   345
         Left            =   5220
         TabIndex        =   21
         Top             =   1020
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Sheet"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Location"
         Height          =   255
         Left            =   150
         TabIndex        =   19
         Top             =   690
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Select Mapping"
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Execution"
      Height          =   915
      Left            =   30
      TabIndex        =   0
      Top             =   9030
      Width           =   17145
      Begin VB.CommandButton cmdupload 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Upload"
         Height          =   495
         Left            =   14040
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Exit"
         Height          =   495
         Left            =   15540
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtnew 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1905
      End
      Begin VB.TextBox txtexisting 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   570
         Width           =   1905
      End
      Begin VB.TextBox txtlead 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   210
         Width           =   1905
      End
      Begin VB.Label Label15 
         Caption         =   "New Data :"
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label16 
         Caption         =   "Existing :"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   570
         Width           =   825
      End
      Begin VB.Label Label17 
         Caption         =   "Total Lead :"
         Height          =   285
         Left            =   3210
         TabIndex        =   6
         Top             =   270
         Width           =   1395
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6465
      Left            =   0
      TabIndex        =   23
      Top             =   2520
      Width           =   17235
      _ExtentX        =   30401
      _ExtentY        =   11404
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "View Data upload    "
      TabPicture(0)   =   "frm_upload_payment.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstview"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Cboexecelmap"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "View Mapping     "
      TabPicture(1)   =   "frm_upload_payment.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstmapping"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "History Upload      "
      TabPicture(2)   =   "frm_upload_payment.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtnumrowshst"
      Tab(2).Control(1)=   "lsthst"
      Tab(2).Control(2)=   "Label11"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Error In Excel        "
      TabPicture(3)   =   "frm_upload_payment.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command2"
      Tab(3).Control(1)=   "txtfound"
      Tab(3).Control(2)=   "lsterror"
      Tab(3).Control(3)=   "Label12"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Lead To Database      "
      TabPicture(4)   =   "frm_upload_payment.frx":0074
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).ControlCount=   1
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export to Excel"
         Height          =   345
         Left            =   -74910
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   390
         Width           =   2115
      End
      Begin VB.ComboBox Cboexecelmap 
         Height          =   315
         Left            =   2820
         TabIndex        =   30
         Top             =   990
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox txtnumrowshst 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -59760
         TabIndex        =   29
         Top             =   6000
         Width           =   1605
      End
      Begin VB.TextBox txtfound 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -60030
         TabIndex        =   28
         Top             =   5970
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "Lead Source"
         Height          =   6015
         Left            =   -74940
         TabIndex        =   24
         Top             =   390
         Width           =   16845
         Begin VB.Frame Frame2 
            Caption         =   "View Lead To be Insert to database"
            Height          =   5775
            Left            =   6570
            TabIndex        =   40
            Top             =   180
            Width           =   10155
            Begin VB.TextBox txtlead_masuk 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   5430
               Width           =   1245
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   5145
               Left            =   150
               TabIndex        =   42
               Top             =   270
               Width           =   9885
               _ExtentX        =   17436
               _ExtentY        =   9075
               _Version        =   393216
               HeadLines       =   1
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label13 
               Caption         =   "Rows:"
               Height          =   255
               Left            =   240
               TabIndex        =   43
               Top             =   5460
               Width           =   795
            End
         End
         Begin VB.TextBox txtrowssource 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   2190
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   5550
            Width           =   1245
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   5235
            Left            =   120
            TabIndex        =   26
            Top             =   270
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   9234
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label7 
            Caption         =   "Rows :"
            Height          =   255
            Left            =   1410
            TabIndex        =   27
            Top             =   5580
            Width           =   555
         End
      End
      Begin MSComctlLib.ListView lstview 
         Height          =   6015
         Left            =   30
         TabIndex        =   32
         Top             =   360
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   10610
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
      Begin MSComctlLib.ListView lstmapping 
         Height          =   5385
         Left            =   -74940
         TabIndex        =   33
         Top             =   420
         Width           =   16485
         _ExtentX        =   29078
         _ExtentY        =   9499
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lsthst 
         Height          =   5565
         Left            =   -74940
         TabIndex        =   34
         Top             =   390
         Width           =   17085
         _ExtentX        =   30136
         _ExtentY        =   9816
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
      Begin MSComctlLib.ListView lsterror 
         Height          =   4995
         Left            =   -74940
         TabIndex        =   35
         Top             =   780
         Width           =   16995
         _ExtentX        =   29977
         _ExtentY        =   8811
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
      Begin VB.Label Label8 
         Caption         =   "Jumlah Data"
         Height          =   255
         Left            =   -63150
         TabIndex        =   38
         Top             =   5520
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label11 
         Caption         =   "Num Of Rows :"
         Height          =   255
         Left            =   -60930
         TabIndex        =   37
         Top             =   6060
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Found :"
         Height          =   255
         Left            =   -60780
         TabIndex        =   36
         Top             =   6000
         Width           =   735
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup Upload"
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
      Left            =   570
      TabIndex        =   39
      Top             =   60
      Width           =   3585
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   30
      Picture         =   "frm_upload_payment.frx":0090
      Stretch         =   -1  'True
      Top             =   30
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   465
      Left            =   0
      Picture         =   "frm_upload_payment.frx":0B9A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17220
   End
End
Attribute VB_Name = "Form_upload_payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public M_XLSCONN As New ADODB.Connection
Public Error As String
Private Sub cbocampaign_KeyPress(KeyAscii As Integer)
Dim OBJRECORD As New ADODB.Recordset
Dim clscampaign As New clscampaign
If KeyAscii = 13 Then
   Set clscampaign = New clscampaign
   Set OBJRECORD = clscampaign.FindCampaign(cbocampaign.Text)
   If OBJRECORD.RecordCount > 0 Then
     txtdescription.Text = IIf(IsNull(OBJRECORD!keterangan), "", OBJRECORD!keterangan)
    Else
        txtdescription.Text = ""
   End If
End If
Set clscampaign = Nothing
Set OBJRECORD = Nothing
End Sub

Private Sub cbocampaign_LostFocus()
cbocampaign_KeyPress (13)
End Sub

Private Sub cboket_Click()
Dim OBJRECORD As ADODB.Recordset
    Dim CMDSQL As String
    
    'Mengisi data ke combo campaigncode
    CMDSQL = "select * from  mandiri.tbldivisi where    nm_divisi='"
    CMDSQL = CMDSQL + cboket.Text + "'"
    
    Set OBJRECORD = New ADODB.Recordset
    OBJRECORD.CursorLocation = adUseClient
    
    OBJRECORD.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not OBJRECORD.EOF Then
        cboproduct.Text = IIf(IsNull(OBJRECORD("kddivisi")), "", OBJRECORD("kddivisi"))
    End If
    
    Set OBJRECORD = Nothing
End Sub

Private Sub cbomap_Click()
    findFx cbomap.Text
End Sub

Private Sub cbomap_DropDown()
    loadCboMap
End Sub

Private Sub cbomap_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub cboproduct_Click()
Dim OBJRECORD As ADODB.Recordset
    Dim CMDSQL As String
    
    'Mengisi data ke combo campaigncode
    CMDSQL = "select * from  mandiri.tbldivisi where kddivisi='"
    CMDSQL = CMDSQL + cboproduct.Text + "'"
    
    Set OBJRECORD = New ADODB.Recordset
    OBJRECORD.CursorLocation = adUseClient
    
    OBJRECORD.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not OBJRECORD.EOF Then
        cboket.Text = IIf(IsNull(OBJRECORD("nm_divisi")), "", OBJRECORD("nm_divisi"))
    End If
    
    Set OBJRECORD = Nothing
End Sub

Private Sub cbosheet_Click()
lblstatus.Caption = ""
If txtlocation.Text <> "" Then

If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 8.0;"
    
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & cbosheet.Text & "] "
        M_objrs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
       M_OBJCONN.Execute "delete from tbl_temp_field_payment "
        If M_objrs.EOF And M_objrs.BOF Then Exit Sub
            For i = 0 To M_objrs.fields.Count - 1
                On Error Resume Next
                STRSQL = "insert into tbl_temp_field_payment (nama_field) values ('" + M_objrs.fields(i).Name + "')"
               M_OBJCONN.Execute (STRSQL)
               lblstatus.Caption = "Field Terdefinisi"
            Next i
    Set M_objrs = Nothing
End If

End Sub

Private Sub CmdBrowse_Click()
  Dim dir_listbulantem$
    With CommonDialog1
        .DialogTitle = "Import From File"
        .Filter = "Excel Files|*.xls"
        .ShowOpen
    End With
    txtlocation.Text = ""
    If CommonDialog1.FileName = "" Then Exit Sub
    txtlocation.Text = CommonDialog1.FileName
    If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 8.0;"
    Set M_objrs = M_XLSCONN.OpenSchema(adSchemaTables)
    cbosheet.CLEAR
    If M_objrs.EOF And M_objrs.BOF Then Exit Sub
    While Not M_objrs.EOF
        cbosheet.AddItem IIf(IsNull(M_objrs!table_name), "", M_objrs!table_name)
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
   Set M_XLSCONN = Nothing
End Sub
Private Sub cmdcreatemap_Click()
   Form_setting_upload_payment.Show 1
End Sub
Public Sub loadCboMap()
    cbomap.CLEAR
    ssql = "select DISTINCT(kode_source) from mandiri.tbl_setting_upload_payment  where (kode_source is not null or kode_source<>'')"
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_objrs.EOF
        cbomap.AddItem IIf(IsNull(M_objrs("kode_source")), "", M_objrs("kode_source"))
        M_objrs.MoveNext
    Wend
 Set M_objrs = Nothing
End Sub
Public Sub create_header_mapping()
    lstview.ColumnHeaders.ADD 1, , "Source Field", 10 * TXT
    lstview.ColumnHeaders.ADD 2, , "Destination Filed", 15 * TXT
    lstview.ColumnHeaders.ADD 3, , "Length", 15 * TXT
    lstview.ColumnHeaders.ADD 4, , "Type Data", 15 * TXT
End Sub
Public Sub findFx(ByVal xCodeMap)
Dim list As ListItem
    sStrsql = " select nama_kolom,field_destination,character_maximum_length,data_type from ( "
    sStrsql = sStrsql + " select * from ( "
    sStrsql = sStrsql + " SELECT column_name as nama_kolom,character_maximum_length,data_type From information_schema.Columns WHERE table_name='tbllunas'"
    sStrsql = sStrsql + " and data_type in ('character varying','numeric','bigint','integer','timestamp without time zone','text') ORDER BY ordinal_position) as tblbaru "
    sStrsql = sStrsql + " full join  ( "
    sStrsql = sStrsql + "  select field_source,field_destination from mandiri.tbl_setting_upload_payment where kode_source='" + xCodeMap + "' ) "
    sStrsql = sStrsql + " as tbldua on tblbaru.nama_kolom =tbldua.field_source) as tblbaru "

  
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        lstview.ListItems.CLEAR
        While Not M_objrs.EOF
            Set list = lstview.ListItems.ADD(, , IIf(IsNull(M_objrs!nama_kolom), "", M_objrs!nama_kolom))
                list.SubItems(1) = IIf(IsNull(M_objrs!field_destination), "", M_objrs!field_destination)
                list.SubItems(2) = IIf(IsNull(M_objrs!character_maximum_length), "", M_objrs!character_maximum_length)
                list.SubItems(3) = IIf(IsNull(M_objrs!data_type), "", M_objrs!data_type)
            M_objrs.MoveNext
        Wend
   
        Set M_objrs = Nothing
           

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdProses_Click()

 Dim mobjtemp As New ADODB.Recordset
   'cek map source sudah di isi apa belum
   
   
  
   
   
   If cbomap.Text = "" Then
       MsgBox "Map Source  Belum di select ", vbOKOnly, "Information"
        cbomap.SetFocus
        Exit Sub
   End If
   
   'CEK FIELDNYA TERDEFINISI ATAU TIDAK
   
   If lblstatus.Caption = "" Then
        MsgBox "Field tidak terdefinisi mohon browse ulang excel ", vbOKOnly, "Information"
        cmdproses.Enabled = True
        Exit Sub
        
   End If
   
   If cekMANDATORTY = False Then
            MsgBox "Cek Field Mandatory Harus Camapaign and No kartu ", vbOKOnly, "Information"
    Exit Sub
           
    End If
           
    'VERIFIKASI FIELD YANG TERDEFINISI
    cmdproses.Enabled = False
    If cekmapping_excel = False Then
           MsgBox "Verifikasi Mapping Gagal karena field di mapping tidak terdefinisi di excel ", vbOKOnly, "Information"
           SSTab1.Tab = 1
           cmdproses.Enabled = True
           Label5.Caption = "Tidak Bisa Upload"
           Exit Sub
    End If
    Call cekStrukturField
    Set mobjtemp = New ADODB.Recordset
    mobjtemp.CursorLocation = adUseClient
    
    mobjtemp.Open "select * from mandiri.tbl_upload_temp_payment", M_OBJCONN, adOpenDynamic, adLockOptimistic
 '   Text1.Text = mobjtemp.RecordCount
    Set DataGrid1.DATASOURCE = mobjtemp
    cmdproses.Enabled = True
    
End Sub

Private Sub CmdUpload_Click()
Dim list As ListItem
Dim jRow As Double
Dim ncount As Integer
Dim njmlExitst As Double
Dim njmlNew As Double
Dim OBJRECORD As New ADODB.Recordset
Dim clscampaign As New clscampaign


'If Text1.Text = "" Or Text1.Text = "0" Then
'        MsgBox "Tidak Ada Data Yang diupload", vbOKOnly, "Information"
'        Exit Sub
'End If

'sintak update dulu data yang sama

  

  
If Val(txtlead_masuk.Text) = 0 And Val(txtexisting.Text) = 0 Then
    MsgBox "Tidak ada record yang diupload", vbInformation + vbOKOnly, "Information"
    SSTab1.Tab = 4
    txtlead_masuk.SetFocus
    Exit Sub


End If

If Label5.Caption = "Tidak Bisa Upload" Then
    MsgBox "Field di excel tidak sama dengan mapping yang telah dibuat", vbInformation + vbOKOnly, "Information"
    SSTab1.Tab = 1
    Exit Sub
End If

If lsterror.ListItems.Count <> 0 Then
            MsgBox "Isi data diexcel tidak sama dengan type didatabase", vbInformation + vbOKOnly, "Information"
       SSTab1.Tab = 3
        Exit Sub


End If

strfieldupdate = ""
strfieldheaderupdate = ""
strinsert = ""
  ncount = 1
  For jRow = 1 To lstview.ListItems.Count
        If Len(lstview.ListItems(jRow).SubItems(1)) > 0 Then
                If ncount = 1 Then
                    strfieldupdate = lstview.ListItems(jRow).Text + "=a." + lstview.ListItems(jRow).Text
                    strfieldheaderupdate = "tblpay." + lstview.ListItems(jRow).Text + ""
                    strinsert = lstview.ListItems(jRow).Text + ""
                    ncount = 2
                Else
                    strfieldupdate = strfieldupdate + " ," + lstview.ListItems(jRow).Text + "=a." + lstview.ListItems(jRow).Text
                    strfieldheaderupdate = strfieldheaderupdate + ",tblpay." + lstview.ListItems(jRow).Text
                    strinsert = strinsert + "," + lstview.ListItems(jRow).Text
                End If
                    
        End If
    Next jRow

'update tbl_mst_performance set nbulan=a.nbulan



Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient

    
    STRSQL = "  select " + strinsert + " from  mandiri.tbl_upload_temp_payment where (F_FLAG IS NULL OR F_FLAG=0) "
    M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

'insert  ke tbl_mst_performance
If M_objrs.RecordCount <> 0 Then
njmlNew = M_objrs.RecordCount
  If MsgBox("New Data :" + CStr(njmlNew) + vbCrLf + "", vbQuestion + vbYesNo, "Question") = vbYes Then
    If strinsert <> "" Then
        STRSQL = "insert into mandiri.tbllunas (" + strinsert + ")"
        STRSQL = STRSQL + "  select " + strinsert + " from  mandiri.tbl_upload_temp_payment where (F_FLAG IS NULL OR F_FLAG=0) "
        M_OBJCONN.Execute (STRSQL)
        
        

        MsgBox "Data Telah Di Upload sebanyak : " + CStr(njmlNew) + "", vbOKOnly, "Information"
        Set list = lsthst.ListItems.ADD(, , MDIForm1.TxtUsername.Text)
        list.SubItems(1) = MDIForm1.txtlevel.Text
        list.SubItems(2) = Format(MDIForm1.TDBDate1.Value, "dd/mm/yyyy")
        list.SubItems(3) = Replace(txtlocation.Text, "\", "/")
        list.SubItems(4) = Replace(cbosheet.Text, "$", "")
        list.SubItems(5) = Val(txtrowssource.Text)
        list.SubItems(6) = "Insert New Data"
        list.SubItems(7) = CStr(Val(njmlNew))
        
        
        
     STRSQL = "insert into mandiri.tbl_hst_upload (userid,nama,location_file,Sheet,lead,eksekusi,jml_row) values ("
     STRSQL = STRSQL + "'" + MDIForm1.TxtUsername.Text + "','" + MDIForm1.txtlevel.Text + "','" + Replace(txtlocation.Text, "\", "/") + "',"
     STRSQL = STRSQL + "'" + Replace(Replace(cbosheet.Text, "$", ""), "'", "") + "'," + CStr(Val(txtrowssource.Text)) + ",'Insert New Data'," + CStr(Val(njmlNew)) + ")"
     M_OBJCONN.Execute (STRSQL)
    
    

    End If
End If
End If
Set M_objrs = Nothing


           
            
  
    
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
isi_dataSTATUS ""
End Sub

Private Sub Form_Load()
    create_header_mapping
    create_header_mapping_verify
    create_header_line_error
    create_header_hst_upload
    load_hst_upload
    'loadcbocampaign
  '  isicombo_product
End Sub
Public Sub create_header_mapping_verify()
    lstmapping.ColumnHeaders.ADD 1, , "Source Field", 5 * TXT
    lstmapping.ColumnHeaders.ADD 2, , "Destination Field", 15 * TXT
    lstmapping.ColumnHeaders.ADD 3, , "Wrong Destination Field", 15 * TXT
End Sub

Public Sub findFxcek(ByVal xCodeMap)
Dim list As ListItem

    sStrsql = " select nama_kolom,field_destination from ( "
    sStrsql = sStrsql + " select * from ( "
    sStrsql = sStrsql + " SELECT column_name as nama_kolom From information_schema.Columns WHERE table_name='mgm'"
    sStrsql = sStrsql + " and substring(column_name,1,2) in ('n_','v_','d_') ORDER BY ordinal_position) as tblbaru "
    sStrsql = sStrsql + " full join  ( "
    sStrsql = sStrsql + "  select field_source,field_destination from tbl_setting_upload where substring(field_source,1,2) in ('n_','v_','d_') and kode_source='" + xCodeMap + "' ) "
    sStrsql = sStrsql + " as tbldua on tblbaru.nama_kolom =tbldua.field_source) as tblbaru "

  
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        lstview.ListItems.CLEAR
        While Not M_objrs.EOF
            Set list = lstview.ListItems.ADD(, , IIf(IsNull(M_objrs!nama_kolom), "", M_objrs!nama_kolom))
                list.SubItems(1) = IIf(IsNull(M_objrs!field_destination), "", M_objrs!field_destination)
            M_objrs.MoveNext
        Wend
        Set M_objrs = Nothing

End Sub
Public Function cekmapping_excel() As Boolean

    STRSQL = " select * from ( "
    STRSQL = STRSQL + " select nama_kolom,field_destination from "
    STRSQL = STRSQL + " (select * from ( "
    STRSQL = STRSQL + " SELECT column_name as nama_kolom From information_schema.Columns WHERE table_name='tbllunas'"
    STRSQL = STRSQL + " and data_type in ('character varying','numeric','bigint','integer','timestamp without time zone','text')  ORDER BY ordinal_position) as tblbaru  full join"
    STRSQL = STRSQL + " (   select field_source,field_destination from tbl_setting_upload_payment  where kode_source='" + cbomap.Text + "')"
    STRSQL = STRSQL + " as tbldua on tblbaru.nama_kolom =tbldua.field_source) as tblbaru where (field_destination is not null or field_destination<>'') ) as tblsatu"
    STRSQL = STRSQL + " Left Join ( select * from tbl_temp_field_payment  ) as tbldua   on tblsatu.field_destination=tbldua.nama_field"
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
       If M_objrs.RecordCount = 0 Then
                    stidak = "1"
       End If
       lstmapping.ListItems.CLEAR
        While Not M_objrs.EOF
            Set list = lstmapping.ListItems.ADD(, , IIf(IsNull(M_objrs!nama_kolom), "", M_objrs!nama_kolom))
                list.SubItems(1) = IIf(IsNull(M_objrs!field_destination), "", M_objrs!field_destination)
                If IIf(IsNull(M_objrs!nama_field), "", M_objrs!nama_field) = "" Then
                 list.SubItems(2) = "Tidak Ada dalam Mapping"
                    stidak = "1"
                    Else
                    list.SubItems(2) = "ADA"
                End If
            M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
    If stidak = "1" Then
        cekmapping_excel = False
    Else
           cekmapping_excel = True
    End If
    
End Function
Public Sub create_header_line_error()
    lsterror.ColumnHeaders.ADD 1, , "[Line/Rows]", 10 * TXT
    lsterror.ColumnHeaders.ADD 2, , "Description Error", 15 * TXT
End Sub
Public Sub create_header_hst_upload()
    lsthst.ColumnHeaders.ADD 1, , "Officer ID", 5 * TXT
    lsthst.ColumnHeaders.ADD 2, , "Officer Name", 15 * TXT
    lsthst.ColumnHeaders.ADD 3, , "Upload Date", 15 * TXT
    lsthst.ColumnHeaders.ADD 4, , "location", 15 * TXT
    lsthst.ColumnHeaders.ADD 5, , "Sheet", 15 * TXT
    lsthst.ColumnHeaders.ADD 6, , "Total Lead", 15 * TXT
    lsthst.ColumnHeaders.ADD 7, , "Execution ", 15 * TXT
    lsthst.ColumnHeaders.ADD 8, , "Number Of row", 15 * TXT

End Sub
Public Sub load_hst_upload()
Dim M_objrs   As New ADODB.Recordset
Dim list As ListItem
Dim no As Double
sStrsql = "select * from tbl_hst_upload "
Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    no = 0
    lsthst.ListItems.CLEAR
    txtnumrowshst.Text = M_objrs.RecordCount
    While Not M_objrs.EOF
        no = no + 1
        Set list = lsthst.ListItems.ADD(, , IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID))
            list.SubItems(1) = IIf(IsNull(M_objrs!nama), "", M_objrs!nama)
            list.SubItems(2) = Format(IIf(IsNull(M_objrs!tgl_upload), "", M_objrs!tgl_upload), "dd/mm/yyyy")
            list.SubItems(3) = IIf(IsNull(M_objrs!location_file), "", M_objrs!location_file)
            list.SubItems(4) = IIf(IsNull(M_objrs!Sheet), "", M_objrs!Sheet)
            list.SubItems(5) = IIf(IsNull(M_objrs!lead), "0", M_objrs!lead)
            list.SubItems(6) = IIf(IsNull(M_objrs!eksekusi), "0", M_objrs!eksekusi)
            list.SubItems(7) = IIf(IsNull(M_objrs!jml_row), "0", M_objrs!jml_row)

            
        M_objrs.MoveNext
    Wend
   
Set M_objrs = Nothing
End Sub
Public Sub cekStrukturField()
Dim list As ListItem
Dim i As Integer
Dim ncount As Integer
Dim sType As String
Dim jml As Double
Dim nlimit As Double
Dim sMapdestination As String
Dim smapsource As String
Dim CEKIN As Boolean
Dim m_objdonot As New ADODB.Recordset
Dim m_objmasuk As New ADODB.Recordset
Dim m_objExisting As New ADODB.Recordset
Dim M_objrs As New ADODB.Recordset
Dim M_objdouble As New ADODB.Recordset
On Error Resume Next
 M_OBJCONN.Execute " Drop TABLE Tbl_Upload_Temp_payment"
 ssql = "CREATE TABLE Tbl_Upload_Temp_payment "
 ssql = ssql & "(id serial)"
 M_OBJCONN.Execute (ssql)
 STRSQL = " select nama_kolom,field_destination,data_type,character_maximum_length from (  select * from (  SELECT column_name as nama_kolom From information_schema.Columns"
 STRSQL = STRSQL + " WHERE table_name='tbllunas' and data_type in ('character varying','numeric','bigint','integer','timestamp without time zone','text')    ORDER BY ordinal_position) as tblbaru  full join  (   select field_source,field_destination from tbl_setting_upload_payment where  kode_source='" + cbomap.Text + "' ) "
 STRSQL = STRSQL + "  as tbldua on tblbaru.nama_kolom =tbldua.field_source) as tblbaru"
 STRSQL = STRSQL + " Left Join"
 STRSQL = STRSQL + " (SELECT column_name,data_type ,character_maximum_length From information_schema.Columns WHERE table_name='tbllunas' ORDER BY ordinal_position) as tbltiga"
 STRSQL = STRSQL + " on tblbaru.nama_kolom=tbltiga.column_name"
 Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   ProgressBar1.Max = M_objrs.RecordCount + 1
   
    While Not M_objrs.EOF
   
    
    DoEvents
    ProgressBar1.Value = M_objrs.Bookmark
                 
            nama_kol = IIf(IsNull(M_objrs!nama_kolom), "", M_objrs!nama_kolom)
           
           
            
            data_type = IIf(IsNull(M_objrs!data_type), "", M_objrs!data_type)
            data_length = IIf(IsNull(M_objrs!character_maximum_length), "", M_objrs!character_maximum_length)
            If Trim(data_type) = "character varying" Then
                If data_length = "" Then
                    Strtype = nama_kol + " " + data_type
                    sStrsql = " alter table Tbl_Upload_Temp_payment  add column  " + Strtype
                Else
                    Strtype = nama_kol + " " + data_type + " (" + CStr(data_length) + ")"
                    sStrsql = " alter table Tbl_Upload_Temp_payment  add column  " + Strtype
                End If
                M_OBJCONN.Execute (sStrsql)
            ElseIf Trim(data_type) = "text" Then
                If data_length = "" Then
                    Strtype = nama_kol + " " + data_type
                    sStrsql = " alter table Tbl_Upload_Temp_payment  add column  " + Strtype
                Else
                    Strtype = nama_kol + " " + data_type + " (" + CStr(data_length) + ")"
                    sStrsql = " alter table Tbl_Upload_Temp_payment  add column  " + Strtype
                End If
                
                M_OBJCONN.Execute (sStrsql)
            ElseIf Trim(data_type) = "timestamp without time zone" Then
                Strtype = nama_kol + " " + data_type
                sStrsql = " alter table Tbl_Upload_Temp_payment  add column  " + Strtype
                M_OBJCONN.Execute (sStrsql)
            ElseIf Trim(data_type) = "numeric" Then
                If data_length = "" Then
                    Strtype = nama_kol + " " + data_type
                     sStrsql = " alter table Tbl_Upload_Temp_payment  add column  " + Strtype
                    Else
                       Strtype = nama_kol + " " + data_type + " (" + CStr(data_length) + ")"
                        sStrsql = " alter table Tbl_Upload_Temp_payment  add column  " + Strtype
                End If
                M_OBJCONN.Execute (sStrsql)
            
             ElseIf Trim(data_type = "bigint") Then
                If data_length = "" Then
                    Strtype = nama_kol + " " + data_type
                     sStrsql = " alter table Tbl_Upload_Temp_payment  add column  " + Strtype
                    Else
                       Strtype = nama_kol + " " + data_type + " (" + CStr(data_length) + ")"
                        sStrsql = " alter table Tbl_Upload_Temp_payment  add column  " + Strtype
                End If
                M_OBJCONN.Execute (sStrsql)
            ElseIf Trim(data_type) = "integer" Then
                If data_length = "" Then
                    Strtype = nama_kol + " " + data_type
                     sStrsql = " alter table Tbl_Upload_Temp_payment  add column  " + Strtype
                    Else
                       Strtype = nama_kol + " " + data_type + " (" + CStr(data_length) + ")"
                        sStrsql = " alter table Tbl_Upload_Temp_payment  add column  " + Strtype
                End If
                M_OBJCONN.Execute (sStrsql)
            End If
        M_objrs.MoveNext
    Wend
    
    
    
        sStrsql = " alter table Tbl_Upload_Temp_payment  add column f_flag numeric"
        M_OBJCONN.Execute (sStrsql)

    
    
    
    
    
    Set M_objrs = Nothing
    
    
    ssql = "SELECT * FROM [" & cbosheet.Text & "]   "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
    lsterror.ListItems.CLEAR
    ProgressBar1.Max = M_objrs.RecordCount + 1
    jml = 0
    txtrowssource.Text = M_objrs.RecordCount
    txtlead.Text = txtrowssource.Text
      Set DataGrid2.DATASOURCE = M_objrs
    While Not M_objrs.EOF
    'Set DataGrid2.DataSource = M_OBJRS
'    Debug.Print M_OBJRS!Target
'     If M_OBJRS.Bookmark = 300 Then
'     MsgBox "sds"
'    End If
    
'    If Val(IIf(IsNull(M_OBJRS!Target), 0, M_OBJRS!Target)) > 0 Then
'    MsgBox "sdsd"
    'End If
    
            DoEvents
           Error = ""
           ProgressBar1.Value = M_objrs.Bookmark
           CEKIN = False
           nlimit = 0
           smapsource = ""
           sMapdestination = ""
         
           
           For jRow = 1 To lstview.ListItems.Count
    
                If Len(lstview.ListItems(jRow).SubItems(1)) > 0 Then
                    sType = ""
                    sType = lstview.ListItems(jRow).SubItems(3)
                   
                    
                    nlimit = Val(lstview.ListItems(jRow).SubItems(2))
                    smapsource = lstview.ListItems(jRow).Text
                    sMapdestination = lstview.ListItems(jRow).SubItems(1)
               '     sMapvalue = iif(isnullm_objrs(sMapdestination).Value
                    CEKIN = ceklength(sType, nlimit, smapsource, sMapdestination, M_objrs, CEKIN)
                End If
           Next jRow
           
           If CEKIN = True Then
           SSTab1.Tab = 3
                     jml = jml + 1
                    If Len(Error) > 1 Then
                  
                        Set list = lsterror.ListItems.ADD(, , CStr(M_objrs.Bookmark))
                            list.SubItems(1) = Error
                            End If
                            
            End If
                
                
           If CEKIN = False Then
                strfield = ""
                
                ' String Ambil field dimapping
                ncount = 1
                For i = 1 To lstview.ListItems.Count
                  
                    
                    
                    
                    If Len(lstview.ListItems(i).SubItems(1)) > 0 Then
                        If ncount = 1 Then
                        
                            strfield = lstview.ListItems(i).Text
                            
                            If lstview.ListItems(i).SubItems(3) = "character varying" Or lstview.ListItems(i).SubItems(3) = "text" Then
                                If IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))) = True Then
                                    strvalues = "null"
                                Else
                                    strvalues = "'" + Replace(IIf(IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))), "", CStr(M_objrs.fields(lstview.ListItems(i).SubItems(1)))), "'", "`") & "'"
                                End If
                            
                            ElseIf lstview.ListItems(i).SubItems(3) = "numeric" Or lstview.ListItems(i).SubItems(3) = "bigint" Or lstview.ListItems(i).SubItems(3) = "integer" Then
                                   
                                If IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))) = True Then
                                    strvalues = "null"
                                Else
                                    strvalues = "" + CStr(M_objrs.fields(lstview.ListItems(i).SubItems(1)))
                                End If
                            
                            ElseIf lstview.ListItems(i).SubItems(3) = "timestamp without time zone" Or lstview.ListItems(i).SubItems(3) = "timestamp with time zone" Then
                                  If IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))) = True Then
                                        strvalues = "Null"
                                  Else
                                        strvalues = "'" + IIf(IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))), Null, Format(M_objrs.fields(lstview.ListItems(i).SubItems(1)), "yyyy/mm/dd")) & "'"
                                End If
    
                            End If
                            
                            ncount = 2
                        Else
                            strfield = strfield + "," + lstview.ListItems(i).Text
                             
                             If lstview.ListItems(i).SubItems(3) = "character varying" Or lstview.ListItems(i).SubItems(3) = "text" Then
                                If IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))) = True Then
                                    strvalues = strvalues + ",null"
                                Else
                                    strvalues = strvalues + ",'" + Replace(IIf(IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))), "", CStr(M_objrs.fields(lstview.ListItems(i).SubItems(1)))), "'", "`") & "'"
                                End If
                            
                            ElseIf lstview.ListItems(i).SubItems(3) = "numeric" Or lstview.ListItems(i).SubItems(3) = "bigint" Or lstview.ListItems(i).SubItems(3) = "integer" Then
                                   
                                If IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))) = True Then
                                    strvalues = strvalues + ",null"
                                Else
                                    strvalues = strvalues + "," + CStr(M_objrs.fields(lstview.ListItems(i).SubItems(1)))
                                End If
                            
                            ElseIf lstview.ListItems(i).SubItems(3) = "timestamp without time zone" Or lstview.ListItems(i).SubItems(3) = "timestamp with time zone" Then
                                  If IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))) = True Then
                                        strvalues = strvalues + ",Null"
                                  Else
                                        strvalues = strvalues + ",'" + IIf(IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))), Null, Format(M_objrs.fields(lstview.ListItems(i).SubItems(1)), "yyyy/mm/dd")) & "'"
                                 End If
    
                            End If
                            
                        
                        End If
                    End If
                Next i
                
                
                If strfield <> "" Then
                        ssqlhead = "INSERT INTO Tbl_Upload_Temp_payment (" + strfield + ") values ( " + strvalues + ")"
                        Debug.Print M_objrs.Bookmark
                        Debug.Print ssqlhead
                        M_OBJCONN.Execute (ssqlhead)
                End If
                
           End If
           
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    
      'MOBILENO (Add 0)
            
         
            
            
         
'
         
            
    
            STRSQL = " update   tbl_upload_temp_payment set  f_flag=1 where id in (select Tbl_Upload_Temp_payment.id  from tbllunas inner join "
            STRSQL = STRSQL + " Tbl_Upload_Temp_payment  on tbllunas.custid= Tbl_Upload_Temp_payment.custid  and date(tbllunas.paydate)=date(Tbl_Upload_Temp_payment.paydate) AND tbllunas.payment=Tbl_Upload_Temp_payment.payment)"
            M_OBJCONN.Execute (STRSQL)
'
         
    
'            STRSQL = " update Tbl_Upload_Temp set f_flag_donot =1 where v_no_kartu in (select no_kartu from tbldonotcall ) "
'            M_OBJCONN.Execute (STRSQL)
            
    
    'cek existing data
      'STRSQL = " select  tbllunas.* from tbllunas,tbl_upload_temp_payment"
      'STRSQL = STRSQL + " where tbllunas.custid=tbl_upload_temp_payment.custid and date(tbllunas.paydate)=date(tbl_upload_temp_payment.paydate)"
     
     STRSQL = " select * from tbl_upload_temp_payment  where f_flag=1"


      Set m_objExisting = New ADODB.Recordset
          m_objExisting.CursorLocation = adUseClient
          m_objExisting.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
          txtexisting.Text = m_objExisting.RecordCount
    'cek data new
    
    STRSQL = " select custid,paydate,agent,sum(payment) as payment from tbl_upload_temp_payment where f_flag is null or f_flag=0  group by custid,paydate,agent "
    
    Set m_objmasuk = New ADODB.Recordset
        m_objmasuk.CursorLocation = adUseClient
        m_objmasuk.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
        txtlead_masuk.Text = m_objmasuk.RecordCount
        Set DataGrid1.DATASOURCE = m_objmasuk
        
        txtnew.Text = txtlead_masuk
    'cek jumlah data yang ke donot call
  
End Sub
Public Function ceklength(sTypeData As String, nlimit, sMapSourceFieldname As String, sMapdestination As String, rsTemp1 As ADODB.Recordset, prm As Boolean) As Boolean
ceklength = prm

If sTypeData = "character varying" Then
    If nlimit > 0 Then
        If Len(rsTemp1(sMapdestination).Value) > nlimit Then
            ceklength = True
            If Len(Error) > 0 Then
                Error = Error + vbCrLf + "value : " + CStr(rsTemp1(sMapdestination).Value) + " kolom  " + sMapdestination + " (Excel) melebihi " + CStr(nlimit) + " Character  " + sMapSourceFieldname + " (database)"
            Else
               Error = "value : " + CStr(rsTemp1(sMapdestination).Value) + " kolom  " + sMapdestination + " (Excel) melebihi " + CStr(nlimit) + " Character  " + sMapSourceFieldname + " (database)"
            End If
        
        End If
    End If
End If


If sTypeData = "timestamp without time zone" Or sTypeData = "timestamp with time zone" Then
    If Len(rsTemp1(sMapdestination).Value) > 0 Then
        If IsDate(rsTemp1(sMapdestination).Value) = False Then
        ceklength = True
            If Len(Error) > 0 Then
                     Error = Error + vbCrLf + " value : " + CStr(rsTemp1(sMapdestination).Value) + " kolom  " + sMapdestination + " (Excel) tidak sama dengan format tanggal " + sMapSourceFieldname + " (Database)"
            Else
                    Error = " value : " + CStr(rsTemp1(sMapdestination).Value) + " kolom  " + sMapdestination + " (Excel) tidak sama dengan format tanggal " + sMapSourceFieldname + " (Database)"
            End If
        End If
    End If

End If


If sTypeData = "numeric" Or sTypeData = "bigint" Or sTypeData = "integer" Then
'Debug.Print CStr(rsTemp1.Bookmark)
    If Len(rsTemp1(sMapdestination).Value) > 0 Then
        If IsNumeric(rsTemp1(sMapdestination).Value) = False Then
        ceklength = True
            If Len(Error) > 0 Then
                     Error = Error + " value : " + CStr(rsTemp1.fields(sMapdestination).Value) + " kolom  " + sMapdestination + " (Excel) tidak sama dengan format angka " + sMapSourceFieldname + " (Database)"
            Else
                     Error = " value : " + CStr(rsTemp1(sMapdestination).Value) + " kolom  " + sMapdestination + " (Excel) tidak sama dengan format angka " + sMapSourceFieldname + " (Database)"
            End If
    End If
    End If
End If
End Function
Private Sub isi_dataSTATUS(STRSQL As String)
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim cmdsql_update As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i As Integer
    Dim m_msgbox As String
    
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
    
    
    
    
    'Set excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
    
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
objSheet.Cells(1, 1).ColumnWidth = 30
objSheet.Cells(1, 1).Value = "[Line/Rows]"
objSheet.Cells(1, 2).ColumnWidth = 30
objSheet.Cells(1, 2).Value = "Description"

n = 1
  For i = 1 To lsterror.ListItems.Count
    n = n + 1
    objSheet.Cells(n, 1).ColumnWidth = 30
    objSheet.Cells(n, 1).Value = lsterror.ListItems(i).Text
    objSheet.Cells(n, 2).ColumnWidth = 30
    objSheet.Cells(n, 2).Value = lsterror.ListItems(i).SubItems(1)
  Next i
  
    objBook.SaveAs TxtPath.Text, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
 
Salah:
    Exit Sub
End Sub
Public Function cekMANDATORTY() As Boolean
    cekMANDATORTY = False
    i = 1
 
    For i = 1 To lstview.ListItems.Count

    If lstview.ListItems(i).Text = "custid" Then
        If lstview.ListItems(i).SubItems(1) = "" Then
            Exit Function
        End If
    End If
    
    If lstview.ListItems(i).Text = "paydate" Then
        If lstview.ListItems(i).SubItems(1) = "" Then
            Exit Function
        End If
    End If
    
    cekMANDATORTY = True
    Next i
End Function
'Private Sub isicombo_product()
'    Dim OBJRECORD As New ADODB.Recordset
'    Dim cmdsql As String
'    cmdsql = "select * from  tbldivisi "
'    Set OBJRECORD = New ADODB.Recordset
'    OBJRECORD.CursorLocation = adUseClient
'
'    OBJRECORD.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    cboproduct.CLEAR
'    cboket.CLEAR
'    While Not OBJRECORD.EOF
'        cboproduct.AddItem IIf(IsNull(OBJRECORD("kddivisi")), "", OBJRECORD("kddivisi"))
'        cboket.AddItem IIf(IsNull(OBJRECORD("nm_divisi")), "", OBJRECORD("nm_divisi"))
'        OBJRECORD.MoveNext
'    Wend
'    Set OBJRECORD = Nothing
'
'End Sub


