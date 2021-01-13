VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_swap 
   Caption         =   "[Swap Data  Account, Curball - Curpri] dan [Upload Class Mapping]"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14640
   LinkTopic       =   "Form2"
   ScaleHeight     =   7785
   ScaleWidth      =   14640
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7785
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   13732
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Upload"
      TabPicture(0)   =   "form_swap.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdproses(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdproses(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DataGrid1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtcount"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CommonDialog1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "History Swap"
      TabPicture(1)   =   "form_swap.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtrowhst_del"
      Tab(1).Control(1)=   "cmdproses(0)"
      Tab(1).Control(2)=   "ListView1"
      Tab(1).Control(3)=   "Label9"
      Tab(1).ControlCount=   4
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3120
         Top             =   7200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         Caption         =   "Catetan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   180
         TabIndex        =   22
         Top             =   5580
         Width           =   14235
         Begin VB.Label LblPerhatian 
            Caption         =   "Jangan lupa pilih jenis SWAP data!"
            Height          =   615
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   13815
         End
      End
      Begin VB.TextBox txtrowhst_del 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   -62040
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   7290
         Width           =   1245
      End
      Begin VB.TextBox txtcount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   6990
         Width           =   1425
      End
      Begin VB.CommandButton cmdproses 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Refresh"
         Height          =   345
         Index           =   0
         Left            =   -74820
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   420
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6255
         Left            =   -74850
         TabIndex        =   14
         Top             =   840
         Width           =   14355
         _ExtentX        =   25321
         _ExtentY        =   11033
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3495
         Left            =   180
         TabIndex        =   13
         Top             =   2010
         Width           =   14265
         _ExtentX        =   25162
         _ExtentY        =   6165
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
      Begin VB.CommandButton cmdproses 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Exit"
         Height          =   495
         Index           =   2
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton cmdproses 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Swap"
         Height          =   495
         Index           =   1
         Left            =   12210
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6960
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Upload Data"
         Height          =   1515
         Left            =   150
         TabIndex        =   1
         Top             =   360
         Width           =   14265
         Begin VB.ComboBox CmbJenis 
            Height          =   315
            ItemData        =   "form_swap.frx":0038
            Left            =   1380
            List            =   "form_swap.frx":0045
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1140
            Width           =   2595
         End
         Begin VB.TextBox TxtPath 
            Height          =   315
            Left            =   10560
            TabIndex        =   5
            Top             =   180
            Visible         =   0   'False
            Width           =   3555
         End
         Begin VB.ComboBox cbosheet 
            Height          =   315
            Left            =   1380
            TabIndex        =   4
            Top             =   720
            Width           =   2565
         End
         Begin VB.CommandButton cmdbrowse 
            BackColor       =   &H00C0FFC0&
            Caption         =   "...."
            Height          =   315
            Left            =   9870
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   300
            Width           =   555
         End
         Begin VB.TextBox txtlocation 
            Height          =   315
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   330
            Width           =   8445
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   360
            Left            =   4140
            TabIndex        =   6
            Top             =   720
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label Label6 
            Caption         =   "Jenis:"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   1200
            Width           =   1035
         End
         Begin VB.Label Label5 
            Height          =   345
            Left            =   7590
            TabIndex        =   10
            Top             =   1080
            Width           =   1365
         End
         Begin VB.Label lblstatus 
            Height          =   345
            Left            =   5220
            TabIndex        =   9
            Top             =   1020
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Sheet"
            Height          =   255
            Left            =   180
            TabIndex        =   8
            Top             =   750
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "Location"
            Height          =   255
            Left            =   150
            TabIndex        =   7
            Top             =   330
            Width           =   795
         End
      End
      Begin VB.Label Label9 
         Caption         =   "Rows:"
         Height          =   255
         Left            =   -62550
         TabIndex        =   19
         Top             =   7290
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Count"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   7050
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form_swap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_XLSCONN As New ADODB.Connection

Private Sub cbosheet_Click()

Dim OBJRECORD As New ADODB.Recordset
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    
    ssql = "SELECT * FROM [" & cbosheet.Text & "] "
    RSTEMP.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
    Set RSTEMP = Nothing
    
     Set OBJRECORD = New ADODB.Recordset
        OBJRECORD.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & cbosheet.Text & "] "
        DoEvents
        OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Set DataGrid1.DATASOURCE = OBJRECORD
        txtcount.Text = OBJRECORD.RecordCount

End Sub

Private Sub CmbJenis_Click()
    If CmbJenis.Text = "SWAP ACCOUNT" Then
        LblPerhatian.Caption = "Untuk Swap Account, pastikan di excel anda ada custid dan agent!"
    ElseIf CmbJenis.Text = "SWAP CURBALL - CURPRI" Then
        LblPerhatian.Caption = "Untuk Swap Curball - Curpri, pastikan di excel anda ada custid, curball dan curpri!"
    ElseIf CmbJenis.Text = "UPLOAD CLASS MAPPING" Then
        LblPerhatian.Caption = "Untuk upload CLASS MAPPING, pastikan di excel anda ada custid, cardno dan class!"
    End If
End Sub

Private Sub CmdBrowse_Click()
 With CommonDialog1
            .DialogTitle = "Import From File"
            .Filter = "Excel Files|*.xls"
            .ShowOpen
        End With
        
        txtlocation.Text = CommonDialog1.FileName
        If CommonDialog1.FileName = "" Then Exit Sub
        If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
                M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & txtlocation.Text & ";Extended Properties=Excel 8.0;"
        Set RSTEMP = M_XLSCONN.OpenSchema(adSchemaTables)
        cbosheet.CLEAR
        If RSTEMP.EOF And RSTEMP.BOF Then Exit Sub
        While Not RSTEMP.EOF
            cbosheet.AddItem IIf(IsNull(RSTEMP!table_name), "", RSTEMP!table_name)
            RSTEMP.MoveNext
        Wend
        Set RSTEMP = Nothing
        
End Sub

Private Sub cmdProses_Click(Index As Integer)
Dim M_objrs   As New ADODB.Recordset
Dim list As ListItem
Select Case Index
Case 0
Dim no As Double
sStrsql = "select * from mandiri.tbl_hst_swap"
Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    no = 0
    ListView1.ListItems.CLEAR
    txtrowhst_del.Text = M_objrs.RecordCount
    While Not M_objrs.EOF
        DoEvents
        no = no + 1
        Set list = ListView1.ListItems.ADD(, , Format(IIf(IsNull(M_objrs!tglswap), "", M_objrs!tglswap), "dd/mm/yyyy hh:nn:ss"))
            list.SubItems(1) = IIf(IsNull(M_objrs!swap_path), "", M_objrs!swap_path)
            list.SubItems(2) = IIf(IsNull(M_objrs!jml_swap), "", M_objrs!jml_swap)
            list.SubItems(3) = IIf(IsNull(M_objrs!user_excecute), "", M_objrs!user_excecute)
        M_objrs.MoveNext
    Wend
Set M_objrs = Nothing



Case 1
   If CmbJenis.Text = "" Then
        MsgBox "Anda belum menentukan jenis Swap atau Upload data!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
   End If
   
    If CmbJenis.Text = "SWAP ACCOUNT" Then
        'Untuk Swap Account
        InsertToUpload
    ElseIf CmbJenis.Text = "SWAP CURBALL - CURPRI" Then
        'Untuk Swap Curball Curpri
        SwapCurballCurpri
    ElseIf CmbJenis.Text = "UPLOAD CLASS MAPPING" Then
        UploadClassMap
    End If
   
   
Case 2
    Unload Me
End Select

End Sub
Private Function InsertToUpload() As Boolean
    Dim RSTEMP As New ADODB.Recordset
    On Error GoTo KE
    Set RSTEMP = New ADODB.Recordset
        RSTEMP.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & cbosheet.Text & "] "
        DoEvents
        RSTEMP.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Set DataGrid1.DATASOURCE = RSTEMP
        txtcount.Text = RSTEMP.RecordCount
        ProgressBar1.Visible = True
        ProgressBar1.Max = RSTEMP.RecordCount + 1
         While Not RSTEMP.EOF
           DoEvents
           ProgressBar1.Value = RSTEMP.Bookmark
                strnokartu = IIf(IsNull(RSTEMP!CustId), "", RSTEMP!CustId)
                stragent = IIf(IsNull(RSTEMP!AGENT), "", RSTEMP!AGENT)
                ' Ditambahin spv_allow = 1 untuk data expired 20 Nov 2013
                ' Tambah app_claim di null 30 Mei 2014
                ' SET AGENT ASLI KETIKA SWAP ACCOUNT
'                Strsql = " UPDATE MGM SET agent= '" + stragent + "',spv_allow=now(),app_claim=null WHERE  custid  ='" + strnokartu + "'"
                STRSQL = " UPDATE mandiri.MGM SET agent= '" + stragent + "',agent_asli= '" + stragent + "',spv_allow=now(),app_claim=null WHERE  custid  ='" + strnokartu + "'"
                M_OBJCONN.Execute (STRSQL)
        
                RSTEMP.MoveNext
            Wend
           
            Set RSTEMP = Nothing
            MsgBox "Data telah di SWAP dengan sukses!", vbInformation + vbOKOnly, "Information"
            STRSQL = "insert into mandiri.tbl_hst_swap (jml_swap,user_excecute, swap_path) "
            STRSQL = STRSQL + " values ('" + txtcount.Text + "','" + MDIForm1.TxtUsername.Text + "','" + Replace(txtlocation.Text, "\", "/") + "')"
            M_OBJCONN.Execute (STRSQL)
            cbosheet.Text = ""
            txtcount.Text = ""
            txtlocation.Text = ""
            If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close

Exit Function
KE:
    MsgBox "keterangan eror di " + err.Description + "di sumber" + err.Source
    
End Function

Private Function SwapCurballCurpri() As Boolean
    Dim RSTEMP As New ADODB.Recordset
    On Error GoTo KE
    Set RSTEMP = New ADODB.Recordset
        RSTEMP.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & cbosheet.Text & "] "
        DoEvents
        RSTEMP.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Set DataGrid1.DATASOURCE = RSTEMP
        txtcount.Text = RSTEMP.RecordCount
        ProgressBar1.Visible = True
        ProgressBar1.Max = RSTEMP.RecordCount + 1
        Dim strnokartu As String
        Dim strcurbal As String
        Dim strcurpri As String
         While Not RSTEMP.EOF
           DoEvents
           ProgressBar1.Value = RSTEMP.Bookmark
                strnokartu = IIf(IsNull(RSTEMP!CustId), "", RSTEMP!CustId)
                strcurbal = IIf(IsNull(RSTEMP!curball), "null", "'" & RSTEMP!curball & "'")
                strcurpri = IIf(IsNull(RSTEMP!curpri), "null", "'" & RSTEMP!curpri & "'")
                
                
                    STRSQL = "update mandiri.mgm set curbal=" + CStr(strcurbal)
                    STRSQL = STRSQL + ", curpri="
                    STRSQL = STRSQL + CStr(strcurpri)
                    STRSQL = STRSQL + " where custid='"
                    STRSQL = STRSQL + Trim(strnokartu) + "'"
                M_OBJCONN.Execute (STRSQL)
                RSTEMP.MoveNext
            Wend
           
            Set RSTEMP = Nothing
            MsgBox "Swap Curball-Curpri sukses!", vbInformation + vbOKOnly, "Information"
            STRSQL = "insert into mandiri.tbl_hst_swap_curbal_curpri (jml_swap,user_excecute, swap_path) "
            STRSQL = STRSQL + " values ('" + txtcount.Text + "','" + MDIForm1.TxtUsername.Text + "','" + Replace(txtlocation.Text, "\", "/") + "')"
            M_OBJCONN.Execute (STRSQL)
            cbosheet.Text = ""
            txtcount.Text = ""
            txtlocation.Text = ""
            If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close

Exit Function
KE:
    MsgBox "Keterangan eror di " + err.Description + "di sumber" + err.Source
    
End Function


Private Function UploadClassMap() As Boolean
    Dim RSTEMP As New ADODB.Recordset
    Dim strcurbal_map As String
    
    On Error GoTo KE
    Set RSTEMP = New ADODB.Recordset
        RSTEMP.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & cbosheet.Text & "] "
        DoEvents
        RSTEMP.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Set DataGrid1.DATASOURCE = RSTEMP
        txtcount.Text = RSTEMP.RecordCount
        ProgressBar1.Visible = True
        ProgressBar1.Max = RSTEMP.RecordCount + 1
         While Not RSTEMP.EOF
           DoEvents
           ProgressBar1.Value = RSTEMP.Bookmark
                strnokartu = IIf(IsNull(RSTEMP!CustId), "", RSTEMP!CustId)
                strcardno = IIf(IsNull(RSTEMP!cardno), "", RSTEMP!cardno)
                strclass = IIf(IsNull(RSTEMP!Class), "", RSTEMP!Class)
                strcurbal_map = IIf(IsNull(RSTEMP!curball), "", RSTEMP!curball)
                
                STRSQL = "insert into mandiri.tbldetailmapping (custid,cardno,class,balance) values ('"
                STRSQL = STRSQL + CStr(strnokartu) + "','"
                STRSQL = STRSQL + CStr(strcardno) + "','"
                STRSQL = STRSQL + CStr(strclass) + "','"
                STRSQL = STRSQL + CStr(strcurbal_map) + "')"
                M_OBJCONN.Execute (STRSQL)
                RSTEMP.MoveNext
            Wend
           
            Set RSTEMP = Nothing
            MsgBox "Upload  class mapping sukses!", vbInformation + vbOKOnly, "Information"
            STRSQL = "insert into mandiri.tbl_hst_upload_class_map (jml_swap,user_excecute, swap_path) "
            STRSQL = STRSQL + " values ('" + txtcount.Text + "','" + MDIForm1.TxtUsername.Text + "','" + Replace(txtlocation.Text, "\", "/") + "')"
            M_OBJCONN.Execute (STRSQL)
            cbosheet.Text = ""
            txtcount.Text = ""
            txtlocation.Text = ""
            If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close

Exit Function
KE:
    MsgBox "Keterangan eror di " + err.Description + "di sumber" + err.Source
    
End Function


Private Sub Form_Load()
ListView1.ColumnHeaders.ADD 1, , "Tgl Swap", 10 * TXT
ListView1.ColumnHeaders.ADD 2, , "Direcotry Swap", 10 * TXT
ListView1.ColumnHeaders.ADD 3, , "Jumlah Swap", 10 * TXT
ListView1.ColumnHeaders.ADD 4, , "User Exceute", 10 * TXT
End Sub
