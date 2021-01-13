VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmRestoreDelete 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Delete dan Restore Account"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13065
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8115
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   14314
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Delete Account"
      TabPicture(0)   =   "FrmRestoreDelete.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CommonDialog1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Restore Account"
      TabPicture(1)   =   "FrmRestoreDelete.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2640
         Top             =   7200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         Caption         =   "Restore Account"
         Height          =   7395
         Left            =   -74820
         TabIndex        =   17
         Top             =   540
         Width           =   12435
         Begin VB.Frame Frame3 
            Caption         =   "Restore From Tabel Backup"
            Height          =   6855
            Left            =   360
            TabIndex        =   18
            Top             =   360
            Width           =   11655
            Begin VB.CommandButton Command1 
               Caption         =   "Refresh data tabel backup ..."
               Height          =   315
               Left            =   6000
               TabIndex        =   41
               Top             =   1260
               Width           =   2595
            End
            Begin MSComctlLib.ProgressBar Pb1 
               Height          =   375
               Left            =   3420
               TabIndex        =   40
               Top             =   5880
               Width           =   7995
               _ExtentX        =   14102
               _ExtentY        =   661
               _Version        =   393216
               BorderStyle     =   1
               Appearance      =   0
            End
            Begin VB.CommandButton CmdCariData 
               Caption         =   "&Cari data >>"
               Height          =   495
               Left            =   9060
               TabIndex        =   39
               Top             =   2460
               Width           =   1935
            End
            Begin VB.CommandButton CmdRestoreFromExcelTabel 
               Caption         =   "&Restore"
               Height          =   435
               Left            =   9840
               TabIndex        =   38
               Top             =   6300
               Width           =   1575
            End
            Begin VB.TextBox TxtDataDitemukan 
               Height          =   285
               Left            =   4500
               TabIndex        =   36
               Text            =   "0"
               Top             =   6300
               Width           =   975
            End
            Begin VB.TextBox TxtJmlRestoreSource 
               Height          =   285
               Left            =   1380
               TabIndex        =   32
               Text            =   "0"
               Top             =   6300
               Width           =   1335
            End
            Begin VB.ComboBox CboSheetLocationRestore 
               Height          =   315
               Left            =   1440
               TabIndex        =   24
               Top             =   780
               Width           =   2565
            End
            Begin VB.CommandButton CmdRestoreLocation 
               BackColor       =   &H00C0FFC0&
               Caption         =   "...."
               Height          =   315
               Left            =   7320
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   360
               Width           =   555
            End
            Begin VB.TextBox TxtLocationRestore 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   360
               Width           =   5865
            End
            Begin MSComctlLib.ListView LvTblBck 
               Height          =   1440
               Left            =   900
               TabIndex        =   26
               Top             =   1620
               Width           =   7740
               _ExtentX        =   13653
               _ExtentY        =   2540
               View            =   3
               LabelEdit       =   1
               SortOrder       =   -1  'True
               Sorted          =   -1  'True
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
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
            Begin MSDataGridLib.DataGrid DataGrid_SourceCustid 
               Height          =   2595
               Left            =   180
               TabIndex        =   29
               Top             =   3540
               Width           =   2985
               _ExtentX        =   5265
               _ExtentY        =   4577
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
            Begin MSDataGridLib.DataGrid DataGrid_DataDitemukan 
               Height          =   2295
               Left            =   3420
               TabIndex        =   33
               Top             =   3540
               Width           =   8025
               _ExtentX        =   14155
               _ExtentY        =   4048
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
            Begin VB.Label Label17 
               Caption         =   "*) Jika anda yakin data yang di list adalah data yang benar untuk di restore, klik tombol restore."
               Height          =   495
               Left            =   5760
               TabIndex        =   37
               Top             =   6240
               Width           =   3915
            End
            Begin VB.Label Label16 
               Caption         =   "Jumlah data:"
               Height          =   195
               Left            =   3360
               TabIndex        =   35
               Top             =   6360
               Width           =   1095
            End
            Begin VB.Label Label15 
               Caption         =   "List data yang ditemukan dalam tabel backup:"
               Height          =   195
               Left            =   3420
               TabIndex        =   34
               Top             =   3300
               Width           =   4035
            End
            Begin VB.Label Label14 
               Caption         =   "Jumlah data:"
               Height          =   195
               Left            =   240
               TabIndex        =   31
               Top             =   6360
               Width           =   1095
            End
            Begin VB.Label Label13 
               Caption         =   "Source custid dari excel:"
               Height          =   195
               Left            =   180
               TabIndex        =   30
               Top             =   3300
               Width           =   2175
            End
            Begin VB.Line Line1 
               X1              =   180
               X2              =   11460
               Y1              =   3240
               Y2              =   3240
            End
            Begin VB.Label LblTblBCK 
               Height          =   195
               Left            =   1920
               TabIndex        =   28
               Top             =   1320
               Width           =   6615
            End
            Begin VB.Label Label12 
               Caption         =   "* Pilih salah satu source tabel backup dimana data tersebut ada, kemudian klik tombol cari data!"
               Height          =   795
               Left            =   8760
               TabIndex        =   27
               Top             =   1620
               Width           =   2475
            End
            Begin VB.Label Label11 
               Caption         =   "Source Tabel Backup"
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   1320
               Width           =   1695
            End
            Begin VB.Label Label10 
               Caption         =   "Sheet"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   840
               Width           =   795
            End
            Begin VB.Label Label9 
               Caption         =   "* File Excel (.xls) dengan isi kolom adalah custid"
               Height          =   315
               Left            =   7920
               TabIndex        =   22
               Top             =   360
               Width           =   3375
            End
            Begin VB.Label Label8 
               Caption         =   "Source custid"
               Height          =   315
               Left            =   120
               TabIndex        =   19
               Top             =   360
               Width           =   1035
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Delete Account"
         Height          =   6135
         Left            =   240
         TabIndex        =   2
         Top             =   540
         Width           =   12495
         Begin VB.TextBox txtbckup 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1620
            TabIndex        =   9
            Top             =   1200
            Width           =   5355
         End
         Begin VB.CommandButton cmddel 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Delete"
            Height          =   495
            Left            =   10680
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   5520
            Width           =   1455
         End
         Begin VB.ComboBox cbosheet 
            Height          =   315
            Left            =   1260
            TabIndex        =   7
            Text            =   "cbosheet"
            Top             =   660
            Width           =   2565
         End
         Begin VB.CommandButton cmdbrowse 
            BackColor       =   &H00C0FFC0&
            Caption         =   "...."
            Height          =   315
            Left            =   7200
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   300
            Width           =   555
         End
         Begin VB.TextBox txtlocation 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   300
            Width           =   5865
         End
         Begin VB.CommandButton CmdBackup 
            BackColor       =   &H00C0FFC0&
            Caption         =   "...."
            Height          =   315
            Left            =   7200
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1200
            Width           =   555
         End
         Begin VB.TextBox txtcount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   3
            Text            =   "0"
            Top             =   5460
            Width           =   1425
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3495
            Left            =   120
            TabIndex        =   10
            Top             =   1920
            Width           =   12045
            _ExtentX        =   21246
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
         Begin VB.Label Label6 
            Caption         =   "Backup File Name"
            Height          =   255
            Left            =   180
            TabIndex        =   16
            Top             =   1260
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Sheet"
            Height          =   255
            Left            =   180
            TabIndex        =   15
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label3 
            Caption         =   "Source Custid"
            Height          =   255
            Left            =   180
            TabIndex        =   14
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "* File Excel (.xls) dengan isi kolom adalah custid"
            Height          =   315
            Left            =   7860
            TabIndex        =   13
            Top             =   300
            Width           =   4335
         End
         Begin VB.Label Label5 
            Caption         =   "Jumlah Data:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   5520
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "List data yang akan di delete:"
            Height          =   195
            Left            =   180
            TabIndex        =   11
            Top             =   1680
            Width           =   2115
         End
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Delete and Restore Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13035
   End
End
Attribute VB_Name = "FrmRestoreDelete"
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

Private Sub CmdBackup_Click()
    CommonDialog1.ShowSave
    txtbckup.Text = CommonDialog1.FileName
    
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

Private Sub CmdCariData_Click()
    Dim OBJRECORD As ADODB.Recordset
    Dim CustId As String
    Dim W As Integer
    
    On Error GoTo salah
    
    If TxtLocationRestore.Text = "" Then
        MsgBox "Source excel masih kosong!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    If CboSheetLocationRestore.Text = "" Then
        MsgBox "Source sheet excel masih kosong!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
  
    If LblTblBCK.Caption = "" Then
        MsgBox "Anda belum memilih tabel backup!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    CustId = ""
    
    'If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
     'Buat Ambil Data custid dari file excel
     Set OBJRECORD = New ADODB.Recordset
        OBJRECORD.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & CboSheetLocationRestore.Text & "] "
        DoEvents
        OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Set DataGrid_SourceCustid.DATASOURCE = OBJRECORD
        TxtJmlRestoreSource.Text = OBJRECORD.RecordCount
        
    If Val(TxtJmlRestoreSource.Text) = 0 Then
        MsgBox "Tidak ada source custid dari excel!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    While Not OBJRECORD.EOF
        If CustId = "" Then
            CustId = "'" + OBJRECORD("custid") + "'"
        Else
            CustId = CustId + ",'" + OBJRECORD("custid") + "'"
        End If
        OBJRECORD.MoveNext
    Wend
    
        
    'Buat ambil data dari tabel backup yang dipilih
    'If M_OBJCONN.state = adStateOpen Then M_OBJCONN.Close
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql As String
    
    cmdsql = " select * from " + Trim(LblTblBCK.Caption) + " where "
    cmdsql = cmdsql + "custid in (" + CustId + ")"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    Set DataGrid_DataDitemukan.DATASOURCE = M_objrs
    TxtDataDitemukan.Text = M_objrs.RecordCount
    Exit Sub
salah:
    MsgBox "Ada error: " & err, vbOKOnly + vbExclamation, "Peringatan"
End Sub

Private Sub cmddel_Click()
    Dim a As String
    
    a = MsgBox("Anda yakin akan menghapus data?", vbYesNo + vbInformation, "Konfirmasi")
    If a = vbNo Then
        Exit Sub
    End If
    
    If txtlocation.Text = "" Then
        MsgBox "Source custid masih kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    If cbosheet.Text = "" Then
        MsgBox "Sheet masih kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    If txtbckup.Text = "" Then
        MsgBox "Lokasi backup file masih kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    Dim cek_record As New ADODB.Recordset
    Dim NmFile As String
    Dim OBJRECORD As New ADODB.Recordset
    
'    Set RSTEMP = New ADODB.Recordset
'    RSTEMP.CursorLocation = adUseClient
'
'    ssql = "SELECT * FROM [" & cbosheet.Text & "] "
'    RSTEMP.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
'    Set RSTEMP = Nothing
    
     Set OBJRECORD = New ADODB.Recordset
        OBJRECORD.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & cbosheet.Text & "] "
        DoEvents
        OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        
    If OBJRECORD.RecordCount = 0 Then
        MsgBox "Tidak ada Data yang didelete!", vbInformation + vbOKOnly, "Pesan"
        Exit Sub
    End If
    
    Dim CustId As String
    CustId = ""
    While Not OBJRECORD.EOF
        If CustId = "" Then
            CustId = "'" + IIf(IsNull(OBJRECORD("custid")), "", OBJRECORD("custid")) + "'"
        Else
            CustId = CustId + ",'" + IIf(IsNull(OBJRECORD("custid")), "", OBJRECORD("custid")) + "'"
        End If
        OBJRECORD.MoveNext
    Wend
    
    
    NmFile = "bckupupload_del_" + Format(MDIForm1.TDBDate1, "ddmmyyyy") + "_" + Format(Time, "hhmmss")
    strQuery = " select * from mgm where custid in (" + CustId + ")"
    
    STRSQL = "create table  " + NmFile + "  as  " + strQuery
    M_OBJCONN.Execute (STRSQL)
    
    
    Set cek_record = New ADODB.Recordset
    cek_record.CursorLocation = adUseClient
    'cek_record.Open "select distinct table_name from information_schema.columns where table_catalog='ritcard' and table_schema='public' and table_name ='" + NmFile + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
    cek_record.Open "select distinct table_name from information_schema.columns where  table_schema='public' and table_name ='" + NmFile + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If cek_record.BOF And cek_record.EOF Then
        MsgBox "Record gagal di backup ....... !"
        Exit Sub
    Else
        STRSQL = " insert into tbl_hst_upload_del(path_excel,path_didb,user_excecute)  values ('" + Replace(txtbckup.Text, "\", "/") + "','" + NmFile + "','" + MDIForm1.txtusername.Text + "')"
        M_OBJCONN.Execute (STRSQL)
        'Buat Backup ke File Excel
        isi_data txtbckup.Text, strQuery
        M_OBJCONN.Execute ("delete from mgm where custid in (" + CustId + ")")
        MsgBox "Data Berhasil dihapus!", vbInformation + vbOKOnly, "Informasi"
        MsgBox "Jangan lupa catat di dalam log: Nama Tabel Backup -> " & NmFile & " , Nama File Backup -> " & txtbckup.Text, vbOKOnly + vbInformation, "Informasi"
        txtlocation.Text = ""
        txtbckup.Text = ""
        cbosheet.Text = ""
        txtcount.Text = "0"
        cbosheet.CLEAR
        If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
    End If
End Sub

'Buat Export Ke Excel
Private Sub isi_data(sPath As String, ssql)
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql As String
    Dim listitem As listitem
    Dim cmdsql_update As String
    Dim objExcel As Excel.Application
    Dim objBook  As Excel.Workbook
    Dim objSheet As Excel.Worksheet
    Dim i As Integer
    Dim m_msgbox As String
    
    i = 1
    
   
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic

    
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data backup tidak ada !", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
   
    
    
    
    
    'Set excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
    
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
    
    
    On Error GoTo salah
    'Proses pengsisian nama field ke excel
    Dim x, Y    As Integer
        If M_objrs.state = 1 Then
            x = 0
            Y = M_objrs.fields().Count - 1
            Do Until x > Y
                DoEvents
                objSheet.Cells(1, i).Value = CStr(M_objrs.fields(x).Name)
                i = i + 1
                x = x + 1
            Loop
        End If
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs sPath, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
salah:
    Exit Sub
End Sub

Private Sub CmdRestoreFromExcelTabel_Click()
    Dim OBJRECORD As ADODB.Recordset
    Dim CustId As String
    Dim W As Integer
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    
    On Error GoTo salah
    
    If TxtLocationRestore.Text = "" Then
        MsgBox "Source excel masih kosong!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    If CboSheetLocationRestore.Text = "" Then
        MsgBox "Source sheet excel masih kosong!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
  
    If LblTblBCK.Caption = "" Then
        MsgBox "Anda belum memilih tabel backup!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    
    If Val(TxtJmlRestoreSource.Text) = 0 Then
        MsgBox "Data custid yang akan di restore tidak ada! Coba source custid di file excel!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If Val(TxtDataDitemukan.Text) = 0 Then
        MsgBox "Data tidak ditemukan dalam tabel backup! Coba cek tabel backup yang lain!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Cek dulu apakah ada data yang sama
     Set OBJRECORD = New ADODB.Recordset
        OBJRECORD.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & CboSheetLocationRestore.Text & "] "
        DoEvents
        OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Set DataGrid_SourceCustid.DATASOURCE = OBJRECORD
        TxtJmlRestoreSource.Text = OBJRECORD.RecordCount
        
    If Val(TxtJmlRestoreSource.Text) = 0 Then
        MsgBox "Tidak ada source custid dari excel!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    While Not OBJRECORD.EOF
        If CustId = "" Then
            CustId = "'" + OBJRECORD("custid") + "'"
        Else
            CustId = CustId + ",'" + OBJRECORD("custid") + "'"
        End If
        OBJRECORD.MoveNext
    Wend
    
    cmdsql = "select * from mgm where custid in (" + CustId + ")"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
       MsgBox "Mohon maaf data yang di restore ada di dalam tabel utama MGM! Data gagal dihapus!", vbOKOnly + vbInformation, "Informasi"
       Dim CustidSama As String
       CustidSama = ""
       While Not M_objrs.EOF
        If CustidSama = "" Then
            CustidSama = M_objrs("custid")
        Else
            CustidSama = CustidSama + ", " + M_objrs("custid")
        End If
        M_objrs.MoveNext
       Wend
       MsgBox "Custid yang sama ada di tabel MGM adalah: " & CustidSama, vbOKOnly + vbInformation, "Informasi"
       Set M_objrs = Nothing
       Exit Sub
    End If
    Set M_objrs = Nothing
    
    'Lakukan konfirmasi
    a = MsgBox("Anda yakin akan melakukan Restore data?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Proses Restore dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Proses Restore
    cmdsql = "insert into mgm "
    cmdsql = cmdsql + "select * from " + LblTblBCK.Caption + " where custid in ("
    cmdsql = cmdsql + CustId + ")"
    
    M_OBJCONN.Execute cmdsql
    MsgBox "Data berhasil di restore sebanyak: " + CStr(TxtDataDitemukan.Text), vbOKOnly + vbInformation, "Informasi"
    
    TxtLocationRestore.Text = ""
    CboSheetLocationRestore.CLEAR
    CboSheetLocationRestore.Text = ""
    LblTblBCK.Caption = ""
    TxtJmlRestoreSource.Text = 0
    TxtDataDitemukan.Text = 0
    
    If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
    'If M_OBJCONN.state = adStateOpen Then M_OBJCONN.Close
    
    Exit Sub
salah:
    MsgBox "Ada error: " & err, vbOKOnly + vbExclamation, "Peringatan"
    
End Sub

Private Sub CmdRestoreLocation_Click()
    With CommonDialog1
            .DialogTitle = "Import From File"
            .Filter = "Excel Files|*.xls"
            .ShowOpen
        End With
        
        TxtLocationRestore.Text = CommonDialog1.FileName
        If CommonDialog1.FileName = "" Then Exit Sub
        If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
                M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & TxtLocationRestore.Text & ";Extended Properties=Excel 8.0;"
        Set RSTEMP = M_XLSCONN.OpenSchema(adSchemaTables)
        cbosheet.CLEAR
        If RSTEMP.EOF And RSTEMP.BOF Then Exit Sub
        While Not RSTEMP.EOF
            CboSheetLocationRestore.AddItem IIf(IsNull(RSTEMP!table_name), "", RSTEMP!table_name)
            RSTEMP.MoveNext
        Wend
        Set RSTEMP = Nothing
End Sub

Private Sub Command1_Click()
    Call Isi_Tabel_Backup
End Sub

Private Sub Form_Load()
    Call HeaderListBackup
    Call Isi_Tabel_Backup
End Sub

Private Sub HeaderListBackup()
    LvTblBck.ColumnHeaders.ADD 1, , "Id ", 1000
    LvTblBck.ColumnHeaders.ADD 2, , "Tgl.Backup", 3000
    LvTblBck.ColumnHeaders.ADD 3, , "Nama Tabel Backup", 5000
    LvTblBck.ColumnHeaders.ADD 4, , "User eksekusi", 4000
    LvTblBck.ColumnHeaders.ADD 5, , "Path backup excel", 5000
End Sub

Private Sub Isi_Tabel_Backup()
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql As String
    Dim listitem As listitem
    
    cmdsql = "select distinct tbl_hst_upload_del.id, tbl_hst_upload_del.path_didb,"
    cmdsql = cmdsql + "tbl_hst_upload_del.path_excel, "
    cmdsql = cmdsql + "tbl_hst_upload_del.tgl_execute, tbl_hst_upload_del.user_excecute "
    cmdsql = cmdsql + "from information_schema.columns as ic,tbl_hst_upload_del "
    cmdsql = cmdsql + " where ic.table_schema='public' and ic.table_name=tbl_hst_upload_del.path_didb "
    cmdsql = cmdsql + " order by tbl_hst_upload_del.tgl_execute desc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
        Set M_objrs = Nothing
    End If
    LvTblBck.ListItems.CLEAR
    While Not M_objrs.EOF
        Set listitem = LvTblBck.ListItems.ADD(, , M_objrs("id"))
            listitem.SubItems(1) = IIf(IsNull(M_objrs("tgl_execute")), "", Format(M_objrs("tgl_execute"), "yyyy-mm-dd hh:mm:ss"))
            listitem.SubItems(2) = IIf(IsNull(M_objrs("path_didb")), "", M_objrs("path_didb"))
            listitem.SubItems(3) = IIf(IsNull(M_objrs("user_excecute")), "", M_objrs("user_excecute"))
            listitem.SubItems(4) = IIf(IsNull(M_objrs("path_excel")), "", M_objrs("path_excel"))
        M_objrs.MoveNext
    Wend
End Sub



Private Sub LvTblBck_Click()
    If LvTblBck.ListItems.Count = 0 Then
        MsgBox "Tabel backup tidak ada!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    LblTblBCK.Caption = IIf(IsNull(LvTblBck.SelectedItem.SubItems(2)), "", LvTblBck.SelectedItem.SubItems(2))
    
End Sub
