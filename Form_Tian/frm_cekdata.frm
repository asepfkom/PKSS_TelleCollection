VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_cekdata 
   Caption         =   "Cek Data"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14250
   LinkTopic       =   "Form3"
   ScaleHeight     =   8400
   ScaleWidth      =   14250
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8715
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   14190
      _ExtentX        =   25030
      _ExtentY        =   15372
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Query Analyzer  "
      TabPicture(0)   =   "frm_cekdata.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdExportXLSX"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CmdSearchBaru(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "History      "
      TabPicture(1)   =   "frm_cekdata.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image3(1)"
      Tab(1).Control(1)=   "Line1(1)"
      Tab(1).Control(2)=   "Line1(4)"
      Tab(1).Control(3)=   "Label1(2)"
      Tab(1).Control(4)=   "txtjmlrow(1)"
      Tab(1).Control(5)=   "DataGrid4"
      Tab(1).Control(6)=   "CmdSearchBaru(1)"
      Tab(1).ControlCount=   7
      Begin VB.CommandButton CmdSearchBaru 
         Height          =   360
         Index           =   0
         Left            =   12420
         Picture         =   "frm_cekdata.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   330
         Width           =   1785
      End
      Begin VB.TextBox txtjmlrow 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   -62670
         MaxLength       =   20
         TabIndex        =   19
         Top             =   6405
         Width           =   1785
      End
      Begin VB.TextBox Text1 
         Height          =   3075
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   360
         Width           =   12390
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export to Excel"
         Height          =   345
         Left            =   12465
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.CommandButton CmdSearchBaru 
         Height          =   360
         Index           =   1
         Left            =   -74865
         Picture         =   "frm_cekdata.frx":0626
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   405
         Width           =   1785
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remarks "
         Height          =   1035
         Left            =   30
         TabIndex        =   3
         Top             =   2400
         Visible         =   0   'False
         Width           =   12315
         Begin VB.TextBox Text2 
            Height          =   735
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   12150
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Delete"
         Height          =   345
         Left            =   12480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1140
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.CommandButton cmdExportXLSX 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export to Excel (.xlsx)"
         Height          =   345
         Left            =   12480
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   1620
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   5190
         Left            =   -74865
         TabIndex        =   6
         Top             =   900
         Width           =   13830
         _ExtentX        =   24395
         _ExtentY        =   9155
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
      Begin TabDlg.SSTab SSTab2 
         Height          =   3570
         Left            =   0
         TabIndex        =   8
         Top             =   3465
         Width           =   14145
         _ExtentX        =   24950
         _ExtentY        =   6297
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Result Query"
         TabPicture(0)   =   "frm_cekdata.frx":0C14
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1(3)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Line1(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Image3(2)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label1(1)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Cd_save"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "DataGrid1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "TxtPath"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtjmlrow(0)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Table Definition"
         TabPicture(1)   =   "frm_cekdata.frx":0C30
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).Control(1)=   "Frame2"
         Tab(1).ControlCount=   2
         Begin VB.TextBox txtjmlrow 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   0
            Left            =   12150
            MaxLength       =   20
            TabIndex        =   15
            Top             =   3015
            Width           =   1785
         End
         Begin VB.Frame Frame1 
            Caption         =   "Table Name"
            Height          =   3120
            Left            =   -74955
            TabIndex        =   12
            Top             =   360
            Width           =   5055
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   2760
               Left            =   90
               TabIndex        =   13
               Top             =   315
               Width           =   4830
               _ExtentX        =   8520
               _ExtentY        =   4868
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
         End
         Begin VB.Frame Frame2 
            Caption         =   "Field Name"
            Height          =   3075
            Left            =   -69690
            TabIndex        =   10
            Top             =   360
            Width           =   5010
            Begin MSDataGridLib.DataGrid DataGrid3 
               Height          =   2760
               Left            =   45
               TabIndex        =   11
               Top             =   315
               Width           =   4830
               _ExtentX        =   8520
               _ExtentY        =   4868
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
         End
         Begin VB.TextBox TxtPath 
            Enabled         =   0   'False
            Height          =   330
            Left            =   1725
            TabIndex        =   9
            Top             =   3060
            Width           =   9615
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   2310
            Left            =   90
            TabIndex        =   14
            Top             =   495
            Width           =   14010
            _ExtentX        =   24712
            _ExtentY        =   4075
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
         Begin MSComDlg.CommonDialog Cd_save 
            Left            =   1350
            Top             =   3060
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "*.xlsx"
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "File Save :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   3120
            Width           =   960
         End
         Begin VB.Image Image3 
            Height          =   18630
            Index           =   2
            Left            =   0
            Picture         =   "frm_cekdata.frx":0C4C
            Top             =   360
            Width           =   26295
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   0
            Left            =   11340
            TabIndex        =   17
            Top             =   3060
            Width           =   810
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000080FF&
            BorderWidth     =   2
            Index           =   0
            X1              =   135
            X2              =   13950
            Y1              =   2925
            Y2              =   2925
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "File Save :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   3
            Left            =   135
            TabIndex        =   16
            Top             =   3105
            Width           =   960
         End
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   1
         Left            =   -75000
         Picture         =   "frm_cekdata.frx":8256
         Top             =   315
         Width           =   26295
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   0
         Left            =   -90
         Picture         =   "frm_cekdata.frx":F860
         Top             =   330
         Width           =   26295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -75000
         X2              =   -60825
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         Index           =   4
         X1              =   -75000
         X2              =   -60870
         Y1              =   6315
         Y2              =   6315
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   2
         Left            =   -63480
         TabIndex        =   21
         Top             =   6450
         Width           =   810
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   1
      Left            =   120
      Picture         =   "frm_cekdata.frx":16E6A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   420
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Query Cek Data "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   1
      Left            =   615
      TabIndex        =   22
      Top             =   300
      Width           =   3585
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   -1920
      Picture         =   "frm_cekdata.frx":17974
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19980
   End
End
Attribute VB_Name = "frm_cekdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Public Sub loadTabel()
Dim clsQuery As clsQuery_analyzer

Set clsQuery = New clsQuery_analyzer
Set M_objrs = clsQuery.getAllTable()
Set DataGrid2.DATASOURCE = M_objrs
Set M_objrs = Nothing
Set clsQuery = Nothing
End Sub

Private Sub cmdExportXLSX_Click()
    isi_data_export
End Sub

Private Sub isi_data_export()
    On Error GoTo Salah
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim cmdsql_update As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i As Integer
    Dim m_msgbox As String
    Dim clsQuery As clsQuery_analyzer

    i = 1
    Set clsQuery = New clsQuery_analyzer

DoEvents
    Set M_objrs = clsQuery.getQuerys(Text1.text)
    'Jika data tidak ada, maka keluar dari fungsi ini!
    If M_objrs.RecordCount = 0 Then
        MsgBox "Query Blank", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If

Form_Save:
    Cd_save.ShowSave
    Cd_save.DefaultExt = "xlsx"
    TxtPath.text = Cd_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.text = Empty Then
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
    
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
'    If WriteRecordsetToCSv(M_OBJRS, TXTPATH.Text + ".CSV", ",") Then
'        MsgBox " export berhasil"
'    End If
    

 Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
    
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
    
    
    On Error GoTo Salah
    'Proses pengsisian nama field ke excel
    Dim X, Y    As Integer
    If M_objrs.State = 1 Then
        X = 0
        Y = M_objrs.fields().Count - 1
        Do Until X > Y
            DoEvents
            objSheet.Cells(1, i).Value = CStr(M_objrs.fields(X).Name)
            objSheet.Cells(1, i).HorizontalAlignment = -4108
            i = i + 1
            X = X + 1
        Loop
    End If
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A2
    
    'objBook.SaveAs TxtPath.Text, xlWorkbookNormal
    objBook.SaveAs TxtPath.text, xlOpenXMLWorkbook
    
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
 
Salah:
    Exit Sub
End Sub
Private Sub CmdSearchBaru_Click(Index As Integer)
Dim clsQuery As clsQuery_analyzer
Dim M_OBJRS3 As ADODB.Recordset
Set clsQuery = New clsQuery_analyzer
Set M_OBJRS3 = New ADODB.Recordset
Select Case Index
    Case 0
        On Error GoTo ke
        'Set m_objrs1 = clsQuery.getQuerys(Text1.Text)
        'Set m_objrs1 = New ADODB.Recordset
        Call ConnectRS(rs)
        Call UnConnectRs(rs)
        DoEvents
        rs.Open Text1.text
        'M_OBJCONN.Execute Text1.Text
        txtjmlrow(0).text = ""
        txtjmlrow(0).text = rs.RecordCount
        Set DataGrid1.DATASOURCE = rs
        'Set m_objrs1 = Nothing
        SSTab2.Tab = 0
        'Set clsQuery = Nothing
        
    Case 1
        On Error GoTo ke
        Set M_objrs = clsQuery.getLoghst
        txtjmlrow(1).text = ""
        txtjmlrow(1).text = M_objrs.RecordCount
        Set DataGrid4.DATASOURCE = M_objrs
        Set M_objrs = Nothing
        SSTab2.Tab = 0
        Set clsQuery = Nothing
End Select
Exit Sub
ke:
MsgBox "Query Wrong", vbInformation + vbOKOnly, "Pesan"
End Sub

Private Sub Command1_Click()
If Text2.text = "" Then
   MsgBox "Anda Harus Beri alasan untuk delete data"
   Exit Sub
End If
End Sub

Private Sub Command2_Click()
ExportToexcel
End Sub

Private Sub DataGrid2_Click()
Dim clsQuery As clsQuery_analyzer
Set clsQuery = New clsQuery_analyzer
Set M_objrs = clsQuery.getFieldTable(DataGrid2.Columns(0).text)
Set DataGrid3.DATASOURCE = M_objrs
Set M_objrs = Nothing
Set clsQuery = Nothing
End Sub
Private Sub Form_Load()
Call ConnectRS(rs)
    loadTabel
End Sub
Public Sub ExportToexcel()
'On Error GoTo SALAH
Dim clsQuery As clsQuery_analyzer
   ' Dim objExcel        As Excel.Application
   ' Dim objBook         As Excel.Workbook
   ' Dim objSheet        As Excel.Worksheet
    Dim i As Integer
    Dim m_msgbox As String
    
    i = 1
   Set clsQuery = New clsQuery_analyzer

DoEvents
   Set M_objrs = clsQuery.getQuerys(Text1.text)
       
    'Jika data tidak ada, maka keluar dari fungsi ini!
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data submit tidak ada!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If

   
Form_Save:
    Cd_save.ShowSave
    TxtPath.text = Cd_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.text = Empty Then
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
    'Proses pengsisian nama field ke excel
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
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs TxtPath.text, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    

    sStrsql = "insert into tbllogexport(tbllogexport_filename,tbllogexport_command, tbllogexport_rows,  tbllogexport_nama_user) values (  "
    sStrsql = sStrsql + "'" + Replace(TxtPath.text, "/", "\") + "','" + Replace(Text1.text, "'", "`") + "','" + CStr(M_objrs.RecordCount) + "','" + MDIForm1.TxtUsername.text + "')"
    M_OBJCONN.Execute (sStrsql)
  '  objExcel.Quit
   ' Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
Exit Sub
Salah:
MsgBox err.Description, vbInformation + vbOKOnly, "Pesan"
    Exit Sub

End Sub

