VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Frm_Upload_Data 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Upload Data Customer"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   17280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   17280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Upload Data"
      Height          =   2025
      Left            =   0
      TabIndex        =   13
      Top             =   480
      Width           =   17235
      Begin VB.TextBox TxtPath 
         Height          =   315
         Left            =   9840
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   1020
         Width           =   1095
      End
      Begin VB.ComboBox cbosheet 
         Height          =   315
         Left            =   1380
         TabIndex        =   19
         Top             =   1020
         Width           =   2565
      End
      Begin VB.CommandButton cmdbrowse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "...."
         Height          =   315
         Left            =   9870
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   630
         Width           =   555
      End
      Begin VB.CommandButton cmdcreatemap 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Create Map"
         Height          =   285
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   270
         Width           =   1155
      End
      Begin VB.TextBox txtlocation 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   630
         Width           =   8445
      End
      Begin VB.ComboBox cbomap 
         Height          =   315
         ItemData        =   "Frm_Upload_Data.frx":0000
         Left            =   1380
         List            =   "Frm_Upload_Data.frx":0002
         TabIndex        =   15
         Top             =   300
         Width           =   2595
      End
      Begin VB.TextBox TXTDESCRIPTION_BATCH 
         Height          =   315
         Left            =   1410
         TabIndex        =   14
         Top             =   1470
         Width           =   9495
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   11340
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   360
         Left            =   5220
         TabIndex        =   22
         Top             =   1020
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComDlg.CommonDialog Cd_save 
         Left            =   12120
         Top             =   420
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Ms. Excel 97/2000/XP|*.xls"
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   11970
         Top             =   1110
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label5 
         Height          =   345
         Left            =   7590
         TabIndex        =   28
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label lblstatus 
         Height          =   345
         Left            =   5220
         TabIndex        =   27
         Top             =   1020
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Sheet"
         Height          =   255
         Left            =   180
         TabIndex        =   26
         Top             =   1050
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Location"
         Height          =   255
         Left            =   150
         TabIndex        =   25
         Top             =   630
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Select Mapping"
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label Label14 
         Caption         =   "Ket_Batch"
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   1500
         Width           =   795
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
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Exit"
         Height          =   495
         Left            =   15540
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtnew 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   255
         Width           =   1905
      End
      Begin VB.TextBox txtexisting 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   570
         Width           =   1905
      End
      Begin VB.TextBox txtlead 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   210
         Width           =   1905
      End
      Begin VB.TextBox txtdonot 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1905
      End
      Begin VB.TextBox txtdouble 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   8700
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   510
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label15 
         Caption         =   "New Data :"
         Height          =   285
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label16 
         Caption         =   "Existing :"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   570
         Width           =   825
      End
      Begin VB.Label Label17 
         Caption         =   "Total Lead :"
         Height          =   285
         Left            =   3210
         TabIndex        =   10
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label Label18 
         Caption         =   "Total Do Not call :"
         Height          =   285
         Left            =   3210
         TabIndex        =   9
         Top             =   570
         Width           =   1395
      End
      Begin VB.Label Label20 
         Caption         =   "Total Cust card >1 card"
         Height          =   285
         Left            =   6870
         TabIndex        =   8
         Top             =   510
         Visible         =   0   'False
         Width           =   1995
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6465
      Left            =   0
      TabIndex        =   29
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
      TabPicture(0)   =   "Frm_Upload_Data.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstview"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Cboexecelmap"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "View Mapping     "
      TabPicture(1)   =   "Frm_Upload_Data.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstmapping"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "History Upload      "
      TabPicture(2)   =   "Frm_Upload_Data.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label11"
      Tab(2).Control(1)=   "lsthst"
      Tab(2).Control(2)=   "cmdrefresh_hst"
      Tab(2).Control(3)=   "txtnumrowshst"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Error In Excel        "
      TabPicture(3)   =   "Frm_Upload_Data.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label12"
      Tab(3).Control(1)=   "lsterror"
      Tab(3).Control(2)=   "txtfound"
      Tab(3).Control(3)=   "Command2"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Lead To Database      "
      TabPicture(4)   =   "Frm_Upload_Data.frx":0074
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame6"
      Tab(4).Control(1)=   "Frame3"
      Tab(4).Control(2)=   "Frame2"
      Tab(4).ControlCount=   3
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export to Excel"
         Height          =   345
         Left            =   -74910
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   390
         Width           =   2115
      End
      Begin VB.ComboBox Cboexecelmap 
         Height          =   315
         Left            =   2820
         TabIndex        =   55
         Top             =   990
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Frame Frame2 
         Caption         =   "View Lead To be Insert to database"
         Height          =   5895
         Left            =   -71370
         TabIndex        =   51
         Top             =   390
         Width           =   4545
         Begin VB.TextBox txtlead_masuk 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   3060
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   5490
            Width           =   1245
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5145
            Left            =   150
            TabIndex        =   53
            Top             =   270
            Width           =   4275
            _ExtentX        =   7541
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
            Left            =   2520
            TabIndex        =   54
            Top             =   5490
            Width           =   795
         End
      End
      Begin VB.TextBox txtnumrowshst 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -59760
         TabIndex        =   50
         Top             =   6000
         Width           =   1605
      End
      Begin VB.TextBox txtfound 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -60030
         TabIndex        =   49
         Top             =   5970
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "Lead Source"
         Height          =   5895
         Left            =   -74940
         TabIndex        =   45
         Top             =   390
         Width           =   3525
         Begin VB.TextBox txtrowssource 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   2190
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   5550
            Width           =   1245
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   5235
            Left            =   120
            TabIndex        =   47
            Top             =   270
            Width           =   3315
            _ExtentX        =   5847
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
            TabIndex        =   48
            Top             =   5580
            Width           =   555
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Existing Data"
         Height          =   5895
         Left            =   -66810
         TabIndex        =   31
         Top             =   390
         Width           =   8925
         Begin VB.CommandButton cmddel 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Delete"
            Height          =   495
            Left            =   7320
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   5310
            Width           =   1455
         End
         Begin VB.TextBox txtbckup 
            Height          =   375
            Left            =   1830
            TabIndex        =   32
            Top             =   5370
            Width           =   5355
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   4875
            Left            =   90
            TabIndex        =   33
            Top             =   240
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   8599
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Cek Existing"
            TabPicture(0)   =   "Frm_Upload_Data.frx":0090
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Label10"
            Tab(0).Control(1)=   "ListView1"
            Tab(0).Control(2)=   "CndUncekAll"
            Tab(0).Control(3)=   "CmdCekAll"
            Tab(0).Control(4)=   "Txtxadarow"
            Tab(0).ControlCount=   5
            TabCaption(1)   =   "History"
            TabPicture(1)   =   "Frm_Upload_Data.frx":00AC
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Label9"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "ListView2"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "CMD_HST"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "txtrowhst_del"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).ControlCount=   4
            Begin VB.TextBox txtrowhst_del 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Left            =   7350
               Locked          =   -1  'True
               TabIndex        =   38
               Top             =   4410
               Width           =   1245
            End
            Begin VB.TextBox Txtxadarow 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Left            =   -67830
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   4440
               Width           =   1245
            End
            Begin VB.CommandButton CMD_HST 
               BackColor       =   &H00C0FFC0&
               Caption         =   "&Refresh"
               Height          =   435
               Left            =   150
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   420
               Width           =   1455
            End
            Begin VB.CommandButton CmdCekAll 
               Caption         =   "&Cek All"
               Height          =   375
               Left            =   -74880
               TabIndex        =   35
               Top             =   4380
               Width           =   1035
            End
            Begin VB.CommandButton CndUncekAll 
               Caption         =   "&UnCek All"
               Height          =   375
               Left            =   -73860
               TabIndex        =   34
               Top             =   4380
               Width           =   1035
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   3825
               Left            =   -74970
               TabIndex        =   39
               Top             =   450
               Width           =   8565
               _ExtentX        =   15108
               _ExtentY        =   6747
               View            =   3
               LabelEdit       =   1
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin MSComctlLib.ListView ListView2 
               Height          =   3195
               Left            =   60
               TabIndex        =   40
               Top             =   960
               Width           =   8505
               _ExtentX        =   15002
               _ExtentY        =   5636
               View            =   3
               LabelEdit       =   1
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin VB.Label Label9 
               Caption         =   "Rows:"
               Height          =   255
               Left            =   6840
               TabIndex        =   42
               Top             =   4410
               Width           =   795
            End
            Begin VB.Label Label10 
               Caption         =   "Rows:"
               Height          =   255
               Left            =   -68340
               TabIndex        =   41
               Top             =   4440
               Width           =   795
            End
         End
         Begin VB.Label Label6 
            Caption         =   "Backup File Name"
            Height          =   255
            Left            =   330
            TabIndex        =   44
            Top             =   5430
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdrefresh_hst 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Refresh"
         Height          =   435
         Left            =   -74970
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   390
         Width           =   1455
      End
      Begin MSComctlLib.ListView lstview 
         Height          =   6015
         Left            =   30
         TabIndex        =   57
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
         TabIndex        =   58
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
         Height          =   5355
         Left            =   -74940
         TabIndex        =   59
         Top             =   870
         Width           =   17085
         _ExtentX        =   30136
         _ExtentY        =   9446
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
         TabIndex        =   60
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
         TabIndex        =   63
         Top             =   5520
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label11 
         Caption         =   "Num Of Rows :"
         Height          =   255
         Left            =   -60930
         TabIndex        =   62
         Top             =   6060
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Found :"
         Height          =   255
         Left            =   -60780
         TabIndex        =   61
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
      TabIndex        =   64
      Top             =   60
      Width           =   3585
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   30
      Picture         =   "Frm_Upload_Data.frx":00C8
      Stretch         =   -1  'True
      Top             =   30
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   465
      Left            =   0
      Picture         =   "Frm_Upload_Data.frx":0BD2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17220
   End
End
Attribute VB_Name = "Frm_Upload_Data"
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
   Set OBJRECORD = clscampaign.FindCampaign(cbocampaign.text)
   If OBJRECORD.RecordCount > 0 Then
     txtdescription.text = IIf(IsNull(OBJRECORD!keterangan), "", OBJRECORD!keterangan)
    Else
        txtdescription.text = ""
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
    CMDSQL = "select * from  tbldivisi where    nm_divisi='"
    CMDSQL = CMDSQL + cboket.text + "'"
    
    Set OBJRECORD = New ADODB.Recordset
    OBJRECORD.CursorLocation = adUseClient
    
    OBJRECORD.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not OBJRECORD.EOF Then
        cboproduct.text = IIf(IsNull(OBJRECORD("kddivisi")), "", OBJRECORD("kddivisi"))
    End If
    
    Set OBJRECORD = Nothing
End Sub

Private Sub cbomap_Click()
    findFx cbomap.text
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
    CMDSQL = "select * from  tbldivisi where kddivisi='"
    CMDSQL = CMDSQL + cboproduct.text + "'"
    
    Set OBJRECORD = New ADODB.Recordset
    OBJRECORD.CursorLocation = adUseClient
    
    OBJRECORD.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not OBJRECORD.EOF Then
        cboket.text = IIf(IsNull(OBJRECORD("nm_divisi")), "", OBJRECORD("nm_divisi"))
    End If
    
    Set OBJRECORD = Nothing
End Sub

Private Sub cbosheet_Click()
    Dim get_header As String
    
    lblstatus.Caption = ""
    If txtlocation.text <> "" Then
        If M_XLSCONN.State = adStateOpen Then M_XLSCONN.Close
        M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 8.0;"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & cbosheet.text & "] "
        M_objrs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        M_OBJCONN.Execute "delete from tbl_temp_field "
        get_header = ""
        If M_objrs.EOF And M_objrs.BOF Then Exit Sub
        For i = 0 To M_objrs.fields.Count - 1
            On Error Resume Next
            strsql = "insert into tbl_temp_field (nama_field) values ('" + M_objrs.fields(i).Name + "')"
            get_header = get_header & """" & M_objrs.fields(i).Name & """" & " varchar,"
            
            M_OBJCONN.Execute (strsql)
            lblstatus.Caption = "Field Terdefinisi"
        Next i
            get_header = Left(get_header, Len(get_header) - 1)
        Set M_objrs = Nothing
    End If

End Sub

Private Sub CMD_HST_Click()
    load_hst_upload_del
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
   Set M_XLSCONN = Nothing
End Sub

Private Sub CmdCekAll_Click()
    Dim W As Integer
    
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Data yang akan dihapus tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To ListView1.ListItems.Count
        ListView1.ListItems(W).Checked = True
    Next W
End Sub

Private Sub cmdcreatemap_Click()
   Form_setting_upload.Show 1
End Sub
Public Sub loadCboMap()
    cbomap.CLEAR
    ssql = "select DISTINCT(kode_source) from tbl_setting_upload  where (kode_source is not null or kode_source<>'')"
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
    'Buat Ambil Mappingan Upload Data
    Dim list As ListItem
    sStrsql = " select nama_kolom,field_destination,character_maximum_length,data_type from ( "
    sStrsql = sStrsql + " select * from ( "
    sStrsql = sStrsql + " SELECT column_name as nama_kolom,character_maximum_length,data_type From information_schema.Columns WHERE table_name='mgm' AND table_schema='public'"
    sStrsql = sStrsql + " and data_type in ('character varying','numeric','bigint','integer','timestamp without time zone','text') and column_name not in (select nama_kolom  from tbl_tidak_map) ORDER BY ordinal_position) as tblbaru "
    sStrsql = sStrsql + " full join  ( "
    sStrsql = sStrsql + "  select field_source,field_destination from tbl_setting_upload where kode_source='" + xCodeMap + "' ) "
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

Private Sub cmddel_Click()
Dim cek_record As New ADODB.Recordset
Dim bDel As Boolean
If ListView1.ListItems.Count = 0 Then
    MsgBox "Tidak ada Data yang dibackup ", vbInformation + vbOKOnly, "Pesan"
    Exit Sub
End If

If txtbckup.text = "" Then
  MsgBox "Select directory for backup with double click in backup file name", vbInformation + vbOKOnly, "Pesan"
    Exit Sub
End If

bDel = False
strdel = ""
For i = 1 To ListView1.ListItems.Count
    DoEvents
        If ListView1.ListItems(i).Checked = True Then
            If bDel = False Then
                    bDel = True
              strdel = "'" + ListView1.ListItems(i).text + "'"
            Else
                strdel = strdel + ",'" + ListView1.ListItems(i).text + "'"
            End If
        End If
Next i

If strdel = "" Then
        MsgBox "Tak ada Data yang didelete ", vbInformation + vbOKOnly, "Pesan"
        Exit Sub
End If
NmFile = "bckupupload_del_cust" + txtbckup.text
strQuery = " select ID , Name, addrnow, homeno, mobileno, addrpt, officeno, nocard, region, DOB, recsource, CustId, curbal, pay_dt, lastpay from tbl_upload_temp where custid in (" + strdel + ")"
strsql = "create table  " + NmFile + "  as " + strQuery
M_OBJCONN.Execute (strsql)


NmFile1 = "bckupupload_del_cust" + txtbckup.text



Set cek_record = New ADODB.Recordset
cek_record.CursorLocation = adUseClient
cek_record.Open "select distinct table_name from information_schema.columns where table_schema='public' and table_name ='" + NmFile1 + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic

If cek_record.BOF And cek_record.EOF Then
    MsgBox "Record gagal Backup ......."
    Exit Sub
Else

strsql = " insert into tbl_hst_upload_del(path_excel,path_didb,user_excecute)  values ('" + Replace(txtbckup.text, "\", "/") + "','" + NmFile + "','" + MDIForm1.TxtUsername.text + "')"
M_OBJCONN.Execute (strsql)

MsgBox "Data Berhasil dihapus", vbInformation + vbOKOnly, "Pesan"
isi_data txtbckup.text, strQuery
M_OBJCONN.Execute ("delete from tbl_upload_temp where custid in (" + strdel + ")")
M_OBJCONN.Execute ("delete from mgm where custid in (" + strdel + ")")
ListView1.ListItems.CLEAR
End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub CmdProses_Click()
    
    Dim mobjtemp As New ADODB.Recordset
    
    'cek map source sudah di isi apa belum
    If cbomap.text = "" Then
        MsgBox "Map Source  Belum di select ", vbOKOnly, "Information"
        cbomap.SetFocus
        Exit Sub
    End If
   
    'CEK FIELDNYA TERDEFINISI ATAU TIDAK
    If lblstatus.Caption = "" Then
        MsgBox "Field tidak terdefinisi mohon browse ulang excel ", vbOKOnly, "Information"
        cmdProses.Enabled = True
        Exit Sub
    End If
   
    If cekMANDATORTY = False Then
        MsgBox "Cek field mandatory harus tersedia recsource  and custid!", vbOKOnly, "Information"
        Exit Sub
    End If
           
    cmdProses.Enabled = False
    'VERIFIKASI FIELD YANG TERDEFINISI
    If cekmapping_excel = False Then
        MsgBox "Verifikasi Mapping Gagal karena field di mapping tidak terdefinisi di excel ", vbOKOnly, "Information"
        SSTab1.Tab = 1
        cmdProses.Enabled = True
        Label5.Caption = "Tidak Bisa Upload"
        Exit Sub
    End If
    
    Call cekStrukturField
    Set mobjtemp = New ADODB.Recordset
    mobjtemp.CursorLocation = adUseClient
    
    mobjtemp.Open "select * from tbl_upload_temp WHERE (CUSTID IS NOT NULL OR CUSTID<>'' )", M_OBJCONN, adOpenDynamic, adLockOptimistic
 '   Text1.Text = mobjtemp.RecordCount
    Set DataGrid1.DATASOURCE = mobjtemp
    cmdProses.Enabled = True
    
End Sub

Private Sub cmdrefresh_hst_Click()
   load_hst_upload
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

  

  
If Val(txtlead_masuk.text) = 0 And Val(txtexisting.text) = 0 Then
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
                    strfieldupdate = lstview.ListItems(jRow).text + "=a." + lstview.ListItems(jRow).text
                    strfieldheaderupdate = "tbl_upload_temp." + lstview.ListItems(jRow).text + ",tbl_upload_temp.recsource"
                    
                    If lstview.ListItems(jRow).text = "limit" Then
                    strinsert = Chr(34) + lstview.ListItems(jRow).text + Chr(34) + ""
                    Else
                    strinsert = lstview.ListItems(jRow).text + ""
                    End If
                    
                    ncount = 2
                Else
                    strfieldupdate = strfieldupdate + " ," + lstview.ListItems(jRow).text + "=a." + lstview.ListItems(jRow).text
                    strfieldheaderupdate = strfieldheaderupdate + ",tbl_upload_temp." + lstview.ListItems(jRow).text
                    
                  If lstview.ListItems(jRow).text = "limit" Then
                    strinsert = strinsert + "," + Chr(34) + lstview.ListItems(jRow).text + Chr(34)
                  Else
                    strinsert = strinsert + "," + lstview.ListItems(jRow).text
                  End If
                    
                End If
                    
        End If
    Next jRow

'update tbl_mst_performance set nbulan=a.nbulan


Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
      strsql = " select  mgm.* from mgm,tbl_upload_temp "
      strsql = strsql + " where mgm.custid=tbl_upload_temp.custid  and  (tbl_upload_temp.f_flag=1 or tbl_upload_temp.f_flag is null"
      strsql = strsql + " or tbl_upload_temp.f_flag=0 )"
      strsql = strsql + "and  (f_flag_donot is null or f_flag_donot=0) "

    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

If M_objrs.RecordCount <> 0 Then
    njmlExitst = M_objrs.RecordCount
    MsgBox "Data sudah tersedia sebanyak: " + CStr(njmlExitst) + "! Hapus dan backup terlebih dahulu data tersebut kemudian klik upload!", vbOKOnly + vbInformation, "Informasi"
    Exit Sub
    'If MsgBox("Data Sudah Pernah ada sebanyak : " + CStr(njmlExitst) + " Anda yakin " + vbCrLf + "Untuk MengGantikan isi data lama", vbQuestion + vbYesNo, "Question") = vbYes Then
    'End If
End If

Set M_objrs = Nothing

Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient

    
    strsql = "  select " + strinsert + " from  tbl_upload_temp where id not in"
    strsql = strsql + " ( "
    strsql = strsql + " select  tbl_upload_temp.id from mgm,tbl_upload_temp"
    strsql = strsql + " where mgm.custid=tbl_upload_temp.custid ) and (CUSTID IS NOT NULL OR CUSTID <>'') AND "
    strsql = strsql + " (tbl_upload_temp.f_flag=1 or tbl_upload_temp.f_flag is null or tbl_upload_temp.f_flag=0 )   and  (f_flag_donot is null or f_flag_donot=0)"
    'strsql = strsql + " AND (coalesce(homeno,'')<>''or coalesce(mobileno,'')<>'' or coalesce(officeno,'')<>'')"
    
    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

'insert  ke tbl_mst_performance
If M_objrs.RecordCount <> 0 Then
njmlNew = M_objrs.RecordCount
  If MsgBox("New Data :" + CStr(njmlNew) + vbCrLf + "", vbQuestion + vbYesNo, "Question") = vbYes Then
    If strinsert <> "" Then
        strsql = "insert into mgm (" + strinsert + ")"
        strsql = strsql + "  select " + strinsert + " from  tbl_upload_temp where id not in"
        strsql = strsql + " ( "
        strsql = strsql + " select  tbl_upload_temp.id from mgm,tbl_upload_temp"
        strsql = strsql + " where mgm.custid=tbl_upload_temp.custid ) and "
        strsql = strsql + " (tbl_upload_temp.f_flag=1 or tbl_upload_temp.f_flag is null or tbl_upload_temp.f_flag=0 )   and  (f_flag_donot is null or f_flag_donot=0) AND tbl_upload_temp.CUSTID IS NOT NULL"
        'strsql = strsql + " AND (coalesce(homeno,'')<>''or coalesce(mobileno,'')<>'' or coalesce(officeno,'')<>'')"
        
        M_OBJCONN.Execute (strsql)
        
        ' TAMBAHAN Field Class -> 835 di buat otomatis batch decease 28 NOP 2013
        M_OBJCONN.Execute "UPDATE mgm SET RECSOURCE=RECSOURCE||'deceased' WHERE custid IN (SELECT custid FROM tbl_upload_temp WHERE cust_class='835'); "
        ' ----------------------------------------------------------------------
        
        MsgBox "Data Telah Di Upload sebanyak : " + CStr(njmlNew) + "", vbOKOnly, "Information"
        Set list = lsthst.ListItems.ADD(, , MDIForm1.TxtUsername.text)
        list.SubItems(1) = MDIForm1.txtlevel.text
        list.SubItems(2) = Format(MDIForm1.TDBDate1.Value, "dd/mm/yyyy")
        list.SubItems(3) = Replace(txtlocation.text, "\", "/")
        list.SubItems(4) = Replace(cbosheet.text, "$", "")
        list.SubItems(5) = Val(txtrowssource.text)
        list.SubItems(6) = "Insert New Data"
        list.SubItems(7) = CStr(Val(njmlNew))
        
        
        
     strsql = "insert into tbl_hst_upload (userid,nama,location_file,Sheet,lead,eksekusi,jml_row,jenis_upload) values ("
     strsql = strsql + "'" + MDIForm1.TxtUsername.text + "','" + MDIForm1.txtlevel.text + "','" + Replace(txtlocation.text, "\", "/") + "',"
     strsql = strsql + "'" + Replace(Replace(cbosheet.text, "$", ""), "'", "") + "'," + CStr(Val(txtrowssource.text)) + ",'Insert New Data'," + CStr(Val(njmlNew)) + ",'MASTER')"
     M_OBJCONN.Execute (strsql)
    
    strsql = "insert into tbl_dontlist (name,addrnow,homeno,mobileno,addrpt,officeno,nocard,region,dob,recsource,custid,curbal,pay_dt,lastpay)"
    strsql = strsql + " select name,addrnow,homeno,mobileno,addrpt,officeno,nocard,region,dob,recsource,custid,curbal,pay_dt,lastpay from tbl_upload_temp where (coalesce(homeno,'')='' AND coalesce(mobileno,'')='' AND coalesce(officeno,'')='')"
    M_OBJCONN.Execute (strsql)

    End If
End If
End If
Set M_objrs = Nothing




'STRSQL = "SELECT DISTINCT(RECSOURCE), '" + TXTDESCRIPTION_BATCH + "' as batch  FROM tbl_upload_temp "


strsql = "SELECT DISTINCT(RECSOURCE), '" + TXTDESCRIPTION_BATCH + "' as batch  FROM tbl_upload_temp where recsource not in (select kodeds from datasourcetbl) "

Set OBJRECORD = New ADODB.Recordset
OBJRECORD.CursorLocation = adUseClient

OBJRECORD.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic


While Not OBJRECORD.EOF
strsql = " INSERT INTO DATASOURCETBL"
            strsql = strsql + " (KODEDS,"
            strsql = strsql + " STATUS,"
            strsql = strsql + " KETERANGAN,campaign_ket)"
            strsql = strsql + " VALUES"
            strsql = strsql + " ('" + UBAH_QUOTE(IIf(IsNull(OBJRECORD!recsource), "", OBJRECORD!recsource)) + "',"
            strsql = strsql + " '1',"
            strsql = strsql + " '" + UBAH_QUOTE(TXTDESCRIPTION_BATCH.text) + "',"
            strsql = strsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyymmdd")) + "')"
            M_OBJCONN.Execute strsql

          OBJRECORD.MoveNext
Wend


End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub CndUncekAll_Click()
    Dim W As Integer
    
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Data yang akan dihapus tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To ListView1.ListItems.Count
        ListView1.ListItems(W).Checked = False
    Next W
End Sub

Private Sub Command2_Click()
isi_dataSTATUS ""
End Sub

Private Sub Form_Load()
    create_header_mapping
    create_header_mapping_verify
    create_header_line_error
    create_header_hst_upload
  '  load_hst_upload
    
  '  isicombo_product
End Sub
Public Sub create_header_mapping_verify()
    lstmapping.ColumnHeaders.ADD 1, , "Source Field", 5 * TXT
    lstmapping.ColumnHeaders.ADD 2, , "Destination Field", 15 * TXT
    lstmapping.ColumnHeaders.ADD 3, , "Wrong Destination Field", 15 * TXT
    
    
     ListView1.ColumnHeaders.ADD 1, , "Custid", 5 * TXT
     ListView2.ColumnHeaders.ADD 1, , "Tanggal", 5 * TXT
     ListView2.ColumnHeaders.ADD 2, , "File Name in Excel", 5 * TXT
     ListView2.ColumnHeaders.ADD 3, , "Backup in Database", 5 * TXT
     ListView2.ColumnHeaders.ADD 4, , "Action user", 5 * TXT
    
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

    strsql = " select * from ( "
    strsql = strsql + " select nama_kolom,field_destination from "
    strsql = strsql + " (select * from ( "
    strsql = strsql + " SELECT column_name as nama_kolom From information_schema.Columns WHERE table_name='mgm'"
    strsql = strsql + " and data_type in ('character varying','numeric','bigint','integer','timestamp without time zone','text') and column_name not in (select nama_kolom  from tbl_tidak_map) ORDER BY ordinal_position) as tblbaru  full join"
    strsql = strsql + " (   select field_source,field_destination from tbl_setting_upload  where kode_source='" + cbomap.text + "')"
    strsql = strsql + " as tbldua on tblbaru.nama_kolom =tbldua.field_source) as tblbaru where (field_destination is not null or field_destination<>'') ) as tblsatu"
    strsql = strsql + " Left Join ( select * from tbl_temp_field  ) as tbldua   on tblsatu.field_destination=tbldua.nama_field"
    
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
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
sStrsql = "select * from tbl_hst_upload WHERE jenis_upload='MASTER'"
Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    no = 0
    lsthst.ListItems.CLEAR
    txtnumrowshst.text = M_objrs.RecordCount
    While Not M_objrs.EOF
        no = no + 1
        Set list = lsthst.ListItems.ADD(, , IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID))
            list.SubItems(1) = IIf(IsNull(M_objrs!nama), "", M_objrs!nama)
            list.SubItems(2) = Format(IIf(IsNull(M_objrs!tgl_upload), "", M_objrs!tgl_upload), "dd/mm/yyyy")
            list.SubItems(3) = IIf(IsNull(M_objrs!location_file), "", M_objrs!location_file)
            list.SubItems(4) = IIf(IsNull(M_objrs!sheet), "", M_objrs!sheet)
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
Dim m_objExisting As ADODB.Recordset
Dim M_objrs As New ADODB.Recordset
Dim M_objdouble As New ADODB.Recordset

Dim m_existing_new  As ADODB.Recordset

 On Error Resume Next
 M_OBJCONN.Execute " Drop TABLE Tbl_Upload_Temp "
 ssql = "CREATE TABLE Tbl_Upload_Temp "
 ssql = ssql & "(""ID"" serial)"
 M_OBJCONN.Execute (ssql)
 strsql = " select nama_kolom,field_destination,data_type,character_maximum_length from (  select * from (  SELECT column_name as nama_kolom From information_schema.Columns"
 strsql = strsql + " WHERE table_name='mgm' and table_schema='public' and data_type in ('character varying','numeric','bigint','integer','timestamp without time zone','text') and column_name not in (select nama_kolom  from tbl_tidak_map)   ORDER BY ordinal_position) as tblbaru  full join  (   select field_source,field_destination from tbl_setting_upload where  kode_source='" + cbomap.text + "' ) "
 strsql = strsql + "  as tbldua on tblbaru.nama_kolom =tbldua.field_source) as tblbaru"
 strsql = strsql + " Left Join"
 strsql = strsql + " (SELECT column_name,data_type ,character_maximum_length From information_schema.Columns WHERE table_name='mgm' and table_schema='public' ORDER BY ordinal_position) as tbltiga"
 strsql = strsql + " on tblbaru.nama_kolom=tbltiga.column_name"
 Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   ProgressBar1.Max = M_objrs.RecordCount + 1
   
    While Not M_objrs.EOF
   
    
    DoEvents
    ProgressBar1.Value = M_objrs.Bookmark
                 
            nama_kol = IIf(IsNull(M_objrs!nama_kolom), "", M_objrs!nama_kolom)
           
           
            
            data_type = IIf(IsNull(M_objrs!data_type), "", M_objrs!data_type)
            data_length = IIf(IsNull(M_objrs!character_maximum_length), "", M_objrs!character_maximum_length)
             If nama_kol = "limit" Then
               ' MsgBox "AADA"
           End If
           
            If Trim(data_type) = "character varying" Then
                If data_length = "" Then
                    If nama_kol = "limit" Then
                        Strtype = Chr(34) + nama_kol + chr34 + " " + data_type
                        sStrsql = " alter table Tbl_Upload_Temp  add column  " + Strtype
                    Else
                        Strtype = nama_kol + " " + data_type
                        sStrsql = " alter table Tbl_Upload_Temp  add column  " + Strtype
                    End If
                    
                Else
                       If nama_kol = "limit" Then
                            Strtype = Chr(34) + nama_kol + Chr(34) + " " + data_type + " (" + CStr(data_length) + ")"
                            sStrsql = " alter table Tbl_Upload_Temp  add column  " + Strtype
                        Else
                            Strtype = nama_kol + " " + data_type + " (" + CStr(data_length) + ")"
                            sStrsql = " alter table Tbl_Upload_Temp  add column  " + Strtype
                    End If
                    
                End If
                M_OBJCONN.Execute (sStrsql)
            ElseIf Trim(data_type) = "text" Then
                If data_length = "" Then
                    Strtype = nama_kol + " " + data_type
                    sStrsql = " alter table Tbl_Upload_Temp  add column  " + Strtype
                Else
                    Strtype = nama_kol + " " + data_type + " (" + CStr(data_length) + ")"
                    sStrsql = " alter table Tbl_Upload_Temp  add column  " + Strtype
                End If
                
                M_OBJCONN.Execute (sStrsql)
            ElseIf Trim(data_type) = "timestamp without time zone" Then
                Strtype = nama_kol + " " + data_type
                sStrsql = " alter table Tbl_Upload_Temp  add column  " + Strtype
                M_OBJCONN.Execute (sStrsql)
            ElseIf Trim(data_type) = "numeric" Then
                If data_length = "" Then
                    Strtype = nama_kol + " " + data_type
                     sStrsql = " alter table Tbl_Upload_Temp  add column  " + Strtype
                    Else
                       Strtype = nama_kol + " " + data_type + " (" + CStr(data_length) + ")"
                        sStrsql = " alter table Tbl_Upload_Temp  add column  " + Strtype
                End If
                M_OBJCONN.Execute (sStrsql)
            
             ElseIf Trim(data_type = "bigint") Then
                If data_length = "" Then
                    Strtype = nama_kol + " " + data_type
                     sStrsql = " alter table Tbl_Upload_Temp  add column  " + Strtype
                    Else
                       Strtype = nama_kol + " " + data_type + " (" + CStr(data_length) + ")"
                        sStrsql = " alter table Tbl_Upload_Temp  add column  " + Strtype
                End If
                M_OBJCONN.Execute (sStrsql)
            ElseIf Trim(data_type) = "integer" Then
                If data_length = "" Then
                    Strtype = nama_kol + " " + data_type
                     sStrsql = " alter table Tbl_Upload_Temp  add column  " + Strtype
                    Else
                       Strtype = nama_kol + " " + data_type + " (" + CStr(data_length) + ")"
                        sStrsql = " alter table Tbl_Upload_Temp  add column  " + Strtype
                End If
                M_OBJCONN.Execute (sStrsql)
            End If
        M_objrs.MoveNext
    Wend
    
    'sStrsql = " alter table Tbl_Upload_Temp  add column recsource character varying(255)"
    'M_OBJCONN.Execute (sStrsql)
    
    
    
    
    sStrsql = " alter table Tbl_Upload_Temp  add column f_flag numeric"
    M_OBJCONN.Execute (sStrsql)
    
    sStrsql = " alter table Tbl_Upload_Temp  add column f_flag_donot numeric"
    M_OBJCONN.Execute (sStrsql)
    
    
    Set M_objrs = Nothing
    
    
    ssql = "SELECT * FROM [" & cbosheet.text & "]   "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
    lsterror.ListItems.CLEAR
    ProgressBar1.Max = M_objrs.RecordCount + 1
    jml = 0
    txtrowssource.text = M_objrs.RecordCount
    txtlead.text = txtrowssource.text
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
                    sType = lstview.ListItems(jRow).SubItems(3)
                   
                    
                    nlimit = Val(lstview.ListItems(jRow).SubItems(2))
                    smapsource = lstview.ListItems(jRow).text
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
                        If lstview.ListItems(i).text = "limit" Or lstview.ListItems(i).text = "name" Or lstview.ListItems(i).text = "prior" Or lstview.ListItems(i).text = "cycle" Then
                            strfield = Chr(34) + lstview.ListItems(i).text + Chr(34)
                            Else
                            strfield = lstview.ListItems(i).text
                            End If
                            
                            If lstview.ListItems(i).SubItems(3) = "character varying" Or lstview.ListItems(i).SubItems(3) = "text" Then
                                If IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))) = True Then
                                    strvalues = "null"
                                Else
                                    strvalues = "'" + Replace(IIf(IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))), "", CStr(M_objrs.fields(lstview.ListItems(i).SubItems(1)))), "'", "") & "'"
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
                        If lstview.ListItems(i).text = "limit" Or lstview.ListItems(i).text = "name" Or lstview.ListItems(i).text = "prior" Or lstview.ListItems(i).text = "cycle" Then
                            strfield = strfield + "," + Chr(34) + lstview.ListItems(i).text + Chr(34)
                        Else
                        
                            strfield = strfield + "," + lstview.ListItems(i).text
                        End If
                        
                             If lstview.ListItems(i).SubItems(3) = "character varying" Or lstview.ListItems(i).SubItems(3) = "text" Then
                                If IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))) = True Then
                                    strvalues = strvalues + ",null"
                                Else
                                    strvalues = strvalues + ",'" + Replace(IIf(IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))), "", CStr(M_objrs.fields(lstview.ListItems(i).SubItems(1)))), "'", "") & "'"
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
                                        strvalues = strvalues + ",'" + cnull(IIf(IsNull(M_objrs.fields(lstview.ListItems(i).SubItems(1))), Null, Format(M_objrs.fields(lstview.ListItems(i).SubItems(1)), "yyyy/mm/dd"))) & "'"
                                 End If
    
                            End If
                            
                        
                        End If
                    End If
                Next i
                
                
                If strfield <> "" Then
                        ssqlhead = "INSERT INTO tbl_upload_temp (" + strfield + ") values ( " + strvalues + ")"
                        Debug.Print M_objrs.Bookmark
                        Debug.Print ssqlhead
                        M_OBJCONN.Execute (ssqlhead)
                End If
                
           End If
           
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
                  
           ssql = " update Tbl_Upload_Temp set name=replace(name,'''','') where name like '%''%' "
            M_OBJCONN.Execute (ssql)
           ssql = " update Tbl_Upload_Temp set name=replace(name,'.','') where name like '%.%'"
            M_OBJCONN.Execute (ssql)
             ssql = " update Tbl_Upload_Temp set name=replace(name,'*','') where name like '%*%'"
             M_OBJCONN.Execute (ssql)
                
            ssql = "update Tbl_Upload_Temp set name=replace(name,'(','') where name like '%(%'"
            M_OBJCONN.Execute (ssql)
           
            ssql = " update Tbl_Upload_Temp set name=replace(name,')','') where name like '%)%'"
             M_OBJCONN.Execute (ssql)
          
            ssql = "  update Tbl_Upload_Temp set name=replace(name,'/','') where name like '%/%'"
            M_OBJCONN.Execute (ssql)
          
            ssql = " update Tbl_Upload_Temp set name=replace(name,'+','') where name like '%+%';"
            M_OBJCONN.Execute (ssql)
            
            
'            ssql = " update Tbl_Upload_Temp set  homeno =substring(homeno,4,(length(homeno)-3)) where substring(homeno,1,3)='021'"
'            M_OBJCONN.Execute (ssql)
'            ssql = " update Tbl_Upload_Temp set officeno =substring(officeno,4,(length(officeno)-3)) where substring(officeno,1,3)='021'"
'            M_OBJCONN.Execute (ssql)
'           ssql = " update Tbl_Upload_Temp set mobileno =substring(mobileno,4,(length(mobileno)-3)) where substring(mobileno,1,3)='021'"
'            M_OBJCONN.Execute (ssql)
'
'           ssql = " update Tbl_Upload_Temp set mobileno2 =substring(mobileno2,4,(length(mobileno2)-3)) where substring(mobileno2,1,3)='021'; "
'           M_OBJCONN.Execute (ssql)
'           ssql = " update Tbl_Upload_Temp set mobilenoadd1 =substring(mobilenoadd1,4,(length(mobilenoadd1)-3)) where substring(mobilenoadd1,1,3)='021';"
'            M_OBJCONN.Execute (ssql)
'            ssql = "  update Tbl_Upload_Temp set mobilenoadd2 =substring(mobilenoadd2,4,(length(mobilenoadd2)-3)) where substring(mobilenoadd2,1,3)='021';"
'            M_OBJCONN.Execute (ssql)
            
            ssql = " UPDATE Tbl_Upload_Temp SET mobileno='0'||substr(trim(mobileno),3) WHERE mobileno like '62%' AND length(mobileno)>10"
            M_OBJCONN.Execute (ssql)

            'strsql = " update  tbl_upload_temp set  f_flag=2 where id  not in (select min(id) from tbl_upload_temp group by v_no_kartu,v_nama,d_dob) "
            'M_OBJCONN.Execute (strsql)
    
'            STRSQL = " update Tbl_Upload_Temp set f_flag_donot =1 where v_no_kartu in (select no_kartu from tbldonotcall ) "
'            M_OBJCONN.Execute (STRSQL)
            
    
    
    Dim cobadeh As ADODB.Recordset
'    Set cobadeh = New ADODB.Recordset
'    cobadeh.CursorLocation = adUseClient
'    cobadeh.Open "select * from mgm ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    MsgBox cobadeh.RecordCount
    
    
    'cek existing data
      strsql = " select mgm.custid as unik ,mgm.* from mgm,tbl_upload_temp "
      strsql = strsql + " where mgm.custid=tbl_upload_temp.custid and (tbl_upload_temp.f_flag=1 or tbl_upload_temp.f_flag is null "
      strsql = strsql + " or tbl_upload_temp.f_flag=0)"
      strsql = strsql + " and (f_flag_donot is null or f_flag_donot=0) "
'STRSQL = "select * from mgm"
         
      Set m_existing_new = New ADODB.Recordset
          m_existing_new.CursorLocation = adUseClient
          m_existing_new.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
          txtexisting.text = m_existing_new.RecordCount
          ListView1.ListItems.CLEAR
          Txtxadarow.text = m_existing_new.RecordCount
          
          If m_existing_new.RecordCount > 0 Then
            While Not m_existing_new.EOF
                Set list = ListView1.ListItems.ADD(, , IIf(IsNull(m_existing_new!Unik), "", m_existing_new!Unik))
                m_existing_new.MoveNext
            Wend
         End If
         
         Set m_existing_new = Nothing
          
    'cek data new
      strsql = "  select * from  tbl_upload_temp where ""ID"" not in"
      strsql = strsql + " ( "
      strsql = strsql + " select  tbl_upload_temp.""ID"" from mgm,tbl_upload_temp"
      strsql = strsql + " where mgm.custid=tbl_upload_temp.custid ) and (CUSTID IS NOT NULL OR CUSTID<>'') AND   "
      strsql = strsql + " (tbl_upload_temp.f_flag=1 or tbl_upload_temp.f_flag is null or tbl_upload_temp.f_flag=0  )   and  (f_flag_donot is null or f_flag_donot=0)"
      'strsql = strsql + " AND (coalesce(homeno,'')<>''or coalesce(mobileno,'')<>'' or coalesce(officeno,'')<>'')"
    Set m_objmasuk = New ADODB.Recordset
        m_objmasuk.CursorLocation = adUseClient
        m_objmasuk.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        txtlead_masuk.text = m_objmasuk.RecordCount
        Set DataGrid1.DATASOURCE = m_objmasuk
        
        txtnew.text = txtlead_masuk
    'cek jumlah data yang ke donot call
     'strsql = " select * from  tbl_upload_temp  where f_flag_donot =1 "
     strsql = "select name,addrnow,homeno,mobileno,addrpt,officeno,nocard,region,dob,recsource,custid,curbal,pay_dt,lastpay from tbl_upload_temp " 'where (coalesce(homeno,'')='' AND coalesce(mobileno,'')='' AND coalesce(officeno,'')='')"
    Set m_objdonot = New ADODB.Recordset
         m_objdonot.CursorLocation = adUseClient
         m_objdonot.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'         Set DataGrid3.DATASOURCE = m_objdonot
'         txtdonot.text = m_objdonot.RecordCount
         txtdonot.text = "0"
        strsql = " select * from  tbl_upload_temp  where f_flag=2 AND  (f_flag_donot IS NULL OR f_flag_donot=0)  "
     Set M_objdouble = New ADODB.Recordset
        M_objdouble.CursorLocation = adUseClient
        M_objdouble.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        'Set DataGrid4.DATASOURCE = M_objdouble
        txtdouble.text = M_objdouble.RecordCount
        'txtjmldoublee.text = txtdouble.text
'        TXTERROR.text = Val(txtrowssource.text) - (Val(txtnew.text) + Val(txtexisting.text) + Val(txtdonot.text) + Val(txtdouble.text))
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
Private Sub isi_dataSTATUS(strsql As String)
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
    CD_save.ShowSave
    TxtPath.text = CD_save.FileName
    
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
    objSheet.Cells(n, 1).Value = lsterror.ListItems(i).text
    objSheet.Cells(n, 2).ColumnWidth = 30
    objSheet.Cells(n, 2).Value = lsterror.ListItems(i).SubItems(1)
  Next i
  
    objBook.SaveAs TxtPath.text, xlWorkbookNormal
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
        'Syarat untuk upload, custid tidak boleh kosong, jika kosong maka akan keluar dari
        'Fungsi Cek Mandatory ini
        If lstview.ListItems(i).text = "custid" Then
            If lstview.ListItems(i).SubItems(1) = "" Then
                '@@ 16-02-2012 Tamabahan Jika tidak ada custid yang di mapping maka keluar
                cekMANDATORTY = False
                Exit Function
            End If
        End If
        'Syarat untuk upload, recsource tidak boleh kosong, jika kosong maka akan keluar dari
        'Fungsi Cek Mandatory ini
        If lstview.ListItems(i).text = "recsource" Then
            If lstview.ListItems(i).SubItems(1) = "" Then
                '@@ 16-02-2012 Tamabahan Jika tidak ada recsource yang di mapping maka keluar
                cekMANDATORTY = False
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
Private Sub txtbckup_DblClick()
   CD_save.ShowSave
    txtbckup.text = CD_save.FileName
    
End Sub
Private Sub isi_data(spath As String, ssql)
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
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
    objBook.SaveAs spath, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
Salah:
    Exit Sub
End Sub
Public Sub load_hst_upload_del()
Dim M_objrs   As New ADODB.Recordset
Dim list As ListItem
Dim no As Double
sStrsql = "select * from tbl_hst_upload_del"
Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    no = 0
    ListView2.ListItems.CLEAR
    txtrowhst_del.text = M_objrs.RecordCount
    While Not M_objrs.EOF
        DoEvents
        no = no + 1
        Set list = ListView2.ListItems.ADD(, , Format(IIf(IsNull(M_objrs!tgl_execute), "", M_objrs!tgl_execute), "dd/mm/yyyy hh:nn:ss"))
            list.SubItems(1) = IIf(IsNull(M_objrs!path_excel), "", M_objrs!path_excel)
            list.SubItems(2) = IIf(IsNull(M_objrs!path_didb), "", M_objrs!path_didb)
            list.SubItems(3) = IIf(IsNull(M_objrs!user_excecute), "", M_objrs!user_excecute)
        M_objrs.MoveNext
    Wend
Set M_objrs = Nothing
End Sub


