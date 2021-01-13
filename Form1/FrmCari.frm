VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form FrmCari 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Searching Data"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11430
   Icon            =   "FrmCari.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReff 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1305
      MaxLength       =   20
      TabIndex        =   0
      Top             =   360
      Width           =   2670
   End
   Begin VB.TextBox txtCust 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1305
      TabIndex        =   1
      Top             =   720
      Width           =   2625
   End
   Begin VB.TextBox txtKTP 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5025
      TabIndex        =   2
      Top             =   360
      Width           =   2745
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Debitur Info"
      TabPicture(0)   =   "FrmCari.frx":1272
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblTarget(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Address"
      TabPicture(1)   =   "FrmCari.frx":128E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "LblTarget(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Kredit"
      TabPicture(2)   =   "FrmCari.frx":12AA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Phone"
      TabPicture(3)   =   "FrmCari.frx":12C6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Jobs"
      TabPicture(4)   =   "FrmCari.frx":12E2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   10815
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   5880
            Width           =   3045
         End
         Begin MSComctlLib.ListView LstVwJobs 
            Height          =   5445
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   9604
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6135
         Left            =   -74760
         TabIndex        =   23
         Top             =   480
         Width           =   10695
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   5760
            Width           =   3045
         End
         Begin MSComctlLib.ListView LstVwPhone 
            Height          =   5445
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   9604
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   10815
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   5760
            Width           =   3045
         End
         Begin MSComctlLib.ListView LstVwKredit 
            Height          =   5205
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   9181
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H00000000&
         Height          =   6120
         Left            =   0
         TabIndex        =   11
         Top             =   480
         Width           =   10995
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   5760
            Width           =   3045
         End
         Begin MSComctlLib.ListView LstVwDebitur 
            Height          =   5415
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   10800
            _ExtentX        =   19050
            _ExtentY        =   9551
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000004&
         Height          =   6285
         Left            =   -74940
         TabIndex        =   8
         Top             =   405
         Width           =   10875
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   5880
            Width           =   3045
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   11910
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   7980
            Width           =   3045
         End
         Begin MSComctlLib.ListView LstVwAddress 
            Height          =   5325
            Left            =   90
            TabIndex        =   10
            Top             =   360
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   9393
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Label LblTarget 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   0
         Left            =   3315
         TabIndex        =   14
         Top             =   285
         Width           =   4605
      End
      Begin VB.Label LblTarget 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   1
         Left            =   -71820
         TabIndex        =   13
         Top             =   285
         Width           =   9465
      End
   End
   Begin TDBDate6Ctl.TDBDate TDBDOB 
      Height          =   315
      Index           =   0
      Left            =   5040
      TabIndex        =   3
      Top             =   720
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmCari.frx":12FE
      Caption         =   "FrmCari.frx":1416
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmCari.frx":1482
      Keys            =   "FrmCari.frx":14A0
      Spin            =   "FrmCari.frx":14FE
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
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
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin Threed.SSCommand cmdSearch 
      Height          =   360
      Left            =   8400
      TabIndex        =   4
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
      _Version        =   196610
      Font3D          =   5
      MousePointer    =   16
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Search"
      ButtonStyle     =   2
   End
   Begin Threed.SSCommand cmdClear 
      Height          =   360
      Left            =   9240
      TabIndex        =   5
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
      _Version        =   196610
      Font3D          =   5
      MousePointer    =   16
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Clear"
      ButtonStyle     =   2
   End
   Begin Threed.SSCommand cmdKeluar 
      Height          =   360
      Left            =   10080
      TabIndex        =   6
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
      _Version        =   196610
      Font3D          =   5
      MousePointer    =   16
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Keluar"
      ButtonStyle     =   2
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref No. "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   585
      TabIndex        =   20
      Top             =   375
      Width           =   630
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debitur Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KTP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   4545
      TabIndex        =   18
      Top             =   360
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   4545
      TabIndex        =   17
      Top             =   720
      Width           =   345
   End
End
Attribute VB_Name = "FrmCari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSTEMP As ADODB.Recordset
Dim i As Integer
Dim sWhere$
Dim sOperator$

Private Sub CmdClear_Click()
    txtreff.Text = ""
    txtCust.Text = ""
    txtKTP.Text = ""
    TDBDOB(0).Value = False
    txtreff.SetFocus
End Sub

Private Sub cmdkeluar_Click()
    Unload Me
End Sub

Private Sub cmdsearch_Click()
    i = 0
    sWhere = vbNullString
    If txtreff.Text <> vbNullString Then
        Call Condition(sOperator, i)
        sWhere = sOperator & _
                 "Reff_Num LIKE '%" & txtreff & "%' "
    End If
    If txtCust.Text <> vbNullString Then
        Call Condition(sOperator, i)
        sWhere = sOperator & _
                 "Debitur_Name LIKE '%" & txtCust & "%' "
    End If

    If txtKTP.Text <> vbNullString Then
        Call Condition(sOperator, i)
        sWhere = sWhere & sOperator & _
                 "KTP_NUM = '" & txtKTP.Text & "' "
    End If
   
    If TDBDOB(0).ValueIsNull <> True Then
            Call Condition(sOperator, i)
            sWhere = sWhere & sOperator & _
            "Born_Date = '" & Format(TDBDOB(0).Value, "yyyy/mm/dd") & "' "

    End If
    ssql = "SELECT * From mandiri.IDI_DEBITUR_INFO "
    If Len(sWhere) > 1 Then
        ssql = ssql & sWhere
    End If
    ssql = ssql & "ORDER BY Reff_Num "
        
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    
    LstVwDebitur.ListItems.CLEAR
    RSTEMP.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If RSTEMP.EOF = True And RSTEMP.BOF = True Then
        MsgBox "Data Tidak Ditemukan", vbOKOnly + vbCritical, "Pemberitahuan"
        Exit Sub
    End If
    
    RSTEMP.MoveFirst
    While Not RSTEMP.EOF
        Set list = LstVwDebitur.ListItems.ADD(, , RSTEMP.Bookmark)
        list.SubItems(1) = RSTEMP("DebiturInfo_Id")
        list.SubItems(2) = RSTEMP("Reff_Num")
        list.SubItems(3) = IIf(IsNull(RSTEMP("No_Laporan")), "", RSTEMP("No_Laporan"))
        list.SubItems(4) = IIf(IsNull(RSTEMP("Tgl_Laporan")), "", Format(RSTEMP("Tgl_Laporan"), "dd/mm/yyyy"))
        list.SubItems(5) = IIf(IsNull(RSTEMP("Idi_User")), "", RSTEMP("Idi_User"))
        list.SubItems(6) = IIf(IsNull(RSTEMP("Debitur_Name")), "", RSTEMP("Debitur_Name"))
        list.SubItems(7) = IIf(IsNull(RSTEMP("Born_Place")), "", RSTEMP("Born_Place"))
        list.SubItems(8) = IIf(IsNull(RSTEMP("Born_Place2")), "", RSTEMP("Born_Place2"))
        list.SubItems(9) = IIf(IsNull(RSTEMP("Born_Date")), "", Format(RSTEMP("Born_date"), "dd/mm/yyyy"))
        list.SubItems(10) = IIf(IsNull(RSTEMP("Din")), "", RSTEMP("Din"))
        list.SubItems(11) = IIf(IsNull(RSTEMP("NPWP")), "", RSTEMP("NPWP"))
        list.SubItems(12) = IIf(IsNull(RSTEMP("NPWP2")), "", RSTEMP("NPWP2"))
        list.SubItems(13) = IIf(IsNull(RSTEMP("KTP_NUM")), "", RSTEMP("KTP_NUM"))
        list.SubItems(14) = IIf(IsNull(RSTEMP("KTP_NUM2")), "", RSTEMP("KTP_NUM2"))
        list.SubItems(15) = IIf(IsNull(RSTEMP("Passport_Num")), "", RSTEMP("Passport_Num"))
        list.SubItems(16) = IIf(IsNull(RSTEMP("Passport_Num2")), "", RSTEMP("Passport_Num2"))
        list.SubItems(17) = IIf(IsNull(RSTEMP("HTML_Source")), "", RSTEMP("HTML_Source"))
        RSTEMP.MoveNext
    Wend
    Text1.Text = RSTEMP.RecordCount & " Customers"
    Set RSTEMP = Nothing

    Dim sSQL1$
    sSQL1 = "SELECT DebiturInfo_Id From mandiri.IDI_DEBITUR_INFO "
    If Len(sWhere) > 1 Then
        sSQL1 = sSQL1 & sWhere
    End If
    ShowLstVwAddress (sSQL1)
    
    Dim sSQL2$
    sSQL2 = "SELECT DebiturInfo_Id From mandiri.IDI_DEBITUR_INFO "
    If Len(sWhere) > 1 Then
        sSQL2 = sSQL2 & sWhere
    End If
    ShowLstVwKredit (sSQL2)
    
    Dim sSQL3$
    sSQL3 = "SELECT DebiturInfo_Id From IDI_DEBITUR_INFO "
    If Len(sWhere) > 1 Then
        sSQL3 = sSQL3 & sWhere
    End If
    ShowLstVwPhone (sSQL3)

    Dim sSQL4$
    sSQL4 = "SELECT DebiturInfo_Id From IDI_DEBITUR_INFO "
    If Len(sWhere) > 1 Then
        sSQL4 = sSQL4 & sWhere
    End If
    ShowLstVwJobs (sSQL4)
End Sub

Private Sub Form_Load()
    'tdbDob(0).Value = Date
    Call CreateLstVwDebitur
   '' Call ShowLstVwDebitur
    Call CreateLstVwAddress
''    Call ShowLstVwAddress
    Call CreateLstVwKredit
 ''   Call ShowLstVwKredit
    Call CreateLstVwPhone
  ''  Call ShowLstVwPhone
    Call CreateLstVwJobs
 ''   Call ShowLstVwJobs
End Sub

Private Sub CreateLstVwDebitur()
    LstVwDebitur.ColumnHeaders.ADD 1, , "No", 5 * 120
    LstVwDebitur.ColumnHeaders.ADD 2, , "DebiturInfo_Id", 20 * 120
    LstVwDebitur.ColumnHeaders.ADD 3, , "Reff_Num", 20 * 120
    LstVwDebitur.ColumnHeaders.ADD 4, , "No_Laporan", 20 * 120
    LstVwDebitur.ColumnHeaders.ADD 5, , "Tgl_Laporan", 15 * 120
    LstVwDebitur.ColumnHeaders.ADD 6, , "Idi_User", 25 * 120
    LstVwDebitur.ColumnHeaders.ADD 7, , "Debitur_Name", 25 * 120
    LstVwDebitur.ColumnHeaders.ADD 8, , "Born_Place", 25 * 120
    LstVwDebitur.ColumnHeaders.ADD 9, , "Born_Place2", 25 * 120
    LstVwDebitur.ColumnHeaders.ADD 10, , "Born_Date", 15 * 120
    LstVwDebitur.ColumnHeaders.ADD 11, , "Din", 25 * 120
    LstVwDebitur.ColumnHeaders.ADD 12, , "NPWP", 15 * 120
    LstVwDebitur.ColumnHeaders.ADD 13, , "NPWP2", 15 * 120
    LstVwDebitur.ColumnHeaders.ADD 14, , "KTP_NUM ID", 25 * 120
    LstVwDebitur.ColumnHeaders.ADD 15, , "KTP_NUM2", 25 * 120
    LstVwDebitur.ColumnHeaders.ADD 16, , "Passport_Num", 15 * 120
    LstVwDebitur.ColumnHeaders.ADD 17, , "Passport_Num2", 15 * 120
    LstVwDebitur.ColumnHeaders.ADD 18, , "HTML_SOURCE", 15 * 120
End Sub

Private Sub ShowLstVwDebitur()
Dim ssql As String
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    ssql = "SELECT * From IDI_DEBITUR_INFO"
    RSTEMP.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If RSTEMP.EOF = True And RSTEMP.BOF = True Then Exit Sub
    RSTEMP.MoveFirst
    While Not RSTEMP.EOF
        Set list = LstVwDebitur.ListItems.ADD(, , RSTEMP.Bookmark)
        list.SubItems(1) = IIf(IsNull(RSTEMP("DebiturInfo_Id")), "", RSTEMP("DebiturInfo_Id"))
        list.SubItems(2) = RSTEMP("Reff_Num")
        list.SubItems(3) = IIf(IsNull(RSTEMP("No_Laporan")), "", RSTEMP("No_Laporan"))
        list.SubItems(4) = IIf(IsNull(RSTEMP("Tgl_Laporan")), "", Format(RSTEMP("Tgl_Laporan"), "dd/mm/yyyy"))
        list.SubItems(5) = IIf(IsNull(RSTEMP("Idi_User")), "", RSTEMP("Idi_User"))
        list.SubItems(6) = IIf(IsNull(RSTEMP("Debitur_Name")), "", RSTEMP("Debitur_Name"))
        list.SubItems(7) = IIf(IsNull(RSTEMP("Born_Place")), "", RSTEMP("Born_Place"))
        list.SubItems(8) = IIf(IsNull(RSTEMP("Born_Place2")), "", RSTEMP("Born_Place2"))
        list.SubItems(9) = IIf(IsNull(RSTEMP("Born_Date")), "", Format(RSTEMP("Born_date"), "dd/mm/yyyy"))
        list.SubItems(10) = IIf(IsNull(RSTEMP("Din")), "", RSTEMP("Din"))
        list.SubItems(11) = IIf(IsNull(RSTEMP("NPWP")), "", RSTEMP("NPWP"))
        list.SubItems(12) = IIf(IsNull(RSTEMP("NPWP2")), "", RSTEMP("NPWP2"))
        list.SubItems(13) = IIf(IsNull(RSTEMP("KTP_NUM")), "", RSTEMP("KTP_NUM"))
        list.SubItems(14) = IIf(IsNull(RSTEMP("KTP_NUM2")), "", RSTEMP("KTP_NUM2"))
        list.SubItems(15) = IIf(IsNull(RSTEMP("Passport_Num")), "", RSTEMP("Passport_Num"))
        list.SubItems(16) = IIf(IsNull(RSTEMP("Passport_Num2")), "", RSTEMP("Passport_Num2"))
        list.SubItems(17) = IIf(IsNull(RSTEMP("HTML_Source")), "", RSTEMP("HTML_Source"))
        RSTEMP.MoveNext
    Wend
    Text1.Text = RSTEMP.RecordCount & " Customers"

    Set RSTEMP = Nothing
End Sub

Private Sub LstVwAddress_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LstVwAddress.SortKey = ColumnHeader.Index - 1
    LstVwAddress.Sorted = True
End Sub

Private Sub LstVwDebitur_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LstVwDebitur.SortKey = ColumnHeader.Index - 1
    LstVwDebitur.Sorted = True
End Sub

Public Function Condition(ByRef sOperator$, ByRef i As Integer)
    If i = 0 Then
        sOperator = "WHERE "
        i = 1
    ElseIf i = 1 Then
        sOperator = "AND "
        i = 1
    End If
End Function

''Address
Private Sub CreateLstVwAddress()
    LstVwAddress.ColumnHeaders.ADD 1, , "No", 5 * 120
    LstVwAddress.ColumnHeaders.ADD 2, , "IDIADDRESS_ID", 20 * 120
    LstVwAddress.ColumnHeaders.ADD 3, , "DEBITURINFO_ID", 20 * 120
    LstVwAddress.ColumnHeaders.ADD 4, , "ADDRESS_SEQ", 10 * 120
    LstVwAddress.ColumnHeaders.ADD 5, , "ALAMAT", 25 * 120
    LstVwAddress.ColumnHeaders.ADD 6, , "KELURAHAN", 25 * 120
    LstVwAddress.ColumnHeaders.ADD 7, , "KECAMATAN", 25 * 120
    LstVwAddress.ColumnHeaders.ADD 8, , "DATI2", 25 * 120
    LstVwAddress.ColumnHeaders.ADD 9, , "KODEPOS", 25 * 120
    LstVwAddress.ColumnHeaders.ADD 10, , "NEGARA", 15 * 120
    LstVwAddress.ColumnHeaders.ADD 11, , "UPDATE", 25 * 120
 End Sub

Private Sub ShowLstVwAddress(ByVal sWhere$)
Dim ssql As String
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    ssql = "SELECT * From IDI_ADDRESS "
    If Len(sWhere) > 1 Then ssql = ssql & "WHERE DEBITURINFO_ID IN (" & sWhere & ")"
    RSTEMP.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If RSTEMP.EOF = True And RSTEMP.BOF = True Then Exit Sub
    RSTEMP.MoveFirst
    While Not RSTEMP.EOF
        Set list = LstVwAddress.ListItems.ADD(, , RSTEMP.Bookmark)
        list.SubItems(1) = IIf(IsNull(RSTEMP("IDIADDRESS_ID")), "", RSTEMP("IDIADDRESS_ID"))
        list.SubItems(2) = IIf(IsNull(RSTEMP("DEBITURINFO_ID")), "", RSTEMP("DEBITURINFO_ID"))
        list.SubItems(3) = IIf(IsNull(RSTEMP("ADDRESS_SEQ")), "", RSTEMP("ADDRESS_SEQ"))
        list.SubItems(4) = IIf(IsNull(RSTEMP("ALAMAT")), "", RSTEMP("ALAMAT"))
        list.SubItems(5) = IIf(IsNull(RSTEMP("KELURAHAN")), "", RSTEMP("KELURAHAN"))
        list.SubItems(6) = IIf(IsNull(RSTEMP("KECAMATAN")), "", RSTEMP("KECAMATAN"))
        list.SubItems(7) = IIf(IsNull(RSTEMP("DATI2")), "", RSTEMP("DATI2"))
        list.SubItems(8) = IIf(IsNull(RSTEMP("KODEPOS")), "", RSTEMP("KODEPOS"))
        list.SubItems(9) = IIf(IsNull(RSTEMP("NEGARA")), "", RSTEMP("NEGARA"))
        list.SubItems(10) = IIf(IsNull(RSTEMP("UPDATE")), "", Format(RSTEMP("UPDATE"), "dd/mm/yyyy"))
        RSTEMP.MoveNext
    Wend
    Text3.Text = RSTEMP.RecordCount & " Customers"

    Set RSTEMP = Nothing
End Sub

''Kredit
Private Sub CreateLstVwKredit()
    LstVwKredit.ColumnHeaders.ADD 1, , "No", 5 * 120
    LstVwKredit.ColumnHeaders.ADD 2, , "IDIKREDIT_ID", 20 * 120
    LstVwKredit.ColumnHeaders.ADD 3, , "DEBITURINFO_ID", 20 * 120
    LstVwKredit.ColumnHeaders.ADD 4, , "KREDIT_SEQ", 20 * 120
    LstVwKredit.ColumnHeaders.ADD 5, , "PELAPOR", 15 * 120
    LstVwKredit.ColumnHeaders.ADD 6, , "SIFAT", 25 * 120
    LstVwKredit.ColumnHeaders.ADD 7, , "NO_REKENING", 25 * 120
    LstVwKredit.ColumnHeaders.ADD 8, , "REKENING_AKTIF", 25 * 120
    LstVwKredit.ColumnHeaders.ADD 9, , "VALUTA", 25 * 120
    LstVwKredit.ColumnHeaders.ADD 10, , "PERCENT_BUNGA", 15 * 120
    LstVwKredit.ColumnHeaders.ADD 11, , "PLAFON", 25 * 120
    LstVwKredit.ColumnHeaders.ADD 12, , "BAKI_DEBET", 15 * 120
    LstVwKredit.ColumnHeaders.ADD 13, , "TUNGGAKAN_POKOK", 15 * 120
    LstVwKredit.ColumnHeaders.ADD 14, , "FREK", 25 * 120
    LstVwKredit.ColumnHeaders.ADD 15, , "BUNGA_ON", 25 * 120
    LstVwKredit.ColumnHeaders.ADD 16, , "BUNGA_OF", 15 * 120
    LstVwKredit.ColumnHeaders.ADD 17, , "SEKTOR_EKONOMI", 15 * 120
    LstVwKredit.ColumnHeaders.ADD 18, , "JENIS_PENGGUNAAN", 15 * 120
    LstVwKredit.ColumnHeaders.ADD 12, , "KONDISI", 15 * 120
    LstVwKredit.ColumnHeaders.ADD 13, , "TGL_KONDISI", 15 * 120
    LstVwKredit.ColumnHeaders.ADD 14, , "SEBAB_MACET", 25 * 120
    LstVwKredit.ColumnHeaders.ADD 15, , "TGL_MACET", 25 * 120
    LstVwKredit.ColumnHeaders.ADD 16, , "AKAD_AWAL", 15 * 120
    LstVwKredit.ColumnHeaders.ADD 17, , "JATUH_TEMPO", 15 * 120
    LstVwKredit.ColumnHeaders.ADD 18, , "UPDATE", 15 * 120

End Sub

Private Sub ShowLstVwKredit(ByVal sWhere$)
Dim ssql As String
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    ssql = "SELECT * From IDI_KREDIT "
    If Len(sWhere) > 1 Then ssql = ssql & "WHERE DEBITURINFO_ID IN (" & sWhere & ")"
    RSTEMP.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If RSTEMP.EOF = True And RSTEMP.BOF = True Then Exit Sub
    RSTEMP.MoveFirst
    While Not RSTEMP.EOF
        Set list = LstVwKredit.ListItems.ADD(, , RSTEMP.Bookmark)
        list.SubItems(1) = IIf(IsNull(RSTEMP("DebiturInfo_Id")), "", RSTEMP("IDIKREDIT_ID"))
        list.SubItems(2) = IIf(IsNull(RSTEMP("DEBITURINFO_ID")), "", RSTEMP("DEBITURINFO_ID"))
        list.SubItems(3) = IIf(IsNull(RSTEMP("KREDIT_SEQ")), "", RSTEMP("KREDIT_SEQ"))
        list.SubItems(4) = IIf(IsNull(RSTEMP("PELAPOR")), "", RSTEMP("PELAPOR"))
        list.SubItems(5) = IIf(IsNull(RSTEMP("SIFAT")), "", RSTEMP("SIFAT"))
        list.SubItems(6) = IIf(IsNull(RSTEMP("NO_REKENING")), "", RSTEMP("NO_REKENING"))
        list.SubItems(7) = IIf(IsNull(RSTEMP("REKENING_AKTIF")), "", RSTEMP("REKENING_AKTIF"))
        list.SubItems(8) = IIf(IsNull(RSTEMP("VALUTA")), "", RSTEMP("VALUTA"))
        list.SubItems(9) = IIf(IsNull(RSTEMP("PERCENT_BUNGA")), "", RSTEMP("PERCENT_BUNGA"))
        list.SubItems(10) = IIf(IsNull(RSTEMP("PLAFON")), "", RSTEMP("PLAFON"))
        list.SubItems(11) = IIf(IsNull(RSTEMP("BAKI_DEBET")), "", RSTEMP("BAKI_DEBET"))
        list.SubItems(12) = IIf(IsNull(RSTEMP("TUNGGAKAN_POKOK")), "", RSTEMP("TUNGGAKAN_POKOK"))
        list.SubItems(13) = IIf(IsNull(RSTEMP("FREK")), "", RSTEMP("FREK"))
        list.SubItems(14) = IIf(IsNull(RSTEMP("BUNGA_ON")), "", RSTEMP("BUNGA_ON"))
        list.SubItems(15) = IIf(IsNull(RSTEMP("BUNGA_OFF")), "", RSTEMP("BUNGA_OFF"))
        list.SubItems(16) = IIf(IsNull(RSTEMP("SEKTOR_EKONOMI")), "", RSTEMP("SEKTOR_EKONOMI"))
        list.SubItems(17) = IIf(IsNull(RSTEMP("JENIS_PENGGUNAAN")), "", RSTEMP("JENIS_PENGGUNAAN"))
        list.SubItems(18) = IIf(IsNull(RSTEMP("KONDISI")), "", RSTEMP("KONDISI"))
        list.SubItems(19) = IIf(IsNull(RSTEMP("TGL_KONDISI")), "", Format(RSTEMP("TGL_KONDISI"), "dd/mm/yyyy"))
        list.SubItems(20) = IIf(IsNull(RSTEMP("SEBAB_MACET")), "", RSTEMP("SEBAB_MACET"))
        list.SubItems(21) = IIf(IsNull(RSTEMP("TGL_MACET")), "", Format(RSTEMP("TGL_MACET"), "dd/mm/yyyy"))
        list.SubItems(22) = IIf(IsNull(RSTEMP("AKAD_AWAL")), "", Format(RSTEMP("AKAD_AWAL"), "dd/mm/yyyy"))
        list.SubItems(23) = IIf(IsNull(RSTEMP("JATUH_TEMPO")), "", Format(RSTEMP("JATUH_TEMPO"), "dd/mm/yyyy"))
        list.SubItems(24) = IIf(IsNull(RSTEMP("UPDATE")), "", Format(RSTEMP("UPDATE"), "dd/mm/yyyy"))
        RSTEMP.MoveNext
    Wend
    Text4.Text = RSTEMP.RecordCount & " Customers"

    Set RSTEMP = Nothing
End Sub

Private Sub LstVwJobs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LstVwJobs.SortKey = ColumnHeader.Index - 1
    LstVwJobs.Sorted = True
End Sub

Private Sub LstVwKredit_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LstVwKredit.SortKey = ColumnHeader.Index - 1
    LstVwKredit.Sorted = True
End Sub

''Phone
Private Sub CreateLstVwPhone()
    LstVwPhone.ColumnHeaders.ADD 1, , "No", 5 * 120
    LstVwPhone.ColumnHeaders.ADD 2, , "IDIPHONE_ID", 20 * 120
    LstVwPhone.ColumnHeaders.ADD 3, , "DEBITURINFO_ID", 20 * 120
    LstVwPhone.ColumnHeaders.ADD 4, , "PHONE_SEQ", 10 * 120
    LstVwPhone.ColumnHeaders.ADD 5, , "PHONE_NUM", 15 * 120
    LstVwPhone.ColumnHeaders.ADD 6, , "UPDATE", 15 * 120
End Sub

Private Sub ShowLstVwPhone(ByVal sWhere$)
Dim ssql As String
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    ssql = "SELECT * From IDI_PHONES "
    If Len(sWhere) > 1 Then ssql = ssql & "WHERE DEBITURINFO_ID IN (" & sWhere & ")"
    RSTEMP.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If RSTEMP.EOF = True And RSTEMP.BOF = True Then Exit Sub
    RSTEMP.MoveFirst
    While Not RSTEMP.EOF
        Set list = LstVwPhone.ListItems.ADD(, , RSTEMP.Bookmark)
        list.SubItems(1) = IIf(IsNull(RSTEMP("IDIPHONE_ID")), "", RSTEMP("IDIPHONE_ID"))
        list.SubItems(2) = IIf(IsNull(RSTEMP("DEBITURINFO_ID")), "", RSTEMP("DEBITURINFO_ID"))
        list.SubItems(3) = IIf(IsNull(RSTEMP("PHONE_SEQ")), "", RSTEMP("PHONE_SEQ"))
        list.SubItems(4) = IIf(IsNull(RSTEMP("PHONE_NUM")), "", RSTEMP("PHONE_NUM"))
        list.SubItems(5) = IIf(IsNull(RSTEMP("UPDATE")), "", Format(RSTEMP("UPDATE"), "dd/mm/yyyy"))
        RSTEMP.MoveNext
    Wend
    Text5.Text = RSTEMP.RecordCount & " Customers"

    Set RSTEMP = Nothing
End Sub

Private Sub LstVwPhone_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LstVwPhone.SortKey = ColumnHeader.Index - 1
    LstVwPhone.Sorted = True
End Sub

''Jobs
Private Sub CreateLstVwJobs()
    LstVwJobs.ColumnHeaders.ADD 1, , "No", 5 * 120
    LstVwJobs.ColumnHeaders.ADD 2, , "IDIJOBS_ID", 20 * 120
    LstVwJobs.ColumnHeaders.ADD 3, , "DEBITURINFO_ID", 20 * 120
    LstVwJobs.ColumnHeaders.ADD 4, , "JOBS_SEQ", 10 * 120
    LstVwJobs.ColumnHeaders.ADD 5, , "PEKERJAAN", 15 * 120
    LstVwJobs.ColumnHeaders.ADD 6, , "TEMPAT_KERJA", 15 * 120
    LstVwJobs.ColumnHeaders.ADD 7, , "BIDANG_USAHA", 15 * 120
    LstVwJobs.ColumnHeaders.ADD 8, , "UPDATE", 15 * 120

End Sub

Private Sub ShowLstVwJobs(ByVal sWhere$)
Dim ssql As String
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    ssql = "SELECT * From IDI_JOBS "
    If Len(sWhere) > 1 Then ssql = ssql & "WHERE DEBITURINFO_ID IN (" & sWhere & ")"
    RSTEMP.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If RSTEMP.EOF = True And RSTEMP.BOF = True Then Exit Sub
    RSTEMP.MoveFirst
    While Not RSTEMP.EOF
        Set list = LstVwJobs.ListItems.ADD(, , RSTEMP.Bookmark)
        list.SubItems(1) = IIf(IsNull(RSTEMP("IDIJOBS_ID")), "", RSTEMP("IDIJOBS_ID"))
        list.SubItems(2) = IIf(IsNull(RSTEMP("DEBITURINFO_ID")), "", RSTEMP("DEBITURINFO_ID"))
        list.SubItems(3) = IIf(IsNull(RSTEMP("JOBS_SEQ")), "", RSTEMP("JOBS_SEQ"))
        list.SubItems(4) = IIf(IsNull(RSTEMP("PEKERJAAN")), "", RSTEMP("PEKERJAAN"))
        list.SubItems(5) = IIf(IsNull(RSTEMP("TEMPAT_KERJA")), "", RSTEMP("TEMPAT_KERJA"))
        list.SubItems(6) = IIf(IsNull(RSTEMP("BIDANG_USAHA")), "", RSTEMP("BIDANG_USAHA"))
        list.SubItems(7) = IIf(IsNull(RSTEMP("UPDATE")), "", Format(RSTEMP("UPDATE"), "dd/mm/yyyy"))
        RSTEMP.MoveNext
    Wend
    Text6.Text = RSTEMP.RecordCount & " Customers"
    Set RSTEMP = Nothing
End Sub

