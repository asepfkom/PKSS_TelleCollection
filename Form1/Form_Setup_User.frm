VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form_Setup_User 
   Caption         =   "Setup User"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   435
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7035
      Left            =   -15
      TabIndex        =   0
      Top             =   855
      Width           =   13950
      _ExtentX        =   24606
      _ExtentY        =   12409
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Searching       "
      TabPicture(0)   =   "Form_Setup_User.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Line1(0)"
      Tab(0).Control(1)=   "Line1(3)"
      Tab(0).Control(2)=   "Label1(1)"
      Tab(0).Control(3)=   "Image3(0)"
      Tab(0).Control(4)=   "ListViewData"
      Tab(0).Control(5)=   "cmdHapus"
      Tab(0).Control(6)=   "CmdSearchBaru(0)"
      Tab(0).Control(7)=   "cmdTambah"
      Tab(0).Control(8)=   "cmdEdit"
      Tab(0).Control(9)=   "cmdKeluar"
      Tab(0).Control(10)=   "cmbFieldCari"
      Tab(0).Control(11)=   "txtCari"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "History      "
      TabPicture(1)   =   "Form_Setup_User.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(2)"
      Tab(1).Control(1)=   "Line1(1)"
      Tab(1).Control(2)=   "Image3(1)"
      Tab(1).Control(3)=   "ListViewHst"
      Tab(1).Control(4)=   "txtjmlrow(1)"
      Tab(1).Control(5)=   "CmdSearchBaru(1)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "[Entry / Edit]     "
      TabPicture(2)   =   "Form_Setup_User.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Image3(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2(5)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label2(3)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label2(0)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label2(1)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label2(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Line1(2)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "CD1"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmbSeniorManager"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "TxtPathFoto"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Frame1(0)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtNama"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "cmbTeamLeader"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "CmdBatal"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "cmdOk"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "CmdInputFoto"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Frame2"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txtUserId"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txtId"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).ControlCount=   20
      Begin VB.TextBox txtId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   7500
         MaxLength       =   20
         TabIndex        =   34
         Top             =   570
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.TextBox txtUserId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   4035
         MaxLength       =   20
         TabIndex        =   33
         Top             =   585
         Width           =   2010
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   4035
         TabIndex        =   30
         Top             =   1665
         Width           =   4185
         Begin VB.OptionButton optAktive 
            BackColor       =   &H80000004&
            Caption         =   "Active"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   240
            TabIndex        =   32
            Top             =   270
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton optNonAktive 
            BackColor       =   &H80000004&
            Caption         =   "Non Active"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1530
            TabIndex        =   31
            Top             =   270
            Width           =   1290
         End
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
         Index           =   0
         Left            =   -62985
         MaxLength       =   20
         TabIndex        =   18
         Top             =   6660
         Width           =   1785
      End
      Begin VB.CommandButton CmdInputFoto 
         BackColor       =   &H00F1E5DB&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         Picture         =   "Form_Setup_User.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2295
         Width           =   1545
      End
      Begin VB.CommandButton cmdOk 
         Height          =   375
         Left            =   10440
         Picture         =   "Form_Setup_User.frx":0666
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6570
         Width           =   1575
      End
      Begin VB.CommandButton CmdBatal 
         Height          =   375
         Left            =   12120
         Picture         =   "Form_Setup_User.frx":0C8C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6570
         Width           =   1575
      End
      Begin VB.ComboBox cmbTeamLeader 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4035
         TabIndex        =   14
         Top             =   1320
         Width           =   4365
      End
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   4035
         MaxLength       =   50
         TabIndex        =   13
         Top             =   945
         Width           =   4320
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   0
         Left            =   45
         TabIndex        =   12
         Top             =   450
         Width           =   1575
         Begin VB.Image Image_Agent 
            Height          =   1455
            Index           =   0
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton CmdSearchBaru 
         Height          =   360
         Index           =   1
         Left            =   -74955
         Picture         =   "Form_Setup_User.frx":12D2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   405
         Width           =   1515
      End
      Begin VB.TextBox txtCari 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -70905
         TabIndex        =   10
         Top             =   450
         Width           =   2760
      End
      Begin VB.ComboBox cmbFieldCari 
         Height          =   315
         Left            =   -73320
         TabIndex        =   9
         Top             =   450
         Width           =   2400
      End
      Begin VB.CommandButton cmdKeluar 
         BackColor       =   &H00F1E5DB&
         Cancel          =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -62715
         Picture         =   "Form_Setup_User.frx":18C0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   375
         Width           =   1550
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00F1E5DB&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -66450
         Picture         =   "Form_Setup_User.frx":1F06
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   405
         Width           =   1590
      End
      Begin VB.CommandButton cmdTambah 
         BackColor       =   &H00F1E5DB&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68070
         Picture         =   "Form_Setup_User.frx":24FA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   405
         UseMaskColor    =   -1  'True
         Width           =   1550
      End
      Begin VB.CommandButton CmdSearchBaru 
         Height          =   360
         Index           =   0
         Left            =   -74910
         Picture         =   "Form_Setup_User.frx":2B8E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   405
         Width           =   1515
      End
      Begin VB.CommandButton cmdHapus 
         BackColor       =   &H00F1E5DB&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -64815
         Picture         =   "Form_Setup_User.frx":317C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   405
         Width           =   1550
      End
      Begin VB.TextBox TxtPathFoto 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1215
         TabIndex        =   3
         Top             =   6570
         Width           =   9105
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
         Left            =   -63210
         MaxLength       =   20
         TabIndex        =   2
         Top             =   6570
         Width           =   1785
      End
      Begin VB.ComboBox cmbSeniorManager 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Form_Setup_User.frx":375B
         Left            =   4095
         List            =   "Form_Setup_User.frx":375D
         TabIndex        =   1
         Top             =   7155
         Visible         =   0   'False
         Width           =   4395
      End
      Begin MSComctlLib.ListView ListViewData 
         Height          =   5505
         Left            =   -74910
         TabIndex        =   19
         Top             =   990
         Width           =   13830
         _ExtentX        =   24395
         _ExtentY        =   9710
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
      Begin MSComctlLib.ListView ListViewHst 
         Height          =   5550
         Left            =   -75000
         TabIndex        =   20
         Top             =   1005
         Width           =   13785
         _ExtentX        =   24315
         _ExtentY        =   9790
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
      Begin MSComDlg.CommonDialog CD1 
         Left            =   1890
         Top             =   3810
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.jpg"
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   1
         Left            =   -75000
         Top             =   345
         Width           =   26295
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   0
         Left            =   -75090
         Picture         =   "Form_Setup_User.frx":375F
         Top             =   915
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
         Index           =   1
         Left            =   -63795
         TabIndex        =   28
         Top             =   6705
         Width           =   810
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         Index           =   3
         X1              =   -75000
         X2              =   -61185
         Y1              =   6570
         Y2              =   6570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   13815
         Y1              =   6435
         Y2              =   6435
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Status "
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
         Left            =   1800
         TabIndex        =   27
         Top             =   1905
         Width           =   1350
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Supervisor"
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
         Index           =   1
         Left            =   1800
         TabIndex        =   26
         Top             =   1350
         Width           =   1350
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Officer Name"
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
         Left            =   1800
         TabIndex        =   25
         Top             =   990
         Width           =   1350
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Officer ID/ NPK"
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
         Left            =   1800
         TabIndex        =   24
         Top             =   630
         Width           =   1980
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -75000
         X2              =   -61185
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         Index           =   0
         X1              =   -75000
         X2              =   -61185
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Path Foto :"
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
         Index           =   3
         Left            =   135
         TabIndex        =   23
         Top             =   6615
         Width           =   1350
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
         Left            =   -64020
         TabIndex        =   22
         Top             =   6615
         Width           =   810
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Senior Manager"
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
         Height          =   240
         Index           =   5
         Left            =   1845
         TabIndex        =   21
         Top             =   7185
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   2
         Left            =   -2970
         Picture         =   "Form_Setup_User.frx":AD69
         Top             =   330
         Width           =   26295
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup User"
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
      Height          =   225
      Index           =   0
      Left            =   750
      TabIndex        =   29
      Top             =   300
      Width           =   3585
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   300
      Picture         =   "Form_Setup_User.frx":12373
      Stretch         =   -1  'True
      Top             =   240
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   -1980
      Picture         =   "Form_Setup_User.frx":12E7D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19980
   End
End
Attribute VB_Name = "Form_Setup_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sKdlevel As String
Public sUsertype As String
Private Sub cmbFieldCari_DropDown()
    Dim rs As New ADODB.Recordset
    Dim strsql, mwhere As String
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    strsql = "select column_name from information_schema.columns where table_name = 'usertbl' order by ordinal_position"

    rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    cmbFieldCari.CLEAR
    While Not rs.EOF
        cmbFieldCari.AddItem cnull(rs(0))
        rs.MoveNext
    Wend
    
    Set rs = Nothing

End Sub

Private Sub cmbSeniorManager_DropDown()
Call loadNamaByUserType("3", cmbSeniorManager)
End Sub

Private Sub cmbTeamLeader_DropDown()
Call loadNamaByUserType("2", cmbTeamLeader)
End Sub

Private Sub CmdBatal_Click()
Unload Me

End Sub

Private Sub cmdEdit_Click()
    SSTab1.Tab = 2
    If ListViewData.SelectedItem.Text <> Empty Then
        Call load_edit
    End If
    
    
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo ERRORA
    If ListViewData.ListItems.Count <> 0 Then
        If MsgBox("Are You sure to delete user " + ListViewData.SelectedItem.SubItems(2), vbQuestion + vbYesNo, App.Title) = vbYes Then
            CMDSQL = "Insert Into TblLogUserAdm ( UserId, Keterangan, UserType,operation) VALUES ( '" + MDIForm1.TxtUsername.Text + "','Delete user " + ListViewData.SelectedItem.SubItems(1) + " - " + ListViewData.SelectedItem.SubItems(2) + "','" + sUsertype + "','Delete') "
            M_OBJCONN.Execute CMDSQL
            
            strsql = "delete from USERTBL where userid = '" + ListViewData.SelectedItem.SubItems(1) + "'"
            M_OBJCONN.Execute strsql
            MsgBox "Done", vbInformation + vbOKOnly, "TINS"
            load_data_user
        End If
    End If
    
    Exit Sub
ERRORA:
    MsgBox "Can't deleted from database ", vbInformation + vbOKOnly, App.Title
End Sub

Private Sub CmdInputFoto_Click()
 CD1.Action = 1
    'Mengisi txtpathphoto dengan path(lokasi file) yang diambil dari kotak dialog open file
    TxtPathFoto.Text = CD1.FileName
    'Mengisi image sesuai dengan lokasi file yang dipilih
    Image_Agent(0).Picture = LoadPicture(TxtPathFoto.Text)

End Sub

Private Sub cmdkeluar_Click()
Unload Me
End Sub
Private Sub Add_Edit_Officer()
    Dim rs As New ADODB.Recordset
    Dim strsql, mwhere As String
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    strsql = "select * from usertbl"

    If useridAlreadyExist(txtuserid.Text) = True Then
        If MsgBox("Userid already Exist, Are you sure to replace data?", vbQuestion + vbYesNo, "TINS") = vbNo Then
            Exit Sub
        End If
        rs.Open "select * from usertbl where userid = '" + txtuserid.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
    Else
        rs.Open "select * from usertbl limit 1", M_OBJCONN, adOpenDynamic, adLockOptimistic
        rs.AddNew
    
    End If
    
    If rs.EOF Then
        MsgBox "Agent Tidak Ditemukan !!", vbCritical + vbOKOnly, "TINS"
        clearDataTexbox
        load_data_user
        SSTab1.Tab = 0
        Exit Sub
    End If
    
    
    rs!AGENT = txtnama.Text
    rs!USERID = txtuserid.Text
    rs!usertype = sUsertype
    rs!TEAM = cmbTeamLeader.Text
    rs!SPVCODE = cmbTeamLeader.Text
    rs!aktif = IIf(optAktive.Value = True, 1, 0)
    rs!ACCREC = Encrypt(Len(txtuserid.Text), "PASS12345")
    'Rs!AM = cmbSeniorManager.Text
    rs!adminserver = MDIForm1.TxtUsername.Text
    rs!level_name = getLevelName
    rs!kdlevel = sKdlevel
    rs!tglexpired = Format(getdateExpired, "yyyy-mm-dd")
    
    rs.update
    
    Set rs = Nothing
    
    
    'Log Add
    If txtuserid.Text <> "" Then
        CMDSQL = "Insert Into TblLogUserAdm ( UserId, Keterangan, UserType,Operation) VALUES ( '" + MDIForm1.TxtUsername.Text + "','Add New Agent','" + sUsertype + "','Create') "
    Else
        CMDSQL = "Insert Into TblLogUserAdm ( UserId, Keterangan, UserType,operation) VALUES ( '" + MDIForm1.TxtUsername.Text + "','Update Agent','" + sUsertype + "','Update') "
    End If
    M_OBJCONN.Execute CMDSQL
    MsgBox "Done", vbInformation + vbOKOnly, "TINS"
    
    Call clearDataTexbox
    Call load_data_user
    SSTab1.Tab = 0
    Set rs = Nothing

End Sub
Private Function getdateExpired() As Date
    Dim rs As New ADODB.Recordset
    Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        sStrsql = " select current_date + cast('30 day' as interval) as tglexpired "
        rs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
          getdateExpired = cnull(rs!tglexpired)
        End If
    Set rs = Nothing
End Function

Private Function getLevelName() As String
    Dim rs As New ADODB.Recordset
    Dim strsql, mwhere As String
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    strsql = "select * from tbllevel where tbllevel_kdlevel = '" + sKdlevel + "'"
    rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

    If Not rs.EOF Then
        getLevelName = cnull(rs!level_name)
    End If
       
    Set rs = Nothing

End Function
Private Sub loadNamaByUserType(sUsertypeCode As String, sCombo As ComboBox)
    Dim rs As New ADODB.Recordset
    Dim strsql, mwhere As String
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
        strsql = "select distinct(userid) as userid from usertbl where usertype = '" + sUsertypeCode + "'"
    rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

    sCombo.CLEAR
    While Not rs.EOF
        sCombo.AddItem cnull(rs!USERID)
        rs.MoveNext
    Wend
       
    Set rs = Nothing

End Sub
Private Sub cmdOK_Click()
    Call Add_Edit_Officer
End Sub

Private Sub CmdSearchBaru_Click(Index As Integer)
    Call load_data_history
End Sub
Public Sub load_data_history()
    Dim list As ListItem
    Dim rs As New ADODB.Recordset
    Dim strsql, mwhere As String
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    strsql = " select * from TblLogUserAdm where operation in ('Create','Update','Delete') and usertype='" + sUsertype + "' order by idlog desc"
    
    rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    ListViewHst.ListItems.CLEAR
    While Not rs.EOF
        sUserid = IIf(cnull(rs!USERID) = "elin", "ADMINISTRATOR", cnull(rs!USERID))
        Set list = ListViewHst.ListItems.ADD(, , sUserid)
            list.SubItems(1) = cnull(rs!keterangan)
            list.SubItems(2) = cnull(rs!operation)
            list.SubItems(3) = Format(cnull(rs!TGL), "YYYY-MM-DD")
        rs.MoveNext
    Wend
    txtjmlrow(1).Text = rs.RecordCount
    Set rs = Nothing
End Sub
Private Function useridAlreadyExist(kdUserid As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim strsql, mwhere As String
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    strsql = "select * from usertbl where userid = '" + kdUserid + "'"
    
    rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        useridAlreadyExist = True
    Else
        useridAlreadyExist = False
    End If

    Set rs = Nothing


End Function
Private Sub cmdTambah_Click()
    SSTab1.Tab = 2
    Call clearDataTexbox
End Sub


Private Sub Form_Load()
    Call create_header_manager
    Call load_data_user
    SSTab1.Tab = 0
    If sKdlevel = "2" Or sKdlevel = "5" Then
        Label2(1).Visible = False
        cmbTeamLeader.Visible = False
        'Label2(5).Top = 1350
        'cmbSeniorManager.Top = 1320
        Label2(2).Top = 1350
        Frame2.Top = 1320
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set clsuser = Nothing
End Sub
Public Sub clearDataTexbox()
    txtId.Text = Empty
    txtnama.Text = Empty
    txtuserid.Text = Empty
    cmbTeamLeader.Text = Empty
    cmbSeniorManager.Text = Empty
    optAktive.Value = True
    optNonAktive.Value = False
End Sub
Public Sub create_header_manager()
    With ListViewData
            .ColumnHeaders.ADD , , "id", 0 * TXT
            .ColumnHeaders.ADD , , "Userid", 10 * TXT
            .ColumnHeaders.ADD , , "Nama", 10 * TXT
            .ColumnHeaders.ADD , , "Status", 10 * TXT
            If sUsertype = "5" Or sUsertype = "2" Then
                .ColumnHeaders.ADD , , "TeamLeader", 0 * TXT
            Else
                .ColumnHeaders.ADD , , "TeamLeader", 10 * TXT
            End If
            '.ColumnHeaders.ADD , , "RSM", 0 * TXT
    End With
    
    
    With ListViewHst
            .ColumnHeaders.ADD , , "Userid", 10 * TXT
            .ColumnHeaders.ADD , , "Keterangan", 10 * TXT
            .ColumnHeaders.ADD , , "Operation", 10 * TXT
            .ColumnHeaders.ADD , , "Tanggal", 10 * TXT
    End With
End Sub
Public Sub load_data_user()
    Dim list As ListItem
    Dim rs As New ADODB.Recordset
    Dim strsql, mwhere As String
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    strsql = "select * from usertbl"
    mwhere = " where usertype = '" + sUsertype + "'"
    
    If cmbFieldCari.Text <> Empty Then
        mwhere = mwhere + " and " + cmbFieldCari.Text + " ilike  '%" + txtCari.Text + "%'"
    End If
    
    rs.Open strsql + mwhere + " order by agent", M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    ListViewData.ListItems.CLEAR
    While Not rs.EOF
        Set list = ListViewData.ListItems.ADD(, , cnull(rs!ID))
            list.SubItems(1) = cnull(rs!USERID)
            list.SubItems(2) = cnull(rs!AGENT)
            If cnull(rs!aktif) = "1" Then
                list.SubItems(3) = "Aktif"
            Else
                list.SubItems(3) = "Tidak Aktif"
            End If
            list.SubItems(4) = cnull(rs!SPVCODE)
            'list.SubItems(5) = cnull(Rs!AM)
          rs.MoveNext
    Wend
    
    Set rs = Nothing
    
End Sub

Private Sub ListViewData_DblClick()
    Call load_edit
End Sub

Private Sub load_edit()
    Dim rs As New ADODB.Recordset
    Dim strsql, mwhere, sUserid As String
    

    If ListViewData.ListItems.Count <> 0 Then
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        strsql = "select * from usertbl"
        mwhere = " where userid = '" + ListViewData.SelectedItem.SubItems(1) + "'"
        
        rs.Open strsql + mwhere + " order by agent", M_OBJCONN, adOpenDynamic, adLockOptimistic
    
        If rs.RecordCount <> 0 Then
            txtuserid.Text = cnull(rs!USERID)
            txtnama.Text = cnull(rs!AGENT)
            
            If cnull(rs!aktif) = "1" Then
                optAktive.Value = True
            Else
                optNonAktive.Value = True
            End If
            
            cmbTeamLeader.Text = cnull(rs!SPVCODE)
            'cmbSeniorManager.Text = cnull(Rs!AM)
            txtId.Text = cnull(rs!USERID)
            
            SSTab1.Tab = 2
            On Error Resume Next
            sStrsql = "select lo_export(tbluser_foto,'d:/foto_agent/"
            sStrsql = sStrsql + sUserid + ".jpg') "
            sStrsql = sStrsql + " from usertbl where userid='" + sUserid + "'"
            M_OBJCONN.Execute sStrsql
            Image_Agent(0).Picture = LoadPicture("d:\foto_agent\" + sUserid + ".jpg")
        End If
    End If
End Sub
