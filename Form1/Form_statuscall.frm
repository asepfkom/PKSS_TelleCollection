VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_statuscall 
   Caption         =   "Master Status Call"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   14010
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7035
      Left            =   15
      TabIndex        =   0
      Top             =   855
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   12409
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Searching       "
      TabPicture(0)   =   "Form_statuscall.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image3(0)"
      Tab(0).Control(1)=   "Line1(0)"
      Tab(0).Control(2)=   "Line1(3)"
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(4)=   "ListView1(0)"
      Tab(0).Control(5)=   "CmdSearchBaru(0)"
      Tab(0).Control(6)=   "cmdTambah"
      Tab(0).Control(7)=   "cmdEdit"
      Tab(0).Control(8)=   "cmdKeluar"
      Tab(0).Control(9)=   "Combo1(0)"
      Tab(0).Control(10)=   "Text1(0)"
      Tab(0).Control(11)=   "txtjmlrow(0)"
      Tab(0).Control(12)=   "cmdHapus"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "History      "
      TabPicture(1)   =   "Form_statuscall.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtjmlrow(1)"
      Tab(1).Control(1)=   "CmdSearchBaru(1)"
      Tab(1).Control(2)=   "Combo1(1)"
      Tab(1).Control(3)=   "Text1(1)"
      Tab(1).Control(4)=   "ListView1(1)"
      Tab(1).Control(5)=   "Label1(2)"
      Tab(1).Control(6)=   "Line1(4)"
      Tab(1).Control(7)=   "Line1(1)"
      Tab(1).Control(8)=   "Image3(1)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "[Entry / Edit]     "
      TabPicture(2)   =   "Form_statuscall.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Image3(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label2(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label2(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Line1(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label1(5)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label1(3)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtUserId"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtNmAgent"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Option1(1)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Option1(0)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "CmdBatal"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cmdOk"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "cmbgroupcall2"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "cmbgroupcall1"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).ControlCount=   15
      Begin VB.ComboBox cmbgroupcall1 
         Height          =   315
         Left            =   1905
         TabIndex        =   29
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox cmbgroupcall2 
         Height          =   315
         Left            =   2280
         TabIndex        =   28
         Top             =   7185
         Visible         =   0   'False
         Width           =   2655
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
         Left            =   -64425
         Picture         =   "Form_statuscall.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   450
         Width           =   1550
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
         Left            =   -62940
         MaxLength       =   20
         TabIndex        =   25
         Top             =   6615
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
         Index           =   0
         Left            =   -62985
         MaxLength       =   20
         TabIndex        =   16
         Top             =   6660
         Width           =   1785
      End
      Begin VB.CommandButton cmdOk 
         Height          =   375
         Left            =   10440
         Picture         =   "Form_statuscall.frx":0633
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6570
         Width           =   1575
      End
      Begin VB.CommandButton CmdBatal 
         Height          =   375
         Left            =   12120
         Picture         =   "Form_statuscall.frx":0C59
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6570
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
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
         Index           =   0
         Left            =   1905
         TabIndex        =   13
         Top             =   2025
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton Option1 
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
         Index           =   1
         Left            =   3200
         TabIndex        =   12
         Top             =   2025
         Width           =   1290
      End
      Begin VB.TextBox txtNmAgent 
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
         Left            =   1905
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1065
         Width           =   4320
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
         Left            =   1905
         MaxLength       =   20
         TabIndex        =   10
         Top             =   585
         Width           =   2010
      End
      Begin VB.CommandButton CmdSearchBaru 
         Height          =   360
         Index           =   1
         Left            =   -74955
         Picture         =   "Form_statuscall.frx":129F
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   405
         Width           =   1515
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   -73335
         TabIndex        =   8
         Top             =   450
         Width           =   3030
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Index           =   1
         Left            =   -70275
         TabIndex        =   7
         Top             =   450
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Index           =   0
         Left            =   -70905
         TabIndex        =   6
         Top             =   450
         Width           =   3120
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   -73335
         TabIndex        =   5
         Top             =   450
         Width           =   2355
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
         Picture         =   "Form_statuscall.frx":188D
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   450
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
         Left            =   -66180
         Picture         =   "Form_statuscall.frx":1ED3
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   450
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
         Left            =   -67800
         Picture         =   "Form_statuscall.frx":24C7
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         UseMaskColor    =   -1  'True
         Width           =   1550
      End
      Begin VB.CommandButton CmdSearchBaru 
         Height          =   360
         Index           =   0
         Left            =   -74955
         Picture         =   "Form_statuscall.frx":2B5B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   405
         Width           =   1515
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5505
         Index           =   0
         Left            =   -75000
         TabIndex        =   17
         Top             =   990
         Width           =   13830
         _ExtentX        =   24395
         _ExtentY        =   9710
         View            =   3
         LabelEdit       =   1
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   5505
         Index           =   1
         Left            =   -75000
         TabIndex        =   18
         Top             =   945
         Width           =   13920
         _ExtentX        =   24553
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
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Call "
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
         TabIndex        =   31
         Top             =   1560
         Width           =   1980
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Call 2"
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
         Index           =   5
         Left            =   510
         TabIndex        =   30
         Top             =   7185
         Visible         =   0   'False
         Width           =   1980
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
         Left            =   -63750
         TabIndex        =   26
         Top             =   6660
         Width           =   810
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         Index           =   4
         X1              =   -74955
         X2              =   -61140
         Y1              =   6525
         Y2              =   6525
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
         TabIndex        =   22
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
         Left            =   135
         TabIndex        =   21
         Top             =   2025
         Width           =   1350
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Status call name"
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
         Left            =   135
         TabIndex        =   20
         Top             =   1110
         Width           =   1350
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Key status call"
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
         Left            =   135
         TabIndex        =   19
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
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   0
         Left            =   -75090
         Picture         =   "Form_statuscall.frx":3149
         Top             =   315
         Width           =   26295
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   2
         Left            =   0
         Picture         =   "Form_statuscall.frx":A753
         Top             =   315
         Width           =   26295
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   1
         Left            =   -75000
         Picture         =   "Form_statuscall.frx":11D5D
         Top             =   315
         Width           =   26295
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   -15
      Picture         =   "Form_statuscall.frx":19367
      Stretch         =   -1  'True
      Top             =   240
      Width           =   420
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Master Status Call"
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
      Index           =   0
      Left            =   435
      TabIndex        =   24
      Top             =   300
      Width           =   3585
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   -1980
      Picture         =   "Form_statuscall.frx":19E71
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19980
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   1
      Left            =   3975
      Picture         =   "Form_statuscall.frx":1F2DC
      Stretch         =   -1  'True
      Top             =   240
      Width           =   420
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Master Officer"
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
      Left            =   4470
      TabIndex        =   23
      Top             =   300
      Width           =   3585
   End
End
Attribute VB_Name = "Form_statuscall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Clsstatuscall As Clsstatuscall
Private Sub cmbgroupcall1_DropDown()
Dim M_objrs As ADODB.Recordset
        Dim CMDSQL As String
    
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        CMDSQL = " SELECT distinct grp_call from tblstatuscall where grp_call not in ('')"
           M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    cmbgroupcall1.CLEAR
    While Not M_objrs.EOF
        cmbgroupcall1.AddItem IIf(IsNull(M_objrs!grp_call), "", M_objrs!grp_call)
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing

End Sub
Private Sub cmbgroupcall2_DropDown()
Dim M_objrs As ADODB.Recordset
        Dim CMDSQL As String
    
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        CMDSQL = " SELECT distinct grp_call2 from tblstatuscall "
           M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    cmbgroupcall2.CLEAR
    While Not M_objrs.EOF
        cmbgroupcall2.AddItem IIf(IsNull(M_objrs!grp_call2), "", M_objrs!grp_call2)
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing

End Sub

Private Sub CmdBatal_Click()
Unload Me

End Sub

Private Sub cmdEdit_Click()
    SSTab1.Tab = 2
    releaseControl
    ListView1_DblClick (0)
    
End Sub

Private Sub cmdHapus_Click()
     If ListView1(0).ListItems.Count <> 0 Then
         If MsgBox("Are You sure to delete this data", vbQuestion + vbYesNo, App.Title) = vbYes Then
                If Clsstatuscall.deleteStatuscall(ListView1(0).SelectedItem.SubItems(1), MDIForm1.TxtUsername.text) = True Then
                            MsgBox "Has been deleted", vbInformation + vbOKOnly, App.Title
                                 ListView1(0).ListItems.Remove ListView1(0).SelectedItem.Index
                    Else
                               MsgBox "Can't deleted from database ", vbInformation + vbOKOnly, App.Title
                    End If

          End If
    End If
End Sub
Private Sub cmdkeluar_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim straktif As String
Dim strAllow As String
On Error GoTo hell
 ' simpan /edit data
    If txtuserid.text = Empty Then
        MsgBox "Kode Status Call Can't be Blank", vbInformation + vbOKOnly, App.Title
        txtuserid.SetFocus
        Exit Sub
     End If
                    
    If txtNmAgent.text = Empty Then
        MsgBox "Status Call Name Can't be Blank", vbInformation + vbOKOnly, App.Title
        txtnama.SetFocus
        Exit Sub
    End If
    
    If cmbgroupcall1.text = "" Then
        MsgBox "Group Call Can't be Blank", vbInformation + vbOKOnly, App.Title
        cmbgroupcall1.SetFocus
        Exit Sub
     End If

'    If cmbgroupcall3.Text = "" Then
'        MsgBox "Group Call Can't be Blank", vbInformation + vbOKOnly, App.Title
'        cmbgroupcall3.SetFocus
'        Exit Sub
'     End If

 'cek data dupicate berdasarkan userid
 
 ' set variabel menjadi record
    Set OBJRECORD = Clsstatuscall.findKdstatus(txtuserid.text)
 'cek data ada atau tidak
    If OBJRECORD.RecordCount > 0 Then
            If MsgBox("Data Already Exist, Do you want to replace to existing account ", vbQuestion + vbYesNo, App.Title) = vbYes Then
                 If Option1(0).Value = True Then
                     straktif = "1"
                 Else
                     straktif = "0"
                 End If
                 
                 
                'encripsi userid  dan password
                 If Clsstatuscall.updateStatusCall(txtuserid.text, txtNmAgent.text, straktif, MDIForm1.TxtUsername.text, cmbgroupcall1.text, cmbgroupcall2.text) = True Then
                     MsgBox "Data has been Replace", vbInformation + vbOKOnly, App.Title
                    
                     On Error Resume Next
                     fill_list_manager
                   
                Else
                    MsgBox "Data Can't Replace because any problem in database", vbInformation + vbOKOnly, App.Title
                End If
                txtuserid.text = IIf(IsNull(OBJRECORD!tblstatuscall_kdstscall), "", OBJRECORD!tblstatuscall_kdstscall)
                txtnama.text = IIf(IsNull(OBJRECORD!tblstatuscall_keterangan), "", OBJRECORD!tblstatuscall_keterangan)
                cmbgroupcall1.text = IIf(IsNull(OBJRECORD!grp_call), "", OBJRECORD!grp_call)
                cmbgroupcall2.text = IIf(IsNull(OBJRECORD!grp_call2), "", OBJRECORD!grp_call2)

                If IIf(IsNull(OBJRECORD!tblstatuscall_kdstatus), "", OBJRECORD!tblstatuscall_kdstatus) = "1" Then
                    Option1(0).Value = True
                Else
                    Option1(1).Value = True
                                   
                End If
                  releaseControl
              End If
    Else
            If Option1(0).Value = True Then
                straktif = "1"
            Else
                straktif = "0"
            End If
            
                 
                            
            If Clsstatuscall.saveStatuscall(txtuserid.text, txtNmAgent.text, straktif, MDIForm1.TxtUsername.text, MDIForm1.txtnama.text, cmbgroupcall1.text, cmbgroupcall2.text) = True Then
                MsgBox "Data has been Insert", vbInformation + vbOKOnly, App.Title
             
                On Error Resume Next
                 fill_list_manager
                    releaseControl
                                       
            Else
                  MsgBox "Data Can't Insert because any problem in database", vbInformation + vbOKOnly, App.Title
            End If
   End If
   Set OBJRECORD = Nothing
   Exit Sub
hell:
   MsgBox err.Description, vbInformation + vbOKOnly, App.Title
   Exit Sub
   
End Sub

Private Sub CmdSearchBaru_Click(Index As Integer)
Dim list As ListItem
Select Case Index
Case 0

    Set M_objrs = Clsstatuscall.FindRecordStatuscall(Combo1(0).text, Text1(0).text)
        ListView1(0).ListItems.CLEAR
        txtjmlrow(0).text = M_objrs.RecordCount
        While Not M_objrs.EOF
            
                Set list = ListView1(0).ListItems.ADD(, , IIf(IsNull(M_objrs!tblstatuscall_id), "", M_objrs!tblstatuscall_id))
                list.SubItems(1) = IIf(IsNull(M_objrs!tblstatuscall_kdstscall), "", M_objrs!tblstatuscall_kdstscall)
                list.SubItems(2) = IIf(IsNull(M_objrs!tblstatuscall_keterangan), "", M_objrs!tblstatuscall_keterangan)
                list.SubItems(3) = IIf(IsNull(M_objrs!grp_call), "", M_objrs!grp_call)
                list.SubItems(4) = IIf(IsNull(M_objrs!grp_call2), "", M_objrs!grp_call2)

                If IIf(IsNull(M_objrs!tblstatuscall_kdstatus), "", M_objrs!tblstatuscall_kdstatus) = "1" Then
                    list.SubItems(5) = "Aktif"
                Else
                    list.SubItems(5) = "Tidak Aktif"
                End If
                
                list.SubItems(6) = IIf(IsNull(M_objrs!tblstatuscall_ketuserwrite), "", M_objrs!tblstatuscall_ketuserwrite)
                
              M_objrs.MoveNext
        Wend
        Warna_Row_Listview Form_statuscall, ListView1(0), &HFFFF80, vbWhite
        Set M_objrs = Nothing
Case 1
    Set M_objrs = Clsstatuscall.FindRecordStatuscallHST(Combo1(1).text, Text1(1).text)
         ListView1(1).ListItems.CLEAR
         txtjmlrow(1).text = M_objrs.RecordCount
        While Not M_objrs.EOF
           Set list = ListView1(1).ListItems.ADD(, , IIf(IsNull(M_objrs!tblstatuscall_hst_id), "", M_objrs!tblstatuscall_hst_id))
                list.SubItems(1) = IIf(IsNull(M_objrs!tblstatuscall_hst_kdstscall), "", M_objrs!tblstatuscall_hst_kdstscall)
                list.SubItems(2) = IIf(IsNull(M_objrs!tblstatuscall_hst_keterangan), "", M_objrs!tblstatuscall_hst_keterangan)
                list.SubItems(3) = IIf(IsNull(M_objrs!grp_call), "", M_objrs!grp_call)
                list.SubItems(4) = IIf(IsNull(M_objrs!grp_call2), "", M_objrs!grp_call2)

                If IIf(IsNull(M_objrs!tblstatuscall_hst_kdstatus), "", M_objrs!tblstatuscall_hst_kdstatus) = "1" Then
                    list.SubItems(5) = "Aktif"
                Else
                    list.SubItems(5) = "Tidak Aktif"
                End If
                list.SubItems(6) = Format(IIf(IsNull(M_objrs("tblstatuscall_hst_tglentry")), "", M_objrs("tblstatuscall_hst_tglentry")), "dd/mm/yyyy")
                list.SubItems(7) = IIf(IsNull(M_objrs("tblstatuscall_hst_action")), "", M_objrs("tblstatuscall_hst_action"))
                list.SubItems(8) = IIf(IsNull(M_objrs("tblstatuscall_hst_nama_user")), "", M_objrs("tblstatuscall_hst_nama_user"))
               
              M_objrs.MoveNext
        Wend
        Warna_Row_Listview Form_statuscall, ListView1(1), &HFFFF80, vbWhite
        Set OBJRECORD = Nothing

End Select

End Sub

Private Sub cmdTambah_Click()
    SSTab1.Tab = 2
    releaseControl
End Sub

Private Sub Combo1_DropDown(Index As Integer)
Select Case Index
    Case 0
        loadCbofield
    Case 1
        loadCbofield
End Select

End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    Set Clsstatuscall = New Clsstatuscall
    create_header_manager
     fill_list_manager
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set clsuser = Nothing
End Sub
Public Sub releaseControl()
    txtNmAgent.text = ""
    txtuserid.text = ""
    cmbgroupcall1.text = ""
    cmbgroupcall2.text = ""
    Option1(0).Value = False
    Option1(1).Value = False
  
End Sub
Public Sub create_header_manager()
    With ListView1(0)
            .ColumnHeaders.ADD 1, , "ID", 5 * TXT
            .ColumnHeaders.ADD 2, , "Kode Status Call", 13 * TXT
            .ColumnHeaders.ADD 3, , "Status Call Name", 13 * TXT
            .ColumnHeaders.ADD 4, , "Group Call 1", 12 * TXT
            .ColumnHeaders.ADD 5, , "Group Call 2", 0 * TXT
            .ColumnHeaders.ADD 6, , "Active", 10 * TXT
            .ColumnHeaders.ADD 7, , "Last User Input", 13 * TXT
    End With
    
    
    With ListView1(1)
            .ColumnHeaders.ADD 1, , "ID", 5 * TXT
            .ColumnHeaders.ADD 2, , "Kode Status Call", 13 * TXT
            .ColumnHeaders.ADD 3, , "Status Call Name", 13 * TXT
            .ColumnHeaders.ADD 4, , "Group Call 1", 12 * TXT
            .ColumnHeaders.ADD 5, , "Group Call 2", 0 * TXT
            .ColumnHeaders.ADD 6, , "Active", 10 * TXT
            .ColumnHeaders.ADD 7, , "Tgl_entry", 15 * TXT
            .ColumnHeaders.ADD 8, , "Action", 10 * TXT
            .ColumnHeaders.ADD 9, , "Last User", 13 * TXT
    End With

End Sub
Public Sub fill_list_manager()
Dim list As ListItem
    Set M_objrs = Clsstatuscall.FindRecordStatuscall()
        ListView1(0).ListItems.CLEAR
        txtjmlrow(0).text = M_objrs.RecordCount
        While Not M_objrs.EOF
            Set list = ListView1(0).ListItems.ADD(, , IIf(IsNull(M_objrs!tblstatuscall_id), "", M_objrs!tblstatuscall_id))
                list.SubItems(1) = IIf(IsNull(M_objrs!tblstatuscall_kdstscall), "", M_objrs!tblstatuscall_kdstscall)
                list.SubItems(2) = IIf(IsNull(M_objrs!tblstatuscall_keterangan), "", M_objrs!tblstatuscall_keterangan)
                list.SubItems(3) = IIf(IsNull(M_objrs!grp_call), "", M_objrs!grp_call)
                list.SubItems(4) = IIf(IsNull(M_objrs!grp_call2), "", M_objrs!grp_call2)
                If IIf(IsNull(M_objrs!tblstatuscall_kdstatus), "", M_objrs!tblstatuscall_kdstatus) = "1" Then
                    list.SubItems(5) = "Aktif"
                Else
                    list.SubItems(5) = "Tidak Aktif"
                End If
                list.SubItems(6) = IIf(IsNull(M_objrs!tblstatuscall_ketuserwrite), "", M_objrs!tblstatuscall_ketuserwrite)
              M_objrs.MoveNext
        Wend
        Warna_Row_Listview Form_statuscall, ListView1(0), &HFFFF80, vbWhite
        Set M_objrs = Nothing
        
End Sub
Private Sub ListView1_DblClick(Index As Integer)
Dim sUserid As String
    Select Case Index
        Case 0
        If ListView1(0).ListItems.Count <> 0 Then
           sUserid = ListView1(0).SelectedItem.SubItems(1)
           
           Set M_objrs = Clsstatuscall.FindRecordStatuscall("tblstatuscall_kdstscall", sUserid)
            If M_objrs.RecordCount <> 0 Then
                txtuserid.text = IIf(IsNull(M_objrs!tblstatuscall_kdstscall), "", M_objrs!tblstatuscall_kdstscall)
                txtNmAgent.text = IIf(IsNull(M_objrs!tblstatuscall_keterangan), "", M_objrs!tblstatuscall_keterangan)
                cmbgroupcall1.text = IIf(IsNull(M_objrs!grp_call), "", M_objrs!grp_call)
                cmbgroupcall2.text = IIf(IsNull(M_objrs!grp_call2), "", M_objrs!grp_call2)

                If IIf(IsNull(M_objrs!tblstatuscall_kdstatus), "", M_objrs!tblstatuscall_kdstatus) = "1" Then
                     Option1(0).Value = True
                Else
                      Option1(1).Value = True
                End If
                SSTab1.Tab = 2
            End If
           Set M_objrs = Nothing
        End If
    End Select
    
End Sub
Private Sub txtUserId_KeyPress(KeyAscii As Integer)
Dim OBJRECORD As New ADODB.Recordset
    If KeyAscii = 13 Then
        Set OBJRECORD = Clsstatuscall.FindRecordStatuscall("tblstatuscall_kdstscall", txtuserid.text)
        If OBJRECORD.RecordCount > 0 Then
                txtuserid.text = IIf(IsNull(OBJRECORD!tblstatuscall_kdstscall), "", OBJRECORD!tblstatuscall_kdstscall)
                txtNmAgent.text = IIf(IsNull(OBJRECORD!tblstatuscall_keterangan), "", OBJRECORD!tblstatuscall_keterangan)
                cmbgroupcall1.text = IIf(IsNull(OBJRECORD!grp_call), "", OBJRECORD!grp_call)
                cmbgroupcall2.text = IIf(IsNull(OBJRECORD!grp_call2), "", OBJRECORD!grp_call2)

                If IIf(IsNull(OBJRECORD!tblstatuscall_kdstatus), "", OBJRECORD!tblstatuscall_kdstatus) = "1" Then
                     Option1(0).Value = True
                Else
                      Option1(1).Value = True
                End If
                    
    Else
        txtNmAgent.text = ""
        Option1(0).Value = False
        Option1(1).Value = False
   End If
End If


End Sub
Public Sub loadCbofield()
Dim list As ListItem
Dim i As Integer
    Set M_objrs = Clsstatuscall.FindRecordStatuscall()
      
    Combo1(0).CLEAR
    For i = 0 To M_objrs.fields.Count - 1
       Combo1(0).AddItem M_objrs.fields(i).Name
    Next i
    Set M_objrs = Nothing
    
    Set M_objrs = Clsstatuscall.FindRecordStatuscallHST()
      
    Combo1(1).CLEAR
    For i = 0 To M_objrs.fields.Count - 1
       Combo1(1).AddItem M_objrs.fields(i).Name
    Next i
    Set M_objrs = Nothing
        


End Sub



