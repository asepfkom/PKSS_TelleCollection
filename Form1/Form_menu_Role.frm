VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_menu_Role 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Master Menu Of Rules"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   13965
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
      TabPicture(0)   =   "Form_menu_Role.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CmdSearchBaru(0)"
      Tab(0).Control(1)=   "cmdTambah"
      Tab(0).Control(2)=   "cmdEdit"
      Tab(0).Control(3)=   "cmdKeluar"
      Tab(0).Control(4)=   "txtjmlrow(0)"
      Tab(0).Control(5)=   "ListView1(0)"
      Tab(0).Control(6)=   "Line1(0)"
      Tab(0).Control(7)=   "Line1(3)"
      Tab(0).Control(8)=   "Label1(1)"
      Tab(0).Control(9)=   "Image3(0)"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "History      "
      TabPicture(1)   =   "Form_menu_Role.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image3(1)"
      Tab(1).Control(1)=   "Label1(2)"
      Tab(1).Control(2)=   "Line1(4)"
      Tab(1).Control(3)=   "Line1(1)"
      Tab(1).Control(4)=   "ListView1(1)"
      Tab(1).Control(5)=   "txtjmlrow(1)"
      Tab(1).Control(6)=   "CmdSearchBaru(1)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "[Entry / Edit]     "
      TabPicture(2)   =   "Form_menu_Role.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Image3(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Line1(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label2(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label2(1)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdOk"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "CmdBatal"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtNmAgent"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame1"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Combo2"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cbolevelname"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Frame2"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cbo_level_visible"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      Begin VB.ComboBox cbo_level_visible 
         Height          =   315
         Left            =   1845
         TabIndex        =   29
         Top             =   405
         Width           =   4335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Destination"
         Height          =   4830
         Left            =   8460
         TabIndex        =   27
         Top             =   1170
         Width           =   5370
         Begin MSComctlLib.ListView ListView2 
            Height          =   4515
            Index           =   1
            Left            =   90
            TabIndex        =   28
            Top             =   225
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   7964
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
      End
      Begin VB.ComboBox cbolevelname 
         Height          =   315
         Left            =   1845
         TabIndex        =   26
         Top             =   405
         Width           =   4335
      End
      Begin VB.CommandButton CmdSearchBaru 
         Height          =   360
         Index           =   0
         Left            =   -67665
         Picture         =   "Form_menu_Role.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   450
         Width           =   1515
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   8610
         TabIndex        =   24
         Top             =   405
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Source"
         Height          =   4830
         Left            =   1845
         TabIndex        =   18
         Top             =   1170
         Width           =   6540
         Begin VB.CommandButton cmd 
            Caption         =   "<<"
            Height          =   255
            Index           =   3
            Left            =   5490
            TabIndex        =   23
            Top             =   1050
            Width           =   945
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">>"
            Height          =   255
            Index           =   2
            Left            =   5490
            TabIndex        =   22
            Top             =   795
            Width           =   945
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<"
            Height          =   255
            Index           =   1
            Left            =   5490
            TabIndex        =   21
            Top             =   510
            Width           =   945
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">"
            Height          =   255
            Index           =   0
            Left            =   5490
            TabIndex        =   20
            Top             =   225
            Width           =   945
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   4245
            Index           =   0
            Left            =   135
            TabIndex        =   19
            Top             =   225
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   7488
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
         Left            =   -66000
         Picture         =   "Form_menu_Role.frx":0642
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   450
         UseMaskColor    =   -1  'True
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
         Left            =   -64380
         Picture         =   "Form_menu_Role.frx":0CD6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   450
         Width           =   1590
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
         Height          =   345
         Left            =   -62715
         Picture         =   "Form_menu_Role.frx":12CA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   450
         Width           =   1550
      End
      Begin VB.CommandButton CmdSearchBaru 
         Height          =   360
         Index           =   1
         Left            =   -74955
         Picture         =   "Form_menu_Role.frx":1910
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   405
         Width           =   1515
      End
      Begin VB.TextBox txtNmAgent 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1860
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   765
         Visible         =   0   'False
         Width           =   4320
      End
      Begin VB.CommandButton CmdBatal 
         Height          =   375
         Left            =   12120
         Picture         =   "Form_menu_Role.frx":1EFE
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6570
         Width           =   1575
      End
      Begin VB.CommandButton cmdOk 
         Height          =   375
         Left            =   10440
         Picture         =   "Form_menu_Role.frx":2544
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   6570
         Width           =   1575
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
         TabIndex        =   2
         Top             =   6660
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
         Left            =   -62940
         MaxLength       =   20
         TabIndex        =   1
         Top             =   6615
         Width           =   1785
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5505
         Index           =   0
         Left            =   -75000
         TabIndex        =   10
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
         TabIndex        =   11
         Top             =   945
         Width           =   13920
         _ExtentX        =   24553
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
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Menu"
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
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   17
         Top             =   1215
         Width           =   1080
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
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -75000
         X2              =   -61185
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Level Name"
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
         TabIndex        =   14
         Top             =   405
         Width           =   1350
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
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         Index           =   3
         X1              =   -75000
         X2              =   -61185
         Y1              =   6570
         Y2              =   6570
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
         TabIndex        =   13
         Top             =   6705
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
         Index           =   2
         Left            =   -63750
         TabIndex        =   12
         Top             =   6660
         Width           =   810
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   0
         Left            =   -75090
         Picture         =   "Form_menu_Role.frx":2B6A
         Top             =   315
         Width           =   26295
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   1
         Left            =   -75000
         Picture         =   "Form_menu_Role.frx":A174
         Top             =   315
         Width           =   26295
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   2
         Left            =   0
         Picture         =   "Form_menu_Role.frx":1177E
         Top             =   315
         Width           =   26295
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Master Menu Of Rules"
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
      Left            =   600
      TabIndex        =   16
      Top             =   300
      Width           =   3585
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   1
      Left            =   105
      Picture         =   "Form_menu_Role.frx":18D88
      Stretch         =   -1  'True
      Top             =   240
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   -1980
      Picture         =   "Form_menu_Role.frx":19892
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19980
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
      TabIndex        =   15
      Top             =   300
      Width           =   3585
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   -15
      Picture         =   "Form_menu_Role.frx":1ECFD
      Stretch         =   -1  'True
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "Form_menu_Role"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub loadlevel()
    Dim clslevel As clstbllevel
    Set clslevel = New clstbllevel
    Set M_objrs = clslevel.FindRecordLevel()
    Combo2.CLEAR
    While Not M_objrs.EOF
             Combo2.AddItem IIf(IsNull(M_objrs!tbllevel_kdlevel), "", M_objrs!tbllevel_kdlevel)
            M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    Set clslevel = Nothing
End Sub
Private Sub cbo_level_visible_Click()
    cbolevelname.ListIndex = cbo_level_visible.ListIndex
End Sub
Private Sub cbo_level_visible_DropDown()
    Call loadlevel_name
End Sub
Private Sub cbo_level_visible_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub cbolevelname_Click()
    Dim M_objrs  As New ADODB.Recordset
    
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        strsql = "SELECT * FROM tbllevel WHERE  tbllevel_keterangan ='" + cbolevelname.text + "'"
       M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
       Combo2.text = ""
       
        If Not M_objrs.EOF Then
            txtNmAgent.text = cbolevelname.text
            Combo2.text = IIf(IsNull(M_objrs!tbllevel_kdlevel), "", M_objrs!tbllevel_kdlevel)
            Combo2_Click
        End If
    txtNmAgent.text = cbolevelname.text
End Sub
Private Sub cbolevelname_DropDown()
    loadlevel_name
End Sub
Private Sub cmd_Click(Index As Integer)
    Dim lList As ListItem
    Dim n As Integer
    Select Case Index
        Case 0
        If ListView2(0).ListItems.Count <> 0 Then
                Set lList = ListView2(1).ListItems.ADD(, , ListView2(0).SelectedItem.text)
                    lList.SubItems(1) = ListView2(0).SelectedItem.SubItems(1)
                    lList.SubItems(2) = ListView2(0).SelectedItem.SubItems(2)
                    ListView2(0).ListItems.Remove ListView2(0).SelectedItem.Index
        End If
        
        Case 1
        If ListView2(1).ListItems.Count <> 0 Then
                Set lList = ListView2(0).ListItems.ADD(, , ListView2(1).SelectedItem.text)
                    lList.SubItems(1) = ListView2(1).SelectedItem.SubItems(1)
                    lList.SubItems(2) = ListView2(1).SelectedItem.SubItems(2)
                    ListView2(1).ListItems.Remove ListView2(1).SelectedItem.Index
        End If
        
        Case 2
            n = ListView2(0).ListItems.Count
            For i = 1 To ListView2(0).ListItems.Count
                    Set lList = ListView2(1).ListItems.ADD(, , ListView2(0).ListItems(n).text)
                        lList.SubItems(1) = ListView2(0).ListItems(n).SubItems(1)
                        lList.SubItems(2) = ListView2(0).ListItems(n).SubItems(2)
                        ListView2(0).ListItems.Remove n
                        n = n - 1
            Next
                
        Case 3
            n = ListView2(1).ListItems.Count
            For i = 1 To ListView2(1).ListItems.Count
                    Set lList = ListView2(0).ListItems.ADD(, , ListView2(1).ListItems(n).text)
                        lList.SubItems(1) = ListView2(1).ListItems(n).SubItems(1)
                        lList.SubItems(2) = ListView2(1).ListItems(n).SubItems(2)
                        ListView2(1).ListItems.Remove n
                        n = n - 1
            Next
    End Select
End Sub
Private Sub CmdBatal_Click()
    Unload Me
End Sub
Private Sub cmdEdit_Click()
    ListView1_DblClick (0)
End Sub
Private Sub cmdkeluar_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    On Error GoTo ke
        If Combo2.text = "" Then
            MsgBox "Level Harus di isi", vbInformation + vbOKOnly, App.Title
            Exit Sub
        End If

    If MsgBox("Apakah anda yakin merubah config menu", vbQuestion + vbYesNo, "Question") = vbYes Then
        If ListView2(1).ListItems.Count <> 0 Then
            M_OBJCONN.BeginTrans
            
            M_OBJCONN.Execute "delete from tblmenu_role where tblmenu_role_kdlevel ='" + Combo2.text + "'"
            
            '===asep18/12/2019===='
'            strsql = "select max(tblmenu_role_id) as id from tblmenu_role"
'            Set rs = New ADODB.Recordset
'            rs.CursorLocation = adUseClient
'            rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'            a = cnull(rs!ID)
'
'            M_OBJCONN.Execute "select setval('tblmenu_role_tblmenu_role_id_seq', " & a & ", true); "
            '===================='
            
            For i = 1 To ListView2(1).ListItems.Count
                    sStrsql = ""
                    sStrsql = "insert into tblmenu_role (tblmenu_role_key_menu, tblmenu_role_ket_menu, tblmenu_role_group,tblmenu_role_kdlevel,tblmenu_role_kduser,tblmenu_role_ketuser , tblmenu_role_ketlevel ) values ("
                    sStrsql = sStrsql + "'" + ListView2(1).ListItems(i).text + "','" + ListView2(1).ListItems(i).SubItems(1) + "', '" + ListView2(1).ListItems(i).SubItems(2) + "','" + Combo2.text + "','" + MDIForm1.TxtUsername.text + "','" + MDIForm1.txtnama.text + "','" + txtNmAgent.text + "')"
                    M_OBJCONN.Execute (sStrsql)
                   
            Next
                    sStrsql = "insert into tblmenuhst_role (tblmenuhst_role_key_menu, tblmenuhst_role_ket_menu, tblmenuhst_role_group,tblmenuhst_role_kdlevel, tblmenuhst_role_nama_user, tblmenuhst_role_ket_level  ) select "
                    sStrsql = sStrsql + " tblmenu_role_key_menu, tblmenu_role_ket_menu, tblmenu_role_group,tblmenu_role_kdlevel ,'" + MDIForm1.TxtUsername.text + "' as nm ,   tblmenu_role_ketlevel from  tblmenu_role where   tblmenu_role_kdlevel ='" + Combo2.text + "'  "
                    M_OBJCONN.Execute (sStrsql)
                   
            M_OBJCONN.CommitTrans
            ListView2(1).ListItems.CLEAR
            ListView2(0).ListItems.CLEAR
            txtNmAgent.text = ""
            Combo2.text = ""
            MsgBox "Has Been reload menu", vbInformation + vbOKOnly, App.Title
        End If
    End If
    Exit Sub
ke:
    M_OBJCONN.RollbackTrans
End Sub
Private Sub CmdSearchBaru_Click(Index As Integer)
    Dim clsmenu As clsmenu
    Dim list As ListItem
    Set clsmenu = New clsmenu
    Dim OBJRECORD2 As New ADODB.Recordset
        Select Case Index
            Case 0
             fill_list_manager
             
            Case 1
             Set OBJRECORD2 = clsmenu.FindMenuRoleHst()
                        ListView1(1).ListItems.CLEAR
                        While Not OBJRECORD2.EOF
                                Set list = ListView1(1).ListItems.ADD(, , IIf(IsNull(OBJRECORD2!tblmenuhst_role_key_menu), "", OBJRECORD2!tblmenuhst_role_key_menu))
                                    list.SubItems(1) = IIf(IsNull(OBJRECORD2!tblmenuhst_role_ket_menu), "", OBJRECORD2!tblmenuhst_role_ket_menu)
                                    list.SubItems(2) = IIf(IsNull(OBJRECORD2!tblmenuhst_role_group), "", OBJRECORD2!tblmenuhst_role_group)
                                    list.SubItems(3) = Format(IIf(IsNull(OBJRECORD2!tblmenuhst_role_tglentry), "", OBJRECORD2!tblmenuhst_role_tglentry), "dd-mm-yyyy hh:nn:ss")
                                    list.SubItems(4) = IIf(IsNull(OBJRECORD2!tblmenuhst_role_ket_level), "", OBJRECORD2!tblmenuhst_role_ket_level)
                                    list.SubItems(5) = IIf(IsNull(OBJRECORD2!tblmenuhst_role_nama_user), "", OBJRECORD2!tblmenuhst_role_nama_user)
                                OBJRECORD2.MoveNext
                        Wend
                        
                        Warna_Row_Listview Form_menu_Role, ListView1(1), &HFFFF80, vbWhite
                        txtjmlrow(1).text = OBJRECORD2.RecordCount
                        Set OBJRECORD2 = Nothing
                        Set clsmenu = Nothing
        End Select
End Sub
Private Sub cmdTambah_Click()
    SSTab1.Tab = 2
End Sub
Private Sub Combo2_Change()
    Combo2_Click
End Sub
Private Sub Combo2_Click()
    Dim clslevel As clstbllevel
    Dim clsmenu As clsmenu
    Dim list As ListItem
    Set clslevel = New clstbllevel
    Set clsmenu = New clsmenu
    Dim OBJRECORD As New ADODB.Recordset
    Dim OBJRECORD2 As New ADODB.Recordset

        Set OBJRECORD = clslevel.FindRecordLevel("tbllevel_kdlevel", Combo2.text)
            If OBJRECORD.RecordCount > 0 Then
               txtNmAgent.text = IIf(IsNull(OBJRECORD!tbllevel_keterangan), "", OBJRECORD!tbllevel_keterangan)
                Set OBJRECORD2 = clsmenu.FindMenuSource(Combo2.text)
                ListView2(0).ListItems.CLEAR
                While Not OBJRECORD2.EOF
                        Set list = ListView2(0).ListItems.ADD(, , IIf(IsNull(OBJRECORD2!tblmenu_key_menu), "", OBJRECORD2!tblmenu_key_menu))
                            list.SubItems(1) = IIf(IsNull(OBJRECORD2!tblmenu_ket_menu), "", OBJRECORD2!tblmenu_ket_menu)
                            list.SubItems(2) = IIf(IsNull(OBJRECORD2!tblmenu_group), "", OBJRECORD2!tblmenu_group)
                        OBJRECORD2.MoveNext
                Wend
        
                Set OBJRECORD2 = Nothing
                Set OBJRECORD2 = clsmenu.FindMenuDestination(Combo2.text)
                ListView2(1).ListItems.CLEAR
                While Not OBJRECORD2.EOF
                        Set list = ListView2(1).ListItems.ADD(, , IIf(IsNull(OBJRECORD2!tblmenu_role_key_menu), "", OBJRECORD2!tblmenu_role_key_menu))
                            list.SubItems(1) = IIf(IsNull(OBJRECORD2!tblmenu_role_ket_menu), "", OBJRECORD2!tblmenu_role_ket_menu)
                            list.SubItems(2) = IIf(IsNull(OBJRECORD2!tblmenu_role_group), "", OBJRECORD2!tblmenu_role_group)
                        OBJRECORD2.MoveNext
                Wend
                Set OBJRECORD2 = Nothing
            Else
                txtNmAgent.text = ""
            End If
    Set OBJRECORD = Nothing
    Set clslevel = Nothing
    Set clsmenu = Nothing
End Sub
Private Sub Combo2_DropDown()
    loadlevel
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    Dim clslevel As clstbllevel
    Set clslevel = New clstbllevel
    Dim OBJRECORD As New ADODB.Recordset
        If KeyAscii = 13 Then
            Combo2_Click
        End If
End Sub
Private Sub Form_Load()
    create_header_menu
End Sub
Public Sub create_header_menu()
    With ListView2(0)
            .ColumnHeaders.ADD 1, , "Key menu", 10 * TXT
            .ColumnHeaders.ADD 2, , "Keterangan Menu", 20 * TXT
            .ColumnHeaders.ADD 3, , "Group", 20 * TXT
    End With
    
    With ListView2(1)
            .ColumnHeaders.ADD 1, , "Key menu", 10 * TXT
            .ColumnHeaders.ADD 2, , "Keterangan Menu", 20 * TXT
            .ColumnHeaders.ADD 3, , "Group", 20 * TXT
    End With
    
    With ListView1(1)
            .ColumnHeaders.ADD 1, , "Key menu", 10 * TXT
            .ColumnHeaders.ADD 2, , "Keterangan Menu", 20 * TXT
            .ColumnHeaders.ADD 3, , "Group", 20 * TXT
        
            .ColumnHeaders.ADD 4, , "tgl Entry", 20 * TXT
            .ColumnHeaders.ADD 5, , "Level", 20 * TXT
            .ColumnHeaders.ADD 6, , "User Last", 20 * TXT
    End With
    
    With ListView1(0)
            .ColumnHeaders.ADD 1, , "Kdlevel", 10 * TXT
            .ColumnHeaders.ADD 2, , "Keterangan Level", 0
            .ColumnHeaders.ADD 3, , "Keterangan Level", 20 * TXT
    End With
End Sub
Public Sub fill_list_manager()
    Dim list As ListItem
    Dim clsmenu As clsmenu
    Set clsmenu = New clsmenu
        Set M_objrs = clsmenu.FindMenulvlDist()
            ListView1(0).ListItems.CLEAR
            txtjmlrow(0).text = M_objrs.RecordCount
            While Not M_objrs.EOF
                Set list = ListView1(0).ListItems.ADD(, , IIf(IsNull(M_objrs!tblmenu_role_kdlevel), "", M_objrs!tblmenu_role_kdlevel))
                    list.SubItems(1) = IIf(IsNull(M_objrs!tblmenu_role_ketlevel), "", M_objrs!tblmenu_role_ketlevel)
                    list.SubItems(2) = IIf(IsNull(M_objrs!level_name), "", M_objrs!level_name)
                  M_objrs.MoveNext
            Wend
            Warna_Row_Listview Form_menu_Role, ListView1(0), &HFFFF80, vbWhite
            Set M_objrs = Nothing
            Set clsmenu = Nothing
End Sub
Private Sub ListView1_DblClick(Index As Integer)
    If ListView1(0).ListItems.Count <> 0 Then
        Combo2.text = ListView1(0).SelectedItem.text
        cbolevelname.text = ListView1(0).SelectedItem.SubItems(1)
        cbo_level_visible.text = ListView1(0).SelectedItem.SubItems(2)
        SSTab1.Tab = 2
    End If
End Sub
Public Sub loadlevel_name()
    Dim clslevel As clstbllevel
    Set clslevel = New clstbllevel
    Set M_objrs = clslevel.FindRecordLevel()
    Combo2.CLEAR
    cbolevelname.CLEAR: cbo_level_visible.CLEAR
    While Not M_objrs.EOF
            cbolevelname.AddItem IIf(IsNull(M_objrs!tbllevel_keterangan), "", M_objrs!tbllevel_keterangan)
            '11 Juni 2014 BY IZUDDIN - USER MATRIX
            cbo_level_visible.AddItem IIf(IsNull(M_objrs!level_name), "", M_objrs!level_name)
            ' Combo2.AddItem IIf(IsNull(M_OBJRS!tbllevel_kdlevel), "", M_OBJRS!tbllevel_kdlevel)
            M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    Set clslevel = Nothing
End Sub
