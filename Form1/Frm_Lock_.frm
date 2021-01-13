VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frm_lock 
   Caption         =   "Lock Account Data"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form2"
   ScaleHeight     =   7965
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameEntry 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   2025
      Left            =   0
      TabIndex        =   1
      Top             =   5190
      Width           =   8850
      Begin VB.OptionButton OptNewEntry 
         Caption         =   "New entry "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5355
         TabIndex        =   11
         Top             =   420
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ComboBox CmbNewEntry 
         Height          =   315
         ItemData        =   "Frm_Lock_.frx":0000
         Left            =   3660
         List            =   "Frm_Lock_.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   450
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton OptReguler 
         Caption         =   "Reguler "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5145
         TabIndex        =   9
         Top             =   780
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ComboBox CmbReguler 
         Height          =   315
         ItemData        =   "Frm_Lock_.frx":0010
         Left            =   3660
         List            =   "Frm_Lock_.frx":0017
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.ComboBox CmbSwap 
         Height          =   315
         ItemData        =   "Frm_Lock_.frx":0020
         Left            =   3660
         List            =   "Frm_Lock_.frx":0027
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1260
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton OptSwap 
         Caption         =   "Swap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5670
         TabIndex        =   6
         Top             =   1365
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CheckBox chkcurrent 
         Caption         =   "Current Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   5
         Top             =   1530
         Width           =   1755
      End
      Begin VB.CheckBox chkswap 
         Caption         =   "Swap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   4
         Top             =   1140
         Width           =   1755
      End
      Begin VB.CheckBox chkreguler 
         Caption         =   "Reguler"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   690
         Width           =   1485
      End
      Begin VB.CheckBox chknewentry 
         Caption         =   "New Entry"
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
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   1725
      End
      Begin VB.Label Label2 
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   510
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   900
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label4 
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   585
      End
   End
   Begin VB.CheckBox CheckEntry 
      Caption         =   "Tampilkan data entry:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   5130
      Width           =   2565
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   1931
      _Version        =   196610
      BackColor       =   14737632
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   3120
         TabIndex        =   18
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox CHKLUNASPENDING 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LUNAS PENDING"
         Height          =   465
         Left            =   5970
         TabIndex        =   17
         Top             =   60
         Width           =   1725
      End
      Begin VB.CheckBox CHKACCOUNT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LUNAS COMPLETE"
         Height          =   285
         Left            =   5970
         TabIndex        =   16
         Top             =   570
         Width           =   1845
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "All TeleCollection"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "Pilih TeleCollection"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "Pilih SPV Name"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "TeleCollection Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3825
      Left            =   0
      TabIndex        =   24
      Top             =   1260
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   6747
      _Version        =   196610
      BackColor       =   14737632
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "S P - Settled Payment"
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   47
         Top             =   810
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "RP - Refuse Payment"
         Height          =   255
         Index           =   7
         Left            =   135
         TabIndex        =   46
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "VL-Valid"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OS - On Process"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "P T P - Promise To Pay"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B P - Broken Promise"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   42
         Top             =   120
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "P O P - Progress Of Payment"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   41
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Blank Data"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   40
         Top             =   1170
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   2325
         Left            =   90
         TabIndex        =   28
         Top             =   1410
         Width           =   8625
         Begin VB.CommandButton cmd 
            Caption         =   "<<"
            Height          =   375
            Index           =   3
            Left            =   3960
            TabIndex        =   35
            Top             =   1860
            Width           =   675
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">>"
            Height          =   375
            Index           =   2
            Left            =   3960
            TabIndex        =   34
            Top             =   1470
            Width           =   675
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<"
            Height          =   375
            Index           =   1
            Left            =   3960
            TabIndex        =   33
            Top             =   1080
            Width           =   675
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">"
            Height          =   375
            Index           =   0
            Left            =   3960
            TabIndex        =   32
            Top             =   690
            Width           =   675
         End
         Begin VB.CheckBox chksingle 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Single"
            Height          =   345
            Left            =   1620
            TabIndex        =   31
            Top             =   240
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CheckBox chkmultiple 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Multiple"
            Height          =   345
            Left            =   3150
            TabIndex        =   30
            Top             =   210
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Hapus"
            Height          =   315
            Left            =   3360
            TabIndex        =   29
            Top             =   720
            Width           =   585
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Left            =   90
            TabIndex        =   36
            Top             =   690
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   2778
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
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
         Begin MSComctlLib.ListView ListView2 
            Height          =   1575
            Left            =   4710
            TabIndex        =   37
            Top             =   660
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   2778
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
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
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Mark Up"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   330
            Width           =   1185
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Destination Lead Markup"
            Height          =   255
            Left            =   4800
            TabIndex        =   38
            Top             =   390
            Width           =   2685
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PR-Prospect"
         Height          =   255
         Index           =   9
         Left            =   5700
         TabIndex        =   27
         Top             =   150
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON-On Nego"
         Height          =   255
         Index           =   10
         Left            =   5700
         TabIndex        =   26
         Top             =   450
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SK-Skip"
         Height          =   255
         Index           =   11
         Left            =   5670
         TabIndex        =   25
         Top             =   780
         Width           =   1245
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   0
      Left            =   45
      TabIndex        =   48
      Top             =   7425
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Execute"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   1
      Left            =   1155
      TabIndex        =   49
      Top             =   7455
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Release"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   2
      Left            =   3375
      TabIndex        =   50
      Top             =   7425
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "E&xit"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   3
      Left            =   4485
      TabIndex        =   51
      Top             =   7425
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&SHUT"
      ButtonStyle     =   3
   End
End
Attribute VB_Name = "frm_lock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdsql As String
Dim StsVl As String
Dim StsOS As String
Dim StsON As String
Dim StsSK As String
Dim StsPR As String
Dim StsPTP As String
Dim StsBP As String
Dim StsPOP As String
Dim StsSP As String
Dim StsRP As String
Dim StsOP As String
Dim StsFresh As String
Dim Stsblank As String
Dim Stsuncontact As String
Dim spv As Boolean
'@@ 140710 Tambahan buat blok entry yang diambil dari field entry_date dan pay_dt di mgm
Dim BlokEntry As String
'@@ 18-11-10 Buat Perbaiki Lock Data BlokEntry
Dim StsNewEntry
Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Batch", 10 * 120
    ListView2.ColumnHeaders.ADD 1, , "Batch", 20 * 120
End Sub
Private Sub CheckEntry_Click()
    If FrameEntry.Enabled = False Then
        FrameEntry.Enabled = True
        OptNewEntry.Value = True
    Else
        FrameEntry.Enabled = False
        OptNewEntry.Value = False
        OptReguler.Value = False
        OptSwap.Value = False
    End If
End Sub


Private Sub chkmultiple_Click()
    If chkmultiple.Value = vbChecked Then
        chksingle.Value = vbUnchecked
    End If
End Sub

Private Sub chksingle_Click()
    If chksingle.Value = vbChecked Then
        chkmultiple.Value = vbUnchecked
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
Select Case Index
Case 1
    If ListView2.ListItems.Count <> 0 Then
            Set lList = ListView1.ListItems.ADD(, , ListView2.SelectedItem.Text)
            ListView2.ListItems.Remove ListView2.SelectedItem.Index
    End If
Case 3
    For i = 1 To ListView2.ListItems.Count
                Set lList = ListView1.ListItems.ADD(, , ListView2.SelectedItem.Text)
                ListView2.ListItems.Remove ListView2.SelectedItem.Index
    Next
Case 0
    If ListView1.ListItems.Count <> 0 Then
        Set lList = ListView2.ListItems.ADD(, , ListView1.SelectedItem.Text)
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
    End If

Case 2
    For i = 1 To ListView1.ListItems.Count
            Set lList = ListView2.ListItems.ADD(, , ListView1.SelectedItem.Text)
                   
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
    Next
End Select
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim M_DATA As New CLS_FRMSEARCH
Dim M_objrs As ADODB.Recordset
Select Case Index
Case 0
    If spv = False Then
        Set M_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("USERID")
            Combo1(1).Text = M_objrs("AGENT")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Else
        Set M_objrs = New ADODB.Recordset
        M_objrs.Open "select * from SPVTBL where SPVCODE='" + Combo1(0) + "'", M_OBJCONN, adOpenDynamic, adLockBatchOptimistic
            While Not M_objrs.EOF
                Combo1(0).Text = M_objrs("SPVCODE")
                Combo1(1).Text = M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
    End If
Case 1
    If spv = False Then
        Set M_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("USERID")
            Combo1(1).Text = M_objrs("AGENT")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Else
        Set M_objrs = New ADODB.Recordset
        M_objrs.Open "select * from SPVTBL where SPVNAME='" + Combo1(1) + "'", M_OBJCONN, adOpenDynamic, adLockBatchOptimistic
            While Not M_objrs.EOF
                Combo1(0).Text = M_objrs("SPVCODE")
                Combo1(1).Text = M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
    End If
    
 End Select
 Set M_DATA = Nothing
 Set M_objrs = Nothing
End Sub

Private Sub Command1_Click()
    M_OBJCONN.Execute "UPDATE MGM SET EXCLUDE =NULL WHERE EXCLUDE='" + ListView1.SelectedItem.Text + "'"
    getMarkup
End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim M_DATA As New CLS_FRMSEARCH
    
    Set M_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "")
        While Not M_objrs.EOF
            Combo1(0).AddItem M_objrs("USERID")
            Combo1(1).AddItem M_objrs("AGENT")
            M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
    SSOption1(0).Value = True
    spv = False
    header
    CmbNewEntry.Text = "< 2"
    CmbReguler.Text = "< 2"
    CmbSwap.Text = "> 2"
    getMarkup
End Sub


Private Sub SSCommand1_Click(Index As Integer)
Dim M_objrs As New ADODB.Recordset
Dim sStrsql As String
Dim MWHERE As String
Select Case Index
Case 0
If CHKLUNASPENDING.Value = vbChecked And CHKACCOUNT.Value = vbChecked Then
    If Combo1(1).Text <> Empty And SSOption1(2).Value = True Then
        sStrsql = " agent in (@LUNAS PENDING@,@LUNAS COMPLETE@) AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE LIKE @%SPV%@ ) "
       'mwhere = " WHERE SPVCODE LIKE '%SPV%'"
        sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE=@" + Combo1(0).Text + "@)"
        MWHERE = "Where spvcode='" + Combo1(0).Text + "'"
        STRSQL = "UPDATE usertbl SET dilockoleh='"
        STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
        STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + MWHERE
        M_OBJCONN.Execute STRSQL
    ElseIf Combo1(1).Text = Empty And SSOption1(2).Value = True Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open "SELECT * FROM USERTBL WHERE USERTYPE='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not M_objrs.EOF
                sStrsql = " agent in (@LUNAS PENDING@,@LUNAS COMPLETE@) "
                sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE=@" + M_objrs("SPVCODE") + "@)"
                MWHERE = "Where spvcode='" + M_objrs("SPVCODE") + "'"
                STRSQL = "UPDATE usertbl SET dilockoleh='"
                STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + MWHERE
                M_OBJCONN.Execute STRSQL
                M_objrs.MoveNext
            Wend
            Set M_objrs = Nothing
    Else
    Exit Sub
    End If
    
    MsgBox "Data Berhasil di Blok", vbOKOnly + vbInformation, "Pesan"
    Exit Sub
  ElseIf CHKLUNASPENDING.Value = vbChecked Then
        If Combo1(1).Text <> Empty And SSOption1(2).Value = True Then
          sStrsql = " agent in (@LUNAS PENDING@) AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE LIKE @%SPV%@ ) "
            MWHERE = " WHERE SPVCODE LIKE '%SPV%'"
            sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE=@" + Combo1(0).Text + "@)"
            MWHERE = "Where spvcode='" + Combo1(0).Text + "'"
             STRSQL = "UPDATE usertbl SET dilockoleh='"
             STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
             STRSQL = STRSQL + " lockdarispvbuattl ='" + sStrsql + "'" + MWHERE
            M_OBJCONN.Execute STRSQL
        ElseIf Combo1(1).Text = Empty And SSOption1(2).Value = True Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open "SELECT * FROM USERTBL WHERE USERTYPE='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not M_objrs.EOF
                sStrsql = " agent in (@LUNAS PENDING@) "
                sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE=@" + M_objrs("SPVCODE") + "@)"
                MWHERE = "Where spvcode='" + M_objrs("SPVCODE") + "'"
                STRSQL = "UPDATE usertbl SET dilockoleh='"
                STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + " lockdarispvbuattl ='" + sStrsql + "'" + MWHERE
                M_OBJCONN.Execute STRSQL
                M_objrs.MoveNext
            Wend
            Set M_objrs = Nothing
    Else
        Exit Sub
        
        End If
        
            MsgBox "Data Berhasil di Blok", vbOKOnly + vbInformation, "Pesan"
        Exit Sub
ElseIf CHKACCOUNT.Value = vbChecked Then
           sStrsql = " agent in (@LUNAS COMPLETE@) AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE LIKE @%SPV%@ ) "
            MWHERE = " WHERE SPVCODE LIKE '%SPV%'"
        If Combo1(1).Text <> Empty And SSOption1(2).Value = True Then
            sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE=@" + Combo1(0).Text + "@)"
            MWHERE = "Where spvcode='" + Combo1(0).Text + "'"
            STRSQL = "UPDATE usertbl SET dilockoleh='"
            STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
            STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + MWHERE
            M_OBJCONN.Execute STRSQL
        ElseIf Combo1(1).Text = Empty And SSOption1(2).Value = True Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open "SELECT * FROM USERTBL WHERE USERTYPE='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not M_objrs.EOF
                sStrsql = " agent in (@LUNAS COMPLETE@) "
                sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE=@" + M_objrs("SPVCODE") + "@)"
                MWHERE = "Where spvcode='" + M_objrs("SPVCODE") + "'"
                STRSQL = "UPDATE usertbl SET dilockoleh='"
                STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + MWHERE
                M_OBJCONN.Execute STRSQL
                M_objrs.MoveNext
            Wend
            Set M_objrs = Nothing
    Else
        Exit Sub
        
        End If
        
            MsgBox "Data Berhasil di Blok", vbOKOnly + vbInformation, "Pesan"
        Exit Sub
End If


        
        If SSOption1(0).Value = False And SSOption1(1).Value = False And SSOption1(2).Value = False Then
            MsgBox "Select DCR Name To Proccess OR All"
         Else
                If SSOption1(0).Value Then
                    Call ceksts
                    STRSQL = "UPDATE usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL, F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL "
                    STRSQL = STRSQL + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
                    STRSQL = STRSQL + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,lock_entry_lpd=NULL, lockmarkup=NULL,lockdarispv=NULL Where usertype='1'"
                    M_OBJCONN.Execute (STRSQL)
                    
                    
                    STRSQL = "UPDATE usertbl SET f_flagrender=1, lockdarispv ='"
                    STRSQL = STRSQL + getblock + "',lock_entry_lpd='"
                    STRSQL = STRSQL + BlokEntry + "',dilockoleh='"
                    STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "' "
                    STRSQL = STRSQL + " Where usertype='1'"
                    M_OBJCONN.Execute (STRSQL)
                    
                        If ListView2.ListItems.Count <> 0 Then
                            HLSMARKUP = Replace(GETSELECTMARKUP, "'", "@")
                            sStrsql = " UPDATE USERTBL SET dilockoleh='"
                            sStrsql = sStrsql + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                            sStrsql = sStrsql + "lockmarkup='" + HLSMARKUP + "' FROM ( "
                            sStrsql = sStrsql + " select distinct(agent) AS AGENT FROM MGM where exclude in(" + GETSELECTMARKUP + ")) AS C "
                            sStrsql = sStrsql + " Where USERTBL.USERID = C.agent "
                            M_OBJCONN.Execute (sStrsql)
                        End If
                    MsgBox "Proccess to All DCR Name Done.....!"
                  End If
        
                If SSOption1(1).Value Then
                    If Combo1(0).Text = "" Then
                        MsgBox "Select DCR Name To Proccess..!"
                        Combo1(0).SetFocus
                    Else
                        Call ceksts
                        STRSQL = "UPDATE usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL, F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL "
                        STRSQL = STRSQL + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
                        STRSQL = STRSQL + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,lock_entry_lpd=NULL, lockmarkup=NULL,lockdarispv=NULL Where userid='" + Combo1(0).Text + "'"
                        M_OBJCONN.Execute (STRSQL)
                        STRSQL = "UPDATE usertbl SET  f_flagrender=1,lockdarispv ='" + getblock + "',lock_entry_lpd='"
                        STRSQL = STRSQL + BlokEntry + "',dilockoleh='"
                        STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "' "
                        STRSQL = STRSQL + "Where userid='" + Combo1(0).Text + "'"
                        M_OBJCONN.Execute STRSQL
                        If ListView2.ListItems.Count <> 0 Then
                            HLSMARKUP = Replace(GETSELECTMARKUP, "'", "@")
                            sStrsql = " UPDATE USERTBL SET dilockoleh='"
                            sStrsql = sStrsql + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                            sStrsql = sStrsql + "lockmarkup='" + HLSMARKUP + "' FROM ( "
                            sStrsql = sStrsql + " select distinct(agent) AS AGENT FROM MGM where exclude in(" + GETSELECTMARKUP + ") AND AGENT='" + Combo1(0).Text + "') AS C "
                            sStrsql = sStrsql + " Where USERTBL.USERID = C.agent "
                            M_OBJCONN.Execute (sStrsql)
                        End If
                        
                    
                        MsgBox "Proccess To  " + Combo1(0).Text + "  " + Combo1(1).Text + " Done.....!"
                    End If
                Else
                    If SSOption1(2).Value = True Then
                        If Combo1(0).Text = "" Then
                            MsgBox "Select SPV Name To Proccess..!"
                        Else
                            Call ceksts
                        STRSQL = "UPDATE usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL,F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL "
                        STRSQL = STRSQL + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
                        STRSQL = STRSQL + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,lockmarkup=NULL,lockdarispv=null,lock_entry_lpd=null Where spvcode='" + Combo1(0).Text + "'"
                        M_OBJCONN.Execute (STRSQL)
                            
                       'If CHKLUNASPENDING.Value = vbChecked Then
                        '    STRSQL = "UPDATE usertbl SET f_flagrender=1,lockdarispvbuattl ='" + getblock + "',lock_entry_lpd='"
                         '   STRSQL = STRSQL + BlokEntry + "',fromaccount ='" + cboaccount.Text + "' Where spvcode='" + Combo1(0).Text + "'"
                          '  M_OBJCONN.Execute STRSQL
                       'End If
                       
                       'If CHKACCOUNT.Value = vbChecked Then
                            STRSQL = "UPDATE usertbl SET f_flagrender=1,lockdarispv ='" + getblock + "',lock_entry_lpd='"
                            STRSQL = STRSQL + BlokEntry + "',dilockoleh='"
                            STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "' "
                            STRSQL = STRSQL + "Where spvcode='" + Combo1(0).Text + "'"
                            M_OBJCONN.Execute STRSQL
                       ' End If
                        
                            
                        
                        If ListView2.ListItems.Count <> 0 Then
                            HLSMARKUP = Replace(GETSELECTMARKUP, "'", "@")
                            sStrsql = " UPDATE USERTBL SET dilockoleh='"
                            sStrsql = sStrsql + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                            sStrsql = sStrsql + "lockmarkup='" + HLSMARKUP + "' FROM ( "
                            sStrsql = sStrsql + " select distinct(agent) AS AGENT FROM MGM where exclude in(" + GETSELECTMARKUP + ") AND AGENT IN (SELECT USERID FROM USERTBL WHERE SPVCODE='" + Combo1(0).Text + "')) AS C "
                            sStrsql = sStrsql + " Where USERTBL.USERID = C.agent "
                            M_OBJCONN.Execute (sStrsql)
                        End If
                        
                            
                            MsgBox "Proccess To  " + Combo1(0).Text + "  " + Combo1(1).Text + " Done.....!"
                        End If
                    End If
             End If
        End If
        StsVl = ""
        StsPR = ""
        StsOS = ""
        StsON = ""
        StsSK = ""
        StsOP = ""
       StsPTP = ""
       StsBP = ""
       StsPOP = ""
       StsSP = ""
       StsUC = ""
       StsRP = ""
       StsWO_Date = ""
       StsWO_2009 = ""
       StsWO_2008 = ""
       StsWO_2007 = ""
       StsWO_2006 = ""
       StsWO_2005 = ""
       StsWO_2004 = ""
       StsWO_2003 = ""
       StsWO_2002 = ""
       StsWO_2001 = ""
       StsWO_2000 = ""
       StsWO_1999 = ""
       STRSQL = ""
 
Case 1
     If SSOption1(2).Value = True Then
               
            If CHKLUNASPENDING.Value = vbChecked Or CHKACCOUNT.Value = vbChecked Then
                    
                   Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            
            If Combo1(0).Text = "" Then
                    M_objrs.Open "SELECT * FROM USERTBL WHERE usertype='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            Else
                    M_objrs.Open "SELECT * FROM USERTBL WHERE SPVCODE='" + Combo1(0).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            End If
            
            While Not M_objrs.EOF
                STRSQL = "UPDATE usertbl SET dilockoleh='Clear by:"
                STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + " lockdarispvbuattl=NULL WHERE SPVCODE='" + M_objrs("SPVCODE") + "'"
                M_OBJCONN.Execute STRSQL
                M_objrs.MoveNext
            Wend
            Set M_objrs = Nothing
            MsgBox "Data telah direlease"
           Exit Sub
           End If
            
                If Combo1(0).Text = "" Then
                MsgBox "CLIK DULU COMBO SPV", vbInformation + vbOKOnly, "PESAN"
                Exit Sub
               End If
                    STRSQL = "UPDATE usertbl SET dilockoleh='Clear by:"
                    STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                    STRSQL = STRSQL + " lockdarispv=NULL,lock_entry_lpd=NULL,fromaccount=NULL,lockmarkup=NULL,lockdarispvbuattl=NULL WHERE SPVCODE='" + Combo1(0).Text + "'"
                
     Else
            If SSOption1(1).Value = True Then
                If Combo1(0).Text = "" Then
                    MsgBox "CLIK DULU COMBO NYA", vbInformation + vbOKOnly, "PESAN"
                Exit Sub
                End If
                STRSQL = "UPDATE usertbl SET dilockoleh='Clear by:"
                STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + " lockdarispv=NULL,lock_entry_lpd=NULL,lockmarkup=NULL ,fromaccount=NULL,lockdarispvbuattl=NULL WHERE userid='" + Combo1(0).Text + "'"
            Else
                    STRSQL = "UPDATE usertbl SET lockdarispv=NULL,lock_entry_lpd=NULL,lockmarkup=NULL,dilockoleh='Clear by:"
                    STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "'"
            End If
            
    End If
    If STRSQL <> "" Then
        M_OBJCONN.Execute STRSQL
        MsgBox "Reset Done.....!"
        End If
        
Case 2
        Unload Me

Case 3
        STRSQL1 = "UPDATE  tblshut SET nshut=1 "
        M_OBJCONN.Execute STRSQL1
End Select

End Sub
Sub ceksts()
If Check1(0).Value Then
    StsVl = "VL-"
End If
If Check1(1).Value Then
          
                StsOP = "OS-"
            
End If
If Check1(2).Value Then
            StsPTP = "PTP"
End If
 If Check1(3).Value Then
            StsBP = "BP-"
End If
       If Check1(4).Value Then
            StsPOP = "POP"
 End If
If Check1(5).Value Then
            StsSP = "SP-"
End If

If Check1(7).Value Then
            StsRP = "RP-"
End If

If Check1(6).Value Then
            Stsblank = "anto"
End If

If Check1(9).Value Then
            StsPR = "PR-"
End If

If Check1(10).Value Then
            StsON = "ON-"
End If

If Check1(11).Value Then
            StsSK = "SK-"
End If

BlokEntry = ""
bCheckNewentry = False
bCheckReguler = False
bCheckSwap = False
bCheckCurrent = False

If chknewentry.Value = vbChecked Then
    bCheckNewentry = True
End If


If chkreguler.Value = vbChecked Then
   bCheckReguler = True
End If

If chkswap.Value = vbChecked Then
   bCheckSwap = True
End If


If chkcurrent.Value = vbChecked Then
   bCheckCurrent = True
End If






If bCheckSwap = True And bCheckNewentry = True And bCheckReguler = True And bCheckCurrent = True Then
    BlokEntry = " (date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + " and date_part(''year'',entry_date)=date_part(''year'',now) or " 'NANTI DITAMBAHIN KURUNG TUTUP
    BlokEntry = BlokEntry + " (date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + " and date_part(''year'',pay_dt_update)=date_part(''year'',now)) or "
    BlokEntry = BlokEntry + " (((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
    BlokEntry = BlokEntry + " and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
    BlokEntry = BlokEntry + " pay_dt_update isnull) and "
    BlokEntry = BlokEntry + " date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)) or "
    BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
    BlokEntry = BlokEntry + " )"
    Exit Sub
ElseIf bCheckNewentry = True And bCheckReguler = True And bCheckCurrent = True Then
    BlokEntry = " (date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + " and date_part(''year'',entry_date)=date_part(''year'',now) or " 'INI SEHARUSNYA DITAMBAHIN KURUNG TUTUP
    BlokEntry = BlokEntry + " (date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + " and date_part(''year'',pay_dt_update)=date_part(''year'',now)) or "
    BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
    BlokEntry = BlokEntry + " )"
    Exit Sub
ElseIf bCheckNewentry = True And bCheckSwap = True And bCheckCurrent = True Then
    BlokEntry = " (date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + " and date_part(''year'',entry_date)=date_part(''year'',now) or "
    BlokEntry = BlokEntry + " (((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
    BlokEntry = BlokEntry + "pay_dt_update isnull) and "
    BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)) or "
    BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
    BlokEntry = BlokEntry + " )"
    Exit Sub
ElseIf bCheckReguler = True And bCheckSwap = True And bCheckCurrent = True Then
   BlokEntry = " (date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
   BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)=date_part(''year'',now) or "
   BlokEntry = BlokEntry + " (((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
   BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
   BlokEntry = BlokEntry + "pay_dt_update isnull) and "
   BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
   BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)) or "
   BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
   BlokEntry = BlokEntry + " )"
   Exit Sub
End If




If bCheckNewentry = True Then
    BlokEntry = " date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + "and date_part(''year'',entry_date)=date_part(''year'',now)"
    Exit Sub
End If


If bCheckReguler = True Then
    BlokEntry = " date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)=date_part(''year'',now)"
    Exit Sub
End If


If bCheckSwap = True Then
    BlokEntry = " ((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
    BlokEntry = BlokEntry + "pay_dt_update isnull) and "
    BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)"
    Exit Sub
End If

If bCheckCurrent = True Then
   BlokEntry = " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
    Exit Sub
End If

'@@ 140710 Tambahan buat blok entry

'If OptNewEntry.Value = True Then
'    BlokEntry = " datediff(''month'',entry_date,now) "
'    BlokEntry = BlokEntry + CmbNewEntry.Text
'End If
'
'If OptReguler.Value = True Then
'    BlokEntry = " datediff(''month'',pay_dt,now) "
'    BlokEntry = BlokEntry + CmbReguler.Text
'End If
'
'If OptSwap.Value = True Then
'    BlokEntry = " datediff(''month'',pay_dt,now) "
'    BlokEntry = BlokEntry + CmbSwap.Text
'End If

'@@ 150710 Ubah blok entry
'If OptNewEntry.Value = True Then
'    BlokEntry = " date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + "and date_part(''year'',entry_date)=date_part(''year'',now)"
'End If
'
'If OptReguler.Value = True Then
'    BlokEntry = " date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)=date_part(''year'',now)"
'End If
'
'If OptSwap.Value = True Then
'    BlokEntry = " ((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + "pay_dt_update isnull) and "
'    BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)"
'End If

End Sub
Public Function getblock() As String


                    STRINGBLOK = ""
                    
                    If StsVl <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                            STRINGBLOK = " substring(F_cek_new,1,3) in (@" + StsVl + "@"
                        Else
                            STRINGBLOK = STRINGBLOK + ",@" + StsVl + "@"
                        End If
                    End If
                    
                    If StsPR <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPR + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPR + "@"
                        End If
                    End If
                    
                    If StsPTP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPTP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPTP + "@"
                        End If
                    End If
                    
                    If StsPOP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPOP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPOP + "@"
                        End If
                    End If
                    
                    If StsBP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsBP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsBP + "@"
                        End If
                    End If
                    
                    If StsSP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)  in (@" + StsSP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsSP + "@"
                        End If
                    End If
                    
                    If StsRP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsRP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsRP + "@"
                        End If
                    End If
                    
                    If StsSK <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsSK + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsSK + "@"
                        End If
                    End If
                    
                     If StsON <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsON + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsON + "@"
                        End If
                    End If
                    
                     If StsOP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsOP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsOP + "@"
                        End If
                    End If
                    
                    
                     If Stsblank <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_NEW,1,3)   in (@@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@@"
                        End If
                    End If
                    
                    
                    
                
                    If Len(STRINGBLOK) > 0 Then
                            STRINGBLOK = STRINGBLOK + ")"
                    End If
                    getblock = STRINGBLOK
End Function
Private Sub SSOption1_Click(Index As Integer, Value As Integer)
Dim M_objrs As ADODB.Recordset
Select Case Index
Case 0
        Combo1(0).Enabled = False
        Combo1(1).Enabled = False
Case 1
        Combo1(0).Enabled = True
        Combo1(1).Enabled = True
        Combo1(0).CLEAR
        Combo1(1).CLEAR
        
        '@@221010
'        Dim M_DATA As New CLS_FRMSEARCH
'        Set M_OBJRS = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "")
'            While Not M_OBJRS.EOF
'                Combo1(0).AddItem M_OBJRS("USERID")
'                Combo1(1).AddItem M_OBJRS("AGENT")
'                M_OBJRS.MoveNext
'            Wend
'        Set M_OBJRS = Nothing
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open "select userid,agent from usertbl where usertype='1' order by userid asc", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        While Not M_objrs.EOF
            Combo1(0).AddItem M_objrs("userid")
            Combo1(1).AddItem M_objrs("agent")
            M_objrs.MoveNext
        Wend
        
        'SSOption1(0).Value = True
        spv = False
Case 2
        Combo1(0).Enabled = True
        Combo1(1).Enabled = True
        Combo1(0).CLEAR
        Combo1(1).CLEAR
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Then
            M_objrs.Open "select * from SPVTBL ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        ElseIf UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Or UCase(MDIForm1.txtlevel.Text) = "ADMIN" Then
            M_objrs.Open "select * from SPVTBL", M_OBJCONN, adOpenDynamic, adLockOptimistic
        ElseIf UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        M_objrs.Open "select * from SPVTBL where team='" + MDIForm1.txtusername.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
        End If
            While Not M_objrs.EOF
                Combo1(0).AddItem M_objrs("SPVCODE")
                Combo1(1).AddItem M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
'        SSOption1(0).Value = True
'        SSOption1(1).Value = True
        
End Select

End Sub
Public Sub getMarkup()
Dim list As listitem
Dim RSNEW As New ADODB.Recordset
Set Rs = New ADODB.Recordset
RSNEW.CursorLocation = adUseClient
RSNEW.Open "select distinct(exclude) from mgm WHERE (exclude <>'')", M_OBJCONN, adOpenDynamic, adLockOptimistic
ListView1.ListItems.CLEAR
While Not RSNEW.EOF
Set list = ListView1.ListItems.ADD(, , IIf(IsNull(RSNEW!exclude), "", RSNEW!exclude))
    RSNEW.MoveNext
Wend

End Sub

Public Function GETSELECTMARKUP() As String
Dim J As Integer
Dim TMPSELECTMARKUP As String
GETSELECTMARKUP = ""
For J = 1 To ListView2.ListItems.Count
        If J = 1 Then
            TMPSELECTMARKUP = TMPSELECTMARKUP + Chr(39) + ListView2.ListItems(J).Text + Chr(39)
        Else
            TMPSELECTMARKUP = TMPSELECTMARKUP + "," + Chr(39) + ListView2.ListItems(J).Text + Chr(39)
        End If
    
        
Next J
GETSELECTMARKUP = TMPSELECTMARKUP
End Function

'@@ 18-11-10 ini buat bikin blok lpd (new entry,swap,dll) yang baru
Public Function GetBlockEntry() As String


                    StringBlokEntry = ""
                    
                    If bCheckNewentry = True Then
                        If Len(StringBlokEntry) = 0 Then
                            StringBlokEntry = "( "
                            StringBlokEntry = StringBlokEntry + "date_part(''month'',entry_date) "
                            StringBlokEntry = StringBlokEntry + " between (date_part(''month'',now())- 2 ) "
                            StringBlokEntry = StringBlokEntry + " and date_part(''month'',now())-1 "
                            StringBlokEntry = StringBlokEntry + " and date_part('year',entry_date)=date_part('year',now())"
                            StringBlokEntry = StringBlokEntry + " )"
                        Else
                            STRINGBLOK = STRINGBLOK + ",@" + StsVl + "@"
                        End If
                    End If
                    
                    If StsPR <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPR + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPR + "@"
                        End If
                    End If
                    
                    If StsPTP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPTP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPTP + "@"
                        End If
                    End If
                    
                    If StsPOP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPOP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPOP + "@"
                        End If
                    End If
                    
                    If StsBP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsBP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsBP + "@"
                        End If
                    End If
                    
                    If StsSP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)  in (@" + StsSP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsSP + "@"
                        End If
                    End If
                    
                    If StsRP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsRP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsRP + "@"
                        End If
                    End If
                    
                    If StsSK <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsSK + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsSK + "@"
                        End If
                    End If
                    
                     If StsON <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsON + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsON + "@"
                        End If
                    End If
                    
                     If StsOP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsOP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsOP + "@"
                        End If
                    End If
                    
                    
                     If Stsblank <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_NEW,1,3)   in (@@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@@"
                        End If
                    End If
                    
                    
                    
                
                    If Len(STRINGBLOK) > 0 Then
                            STRINGBLOK = STRINGBLOK + ")"
                    End If
                    GetBlockEntry = STRINGBLOK
End Function
