VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form Form_Menu_Role1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menu Role"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8925
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Editing Menu Role"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton save 
         Caption         =   "Save"
         Height          =   375
         Left            =   4440
         TabIndex        =   51
         Top             =   360
         Width           =   1815
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   11033
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "Master"
         TabPicture(0)   =   "Form_Menu_Role.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Line1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "listsid"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "listdatacomplaint"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "listaccountlunas"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "managedistribusiaccount"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "blokaplikasitins"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "listreportproblemtelepon"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "listreportproblemheadset"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "resetpassword"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "listrequestptp"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "akseslayanantelkom"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "listunvalidnumber"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "sendsmsblastviaexcel"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "approvedandrejectedsms"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "listsmsscript"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "blastsmstext"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "verifysms"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "reportsmsnew"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "paymentpattern"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "formlistconfidence"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "approvalrequestadditionalphone"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "viewlistrequestform"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "cekaccountstatusprogress"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "ubahstatusaccount"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "settargetfromspv"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "scheduleblockdata"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "blacklistnotelpon"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "telecollection"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "UploadForLockAccount"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "swapdata"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "uploaddata"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).ControlCount=   31
         TabCaption(1)   =   "Data Confidence"
         TabPicture(1)   =   "Form_Menu_Role.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "confidencelist"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "monthlyptppayment"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "monthlycpa"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "monthlybp"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Tools"
         TabPicture(2)   =   "Form_Menu_Role.frx":0038
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "filterhidesystem"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "copyfilecpadandokumenpendukung"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "callmonitor"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "deskcollperformancereguler"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "deleterestoremarks"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "averageperformance"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "deskcollperformance"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "reporttempagent"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "uploaddatafreshwo"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "addspecialhistory"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "transferdata"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "aoc"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "listphonereview"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).ControlCount=   13
         Begin Threed.SSCheck uploaddata 
            Height          =   255
            Left            =   -74520
            TabIndex        =   4
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Upload Data"
         End
         Begin Threed.SSCheck swapdata 
            Height          =   255
            Left            =   -74040
            TabIndex        =   5
            Top             =   750
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Swap data"
         End
         Begin Threed.SSCheck UploadForLockAccount 
            Height          =   255
            Left            =   -74520
            TabIndex        =   6
            Top             =   1005
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Upload For Lock Account"
         End
         Begin Threed.SSCheck telecollection 
            Height          =   255
            Left            =   -74520
            TabIndex        =   7
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "TeleCollection"
         End
         Begin Threed.SSCheck blacklistnotelpon 
            Height          =   255
            Left            =   -74520
            TabIndex        =   8
            Top             =   1605
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Black List No Telpon"
         End
         Begin Threed.SSCheck scheduleblockdata 
            Height          =   255
            Left            =   -74520
            TabIndex        =   9
            Top             =   1890
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Schedule Blok Data"
         End
         Begin Threed.SSCheck settargetfromspv 
            Height          =   255
            Left            =   -74520
            TabIndex        =   10
            Top             =   2160
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Set Target From SPV"
         End
         Begin Threed.SSCheck ubahstatusaccount 
            Height          =   255
            Left            =   -74520
            TabIndex        =   11
            Top             =   2475
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Ubah Status Account"
         End
         Begin Threed.SSCheck cekaccountstatusprogress 
            Height          =   255
            Left            =   -74520
            TabIndex        =   12
            Top             =   2760
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Cek Account Status Progress"
         End
         Begin Threed.SSCheck viewlistrequestform 
            Height          =   255
            Left            =   -74520
            TabIndex        =   13
            Top             =   3045
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Viewer List Request Form"
         End
         Begin Threed.SSCheck approvalrequestadditionalphone 
            Height          =   255
            Left            =   -74520
            TabIndex        =   14
            Top             =   3360
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Approval Request Additional Phone"
         End
         Begin Threed.SSCheck formlistconfidence 
            Height          =   255
            Left            =   -74520
            TabIndex        =   15
            Top             =   3675
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Form List Confidence"
         End
         Begin Threed.SSCheck paymentpattern 
            Height          =   255
            Left            =   -74520
            TabIndex        =   16
            Top             =   3960
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Payment Pattern"
         End
         Begin Threed.SSCheck reportsmsnew 
            Height          =   255
            Left            =   -74520
            TabIndex        =   17
            Top             =   4245
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Report SMS NEW"
         End
         Begin Threed.SSCheck verifysms 
            Height          =   255
            Left            =   -74520
            TabIndex        =   18
            Top             =   4515
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Verify SMS"
         End
         Begin Threed.SSCheck blastsmstext 
            Height          =   255
            Left            =   -74520
            TabIndex        =   19
            Top             =   4830
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Blast SMS Text"
         End
         Begin Threed.SSCheck listsmsscript 
            Height          =   255
            Left            =   -74520
            TabIndex        =   20
            Top             =   5115
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "List sms script"
         End
         Begin Threed.SSCheck approvedandrejectedsms 
            Height          =   255
            Left            =   -74520
            TabIndex        =   21
            Top             =   5400
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Approved and rejected sms"
         End
         Begin Threed.SSCheck sendsmsblastviaexcel 
            Height          =   255
            Left            =   -70680
            TabIndex        =   22
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Send SMS Blast Via Excel"
         End
         Begin Threed.SSCheck listunvalidnumber 
            Height          =   255
            Left            =   -70680
            TabIndex        =   23
            Top             =   795
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "List Unvalid Number"
         End
         Begin Threed.SSCheck akseslayanantelkom 
            Height          =   255
            Left            =   -70680
            TabIndex        =   24
            Top             =   1110
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Akses Layanan Telkom"
         End
         Begin Threed.SSCheck listrequestptp 
            Height          =   255
            Left            =   -70680
            TabIndex        =   25
            Top             =   1395
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "List Request PTP"
         End
         Begin Threed.SSCheck resetpassword 
            Height          =   255
            Left            =   -70680
            TabIndex        =   26
            Top             =   1680
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Reset Password"
         End
         Begin Threed.SSCheck listreportproblemheadset 
            Height          =   255
            Left            =   -70680
            TabIndex        =   27
            Top             =   1950
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "List Report Problem Headset"
         End
         Begin Threed.SSCheck listreportproblemtelepon 
            Height          =   255
            Left            =   -70680
            TabIndex        =   28
            Top             =   2265
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "List Report Problem Telepon"
         End
         Begin Threed.SSCheck blokaplikasitins 
            Height          =   255
            Left            =   -70680
            TabIndex        =   29
            Top             =   2550
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "BLok Aplikasi TINS"
         End
         Begin Threed.SSCheck managedistribusiaccount 
            Height          =   255
            Left            =   -70680
            TabIndex        =   30
            Top             =   2835
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Manage Distribusi Account"
         End
         Begin Threed.SSCheck listaccountlunas 
            Height          =   255
            Left            =   -70680
            TabIndex        =   31
            Top             =   3120
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "List Account Lunas"
         End
         Begin Threed.SSCheck listdatacomplaint 
            Height          =   255
            Left            =   -70680
            TabIndex        =   32
            Top             =   3405
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "List Data Complaint"
         End
         Begin Threed.SSCheck listsid 
            Height          =   255
            Left            =   -70680
            TabIndex        =   33
            Top             =   3690
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "List SID"
         End
         Begin Threed.SSCheck monthlybp 
            Height          =   255
            Left            =   -74520
            TabIndex        =   34
            Top             =   600
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Monthly BP (Broken Promise)"
         End
         Begin Threed.SSCheck monthlycpa 
            Height          =   255
            Left            =   -74520
            TabIndex        =   35
            Top             =   885
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Monthly CPA"
         End
         Begin Threed.SSCheck monthlyptppayment 
            Height          =   255
            Left            =   -74520
            TabIndex        =   36
            Top             =   1170
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Monthly PTP - Payment"
         End
         Begin Threed.SSCheck confidencelist 
            Height          =   255
            Left            =   -74520
            TabIndex        =   37
            Top             =   1455
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Confidence List"
         End
         Begin Threed.SSCheck listphonereview 
            Height          =   255
            Left            =   480
            TabIndex        =   38
            Top             =   600
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "List Phone Review"
         End
         Begin Threed.SSCheck aoc 
            Height          =   255
            Left            =   480
            TabIndex        =   39
            Top             =   915
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "AOC"
         End
         Begin Threed.SSCheck transferdata 
            Height          =   255
            Left            =   480
            TabIndex        =   40
            Top             =   1200
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Transfer Data"
         End
         Begin Threed.SSCheck addspecialhistory 
            Height          =   255
            Left            =   480
            TabIndex        =   41
            Top             =   1485
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Add Special History"
         End
         Begin Threed.SSCheck uploaddatafreshwo 
            Height          =   255
            Left            =   480
            TabIndex        =   42
            Top             =   1755
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Upload Data Fresh WO"
         End
         Begin Threed.SSCheck reporttempagent 
            Height          =   255
            Left            =   480
            TabIndex        =   43
            Top             =   2070
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Report Temp Agent"
         End
         Begin Threed.SSCheck deskcollperformance 
            Height          =   255
            Left            =   480
            TabIndex        =   44
            Top             =   2355
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "DeskColl Performance"
         End
         Begin Threed.SSCheck averageperformance 
            Height          =   255
            Left            =   480
            TabIndex        =   45
            Top             =   2640
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Average Performance"
         End
         Begin Threed.SSCheck deleterestoremarks 
            Height          =   255
            Left            =   480
            TabIndex        =   46
            Top             =   2910
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Delete Restore Marks"
         End
         Begin Threed.SSCheck deskcollperformancereguler 
            Height          =   255
            Left            =   480
            TabIndex        =   47
            Top             =   3195
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "DeskColl Performance Reguler"
         End
         Begin Threed.SSCheck callmonitor 
            Height          =   255
            Left            =   480
            TabIndex        =   48
            Top             =   3480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Call Monitor"
         End
         Begin Threed.SSCheck copyfilecpadandokumenpendukung 
            Height          =   255
            Left            =   480
            TabIndex        =   49
            Top             =   3750
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "Copy File CPA dan Dokumen Pendukung"
         End
         Begin Threed.SSCheck filterhidesystem 
            Height          =   255
            Left            =   480
            TabIndex        =   50
            Top             =   4035
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            _Version        =   196610
            Caption         =   "FIlter Hide System"
         End
         Begin VB.Line Line1 
            X1              =   -71040
            X2              =   -71040
            Y1              =   480
            Y2              =   5760
         End
      End
      Begin VB.ComboBox cbleveluser 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   450
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "Level User :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form_Menu_Role1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UD  As Integer
Dim SD  As Integer
Dim UFLA  As Integer
Dim TC  As Integer
Dim BLNT  As Integer
Dim SBD  As Integer
Dim STFS  As Integer
Dim USA  As Integer
Dim CASP  As Integer
Dim VLRF  As Integer
Dim ARAP  As Integer
Dim FLC  As Integer
Dim PP  As Integer
Dim RSN  As Integer
Dim VS  As Integer
Dim BST  As Integer
Dim LSS  As Integer
Dim AARS  As Integer
Dim SSBVE  As Integer
Dim LUN  As Integer
Dim ALT  As Integer
Dim LRP  As Integer
Dim RP  As Integer
Dim LRPH  As Integer
Dim LRPT  As Integer
Dim BAT  As Integer
Dim MDA  As Integer
Dim LAL  As Integer
Dim LDC  As Integer
Dim LS  As Integer
Dim MBP  As Integer
Dim MCPA  As Integer
Dim MPP  As Integer
Dim CL  As Integer
Dim LPR  As Integer
Dim noaoc  As Integer
Dim TD  As Integer
Dim ASH  As Integer
Dim UDFW  As Integer
Dim RTA  As Integer
Dim DP  As Integer
Dim AP  As Integer
Dim DRM  As Integer
Dim DPR  As Integer
Dim CM  As Integer
Dim CFCDDP  As Integer
Dim FHS  As Integer
Dim tingkat As String
Dim query As String
Dim rs_lvtian As New ADODB.Recordset
Private Sub cbleveluser_Click()
    Call cleario
    Call getio
End Sub

Private Sub cbleveluser_DropDown()
    cbleveluser.CLEAR
    sQuery = "SELECT tingkat FROM menurole order by tingkat"
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If RS_Lv.RecordCount > 0 Then
        While Not RS_Lv.EOF
            cbleveluser.AddItem RS_Lv!tingkat
            RS_Lv.MoveNext
        Wend
    End If
End Sub

Public Sub simpancheck()
    Call io
    Call queryio
    Call cleario
End Sub

Public Sub io()
    If uploaddata.Value = -1 Then
                UD = 1
            Else
                UD = 0
    End If
    If swapdata.Value = -1 Then
                SD = 1
            Else
                SD = 0
    End If
    If UploadForLockAccount.Value = -1 Then
                UFLA = 1
            Else
                UFLA = 0
    End If
    If telecollection.Value = -1 Then
                TC = 1
            Else
                TC = 0
    End If
    If blacklistnotelpon.Value = -1 Then
                BLNT = 1
            Else
                BLNT = 0
    End If
    If scheduleblockdata.Value = -1 Then
                SBD = 1
            Else
                SBD = 0
    End If
    If settargetfromspv.Value = -1 Then
                STFS = 1
            Else
                STFS = 0
    End If
    If ubahstatusaccount.Value = -1 Then
                USA = 1
            Else
                USA = 0
    End If
    If cekaccountstatusprogress.Value = -1 Then
                CASP = 1
            Else
                CASP = 0
    End If
    If viewlistrequestform.Value = -1 Then
                VLRF = 1
            Else
                VLRF = 0
    End If
    If approvalrequestadditionalphone.Value = -1 Then
                ARAP = 1
            Else
                ARAP = 0
    End If
    If formlistconfidence.Value = -1 Then
                FLC = 1
            Else
                FLC = 0
    End If
    If paymentpattern.Value = -1 Then
                PP = 1
            Else
                PP = 0
    End If
    If reportsmsnew.Value = -1 Then
                RSN = 1
            Else
                RSN = 0
    End If
    If verifysms.Value = -1 Then
                RSN = 1
            Else
                RSN = 0
    End If
    If blastsmstext.Value = -1 Then
                BST = 1
            Else
                BST = 0
    End If
    If listsmsscript.Value = -1 Then
                LSS = 1
            Else
                LSS = 0
    End If
    If approvedandrejectedsms.Value = -1 Then
                AARS = 1
            Else
                AARS = 0
    End If
    If sendsmsblastviaexcel.Value = -1 Then
                SSBVE = 1
            Else
                SSBVE = 0
    End If
    If listunvalidnumber.Value = -1 Then
                LUN = 1
            Else
                LUN = 0
    End If
    If akseslayanantelkom.Value = -1 Then
                ALT = 1
            Else
                ALT = 0
    End If
    If listrequestptp.Value = -1 Then
                LRP = 1
            Else
                LRP = 0
    End If
    If resetpassword.Value = -1 Then
                RP = 1
            Else
                RP = 0
    End If
    If listreportproblemheadset.Value = -1 Then
                LRPH = 1
            Else
                LRPH = 0
    End If
    If listreportproblemtelepon.Value = -1 Then
                LRPT = 1
            Else
                LRPT = 0
    End If
    If blokaplikasitins.Value = -1 Then
                BAT = 1
            Else
                BAT = 0
    End If
    If managedistribusiaccount.Value = -1 Then
                MDA = 1
            Else
                MDA = 0
    End If
    If listaccountlunas.Value = -1 Then
                LAL = 1
            Else
                LAL = 0
    End If
    If listdatacomplaint.Value = -1 Then
                LDC = 1
            Else
                LDC = 0
    End If
    If listsid.Value = -1 Then
                LS = 1
            Else
                LS = 0
    End If
    If monthlybp.Value = -1 Then
                MBP = 1
            Else
                MBP = 0
    End If
    If monthlycpa.Value = -1 Then
                MCPA = 1
            Else
                MCPA = 0
    End If
    If monthlyptppayment = -1 Then
                MPP = 1
            Else
                MPP = 0
    End If
    If confidencelist.Value = -1 Then
                CL = 1
            Else
                CL = 0
    End If
    If listphonereview.Value = -1 Then
                LPR = 1
            Else
                LPR = 0
    End If
    If aoc.Value = -1 Then
                noaoc = 1
            Else
                noaoc = 0
    End If
    If transferdata.Value = -1 Then
                TD = 1
            Else
                TD = 0
    End If
    If addspecialhistory.Value = -1 Then
                ASH = 1
            Else
                ASH = 0
    End If
    If uploaddatafreshwo.Value = -1 Then
                UDFW = 1
            Else
                UDFW = 0
    End If
    If reporttempagent.Value = -1 Then
                RTA = 1
            Else
                RTA = 0
    End If
    If deskcollperformance.Value = -1 Then
                DP = 1
            Else
                DP = 0
    End If
    If averageperformance.Value = -1 Then
                AP = 1
            Else
                AP = 0
    End If
    If deleterestoremarks.Value = -1 Then
                DRM = 1
            Else
                DRM = 0
    End If
    If deskcollperformancereguler.Value = -1 Then
                DPR = 1
            Else
                DPR = 0
    End If
    If callmonitor.Value = -1 Then
                CM = 1
            Else
                CM = 0
    End If
    If copyfilecpadandokumenpendukung.Value = -1 Then
                CFCDDP = 1
            Else
                CFCDDP = 0
    End If
    If filterhidesystem.Value = -1 Then
                FHS = 1
            Else
                FHS = 0
    End If
End Sub

Public Sub queryio()
    tingkat = cbleveluser.Text
    query = "select * from checkmenurole where tingkat = '" + tingkat + "'"
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If RS_Lv.RecordCount > 0 Then
        Call updateio
        NN = update
    Else
        Call insertio
        NN = Insert
    End If
    M_OBJCONN.Execute query
    MsgBox " Data Berhasil Disave/update"
End Sub

Public Sub updateio()
        query = "update checkmenurole set"
        query = query + " UD  = " + CStr(UD) + ","
        query = query + " SD  = " + CStr(SD) + ","
        query = query + " UFLA = " + CStr(UFLA) + ", "
        query = query + " TC  = " + CStr(TC) + ","
        query = query + " BLNT  = " + CStr(BLNT) + ","
        query = query + " SBD  = " + CStr(SBD) + ","
        query = query + " STFS  = " + CStr(STFS) + ", "
        query = query + " USA  = " + CStr(USA) + ","
        query = query + " CASP  = " + CStr(CASP) + ","
        query = query + " VLRF  = " + CStr(VLRF) + ","
        query = query + " ARAP  = " + CStr(ARAP) + ","
        query = query + " FLC  = " + CStr(FLC) + ","
        query = query + " PP  = " + CStr(PP) + ","
        query = query + " RSN  = " + CStr(RSN) + ","
        query = query + " VS  = " + CStr(VS) + ","
        query = query + " BST  = " + CStr(BST) + ","
        query = query + " LSS  = " + CStr(LSS) + ","
        query = query + " AARS  = " + CStr(AARS) + ","
        query = query + " SSBVE = " + CStr(SSBVE) + ", "
        query = query + " LUN  = " + CStr(LUN) + ","
        query = query + " ALT  = " + CStr(ALT) + ","
        query = query + " LRP  = " + CStr(LRP) + ","
        query = query + " RP  = " + CStr(RP) + ","
        query = query + " LRPH  = " + CStr(LRPH) + ","
        query = query + " LRPT  = " + CStr(LRPT) + ","
        query = query + " BAT  = " + CStr(BAT) + ","
        query = query + " MDA = " + CStr(MDA) + " ,"
        query = query + " LAL = " + CStr(LAL) + " ,"
        query = query + " LDC = " + CStr(LDC) + " ,"
        query = query + " LS  = " + CStr(LS) + ","
        query = query + " MBP = " + CStr(MBP) + ", "
        query = query + " MCPA = " + CStr(MCPA) + ", "
        query = query + " MPP = " + CStr(MPP) + " ,"
        query = query + " CL  = " + CStr(CL) + " ,"
        query = query + " LPR = " + CStr(LPR) + " ,"
        query = query + " AOC = " + CStr(noaoc) + ", "
        query = query + " TD  = " + CStr(TD) + ","
        query = query + " ASH  = " + CStr(ASH) + ","
        query = query + " UDFW = " + CStr(UDFW) + ", "
        query = query + " RTA  = " + CStr(RTA) + ","
        query = query + " DP  = " + CStr(DP) + ","
        query = query + " AP  = " + CStr(AP) + ","
        query = query + " DRM  = " + CStr(DRM) + ","
        query = query + " DPR  = " + CStr(DPR) + ","
        query = query + " CM = " + CStr(CM) + " ,"
        query = query + " CFCDDP = " + CStr(CFCDDP) + ", "
        query = query + " FHS  = " + CStr(FHS) + ""
        query = query + " where tingkat = '" + tingkat + "'"
End Sub

Public Sub insertio()
    query = " insert into checkmenurole values (" + CStr(UD) + "," + CStr(SD) + "," + CStr(UFLA) + "," + CStr(TC) + "," + CStr(BLNT) + "," + CStr(SBD) + "," + CStr(STFS) + "," + CStr(USA) + "," + CStr(CASP) + "," + CStr(VLRF) + "," + CStr(ARAP) + "," + CStr(FLC) + "," + CStr(PP) + "," + CStr(RSN) + "," + CStr(VS) + "," + CStr(BST) + "," + CStr(LSS) + "," + CStr(AARS) + "," + CStr(SSBVE) + "," + CStr(LUN) + "," + CStr(ALT) + "," + CStr(LRP) + "," + CStr(RP) + "," + CStr(LRPH) + "," + CStr(LRPT) + "," + CStr(BAT) + "," + CStr(MDA) + "," + CStr(LAL) + "," + CStr(LDC) + "," + CStr(LS) + "," + CStr(MBP) + "," + CStr(MCPA) + "," + CStr(MPP) + "," + CStr(CL) + "," + CStr(LPR) + "," + CStr(aoc) + "," + CStr(TD) + "," + CStr(ASH) + "," + CStr(UDFW) + "," + CStr(RTA) + "," + CStr(DP) + "," + CStr(AP) + "," + CStr(DRM) + "," + CStr(DPR) + "," + CStr(CM) + "," + CStr(CFCDDP) + "," + CStr(FHS) + ", '" + tingkat + "')"
End Sub

Private Sub save_Click()
    If cbleveluser.Text = "" Then
        MsgBox "Harus Pilih Level User"
        Exit Sub
    End If
    Call simpancheck
End Sub

Private Sub getio()
    query = "Select * from checkmenurole where tingkat = '" + cbleveluser.Text + "'"
    Set rs_lvtian = New ADODB.Recordset
    rs_lvtian.CursorLocation = adUseClient
    rs_lvtian.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If rs_lvtian.RecordCount > 0 Then
        Call checkio
    End If
End Sub

Private Sub checkio()
    If rs_lvtian!UD = 1 Then
        uploaddata.Value = ssCBChecked
    Else
        uploaddata.Value = ssCBUnchecked
    End If
    If rs_lvtian!SD = 1 Then
        swapdata.Value = ssCBChecked
    Else
        swapdata.Value = ssCBUnchecked
    End If
    If rs_lvtian!UFLA = 1 Then
        UploadForLockAccount.Value = ssCBChecked
    Else
        UploadForLockAccount.Value = ssCBUnchecked
    End If
    If rs_lvtian!TC = 1 Then
        telecollection.Value = ssCBChecked
    Else
        telecollection.Value = ssCBUnchecked
    End If
    If rs_lvtian!BLNT = 1 Then
        blacklistnotelpon.Value = ssCBChecked
    Else
        blacklistnotelpon.Value = ssCBUnchecked
    End If
    If rs_lvtian!SBD = 1 Then
        scheduleblockdata.Value = ssCBChecked
    Else
        scheduleblockdata.Value = ssCBUnchecked
    End If
    If rs_lvtian!STFS = 1 Then
        settargetfromspv.Value = ssCBChecked
    Else
        settargetfromspv.Value = ssCBUnchecked
    End If
    If rs_lvtian!USA = 1 Then
        ubahstatusaccount.Value = ssCBChecked
    Else
        ubahstatusaccount.Value = ssCBUnchecked
    End If
    If rs_lvtian!CASP = 1 Then
        cekaccountstatusprogress.Value = ssCBChecked
    Else
        cekaccountstatusprogress.Value = ssCBUnchecked
    End If
    If rs_lvtian!VLRF = 1 Then
        viewlistrequestform.Value = ssCBChecked
    Else
        viewlistrequestform.Value = ssCBUnchecked
    End If
    If rs_lvtian!ARAP = 1 Then
        approvalrequestadditionalphone.Value = ssCBChecked
    Else
        approvalrequestadditionalphone.Value = ssCBUnchecked
    End If
    If rs_lvtian!FLC = 1 Then
        formlistconfidence.Value = ssCBChecked
    Else
        formlistconfidence.Value = ssCBUnchecked
    End If
    If rs_lvtian!PP = 1 Then
        paymentpattern.Value = ssCBChecked
    Else
        paymentpattern.Value = ssCBUnchecked
    End If
    If rs_lvtian!RSN = 1 Then
        reportsmsnew.Value = ssCBChecked
    Else
        reportsmsnew.Value = ssCBUnchecked
    End If
    If rs_lvtian!VS = 1 Then
        verifysms.Value = ssCBChecked
    Else
        verifysms.Value = ssCBUnchecked
    End If
    If rs_lvtian!BST = 1 Then
        blastsmstext.Value = ssCBChecked
    Else
        blastsmstext.Value = ssCBUnchecked
    End If
    If rs_lvtian!LSS = 1 Then
        listsmsscript.Value = ssCBChecked
    Else
        listsmsscript.Value = ssCBUnchecked
    End If
    If rs_lvtian!AARS = 1 Then
        approvedandrejectedsms.Value = ssCBChecked
    Else
        approvedandrejectedsms.Value = ssCBUnchecked
    End If
    If rs_lvtian!SSBVE = 1 Then
        sendsmsblastviaexcel.Value = ssCBChecked
    Else
        sendsmsblastviaexcel.Value = ssCBUnchecked
    End If
    If rs_lvtian!LUN = 1 Then
        listunvalidnumber.Value = ssCBChecked
    Else
        listunvalidnumber.Value = ssCBUnchecked
    End If
    If rs_lvtian!ALT = 1 Then
        akseslayanantelkom.Value = ssCBChecked
    Else
        akseslayanantelkom.Value = ssCBUnchecked
    End If
    If rs_lvtian!LRP = 1 Then
        listrequestptp.Value = ssCBChecked
    Else
        listrequestptp.Value = ssCBUnchecked
    End If
    If rs_lvtian!RP = 1 Then
        resetpassword.Value = ssCBChecked
    Else
        resetpassword.Value = ssCBUnchecked
    End If
    If rs_lvtian!LRPH = 1 Then
        listreportproblemheadset.Value = ssCBChecked
    Else
        listreportproblemheadset.Value = ssCBUnchecked
    End If
    If rs_lvtian!LRPT = 1 Then
        listreportproblemtelepon.Value = ssCBChecked
    Else
        listreportproblemtelepon.Value = ssCBUnchecked
    End If
    If rs_lvtian!BAT = 1 Then
        blokaplikasitins.Value = ssCBChecked
    Else
        blokaplikasitins.Value = ssCBUnchecked
    End If
    If rs_lvtian!MDA = 1 Then
        managedistribusiaccount.Value = ssCBChecked
    Else
        managedistribusiaccount.Value = ssCBUnchecked
    End If
    If rs_lvtian!LAL = 1 Then
        listaccountlunas.Value = ssCBChecked
    Else
        listaccountlunas.Value = ssCBUnchecked
    End If
    If rs_lvtian!LDC = 1 Then
        listdatacomplaint.Value = ssCBChecked
    Else
        listdatacomplaint.Value = ssCBUnchecked
    End If
    If rs_lvtian!LS = 1 Then
        listsid.Value = ssCBChecked
    Else
        listsid.Value = ssCBUnchecked
    End If
    'dataconfidence
    If rs_lvtian!MBP = 1 Then
        monthlybp.Value = ssCBChecked
    Else
        monthlybp.Value = ssCBUnchecked
    End If
    If rs_lvtian!MCPA = 1 Then
        monthlycpa.Value = ssCBChecked
    Else
        monthlycpa.Value = ssCBUnchecked
    End If
    If rs_lvtian!MPP = 1 Then
        monthlyptppayment.Value = ssCBChecked
    Else
        monthlyptppayment.Value = ssCBUnchecked
    End If
    If rs_lvtian!CL = 1 Then
        confidencelist.Value = ssCBChecked
    Else
        confidencelist.Value = ssCBUnchecked
    End If
    If rs_lvtian!LPR = 1 Then
        listphonereview.Value = ssCBChecked
    Else
        listphonereview.Value = ssCBUnchecked
    End If
    If rs_lvtian!aoc = 1 Then
        aoc.Value = ssCBChecked
    Else
        aoc.Value = ssCBUnchecked
    End If
    If rs_lvtian!TD = 1 Then
        transferdata.Value = ssCBChecked
    Else
        transferdata.Value = ssCBUnchecked
    End If
    If rs_lvtian!ASH = 1 Then
        addspecialhistory.Value = ssCBChecked
    Else
        addspecialhistory.Value = ssCBUnchecked
    End If
    If rs_lvtian!UDFW = 1 Then
        uploaddatafreshwo.Value = ssCBChecked
    Else
        uploaddatafreshwo.Value = ssCBUnchecked
    End If
    If rs_lvtian!RTA = 1 Then
        reporttempagent.Value = ssCBChecked
    Else
        reporttempagent.Value = ssCBUnchecked
    End If
    If rs_lvtian!DP = 1 Then
        deskcollperformance.Value = ssCBChecked
    Else
        deskcollperformance.Value = ssCBUnchecked
    End If
    If rs_lvtian!AP = 1 Then
        averageperformance.Value = ssCBChecked
    Else
        averageperformance.Value = ssCBUnchecked
    End If
    If rs_lvtian!DRM = 1 Then
        deleterestoremarks.Value = ssCBChecked
    Else
        deleterestoremarks.Value = ssCBUnchecked
    End If
    If rs_lvtian!DPR = 1 Then
        deskcollperformancereguler.Value = ssCBChecked
    Else
        deskcollperformancereguler.Value = ssCBUnchecked
    End If
    If rs_lvtian!CM = 1 Then
        callmonitor.Value = ssCBChecked
    Else
        callmonitor.Value = ssCBUnchecked
    End If
    If rs_lvtian!CFCDDP = 1 Then
        copyfilecpadandokumenpendukung.Value = ssCBChecked
    Else
        copyfilecpadandokumenpendukung.Value = ssCBUnchecked
    End If
    If rs_lvtian!FHS = 1 Then
        filterhidesystem.Value = ssCBChecked
    Else
        filterhidesystem.Value = ssCBUnchecked
    End If
End Sub

Private Sub cleario()
    uploaddata.Value = ssCBUnchecked
    swapdata.Value = ssCBUnchecked
    UploadForLockAccount.Value = ssCBUnchecked
    telecollection.Value = ssCBUnchecked
    blacklistnotelpon.Value = ssCBUnchecked
    scheduleblockdata.Value = ssCBUnchecked
    settargetfromspv.Value = ssCBUnchecked
    ubahstatusaccount.Value = ssCBUnchecked
    cekaccountstatusprogress.Value = ssCBUnchecked
    viewlistrequestform.Value = ssCBUnchecked
    approvalrequestadditionalphone.Value = ssCBUnchecked
    formlistconfidence.Value = ssCBUnchecked
    paymentpattern.Value = ssCBUnchecked
    reportsmsnew.Value = ssCBUnchecked
    verifysms.Value = ssCBUnchecked
    blastsmstext.Value = ssCBUnchecked
    listsmsscript.Value = ssCBUnchecked
    approvedandrejectedsms.Value = ssCBUnchecked
    sendsmsblastviaexcel.Value = ssCBUnchecked
    listunvalidnumber.Value = ssCBUnchecked
    akseslayanantelkom.Value = ssCBUnchecked
    listrequestptp.Value = ssCBUnchecked
    resetpassword.Value = ssCBUnchecked
    listreportproblemheadset.Value = ssCBUnchecked
    listreportproblemtelepon.Value = ssCBUnchecked
    blokaplikasitins.Value = ssCBUnchecked
    managedistribusiaccount.Value = ssCBUnchecked
    listaccountlunas.Value = ssCBUnchecked
    listdatacomplaint.Value = ssCBUnchecked
    listsid.Value = ssCBUnchecked
    monthlybp.Value = ssCBUnchecked
    monthlycpa.Value = ssCBUnchecked
    monthlyptppayment.Value = ssCBUnchecked
    confidencelist.Value = ssCBUnchecked
    listphonereview.Value = ssCBUnchecked
    aoc.Value = ssCBUnchecked
    transferdata.Value = ssCBUnchecked
    addspecialhistory.Value = ssCBUnchecked
    uploaddatafreshwo.Value = ssCBUnchecked
    reporttempagent.Value = ssCBUnchecked
    deskcollperformance.Value = ssCBUnchecked
    averageperformance.Value = ssCBUnchecked
    deleterestoremarks.Value = ssCBUnchecked
    deskcollperformancereguler.Value = ssCBUnchecked
    callmonitor.Value = ssCBUnchecked
    copyfilecpadandokumenpendukung.Value = ssCBUnchecked
    filterhidesystem.Value = ssCBUnchecked

End Sub
