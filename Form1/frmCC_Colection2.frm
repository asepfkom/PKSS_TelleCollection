VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCC_Colection 
   Caption         =   "Customer Form"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form2"
   Moveable        =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   330
      Width           =   1665
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   135
      TabIndex        =   0
      Top             =   840
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Personal Data"
      TabPicture(0)   =   "frmCC_Colection2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "label1(72)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CustId(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCardNo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "label1(8)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblNoCard"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label11(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "label1(24)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "label1(23)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "label1(21)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "label1(22)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "label1(28)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "label1(27)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "label1(26)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "label1(25)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "label1(73)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "label1(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblNama"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "label1(9)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblNoPay"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label14"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label11(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "label1(3)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label5"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblID"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "label1(4)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label6"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "label1(5)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lblAddr"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label8"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "label1(6)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label27"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lblOfficeAddr"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "label1(7)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lblZIP"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label22"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "label1(10)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "lblPromPA"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label16"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Label18"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "label1(11)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "label1(12)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Label20"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "label1(13)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Label11(0)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "label1(14)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Label25"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "lblBrokenPromised"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "label1(20)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Label11(6)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "label1(19)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Label11(5)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "label1(18)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Label11(4)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "label1(17)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Label11(2)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "lblPayDt"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "lblLastPay"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "lblTtlPay"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "lblAmount"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "lblLcAtm"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "lblLastBill"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "lblOpenDate"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "lblDate"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "lblLimit"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "AOffice2"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "txtOfficeNo2"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "AOffice1"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "txtOfficeNo1"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "AHome1"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "txtHomeNo1"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "AHome2"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "txtHomeNo2"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "lblBD"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "Frame34"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "Frame33"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "Option2"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "Option1"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "Option4"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "Option3"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).ControlCount=   82
      TabCaption(1)   =   "Additional Fields"
      TabPicture(1)   =   "frmCC_Colection2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "label1(63)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "label1(62)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "label1(61)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "label1(60)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "label1(71)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "label1(47)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "label1(49)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "label1(51)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "label1(50)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "label1(70)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "label1(54)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "label1(55)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "label1(53)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "label1(52)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "label1(69)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "label1(56)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "label1(57)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "label1(59)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "label1(58)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "label1(68)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "AHomeAdd2(1)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "AHomeAdd1(0)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtHomeAdd2"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtHomeAdd1"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "AOfficeAdd(3)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "AOfficeAdd(2)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtOfficeAdd2"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtOfficeAdd1"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtMobileAdd2"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "txtMobileAdd1"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "AFaxAdd(5)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "AFaxAdd(4)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "txtFaxAdd2"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "txtFaxAdd1"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).ControlCount=   34
      TabCaption(2)   =   "History"
      TabPicture(2)   =   "frmCC_Colection2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView1(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Results"
      TabPicture(3)   =   "frmCC_Colection2.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "label1(66)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "label1(36)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "label1(35)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "label1(33)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "label1(34)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "label1(67)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "label1(40)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "label1(39)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "label1(38)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "label1(37)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "label1(79)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "label1(0)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "label1(75)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "label1(76)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "label1(77)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "label1(78)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "label1(74)"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "label1(46)"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "label1(45)"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "label1(44)"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "label1(43)"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "label1(42)"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "label1(41)"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "label1(48)"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "txtRemarks"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "cmbTimeSch"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "cmbDateSch"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "TdbPTP"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "txtPayment"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "C_NotContacted"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "cmbDescUn"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).Control(31)=   "cmbUncontacted"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "C_Contacted"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "cmbDescCon"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).Control(34)=   "cmbContacted"
      Tab(3).Control(34).Enabled=   0   'False
      Tab(3).Control(35)=   "C_Payment"
      Tab(3).Control(35).Enabled=   0   'False
      Tab(3).Control(36)=   "cmbDiscount"
      Tab(3).Control(36).Enabled=   0   'False
      Tab(3).Control(37)=   "txtDiscount"
      Tab(3).Control(37).Enabled=   0   'False
      Tab(3).Control(38)=   "txtResultDesc"
      Tab(3).Control(38).Enabled=   0   'False
      Tab(3).Control(39)=   "txtResult"
      Tab(3).Control(39).Enabled=   0   'False
      Tab(3).Control(40)=   "cmbNextAct"
      Tab(3).Control(40).Enabled=   0   'False
      Tab(3).Control(41)=   "cmbPrior"
      Tab(3).Control(41).Enabled=   0   'False
      Tab(3).ControlCount=   42
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   6840
         TabIndex        =   122
         Top             =   5280
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   6855
         TabIndex        =   117
         Top             =   4755
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   3525
         TabIndex        =   112
         Top             =   4785
         Width           =   240
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   3555
         TabIndex        =   107
         Top             =   5145
         Width           =   255
      End
      Begin VB.ComboBox cmbPrior 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCC_Colection2.frx":0070
         Left            =   -73080
         List            =   "frmCC_Colection2.frx":007D
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   5115
         Width           =   1335
      End
      Begin VB.ComboBox cmbNextAct 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73080
         TabIndex        =   80
         Top             =   4395
         Width           =   2295
      End
      Begin VB.TextBox txtResult 
         Height          =   285
         Left            =   -68760
         TabIndex        =   78
         Top             =   3645
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtResultDesc 
         Height          =   285
         Left            =   -68760
         TabIndex        =   77
         Top             =   3285
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtDiscount 
         Height          =   285
         Left            =   -66480
         TabIndex        =   76
         Top             =   3285
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbDiscount 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73080
         TabIndex        =   68
         Top             =   3270
         Width           =   975
      End
      Begin VB.CheckBox C_Payment 
         BackColor       =   &H00C5974B&
         Height          =   255
         Left            =   -74790
         TabIndex        =   66
         Top             =   2910
         Width           =   315
      End
      Begin VB.ComboBox cmbContacted 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73080
         TabIndex        =   61
         Top             =   2115
         Width           =   3060
      End
      Begin VB.ComboBox cmbDescCon 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73080
         TabIndex        =   60
         Top             =   2475
         Width           =   4605
      End
      Begin VB.CheckBox C_Contacted 
         BackColor       =   &H00C5974B&
         Height          =   255
         Left            =   -74850
         TabIndex        =   58
         Top             =   1785
         Width           =   315
      End
      Begin VB.ComboBox cmbUncontacted 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73095
         TabIndex        =   53
         Top             =   960
         Width           =   2340
      End
      Begin VB.ComboBox cmbDescUn 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCC_Colection2.frx":0095
         Left            =   -73095
         List            =   "frmCC_Colection2.frx":0097
         TabIndex        =   52
         Top             =   1320
         Width           =   3285
      End
      Begin VB.CheckBox C_NotContacted 
         BackColor       =   &H00C5974B&
         Height          =   255
         Left            =   -74820
         TabIndex        =   50
         Top             =   615
         Width           =   315
      End
      Begin VB.Frame Frame33 
         BackColor       =   &H00E8BE91&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   7290
         TabIndex        =   5
         Top             =   5070
         Width           =   3150
         Begin VB.OptionButton Option5 
            BackColor       =   &H00E8BE91&
            Height          =   195
            Left            =   2760
            TabIndex        =   6
            Top             =   0
            Width           =   255
         End
         Begin TDBMask6Ctl.TDBMask txtMobileNo2 
            Height          =   495
            Left            =   1680
            TabIndex        =   13
            Top             =   0
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   873
            Caption         =   "frmCC_Colection2.frx":0099
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection2.frx":0105
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   15253137
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&"
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            LookupMode      =   0
            LookupTable     =   ""
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________"
            Value           =   ""
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E8BE91&
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   29
            Left            =   1440
            TabIndex        =   8
            Top             =   0
            Width           =   75
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E8BE91&
            Caption         =   "Mobile Phone II"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   30
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Frame Frame34 
         BackColor       =   &H00E8BE91&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   7290
         TabIndex        =   1
         Top             =   4750
         Width           =   3150
         Begin VB.OptionButton Option6 
            BackColor       =   &H00E8BE91&
            Height          =   195
            Left            =   2760
            TabIndex        =   2
            Top             =   0
            Width           =   375
         End
         Begin TDBMask6Ctl.TDBMask txtMobileNo1 
            Height          =   495
            Left            =   1680
            TabIndex        =   12
            Top             =   0
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   873
            Caption         =   "frmCC_Colection2.frx":0147
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection2.frx":01B3
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   15253137
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&"
            HighlightText   =   0
            IMEMode         =   0
            IMEStatus       =   0
            LookupMode      =   0
            LookupTable     =   ""
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MousePointer    =   0
            MoveOnLRKey     =   0
            OLEDragMode     =   0
            OLEDropMode     =   0
            PromptChar      =   "_"
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "_____________"
            Value           =   ""
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E8BE91&
            Caption         =   "Mobile Phone I"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   31
            Left            =   120
            TabIndex        =   4
            Top             =   0
            Width           =   1260
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E8BE91&
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   32
            Left            =   1440
            TabIndex        =   3
            Top             =   0
            Width           =   75
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5790
         Index           =   1
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   10213
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
         BackColor       =   16436909
         BorderStyle     =   1
         Appearance      =   1
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
      Begin TDBMask6Ctl.TDBMask txtFaxAdd1 
         Height          =   330
         Left            =   -72330
         TabIndex        =   16
         Top             =   2790
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":01F5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":0261
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   16777215
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&&"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "_________"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtFaxAdd2 
         Height          =   330
         Left            =   -72330
         TabIndex        =   17
         Top             =   3150
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":02A3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":030F
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   16777215
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&&"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "_________"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AFaxAdd 
         Height          =   330
         Index           =   4
         Left            =   -73050
         TabIndex        =   18
         Top             =   2790
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":0351
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":03BD
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   1
         AutoConvert     =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   0
         Format          =   "[9999]"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   -1
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[____]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AFaxAdd 
         Height          =   330
         Index           =   5
         Left            =   -73050
         TabIndex        =   19
         Top             =   3150
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":03FF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":046B
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   1
         AutoConvert     =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   0
         Format          =   "[9999]"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   -1
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[____]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtMobileAdd1 
         Height          =   330
         Left            =   -73110
         TabIndex        =   25
         Top             =   3750
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":04AD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":0519
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   16777215
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&&&&&&"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "_____________"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtMobileAdd2 
         Height          =   330
         Left            =   -73110
         TabIndex        =   26
         Top             =   4110
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":055B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":05C7
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   16777215
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&&&&&&"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "_____________"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtOfficeAdd1 
         Height          =   330
         Left            =   -72315
         TabIndex        =   32
         Top             =   1830
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":0609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":0675
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   16777215
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "________"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtOfficeAdd2 
         Height          =   330
         Left            =   -72315
         TabIndex        =   33
         Top             =   2190
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":06B7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":0723
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   16777215
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "________"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AOfficeAdd 
         Height          =   330
         Index           =   2
         Left            =   -73035
         TabIndex        =   34
         Top             =   1830
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":0765
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":07D1
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   1
         AutoConvert     =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   0
         Format          =   "[9999]"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[____]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AOfficeAdd 
         Height          =   330
         Index           =   3
         Left            =   -73035
         TabIndex        =   35
         Top             =   2190
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":0813
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":087F
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   1
         AutoConvert     =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   0
         Format          =   "[9999]"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[____]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtHomeAdd1 
         Height          =   330
         Left            =   -72300
         TabIndex        =   41
         Top             =   825
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":08C1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":092D
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   16777215
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "________"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtHomeAdd2 
         Height          =   330
         Left            =   -72300
         TabIndex        =   42
         Top             =   1185
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":096F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":09DB
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   16777215
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "________"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AHomeAdd1 
         Height          =   330
         Index           =   0
         Left            =   -73020
         TabIndex        =   43
         Top             =   825
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":0A1D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":0A89
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   1
         AutoConvert     =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   0
         Format          =   "[9999]"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[____]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AHomeAdd2 
         Height          =   330
         Index           =   1
         Left            =   -73020
         TabIndex        =   44
         Top             =   1185
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   582
         Caption         =   "frmCC_Colection2.frx":0ACB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":0B37
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   1
         AutoConvert     =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   0
         Format          =   "[9999]"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[____]"
         Value           =   ""
      End
      Begin TDBNumber6Ctl.TDBNumber txtPayment 
         Height          =   345
         Left            =   -73065
         TabIndex        =   69
         Top             =   3630
         Width           =   1965
         _Version        =   65536
         _ExtentX        =   3466
         _ExtentY        =   609
         Calculator      =   "frmCC_Colection2.frx":0B79
         Caption         =   "frmCC_Colection2.frx":0B99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection2.frx":0C05
         Keys            =   "frmCC_Colection2.frx":0C23
         Spin            =   "frmCC_Colection2.frx":0C6D
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999999
         MinValue        =   -99999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBDate6Ctl.TDBDate TdbPTP 
         Height          =   285
         Left            =   -70980
         TabIndex        =   70
         Top             =   3270
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   503
         Calendar        =   "frmCC_Colection2.frx":0C95
         Caption         =   "frmCC_Colection2.frx":0DAD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection2.frx":0E19
         Keys            =   "frmCC_Colection2.frx":0E37
         Spin            =   "frmCC_Colection2.frx":0E95
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   3.54027066542603E-316
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate cmbDateSch 
         Height          =   315
         Left            =   -73080
         TabIndex        =   82
         Top             =   4755
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         Calendar        =   "frmCC_Colection2.frx":0EBD
         Caption         =   "frmCC_Colection2.frx":0FD5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection2.frx":1041
         Keys            =   "frmCC_Colection2.frx":105F
         Spin            =   "frmCC_Colection2.frx":10BD
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
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
         Text            =   "12/09/2006"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   38972
         CenturyMode     =   0
      End
      Begin TDBTime6Ctl.TDBTime cmbTimeSch 
         Height          =   315
         Left            =   -71640
         TabIndex        =   83
         Top             =   4755
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "frmCC_Colection2.frx":10E5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":1151
         Spin            =   "frmCC_Colection2.frx":11A1
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "hh:nn"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "hh:nn"
         HighlightText   =   0
         Hour12Mode      =   1
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxTime         =   0.99999
         MidnightMode    =   0
         MinTime         =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "03:05"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   0.128842592592593
      End
      Begin RichTextLib.RichTextBox txtRemarks 
         Height          =   1095
         Left            =   -70170
         TabIndex        =   90
         Top             =   4800
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   1931
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmCC_Colection2.frx":11C9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TDBDate6Ctl.TDBDate lblBD 
         Height          =   255
         Left            =   8715
         TabIndex        =   105
         Top             =   885
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   450
         Calendar        =   "frmCC_Colection2.frx":1245
         Caption         =   "frmCC_Colection2.frx":135D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection2.frx":13C9
         Keys            =   "frmCC_Colection2.frx":13E7
         Spin            =   "frmCC_Colection2.frx":1445
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15253137
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
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
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   1.07202956713409E-317
         CenturyMode     =   0
      End
      Begin TDBMask6Ctl.TDBMask txtHomeNo2 
         Height          =   300
         Left            =   2355
         TabIndex        =   108
         Top             =   5070
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   529
         Caption         =   "frmCC_Colection2.frx":146D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":14D9
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   15253137
         BorderStyle     =   0
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&&&"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__________"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AHome2 
         Height          =   300
         Left            =   1770
         TabIndex        =   109
         Top             =   5070
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   529
         Caption         =   "frmCC_Colection2.frx":151B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":1587
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   15253137
         BorderStyle     =   0
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[&&&&]"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[____]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtHomeNo1 
         Height          =   300
         Left            =   2340
         TabIndex        =   113
         Top             =   4755
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   529
         Caption         =   "frmCC_Colection2.frx":15C9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":1635
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   15253137
         BorderStyle     =   0
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&&&"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__________"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AHome1 
         Height          =   300
         Left            =   1740
         TabIndex        =   114
         Top             =   4755
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   529
         Caption         =   "frmCC_Colection2.frx":1677
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":16E3
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   15253137
         BorderStyle     =   0
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[&&&&]"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[____]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtOfficeNo1 
         Height          =   315
         Left            =   5775
         TabIndex        =   118
         Top             =   4755
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "frmCC_Colection2.frx":1725
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":1791
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   15253137
         BorderStyle     =   0
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&&&"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   -1
         ShowContextMenu =   0
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__________"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AOffice1 
         Height          =   315
         Left            =   5175
         TabIndex        =   119
         Top             =   4755
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   556
         Caption         =   "frmCC_Colection2.frx":17D3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":183F
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   15253137
         BorderStyle     =   0
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[&&&&]"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[____]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtOfficeNo2 
         Height          =   315
         Left            =   5760
         TabIndex        =   123
         Top             =   5280
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "frmCC_Colection2.frx":1881
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":18ED
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   15253137
         BorderStyle     =   0
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&&&"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__________"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AOffice2 
         Height          =   315
         Left            =   5160
         TabIndex        =   124
         Top             =   5280
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   556
         Caption         =   "frmCC_Colection2.frx":192F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection2.frx":199B
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   15253137
         BorderStyle     =   0
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[&&&&]"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[____]"
         Value           =   ""
      End
      Begin TDBNumber6Ctl.TDBNumber lblLimit 
         Height          =   285
         Left            =   8715
         TabIndex        =   134
         Top             =   1170
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   503
         Calculator      =   "frmCC_Colection2.frx":19DD
         Caption         =   "frmCC_Colection2.frx":19FD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection2.frx":1A69
         Keys            =   "frmCC_Colection2.frx":1A87
         Spin            =   "frmCC_Colection2.frx":1AD1
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15253137
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999999
         MinValue        =   -99999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBDate6Ctl.TDBDate lblDate 
         Height          =   495
         Left            =   1695
         TabIndex        =   139
         Top             =   1830
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   873
         Calendar        =   "frmCC_Colection2.frx":1AF9
         Caption         =   "frmCC_Colection2.frx":1C11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection2.frx":1C7D
         Keys            =   "frmCC_Colection2.frx":1C9B
         Spin            =   "frmCC_Colection2.frx":1CF9
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15253137
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
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
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   3.54031216694028E-316
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate lblOpenDate 
         Height          =   495
         Left            =   5625
         TabIndex        =   154
         Top             =   1725
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   873
         Calendar        =   "frmCC_Colection2.frx":1D21
         Caption         =   "frmCC_Colection2.frx":1E39
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection2.frx":1EA5
         Keys            =   "frmCC_Colection2.frx":1EC3
         Spin            =   "frmCC_Colection2.frx":1F21
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15253137
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
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
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   3.54028054673894E-316
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate lblLastBill 
         Height          =   495
         Left            =   5610
         TabIndex        =   157
         Top             =   2295
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   873
         Calendar        =   "frmCC_Colection2.frx":1F49
         Caption         =   "frmCC_Colection2.frx":2061
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection2.frx":20CD
         Keys            =   "frmCC_Colection2.frx":20EB
         Spin            =   "frmCC_Colection2.frx":2149
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15253137
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
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
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   3.54028845178928E-316
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate lblLcAtm 
         Height          =   495
         Left            =   5595
         TabIndex        =   160
         Top             =   2880
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   873
         Calendar        =   "frmCC_Colection2.frx":2171
         Caption         =   "frmCC_Colection2.frx":2289
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection2.frx":22F5
         Keys            =   "frmCC_Colection2.frx":2313
         Spin            =   "frmCC_Colection2.frx":2371
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15253137
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
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
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   3.54025880785053E-316
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber lblAmount 
         Height          =   345
         Left            =   8880
         TabIndex        =   166
         Top             =   4020
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   609
         Calculator      =   "frmCC_Colection2.frx":2399
         Caption         =   "frmCC_Colection2.frx":23B9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection2.frx":2425
         Keys            =   "frmCC_Colection2.frx":2443
         Spin            =   "frmCC_Colection2.frx":248D
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15253137
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999999
         MinValue        =   -99999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber lblTtlPay 
         Height          =   345
         Left            =   8865
         TabIndex        =   169
         Top             =   3615
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   609
         Calculator      =   "frmCC_Colection2.frx":24B5
         Caption         =   "frmCC_Colection2.frx":24D5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection2.frx":2541
         Keys            =   "frmCC_Colection2.frx":255F
         Spin            =   "frmCC_Colection2.frx":25A9
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15253137
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999999
         MinValue        =   -99999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber lblLastPay 
         Height          =   345
         Left            =   8835
         TabIndex        =   172
         Top             =   3225
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   609
         Calculator      =   "frmCC_Colection2.frx":25D1
         Caption         =   "frmCC_Colection2.frx":25F1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection2.frx":265D
         Keys            =   "frmCC_Colection2.frx":267B
         Spin            =   "frmCC_Colection2.frx":26C5
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15253137
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999999999
         MinValue        =   -99999999999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBDate6Ctl.TDBDate lblPayDt 
         Height          =   495
         Left            =   9075
         TabIndex        =   175
         Top             =   2625
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   873
         Calendar        =   "frmCC_Colection2.frx":26ED
         Caption         =   "frmCC_Colection2.frx":2805
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection2.frx":2871
         Keys            =   "frmCC_Colection2.frx":288F
         Spin            =   "frmCC_Colection2.frx":28ED
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15253137
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
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
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   3.54027066542603E-316
         CenturyMode     =   0
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Pay_dt"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   7755
         TabIndex        =   177
         Top             =   2625
         Width           =   585
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   8955
         TabIndex        =   176
         Top             =   2625
         Width           =   75
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Last Pay"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   7395
         TabIndex        =   174
         Top             =   3225
         Width           =   720
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   8595
         TabIndex        =   173
         Top             =   3225
         Width           =   75
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Ttl_Pay"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   7425
         TabIndex        =   171
         Top             =   3615
         Width           =   630
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   8625
         TabIndex        =   170
         Top             =   3615
         Width           =   75
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Amount_wo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   7440
         TabIndex        =   168
         Top             =   4020
         Width           =   1005
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   8640
         TabIndex        =   167
         Top             =   4020
         Width           =   75
      End
      Begin VB.Label lblBrokenPromised 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5610
         TabIndex        =   165
         Top             =   3495
         Width           =   75
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Broken Promise"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4050
         TabIndex        =   164
         Top             =   3495
         Width           =   1365
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   5490
         TabIndex        =   163
         Top             =   3495
         Width           =   75
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Lc_atmp"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   4035
         TabIndex        =   162
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   5475
         TabIndex        =   161
         Top             =   2880
         Width           =   75
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Last Bill"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4050
         TabIndex        =   159
         Top             =   2295
         Width           =   660
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   5490
         TabIndex        =   158
         Top             =   2295
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   5505
         TabIndex        =   156
         Top             =   1725
         Width           =   75
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Open Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4065
         TabIndex        =   155
         Top             =   1725
         Width           =   915
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "From_PA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4110
         TabIndex        =   153
         Top             =   1455
         Width           =   765
      End
      Begin VB.Label lblPromPA 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5670
         TabIndex        =   152
         Top             =   1455
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   5550
         TabIndex        =   151
         Top             =   1455
         Width           =   75
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Zip"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   150
         Top             =   3150
         Width           =   270
      End
      Begin VB.Label lblZIP 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   149
         Top             =   3150
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   1560
         TabIndex        =   148
         Top             =   3150
         Width           =   75
      End
      Begin VB.Label lblOfficeAddr 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1695
         TabIndex        =   147
         Top             =   2895
         Width           =   75
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Office Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   146
         Top             =   2895
         Width           =   1245
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1575
         TabIndex        =   145
         Top             =   2895
         Width           =   75
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   144
         Top             =   2385
         Width           =   690
      End
      Begin VB.Label lblAddr 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1695
         TabIndex        =   143
         Top             =   2385
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1575
         TabIndex        =   142
         Top             =   2385
         Width           =   75
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "DOB"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   141
         Top             =   1830
         Width           =   390
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1575
         TabIndex        =   140
         Top             =   1830
         Width           =   75
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1710
         TabIndex        =   138
         Top             =   1530
         Width           =   75
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "ID No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   137
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1590
         TabIndex        =   136
         Top             =   1530
         Width           =   75
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Limit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   7395
         TabIndex        =   135
         Top             =   1155
         Width           =   405
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "No Pay"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4125
         TabIndex        =   133
         Top             =   1170
         Width           =   600
      End
      Begin VB.Label lblNoPay 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5685
         TabIndex        =   132
         Top             =   1170
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   5565
         TabIndex        =   131
         Top             =   1170
         Width           =   75
      End
      Begin VB.Label lblNama 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1725
         TabIndex        =   130
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   285
         TabIndex        =   129
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1605
         TabIndex        =   128
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Data Phone Customer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   73
         Left            =   300
         TabIndex        =   127
         Top             =   4500
         Width           =   1890
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   5055
         TabIndex        =   126
         Top             =   5280
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Off Phone II"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   3960
         TabIndex        =   125
         Top             =   5280
         Width           =   1050
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Off Phone I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   3975
         TabIndex        =   121
         Top             =   4755
         Width           =   975
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   5055
         TabIndex        =   120
         Top             =   4755
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   1620
         TabIndex        =   116
         Top             =   4755
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Home Phone I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   300
         TabIndex        =   115
         Top             =   4755
         Width           =   1215
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Home Phone II"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   315
         TabIndex        =   111
         Top             =   5070
         Width           =   1290
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   1635
         TabIndex        =   110
         Top             =   5070
         Width           =   75
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "B_D"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   7395
         TabIndex        =   106
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "No Card"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4125
         TabIndex        =   104
         Top             =   945
         Width           =   705
      End
      Begin VB.Label lblNoCard 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5685
         TabIndex        =   103
         Top             =   945
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   5565
         TabIndex        =   102
         Top             =   945
         Width           =   75
      End
      Begin VB.Label lblCardNo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1725
         TabIndex        =   101
         Top             =   960
         Width           =   75
      End
      Begin VB.Label CustId 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Card No "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   100
         Top             =   960
         Width           =   765
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1605
         TabIndex        =   99
         Top             =   960
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Personal Data Customer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   72
         Left            =   255
         TabIndex        =   98
         Top             =   660
         Width           =   2100
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   48
         Left            =   -70170
         TabIndex        =   91
         Top             =   4560
         Width           =   765
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Next Action"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   41
         Left            =   -74760
         TabIndex        =   89
         Top             =   4395
         Width           =   975
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   42
         Left            =   -73320
         TabIndex        =   88
         Top             =   4395
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Schedule"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   43
         Left            =   -74760
         TabIndex        =   87
         Top             =   4755
         Width           =   780
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   44
         Left            =   -73320
         TabIndex        =   86
         Top             =   4755
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Priority"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   45
         Left            =   -74760
         TabIndex        =   85
         Top             =   5115
         Width           =   615
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   46
         Left            =   -73320
         TabIndex        =   84
         Top             =   5115
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Next Action "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   74
         Left            =   -74850
         TabIndex        =   79
         Top             =   4170
         Width           =   1035
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   78
         Left            =   -73305
         TabIndex        =   75
         Top             =   3630
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   77
         Left            =   -74745
         TabIndex        =   74
         Top             =   3630
         Width           =   750
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   76
         Left            =   -73305
         TabIndex        =   73
         Top             =   3270
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   75
         Left            =   -74745
         TabIndex        =   72
         Top             =   3270
         Width           =   735
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Date PTP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -71895
         TabIndex        =   71
         Top             =   3300
         Width           =   780
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   79
         Left            =   -74430
         TabIndex        =   67
         Top             =   2910
         Width           =   750
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   37
         Left            =   -73320
         TabIndex        =   65
         Top             =   2475
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   38
         Left            =   -74760
         TabIndex        =   64
         Top             =   2475
         Width           =   960
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   39
         Left            =   -73320
         TabIndex        =   63
         Top             =   2115
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Contacted"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   40
         Left            =   -74760
         TabIndex        =   62
         Top             =   2115
         Width           =   870
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Contacted"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   67
         Left            =   -74490
         TabIndex        =   59
         Top             =   1785
         Width           =   870
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Uncontacted"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   -74775
         TabIndex        =   57
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   -73335
         TabIndex        =   56
         Top             =   960
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   35
         Left            =   -74775
         TabIndex        =   55
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   36
         Left            =   -73335
         TabIndex        =   54
         Top             =   1320
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Uncontacted"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   66
         Left            =   -74460
         TabIndex        =   51
         Top             =   615
         Width           =   1050
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Home Phone Additional"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   68
         Left            =   -74850
         TabIndex        =   49
         Top             =   540
         Width           =   1980
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Home Phone I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   58
         Left            =   -74820
         TabIndex        =   48
         Top             =   825
         Width           =   1215
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   59
         Left            =   -73500
         TabIndex        =   47
         Top             =   825
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Home Phone II"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   57
         Left            =   -74820
         TabIndex        =   46
         Top             =   1185
         Width           =   1290
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   56
         Left            =   -73500
         TabIndex        =   45
         Top             =   1185
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Office Phone Additional"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   69
         Left            =   -74895
         TabIndex        =   40
         Top             =   1560
         Width           =   1980
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   52
         Left            =   -73515
         TabIndex        =   39
         Top             =   1830
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Office Phone I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   53
         Left            =   -74835
         TabIndex        =   38
         Top             =   1830
         Width           =   1215
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   55
         Left            =   -73515
         TabIndex        =   37
         Top             =   2190
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Office Phone II"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   54
         Left            =   -74835
         TabIndex        =   36
         Top             =   2190
         Width           =   1290
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Mobile Phone Additional"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   70
         Left            =   -74910
         TabIndex        =   31
         Top             =   3510
         Width           =   2025
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Mobile Phone I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   50
         Left            =   -74910
         TabIndex        =   30
         Top             =   3750
         Width           =   1260
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   51
         Left            =   -73350
         TabIndex        =   29
         Top             =   3750
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Mobile Phone II"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   49
         Left            =   -74910
         TabIndex        =   28
         Top             =   4110
         Width           =   1335
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   47
         Left            =   -73350
         TabIndex        =   27
         Top             =   4110
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Fax Phone Additional"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   71
         Left            =   -74895
         TabIndex        =   24
         Top             =   2535
         Width           =   1785
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   60
         Left            =   -73290
         TabIndex        =   23
         Top             =   3150
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Fax II"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   61
         Left            =   -74850
         TabIndex        =   22
         Top             =   3150
         Width           =   510
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   62
         Left            =   -73290
         TabIndex        =   21
         Top             =   2790
         Width           =   75
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Fax I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   63
         Left            =   -74850
         TabIndex        =   20
         Top             =   2790
         Width           =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   9000
         Y1              =   -3960
         Y2              =   -3960
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Index           =   0
      Left            =   165
      TabIndex        =   9
      Top             =   240
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   741
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
      PictureMaskColor=   -2147483644
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCC_Colection2.frx":2915
      Caption         =   "&Call"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Cancel          =   -1  'True
      Height          =   420
      Index           =   3
      Left            =   9570
      TabIndex        =   10
      Top             =   240
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   741
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
      PictureMaskColor=   -2147483644
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCC_Colection2.frx":3189
      Caption         =   "&Exit"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Index           =   2
      Left            =   8400
      TabIndex        =   11
      Top             =   240
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   741
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
      PictureMaskColor=   -2147483644
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCC_Colection2.frx":32E3
      Caption         =   "&Save"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin VB.Label lblRecsource 
      AutoSize        =   -1  'True
      BackColor       =   &H00E8BE91&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5490
      TabIndex        =   97
      Top             =   450
      Width           =   150
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E8BE91&
      Caption         =   "Recsource"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   80
      Left            =   4050
      TabIndex        =   96
      Top             =   450
      Width           =   885
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E8BE91&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   81
      Left            =   5370
      TabIndex        =   95
      Top             =   450
      Width           =   75
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E8BE91&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   64
      Left            =   5370
      TabIndex        =   94
      Top             =   150
      Width           =   75
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E8BE91&
      Caption         =   "Cust ID "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   65
      Left            =   4050
      TabIndex        =   93
      Top             =   150
      Width           =   720
   End
   Begin VB.Label lblCustId 
      AutoSize        =   -1  'True
      BackColor       =   &H00E8BE91&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5490
      TabIndex        =   92
      Top             =   150
      Width           =   150
   End
End
Attribute VB_Name = "frmCC_Colection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_cust As ADODB.Recordset
Dim m_update As ADODB.Recordset
Dim m_objrs As ADODB.Recordset


Private Sub C_Contacted_Click()
   If C_Contacted.Value Then
      Frame38.Enabled = True
      C_NotContacted.Value = False
      C_Payment.Value = False
   Else
      Frame38.Enabled = False
      cmbContacted.Text = ""
      cmbDescCon.Text = ""
   End If
End Sub

Private Sub C_NotContacted_Click()
   If C_NotContacted.Value Then
      Frame37.Enabled = True
      C_Contacted.Value = False
      C_Payment.Value = False
   Else
      Frame37.Enabled = False
      cmbDescUn.Text = ""
      cmbUncontacted = ""
   End If
End Sub

Private Sub C_Payment_Click()
   If C_Payment.Value Then
      Frame54.Enabled = True
   Else
      Frame54.Enabled = False
      cmbDiscount.Text = ""
   End If
End Sub

Private Sub cmbContacted_Click()
'DESCRIPTION CONTACTED
Dim i As Integer
cmbDescCon.Clear
If Left(cmbContacted.Text, 2) = "RP" Then
   Set m_objrs = New ADODB.Recordset
   m_objrs.Open "Select * from DescContacted where id <= 12", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not m_objrs.EOF
        cmbDescCon.AddItem m_objrs("Description")
        m_objrs.MoveNext
    Wend
Else
    If Left(cmbContacted.Text, 2) = "NA" Then
        Set m_objrs = New ADODB.Recordset
        m_objrs.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not m_objrs.EOF
        cmbDescCon.AddItem m_objrs("Description")
        m_objrs.MoveNext
    Wend
    End If
    If Left(cmbContacted.Text, 2) = "PT" Then
        C_Payment.Value = 1
    End If
End If
Set m_objrs = Nothing
End Sub

Private Sub cmbDiscount_Click()
If lblAmount.Value = 0 Then
txtPayment.Value = 0
Else
txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
txtPayment.Value = lblAmount.Value - (CCur(txtDiscount.Text) * lblAmount.Value)
End If
End Sub

Private Sub cmbUncontacted_Click()
'DESCRIPTION UNCONTACTED
Dim i As Integer
cmbDescUn.Clear
If Left(cmbUncontacted.Text, 2) <> "MV" Then
   Set m_objrs = New ADODB.Recordset
   m_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
         For i = 0 To 2
           cmbDescUn.AddItem m_objrs("Description")
           m_objrs.MoveNext
         Next i
   Set m_objrs = Nothing
Else
   Set m_objrs = New ADODB.Recordset
   m_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
       While Not m_objrs.EOF
           cmbDescUn.AddItem m_objrs("Description")
           m_objrs.MoveNext
       Wend
   Set m_objrs = Nothing
End If
End Sub


Private Sub Form_Load()
   Call HEADER_HISTORY
   Call show_cust
  ' Call Custid_Double
SSTab1.Tab = 0
'CONTACTED

Set m_objrs = New ADODB.Recordset
m_objrs.Open "Select * from ContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not m_objrs.EOF
        cmbContacted.AddItem m_objrs("KdNoProdPresented")
        m_objrs.MoveNext
    Wend
Set m_objrs = Nothing

'UNCONTACTED
Set m_objrs = New ADODB.Recordset
m_objrs.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not m_objrs.EOF
        cmbUncontacted.AddItem m_objrs("KdNoProdPresented")
        m_objrs.MoveNext
    Wend
Set m_objrs = Nothing

'DISCOUNT
Set m_objrs = New ADODB.Recordset
m_objrs.Open "Select * from tblDiscount", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not m_objrs.EOF
        cmbDiscount.AddItem m_objrs("Description")
        m_objrs.MoveNext
    Wend
Set m_objrs = Nothing

'NEXT ACTION
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from StsNextAct", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    cmbNextAct.AddItem m_objrs("NmStsNextAct")
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing


End Sub


Private Sub Option1_Click()
If Option1.Value = True Then
   txtPhone.Text = AHome1.Value & txtHomeNo1.Value
   Option2.Value = False
   Option3.Value = False
   Option4.Value = False
   Option5.Value = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
   txtPhone.Text = AHome2.Value & txtHomeNo2.Value
   Option1.Value = False
   Option3.Value = False
   Option4.Value = False
   Option5.Value = False
End If
End Sub

Private Sub Option3_Click()
   If Option3.Value = True Then
   txtPhone.Text = AOffice1.Value & txtOfficeNo1.Value
   Option2.Value = False
   Option4.Value = False
   Option1.Value = False
   Option5.Value = False
   End If
End Sub

Private Sub Option4_Click()
   If Option4.Value = True Then
   txtPhone.Text = AOffice1.Value & txtOfficeNo1.Value
   Option2.Value = False
   Option3.Value = False
   Option1.Value = False
   Option5.Value = False
End If
End Sub

Private Sub Option5_Click()
 If Option5.Value = True Then
   txtPhone.Text = txtMobileNo2.Value
   Option2.Value = False
   Option3.Value = False
   Option1.Value = False
   Option4.Value = False
   Option6.Value = False
   End If
End Sub

Private Sub Option6_Click()
 If Option6.Value = True Then
   txtPhone.Text = txtMobileNo1.Value
   Option2.Value = False
   Option3.Value = False
   Option1.Value = False
   Option4.Value = False
   Option5.Value = False
   End If
End Sub

Private Sub SSCommand1_Click(Index As Integer)
Select Case Index
   Case 0
        If Len(txtPhone.Text) <> 0 Then
            MDIForm1.ActionCTI ("DIAL|973" + txtPhone.Text + "|" + frmCC_Colection.lblCustId.Caption + "|" + frmCC_Colection.lblRecsource.Caption)
        End If
   Case 2
        V_SAVE = CEK_DATA_VALID
        If V_SAVE = False Then
            Exit Sub
        Else
        End If
        If ADD_CUST Then
            'Call CEK_ADD_PELANGGAN
        Else
            Call CEK_UPDATE_PELANGGAN
        End If
   Case 3
      Unload Me
End Select
End Sub
Public Sub show_cust()
Dim listitem As listitem
Dim m_data As New CLS_FRMCUST_CC
Dim m_cust1 As ADODB.Recordset

On Error GoTo HELL:
Set m_cust = New ADODB.Recordset
m_cust.CursorLocation = adUseClient
m_cust.Open "Select * from mgm where custid='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
If Not m_cust.EOF Then
    lblCustId.Caption = IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID"))
    lblRecsource.Caption = IIf(IsNull(m_cust("RECSOURCE")), "", m_cust("RECSOURCE"))
    lblNama.Caption = IIf(IsNull(m_cust("NAME")), "", m_cust("NAME"))
    lblCardNo.Caption = IIf(IsNull(m_cust("NoCard")), "", m_cust("NoCard"))
    lblID.Caption = IIf(IsNull(m_cust("ktpno")), "", m_cust("ktpno"))
    lblDate.Value = IIf(IsNull(m_cust("BIRTHD")), "", Format(m_cust("BIRTHD"), "dd-mmm-yyyy"))
    lblAddr.Caption = IIf(IsNull(m_cust("ADDRNOW")), "", m_cust("ADDRNOW"))
    lblOfficeAddr.Caption = IIf(IsNull(m_cust("ADDRPT")), "", m_cust("ADDRPT"))
    lblZIP.Caption = IIf(IsNull(m_cust("ZIPNOW")), "", m_cust("ZIPNOW"))
    lblNoCard.Caption = IIf(IsNull(m_cust("NoCard")), "", m_cust("NoCard"))
    lblNoPay.Caption = IIf(IsNull(m_cust("NoPay")), "", m_cust("NoPay"))
    lblPromPA.Caption = IIf(IsNull(m_cust("Prom_FA")), "", m_cust("Prom_FA"))
    lblOpenDate.Caption = IIf(IsNull(m_cust("OpenDate")), "", m_cust("OpenDate"))
    lblLastBill.Caption = IIf(IsNull(m_cust("LastBill")), "", m_cust("LastBill"))
    lblLcAtm.Caption = IIf(IsNull(m_cust("LcATMP")), "", m_cust("LcATMP"))
    lblBrokenPromised.Caption = IIf(IsNull(m_cust("BrokenPromise")), "", m_cust("BrokenPromise"))
    lblBD.Caption = IIf(IsNull(m_cust("B_D")), "", m_cust("B_D"))
    lblLimit.Value = IIf(IsNull(m_cust("Limit")), "", m_cust("Limit"))
    lblPayDt.Value = IIf(IsNull(m_cust("Pay_Dt")), "", m_cust("Pay_Dt"))
    lblLastPay.Value = IIf(IsNull(m_cust("LastPay")), "", m_cust("LastPay"))
    lblTtlPay.Value = IIf(IsNull(m_cust("TtlPay")), "", m_cust("TtlPay"))
    lblAmount.Value = IIf(IsNull(m_cust("AmountWo")), "", Format(m_cust("AmountWo"), "##.##0"))
    AHome1.Value = IIf(IsNull(m_cust("AHOMENO")), "", m_cust("AHOMENO"))
    txtHomeNo1.Value = IIf(IsNull(m_cust("HOMENO")), "", m_cust("HOMENO"))
    AHome2.Value = IIf(IsNull(m_cust("AHOMENO2")), "", m_cust("AHOMENO2"))
    txtHomeNo2.Value = IIf(IsNull(m_cust("HOMENO2")), "", m_cust("HOMENO2"))
    AOffice1.Value = IIf(IsNull(m_cust("AOFFICENO")), "", m_cust("AOFFICENO"))
    txtOfficeNo1.Value = IIf(IsNull(m_cust("OFFICENO")), "", m_cust("OFFICENO"))
    AOffice2.Value = IIf(IsNull(m_cust("AOFFICENO2")), "", m_cust("AOFFICENO2"))
    txtOfficeNo2.Value = IIf(IsNull(m_cust("OFFICENO2")), "", m_cust("OFFICENO2"))
    txtMobileNo1.Value = IIf(IsNull(m_cust("MOBILENO")), "", m_cust("MOBILENO"))
    txtMobileNo2.Value = IIf(IsNull(m_cust("MOBILENO2")), "", m_cust("MOBILENO2"))
    'isi data additional
    
    txtHomeAdd1.Value = IIf(IsNull(m_cust("HOMENOADD1")), "", m_cust("HOMENOADD1"))
    
    txtHomeAdd2.Value = IIf(IsNull(m_cust("HOMENOADD2")), "", m_cust("HOMENOADD2"))
    txtOfficeAdd1.Value = IIf(IsNull(m_cust("OFFICENOADD1")), "", m_cust("OFFICENOADD1"))
    txtOfficeAdd2.Value = IIf(IsNull(m_cust("OFFICENOADD2")), "", m_cust("OFFICENOADD2"))
    txtMobileAdd1.Value = IIf(IsNull(m_cust("MOBILENOADD1")), "", m_cust("MOBILENOADD1"))
    txtMobileAdd2.Value = IIf(IsNull(m_cust("MOBILENOADD2")), "", m_cust("MOBILENOADD2"))
    txtFaxAdd1.Value = IIf(IsNull(m_cust("FAXNOADD1")), "", m_cust("FAXNOADD1"))
    txtFaxAdd2.Value = IIf(IsNull(m_cust("FAXNOADD2")), "", m_cust("FAXNOADD2"))
    
    If UCase(MDIForm1.Text2.Text) = "AGENT" Then
        If Len(txtHomeAdd1.Value) > 2 Then
            txtHomeAdd1.ReadOnly = True
        End If
        If Len(txtHomeAdd2.Value) > 2 Then
            txtHomeAdd2.ReadOnly = True
        End If
        If Len(txtOfficeAdd1.Value) > 2 Then
            txtOfficeAdd1.ReadOnly = True
        End If
        If Len(txtOfficeAdd2.Value) > 2 Then
            txtOfficeAdd2.ReadOnly = True
        End If
        If Len(txtMobileAdd1.Value) > 2 Then
            txtMobileAdd1.ReadOnly = True
        End If
        If Len(txtMobileAdd2.Value) > 2 Then
            txtMobileAdd2.ReadOnly = True
        End If
    End If
    cmbNextAct.Text = IIf(IsNull(m_cust("NEXTACT")), "", m_cust("NEXTACT"))
    Select Case m_cust!RECSTATUS
        Case "N"
            C_NotContacted.Value = 1
            cmbUncontacted.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
            cmbDescUn.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
        Case "C"
            C_Contacted.Value = 1
            cmbContacted.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
            cmbDescCon.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
     End Select
        If IIf(IsNull(m_cust!F_CEK), "", m_cust!F_CEK) = "PTP" Then
            C_Payment.Value = 1
            TdbPTP.Value = IIf(IsNull(m_cust!TGLINCOMING), "", m_cust!TGLINCOMING)
            txtPayment.Value = IIf(IsNull(m_cust!TtlPTP), 0, m_cust!TtlPTP)
            cmbDiscount.Text = IIf(IsNull(m_cust!DiscPersen), 0, m_cust!DiscPersen)
        Else
        
        End If
     
End If
Set m_cust1 = m_data.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + lblCustId.Caption + "'")
While Not m_cust1.EOF
    Set listitem = ListView1(1).ListItems.ADD(, , Left(m_cust1("DATETIME"), 4) & "/" & Mid(m_cust1("DATETIME"), 5, 2) & "/" & IIf(IsNull(m_cust1("DATETIME")), "", Mid(m_cust1("DATETIME"), 7, 2)) & " " & IIf(IsNull(m_cust1("DATETIME")), "", Mid(m_cust1("DATETIME"), 9, 2)) & ":" & Right(m_cust1("DATETIME"), 2))
        listitem.SubItems(1) = IIf(IsNull(m_cust1("HST")), "", m_cust1("HST"))
        listitem.SubItems(2) = IIf(IsNull(m_cust1("AGENT")), "", m_cust1("AGENT"))
        listitem.SubItems(3) = IIf(IsNull(m_cust1("KodeDs")), "", m_cust1("KodeDs"))
        listitem.SubItems(4) = IIf(IsNull(m_cust1("f_cek")), "", m_cust1("f_cek"))
m_cust1.MoveNext
Wend
Set m_cust = Nothing
Exit Sub
HELL:
   MsgBox Err.Description
'   Resume
Set m_cust = Nothing
End Sub
Private Sub CEK_UPDATE_PELANGGAN()
Dim m_data As New CLS_FRMCUST_CC_MGM
Dim pStatusHstLstCall As String
On Error GoTo editErr
    M_OBJCONN.BeginTrans
Set m_update = New ADODB.Recordset
   m_update.Open "Select * from mgm where custid='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
        m_update("CUSTID") = lblCustId.Caption
        m_update("NAME") = lblNama.Caption
        m_update("BIRTHD") = lblDate.Value
        m_update("RECSOURCE") = lblRecsource.Caption
        m_update("ADDRNOW") = lblAddr.Caption
        m_update("ZIPNOW") = lblZIP.Caption
        m_update("HOMENO") = txtHomeNo1.Value
        m_update("HOMENO2") = txtHomeNo2.Value
        m_update("OFFICENO") = txtOfficeNo1.Value
        m_update("OFFICENO2") = txtOfficeNo2.Value
        m_update("MOBILENO") = txtMobileNo1.Value
        m_update("MOBILENO2") = txtMobileNo2.Value
        
        'ADDITIONAL PHONE
        m_update("AHOMENOADD1") = AHomeAdd1(0).Value
        m_update("AHOMENOADD2") = AHomeAdd2(1).Value
        m_update("AOFFICENOADD1") = AOfficeAdd(2).Value
        m_update("AOFFICENOADD2") = AOfficeAdd(3).Value
        m_update("AFAXNOADD1") = AFaxAdd(4).Value
        m_update("AFAXNOADD2") = AFaxAdd(5).Value
        m_update("HOMENOADD1") = txtHomeAdd1.Value
        m_update("HOMENOADD2") = txtHomeAdd2.Value
        m_update("OFFICENOADD1") = txtOfficeAdd1.Value
        m_update("OFFICENOADD2") = txtOfficeAdd2.Value
        m_update("MOBILENOADD1") = txtMobileAdd1.Value
        m_update("MOBILENOADD2") = txtMobileAdd2.Value
        m_update("FAXNOADD1") = txtFaxAdd1.Value
        m_update("FAXNOADD2") = txtFaxAdd2.Value
        If UCase(MDIForm1.Text2.Text) = "AGENT" Then
            If Len(txtHomeAdd1.Value) > 2 Then
                txtHomeAdd1.ReadOnly = True
            End If
            If Len(txtHomeAdd2.Value) > 2 Then
                txtHomeAdd2.ReadOnly = True
            End If
            If Len(txtOfficeAdd1.Value) > 2 Then
                txtOfficeAdd1.ReadOnly = True
            End If
            If Len(txtOfficeAdd2.Value) > 2 Then
                txtOfficeAdd2.ReadOnly = True
            End If
            If Len(txtMobileAdd1.Value) > 2 Then
                txtMobileAdd1.ReadOnly = True
            End If
            If Len(txtMobileAdd2.Value) > 2 Then
                txtMobileAdd2.ReadOnly = True
            End If
        End If

        m_update("PRIOR") = cmbPrior.Text
        m_update("ADDRPT") = lblOfficeAddr.Caption
        m_update("AHOMENO") = AHome1.Value
        m_update("AHOMENO2") = AHome2.Value
        m_update("AOFFICENO") = AOffice1.Value
        m_update("AOFFICENO2") = AOffice2.Value
        m_update("TGLCALL") = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
        If Len(IIf(IsNull(m_update!HOMENO), "", m_update!HOMENO)) > 2 Then
            txtHomeNo1.ReadOnly = True
        End If
        m_update("HOMENO2") = txtHomeNo2.Value
        If Len(IIf(IsNull(m_update!HOMENO2), "", m_update!HOMENO2)) > 2 Then
            txtHomeNo2.ReadOnly = True
        End If
        m_update("MOBILENO") = txtMobileNo1.Value
        If Len(IIf(IsNull(m_update!MOBILENO), "", m_update!MOBILENO)) > 2 Then
            txtMobileNo1.ReadOnly = True
        End If
        m_update("MOBILENO2") = txtMobileNo2.Value
        If Len(IIf(IsNull(m_update!MOBILENO2), "", m_update!MOBILENO2)) > 2 Then
            txtMobileNo2.ReadOnly = True
        End If
        
        m_update("OFFICENO") = txtOfficeNo1.Value
        If Len(IIf(IsNull(m_update!OFFICENO), "", m_update!OFFICENO)) > 2 Then
            txtOfficeNo1.ReadOnly = True
        End If
        m_update("OFFICENO2") = txtOfficeNo2.Value
        If Len(IIf(IsNull(m_update!OFFICENO2), "", m_update!OFFICENO2)) > 2 Then
            txtOfficeNo2.ReadOnly = True
            
         If Len(IIf(IsNull(m_update!HOMENO), "", m_update!HOMENO)) > 2 Then
            txtHomeNo1.ReadOnly = True
        End If
        End If
        If C_Contacted.Value Then
            m_update("RECSTATUS") = "C"
               pStatusLstCall = cmbContacted.Text
               txtResult.Text = pStatusLstCall
               pStatusLstCalldesc = cmbDescCon.Text
               txtResultDesc.Text = pStatusLstCalldesc
               m_update!F_CEK = Left(cmbContacted.Text, 3) & Left(cmbDescCon.Text, 1)
            Else
                If C_NotContacted.Value Then
                    m_update("RECSTATUS") = "N"
                    pStatusLstCall = cmbUncontacted.Text
                    txtResult.Text = pStatusLstCall
                    pStatusLstCalldesc = cmbDescUn.Text
                    txtResultDesc.Text = pStatusLstCalldesc
                    m_update!F_CEK = Left(cmbUncontacted.Text, 3) & Left(cmbDescUn.Text, 2)
                Else
                    m_update!F_CEK = ""
                End If
        End If
        If C_Payment.Value Then
            m_update!TGLINCOMING = Format(TdbPTP.Value, "yyyy/mm/dd")
            m_update!TtlPTP = txtPayment.Value
            m_update!DiscPersen = cmbDiscount.Text
        Else
            m_update!TGLINCOMING = Null
            m_update!TtlPTP = 0
            m_update!DiscPersen = 0
        End If
        If Trim(UCase(IIf(IsNull(m_update("KETHSLKERJA")), "", m_update("KETHSLKERJA")))) = Trim(UCase(pStatusLstCall)) Then
        Else
            m_update("KETHSLKERJA") = pStatusLstCall
            m_update("TGLSTATUS") = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
        End If
        pStatusHstLstCall = m_update("KETHSLKERJA")
        
        m_update("KETHSLKERJADESC") = txtResultDesc.Text
        m_update("PRIOR") = cmbPrior.Text
        m_update("NEXTACT") = cmbNextAct.Text
        m_update("REMARKS") = txtRemarks.Text
        m_update!NEXTACTDATE = Format(cmbDateSch.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
        
    m_update.UPDATE
    
'M_DATA.UPDATE_CUSTOMER_BARU M_OBJCONN, KETHSLKERJA, STATUS_FIELD_LAMA, M_CALL, M_STATUS, DOK1
If C_NotContacted.Value = 1 Then
    If txtRemarks.Text <> Empty Or cmbNextAct.Text <> Empty Then
        m_data.ADD_HISTORY M_OBJCONN, lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, MDIForm1.Text1.Text, "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(m_update!F_CEK), "", m_update!F_CEK))
    End If
Else
    m_data.ADD_HISTORY M_OBJCONN, lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, MDIForm1.Text1.Text, "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(m_update!F_CEK), "", m_update!F_CEK))
End If
M_OBJCONN.CommitTrans
MsgBox "Data Sudah Tersimpan", vbInformation + vbOKOnly, "Sukses"
        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(4) = Format(cmbDateSch.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(5) = cmbNextAct.Text
        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(6) = txtRemarks.Text
        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(10) = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd")
        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11) = pStatusHstLstCall
pStatusLstCall = ""
pStatusHstLstCall = ""
txtRemarks.Text = Empty
'cmbNextAct.Text = Empty
'Unload Me
Set m_data = Nothing
Exit Sub
editErr:
    M_OBJCONN.RollbackTrans
    MsgBox Err.Description
    Resume
End Sub


Private Sub HEADER_HISTORY()
    ListView1(1).ColumnHeaders.ADD 1, , "Tanggal Jam", 15 * TXT
    ListView1(1).ColumnHeaders.ADD 2, , "History", 30 * TXT
    ListView1(1).ColumnHeaders.ADD 3, , "Agent", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 4, , "Sts Call", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 5, , "Sts Call1", 20 * TXT
End Sub


Private Function CEK_DATA_VALID() As Boolean
Dim m_msgbox As Variant

      If C_NotContacted.Value = 1 Then
        If cmbUncontacted.Text = Empty Then
            CEK_DATA_VALID = False
            MsgBox "Not Contacted Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
            SSTab1.Tab = 3
            Exit Function
        Else
                  If cmbDescUn.Text = Empty Then
                     MsgBox "Not Contacted Description harus diisi", vbCritical + vbOKOnly, "Peringatan"
                     Exit Function
                  End If
                  If txtRemarks.Text = "" And cmbNextAct.Text = "" Then
                       CEK_DATA_VALID = False
                        MsgBox "Remarks Atau Next Action Harus Diisi...!!!", vbCritical + vbOKOnly, "Peringatan"
                        SSTab1.Tab = 3
                        Exit Function
                  End If
        End If
    End If
    
    
    If ADD_CUST = True Then
        
    Else
        If cmbDateSch.ValueIsNull = True Or cmbTimeSch.ValueIsNull = True Then
            CEK_DATA_VALID = False
            MsgBox "Tanggal Schedule Harus Di isi", vbCritical + vbOKOnly, "Peringatan"
            SSTab1.Tab = 3
            Exit Function
        End If
        If C_NotContacted.Value = 0 And C_Contacted.Value = 0 Then
                    CEK_DATA_VALID = False
                    MsgBox "Status Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
                    SSTab1.Tab = 3
                    Exit Function
        End If
            If C_NotContacted.Value = 1 Then
                txtRemarks.Text = txtRemarks.Text
                Else
                    If txtRemarks.Text = "" And cmbNextAct.Text = "" Then
                        CEK_DATA_VALID = False
                        MsgBox "Remarks Atau Next Action Harus Diisi...!!!", vbCritical + vbOKOnly, "Peringatan"
                        SSTab1.Tab = 3
                        Exit Function
                    End If
            End If
    End If
CEK_DATA_VALID = True
End Function

Public Sub Custid_Double()
Set m_cust = New ADODB.Recordset
m_cust.CursorLocation = adUseClient
m_cust.Open "Select * from mgm where KTPNO='" & lblNoCard.Caption & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set DataGrid1.DATASOURCE = m_cust
Set m_cust = Nothing
End Sub

Private Sub txtHomeAdd1_Click()
    txtPhone.Text = txtHomeAdd1.Value
End Sub

Private Sub txtHomeAdd2_Click()
    txtPhone.Text = txtHomeAdd2.Value
End Sub

Private Sub txtOfficeAdd1_Click()
    txtPhone.Text = txtOfficeAdd1.Value
End Sub

Private Sub txtOfficeAdd2_Click()
    txtPhone.Text = txtOfficeAdd2.Value
End Sub

Private Sub txtMobileAdd1_Click()
    txtPhone.Text = txtMobileAdd1.Value
End Sub

Private Sub txtMobileAdd2_Click()
    txtPhone.Text = txtMobileAdd2.Value
End Sub
