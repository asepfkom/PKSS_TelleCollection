VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Frm_Request 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form Request"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10065
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Request"
      TabPicture(0)   =   "Frm_Request.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SSTab2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TxtAgent"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtCustid"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TxtNamaCH"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Remarks request"
      TabPicture(1)   =   "Frm_Request.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin TabDlg.SSTab SSTab3 
         Height          =   4815
         Left            =   -74940
         TabIndex        =   46
         Top             =   420
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   8493
         _Version        =   393216
         Style           =   1
         Tabs            =   6
         TabHeight       =   520
         TabCaption(0)   =   "Remarks PUM"
         TabPicture(0)   =   "Frm_Request.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LvRemarksPUM"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Remarks EC"
         TabPicture(1)   =   "Frm_Request.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "LvRemarksEC"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Remarks BS"
         TabPicture(2)   =   "Frm_Request.frx":0070
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "LvRemarksBS"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Remarks RS"
         TabPicture(3)   =   "Frm_Request.frx":008C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "LvRemarksRS"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Remarks OST"
         TabPicture(4)   =   "Frm_Request.frx":00A8
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "LvRemarksOST"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Remarks Problem"
         TabPicture(5)   =   "Frm_Request.frx":00C4
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "LvRemarksProblem"
         Tab(5).ControlCount=   1
         Begin MSComctlLib.ListView LvRemarksPUM 
            Height          =   3840
            Left            =   120
            TabIndex        =   47
            Top             =   780
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   6773
            View            =   3
            LabelEdit       =   1
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
         Begin MSComctlLib.ListView LvRemarksEC 
            Height          =   3840
            Left            =   -74820
            TabIndex        =   48
            Top             =   780
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   6773
            View            =   3
            LabelEdit       =   1
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
         Begin MSComctlLib.ListView LvRemarksBS 
            Height          =   3840
            Left            =   -74820
            TabIndex        =   49
            Top             =   780
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   6773
            View            =   3
            LabelEdit       =   1
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
         Begin MSComctlLib.ListView LvRemarksRS 
            Height          =   3840
            Left            =   -74820
            TabIndex        =   50
            Top             =   780
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   6773
            View            =   3
            LabelEdit       =   1
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
         Begin MSComctlLib.ListView LvRemarksOST 
            Height          =   3840
            Left            =   -74820
            TabIndex        =   51
            Top             =   780
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   6773
            View            =   3
            LabelEdit       =   1
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
         Begin MSComctlLib.ListView LvRemarksProblem 
            Height          =   3840
            Left            =   -74880
            TabIndex        =   52
            Top             =   780
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   6773
            View            =   3
            LabelEdit       =   1
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
      End
      Begin VB.TextBox TxtNamaCH 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         TabIndex        =   18
         Top             =   420
         Width           =   1935
      End
      Begin VB.TextBox TxtCustid 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   12
         Top             =   420
         Width           =   1935
      End
      Begin VB.TextBox TxtAgent 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1020
         TabIndex        =   11
         Top             =   420
         Width           =   1935
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4155
         Left            =   240
         TabIndex        =   1
         Top             =   780
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7329
         _Version        =   393216
         Style           =   1
         Tabs            =   6
         TabHeight       =   520
         TabCaption(0)   =   "PUM"
         TabPicture(0)   =   "Frm_Request.frx":00E0
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FramePUM"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "EC"
         TabPicture(1)   =   "Frm_Request.frx":00FC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "BS"
         TabPicture(2)   =   "Frm_Request.frx":0118
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "RS"
         TabPicture(3)   =   "Frm_Request.frx":0134
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame3"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "OST"
         TabPicture(4)   =   "Frm_Request.frx":0150
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame4"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Problem..."
         TabPicture(5)   =   "Frm_Request.frx":016C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame5"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).ControlCount=   1
         Begin VB.Frame Frame5 
            Caption         =   "Form Problem"
            Height          =   3375
            Left            =   -74640
            TabIndex        =   41
            Top             =   660
            Width           =   8715
            Begin VB.CommandButton CmdSendProblem 
               Caption         =   "Send &Problem to Admin"
               Height          =   495
               Left            =   6360
               TabIndex        =   43
               Top             =   2580
               Width           =   2175
            End
            Begin VB.TextBox TxtNoteProblem 
               Appearance      =   0  'Flat
               Height          =   795
               Left            =   1740
               TabIndex        =   42
               Top             =   240
               Width           =   4395
            End
            Begin VB.Label Label16 
               Caption         =   "Your problem:"
               Height          =   555
               Left            =   360
               TabIndex        =   44
               Top             =   300
               Width           =   1035
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Form OST"
            Height          =   3375
            Left            =   -74580
            TabIndex        =   35
            Top             =   660
            Width           =   8715
            Begin VB.TextBox TxtAddrOST 
               Appearance      =   0  'Flat
               Height          =   795
               Left            =   1740
               TabIndex        =   40
               Top             =   240
               Width           =   4395
            End
            Begin VB.TextBox TxtNotesOST 
               Appearance      =   0  'Flat
               Height          =   795
               Left            =   1740
               TabIndex        =   37
               Top             =   1080
               Width           =   4395
            End
            Begin VB.CommandButton CMdSendOST 
               Caption         =   "Send &RS to Admin"
               Height          =   495
               Left            =   6360
               TabIndex        =   36
               Top             =   2580
               Width           =   2175
            End
            Begin VB.Label Label15 
               Caption         =   "Address Request:"
               Height          =   555
               Left            =   360
               TabIndex        =   39
               Top             =   300
               Width           =   1035
            End
            Begin VB.Label Label14 
               Caption         =   "Note:"
               Height          =   255
               Left            =   360
               TabIndex        =   38
               Top             =   1080
               Width           =   1035
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Form RS"
            Height          =   3375
            Left            =   -74640
            TabIndex        =   27
            Top             =   660
            Width           =   8715
            Begin VB.CommandButton CmdSendRS 
               Caption         =   "Send &RS to Admin"
               Height          =   495
               Left            =   6360
               TabIndex        =   34
               Top             =   2580
               Width           =   2175
            End
            Begin VB.TextBox TxtNoteRS 
               Appearance      =   0  'Flat
               Height          =   795
               Left            =   1740
               TabIndex        =   33
               Top             =   1080
               Width           =   4395
            End
            Begin TDBNumber6Ctl.TDBNumber TxtTotalPaymentRS 
               Height          =   255
               Left            =   1800
               TabIndex        =   28
               Top             =   420
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   450
               Calculator      =   "Frm_Request.frx":0188
               Caption         =   "Frm_Request.frx":01A8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Frm_Request.frx":0214
               Keys            =   "Frm_Request.frx":0232
               Spin            =   "Frm_Request.frx":027C
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   0
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###,##0"
               EditMode        =   0
               Enabled         =   0
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
               ReadOnly        =   -1
               Separator       =   ","
               ShowContextMenu =   -1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber TxtInstallmentPeriodRS 
               Height          =   255
               Left            =   1800
               TabIndex        =   31
               Top             =   720
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   450
               Calculator      =   "Frm_Request.frx":02A4
               Caption         =   "Frm_Request.frx":02C4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Frm_Request.frx":0330
               Keys            =   "Frm_Request.frx":034E
               Spin            =   "Frm_Request.frx":0398
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   0
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###,##0"
               EditMode        =   0
               Enabled         =   0
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
               ReadOnly        =   -1
               Separator       =   ","
               ShowContextMenu =   -1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin VB.Label Label13 
               Caption         =   "Note:"
               Height          =   255
               Left            =   360
               TabIndex        =   32
               Top             =   1080
               Width           =   1035
            End
            Begin VB.Label Label12 
               Caption         =   "Installment period:"
               Height          =   255
               Left            =   360
               TabIndex        =   30
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label11 
               Caption         =   "Total Payment"
               Height          =   255
               Left            =   360
               TabIndex        =   29
               Top             =   420
               Width           =   1035
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Form BS"
            Height          =   3135
            Left            =   -74580
            TabIndex        =   19
            Top             =   780
            Width           =   8775
            Begin VB.CommandButton CmdSendBS 
               Caption         =   "Send &BS to Admin"
               Height          =   495
               Left            =   6480
               TabIndex        =   26
               Top             =   2520
               Width           =   2175
            End
            Begin VB.TextBox TxtNoteBS 
               Appearance      =   0  'Flat
               Height          =   795
               Left            =   1260
               TabIndex        =   25
               Top             =   900
               Width           =   4395
            End
            Begin TDBMask6Ctl.TDBMask TxtYearBS 
               Height          =   255
               Left            =   2580
               TabIndex        =   23
               Top             =   540
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   450
               Caption         =   "Frm_Request.frx":03C0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "Frm_Request.frx":042C
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               AllowSpace      =   -1
               AutoConvert     =   -1
               BackColor       =   -2147483643
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   -1
               DataProperty    =   0
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "####"
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
               Text            =   "____"
               Value           =   ""
            End
            Begin VB.ComboBox TxtMonthBS 
               Height          =   315
               ItemData        =   "Frm_Request.frx":046E
               Left            =   960
               List            =   "Frm_Request.frx":0496
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   480
               Width           =   915
            End
            Begin VB.Label Label10 
               Caption         =   "Note:"
               Height          =   255
               Left            =   300
               TabIndex        =   24
               Top             =   900
               Width           =   855
            End
            Begin VB.Label Label9 
               Caption         =   "Year:"
               Height          =   255
               Left            =   1980
               TabIndex        =   22
               Top             =   540
               Width           =   615
            End
            Begin VB.Label Label8 
               Caption         =   "Month:"
               Height          =   255
               Left            =   300
               TabIndex        =   20
               Top             =   540
               Width           =   615
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Form EC"
            Height          =   3375
            Left            =   -74640
            TabIndex        =   13
            Top             =   600
            Width           =   8715
            Begin VB.CommandButton CmdSendEC 
               Caption         =   "Send &EC to Admin"
               Height          =   495
               Left            =   6300
               TabIndex        =   16
               Top             =   2700
               Width           =   2175
            End
            Begin VB.TextBox TxtNoteEc 
               Appearance      =   0  'Flat
               Height          =   795
               Left            =   1380
               TabIndex        =   15
               Top             =   420
               Width           =   4395
            End
            Begin VB.Label Label7 
               Caption         =   "Note:"
               Height          =   255
               Left            =   420
               TabIndex        =   14
               Top             =   420
               Width           =   855
            End
         End
         Begin VB.Frame FramePUM 
            Caption         =   "Form PUM"
            Height          =   3375
            Left            =   300
            TabIndex        =   2
            Top             =   660
            Width           =   8715
            Begin MSWinsockLib.Winsock WinsockSendReq 
               Index           =   0
               Left            =   7440
               Top             =   420
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin VB.TextBox TxtPaymentDatePUM 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   1560
               TabIndex        =   45
               Top             =   1080
               Width           =   1935
            End
            Begin VB.CommandButton CmdPUM 
               Caption         =   "Send &PUM to Admin"
               Height          =   495
               Left            =   6300
               TabIndex        =   8
               Top             =   2700
               Width           =   2175
            End
            Begin VB.TextBox TxtNotePUM 
               Appearance      =   0  'Flat
               Height          =   795
               Left            =   1560
               TabIndex        =   7
               Top             =   1440
               Width           =   4395
            End
            Begin TDBNumber6Ctl.TDBNumber TXtAmountWoPUM 
               Height          =   255
               Left            =   1560
               TabIndex        =   3
               Top             =   780
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   450
               Calculator      =   "Frm_Request.frx":04C1
               Caption         =   "Frm_Request.frx":04E1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "Frm_Request.frx":054D
               Keys            =   "Frm_Request.frx":056B
               Spin            =   "Frm_Request.frx":05B5
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   0
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###,##0"
               EditMode        =   0
               Enabled         =   0
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
               ReadOnly        =   -1
               Separator       =   ","
               ShowContextMenu =   -1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin VB.Label Label5 
               Caption         =   "Note:"
               Height          =   255
               Left            =   480
               TabIndex        =   6
               Top             =   1440
               Width           =   1035
            End
            Begin VB.Label Label4 
               Caption         =   "Payment:"
               Height          =   255
               Left            =   480
               TabIndex        =   5
               Top             =   1140
               Width           =   1035
            End
            Begin VB.Label Label3 
               Caption         =   "Amount Wo:"
               Height          =   255
               Left            =   480
               TabIndex        =   4
               Top             =   780
               Width           =   1035
            End
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Nama CH:"
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Custid:"
         Height          =   255
         Left            =   3060
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Agent:"
         Height          =   255
         Left            =   420
         TabIndex        =   9
         Top             =   420
         Width           =   855
      End
   End
End
Attribute VB_Name = "Frm_Request"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jml As Integer

Private Sub CmdPUM_Click()
    Dim STRSQL As String
    
    STRSQL = "insert into tbl_req_pum (agent,custid,amountwo,"
    STRSQL = STRSQL + "tgl_req,payment_date,remarks_agent,status) values ('"
    STRSQL = STRSQL + Trim(TxtAgent.Text) + "','"
    STRSQL = STRSQL + Trim(TxtCustid.Text) + "','"
    STRSQL = STRSQL + CStr(IIf(IsNull(TXtAmountWoPUM.Value), "0", TXtAmountWoPUM.Value)) + "','"
    STRSQL = STRSQL + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "','"
    STRSQL = STRSQL + Trim(TxtPaymentDatePUM.Text) + "','"
    STRSQL = STRSQL + IIf(IsNull(TxtNotePUM.Text), "", Trim(TxtNotePUM.Text)) + "','0')"
 
    
    M_OBJCONN.Execute STRSQL
    MsgBox "Data PUM untuk custid:" & TxtCustid.Text & " berhasil dikirim ke admin!", vbOKOnly + vbInformation, "Informasi"
    
    SendRequest MDIForm1.TxtUsername.Text & "-Send PUM"
End Sub

Private Sub CmdSendBS_Click()
    Dim STRSQL As String
    
    STRSQL = "insert into tbl_req_bs (agent,custid,nama,tgl_req_bs,"
    STRSQL = STRSQL + "year_bs,month_bs,remarks_agent,status) values ('"
    STRSQL = STRSQL + Trim(TxtAgent.Text) + "','"
    STRSQL = STRSQL + Trim(TxtCustid.Text) + "','"
    STRSQL = STRSQL + Trim(TxtNamaCH.Text) + "','"
    STRSQL = STRSQL + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "','"
    STRSQL = STRSQL + CStr(TxtYearBS.Value) + "','"
    STRSQL = STRSQL + Trim(TxtMonthBS.Text) + "','"
    STRSQL = STRSQL + Trim(TxtNoteBS.Text) + "','0')"
    
    M_OBJCONN.Execute STRSQL
    MsgBox "Data BS untuk custid:" & TxtCustid.Text & " berhasil dikirim ke admin!", vbOKOnly + vbInformation, "Informasi"
    
    SendRequest MDIForm1.TxtUsername.Text & "-Send BS"
End Sub

Private Sub CmdSendEC_Click()
    Dim STRSQL As String
    
    STRSQL = "insert into tbl_req_ec (agent,custid,tgl_req_ec,nama,status,remarks_agent) values ('"
    STRSQL = STRSQL + Trim(TxtAgent.Text) + "','"
    STRSQL = STRSQL + Trim(TxtCustid.Text) + "','"
    STRSQL = STRSQL + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "','"
    STRSQL = STRSQL + Trim(TxtNamaCH.Text) + "','0','"
    STRSQL = STRSQL + IIf(IsNull(TxtNoteEc.Text), "", Trim(TxtNoteEc.Text)) + "')"
    
    M_OBJCONN.Execute STRSQL
    MsgBox "Data EC untuk custid:" & TxtCustid.Text & " berhasil dikirim ke admin!", vbOKOnly + vbInformation, "Informasi"

    SendRequest MDIForm1.TxtUsername.Text & "-Send EC"
End Sub

Private Sub CekCPA()
    Dim STRSQL As String
    Dim M_objrs As ADODB.Recordset
    
    STRSQL = "select * from tblcpa where vcustid='" + Trim(TxtCustid.Text) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount > 0 Then
        TxtNoteRS.Enabled = True
        CmdSendRS.Enabled = True
        
        TxtTotalPaymentRS.Value = IIf(IsNull(M_objrs("nttlpayment")), "0", M_objrs("nttlpayment"))
        TxtInstallmentPeriodRS.Value = IIf(IsNull(M_objrs("nperiod")), "0", M_objrs("nperiod"))
    Else
        TxtNoteRS.Enabled = False
        CmdSendRS.Enabled = False
    End If
    Set M_objrs = Nothing
End Sub

Private Sub CMdSendOST_Click()
    Dim STRSQL As String
    
    STRSQL = "insert into tbl_req_ost (agent,custid,tgl_req_ost,addr,remarks_agent,status) values ('"
    STRSQL = STRSQL + Trim(TxtAgent.Text) + "','"
    STRSQL = STRSQL + Trim(TxtCustid.Text) + "','"
    STRSQL = STRSQL + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "','"
    STRSQL = STRSQL + Trim(TxtAddrOST.Text) + "','"
    STRSQL = STRSQL + IIf(IsNull(TxtNotesOST.Text), "", Trim(TxtNotesOST.Text)) + "','0')"
    
    
    M_OBJCONN.Execute STRSQL
    MsgBox "Data OST untuk custid:" & TxtCustid.Text & " berhasil dikirim ke admin!", vbOKOnly + vbInformation, "Informasi"
    
    SendRequest MDIForm1.TxtUsername.Text & "-Send OST"
End Sub

Private Sub CmdSendProblem_Click()
    Dim STRSQL As String
    
    STRSQL = "insert into tbl_req_problem (agent,nama_agent,tgl,problem,custid,status) values ('"
    STRSQL = STRSQL + Trim(TxtAgent.Text) + "','"
    STRSQL = STRSQL + Trim(MDIForm1.txtnama.Text) + "','"
    STRSQL = STRSQL + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "','"
    STRSQL = STRSQL + Trim(TxtNoteProblem.Text) + "','"
    STRSQL = STRSQL + Trim(TxtCustid.Text) + "','0')"
    
    M_OBJCONN.Execute STRSQL
    MsgBox "Data Problem untuk custid:" & TxtCustid.Text & " berhasil dikirim ke admin!", vbOKOnly + vbInformation, "Informasi"
    
    SendRequest MDIForm1.TxtUsername.Text & "-Send Problem"
End Sub

Private Sub CmdSendRS_Click()
    Dim STRSQL As String
    
    STRSQL = "insert into tbl_req_rs (agent,custid,tgl_req_rs,tot_payment,installment_period,remarks_agent,status) values ('"
    STRSQL = STRSQL + Trim(TxtAgent.Text) + "','"
    STRSQL = STRSQL + Trim(TxtCustid.Text) + "','"
    STRSQL = STRSQL + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "','"
    STRSQL = STRSQL + IIf(IsNull(TxtTotalPaymentRS.Value), "0", CStr(TxtTotalPaymentRS.Value)) + "','"
    STRSQL = STRSQL + IIf(IsNull(TxtInstallmentPeriodRS.Value), "0", CStr(TxtInstallmentPeriodRS.Value)) + "','"
    STRSQL = STRSQL + IIf(IsNull(TxtNoteRS.Text), "", Trim(TxtNoteRS.Text)) + "','0')"
    
    M_OBJCONN.Execute STRSQL
    MsgBox "Data RS untuk custid:" & TxtCustid.Text & " berhasil dikirim ke admin!", vbOKOnly + vbInformation, "Informasi"

    SendRequest MDIForm1.TxtUsername.Text & "-Send RS"

    
End Sub



Private Sub Form_Activate()
    Call CekCPA
    Call HeaderPUM
    Call IsiRemarksPUM
    Call HeaderEC
    Call IsiRemarksEC
    Call HeaderBS
    Call IsiRemarksBS
    Call HeaderRS
    Call IsiRemarksRS
    Call HeaderOST
    Call IsiRemarksOST
    Call HeaderProblem
    Call IsiRemarksProblem
    
   
End Sub

Private Sub HeaderPUM()
    LvRemarksPUM.ColumnHeaders.ADD , , "Tanggal Request", 1700
    LvRemarksPUM.ColumnHeaders.ADD , , "Custid", 1000
    LvRemarksPUM.ColumnHeaders.ADD , , "Agent", 1000
    LvRemarksPUM.ColumnHeaders.ADD , , "Amountwo", 1000
    LvRemarksPUM.ColumnHeaders.ADD , , "Payment Date", 1000
    LvRemarksPUM.ColumnHeaders.ADD , , "Remarks PUM", 4000
    LvRemarksPUM.ColumnHeaders.ADD , , "Remarks By Admin", 4000
End Sub
Private Sub IsiRemarksPUM()
    Dim ListItem As ListItem
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    CMDSQL = "select * from tbl_req_pum where custid='"
    CMDSQL = CMDSQL + Trim(TxtCustid.Text) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvRemarksPUM.ListItems.ADD(, , Format(M_objrs("tgl_req"), "yyyy-mm-dd"))
                ListItem.SubItems(1) = M_objrs("custid")
                ListItem.SubItems(2) = M_objrs("agent")
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("amountwo")), "0", M_objrs("amountwo"))
                ListItem.SubItems(4) = IIf(IsNull(M_objrs("payment_date")), "", Format(M_objrs("payment_date"), "yyyy-mm-dd"))
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("remarks_agent")), "", M_objrs("remarks_agent"))
                ListItem.SubItems(6) = IIf(IsNull(M_objrs("remarks")), "", M_objrs("remarks"))
                
            If M_objrs("status") = "0" Then
                ListItem.ForeColor = vbRed
                ListItem.ListSubItems(1).ForeColor = vbRed
                ListItem.ListSubItems(2).ForeColor = vbRed
                ListItem.ListSubItems(3).ForeColor = vbRed
                ListItem.ListSubItems(4).ForeColor = vbRed
                ListItem.ListSubItems(5).ForeColor = vbRed
                ListItem.ListSubItems(6).ForeColor = vbRed
            Else
                ListItem.ForeColor = vbBlue
                ListItem.ListSubItems(1).ForeColor = vbBlue
                ListItem.ListSubItems(2).ForeColor = vbBlue
                ListItem.ListSubItems(3).ForeColor = vbBlue
                ListItem.ListSubItems(4).ForeColor = vbBlue
                ListItem.ListSubItems(5).ForeColor = vbBlue
                ListItem.ListSubItems(6).ForeColor = vbBlue
            End If
            M_objrs.MoveNext
        Wend
    End If
End Sub

Private Sub HeaderEC()
    LvRemarksEC.ColumnHeaders.ADD , , "Tanggal Request", 1700
    LvRemarksEC.ColumnHeaders.ADD , , "Custid", 1000
    LvRemarksEC.ColumnHeaders.ADD , , "Agent", 1000
    LvRemarksEC.ColumnHeaders.ADD , , "Nama CH", 1000
    LvRemarksEC.ColumnHeaders.ADD , , "Remarks EC", 4000
    LvRemarksEC.ColumnHeaders.ADD , , "Remarks By Admin", 4000
End Sub

Private Sub IsiRemarksEC()
    Dim ListItem As ListItem
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    CMDSQL = "select * from tbl_req_ec where custid='"
    CMDSQL = CMDSQL + Trim(TxtCustid.Text) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvRemarksEC.ListItems.ADD(, , Format(M_objrs("tgl_req_ec"), "yyyy-mm-dd"))
                ListItem.SubItems(1) = M_objrs("custid")
                ListItem.SubItems(2) = M_objrs("agent")
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("nama")), "", M_objrs("nama"))
                ListItem.SubItems(4) = IIf(IsNull(M_objrs("remarks_agent")), "", M_objrs("remarks_agent"))
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("remarks")), "", M_objrs("remarks"))
                
            If M_objrs("status") = "0" Then
                ListItem.ForeColor = vbRed
                ListItem.ListSubItems(1).ForeColor = vbRed
                ListItem.ListSubItems(2).ForeColor = vbRed
                ListItem.ListSubItems(3).ForeColor = vbRed
                ListItem.ListSubItems(4).ForeColor = vbRed
                ListItem.ListSubItems(5).ForeColor = vbRed
            Else
                ListItem.ForeColor = vbBlue
                ListItem.ListSubItems(1).ForeColor = vbBlue
                ListItem.ListSubItems(2).ForeColor = vbBlue
                ListItem.ListSubItems(3).ForeColor = vbBlue
                ListItem.ListSubItems(4).ForeColor = vbBlue
                ListItem.ListSubItems(5).ForeColor = vbBlue
            End If
            M_objrs.MoveNext
        Wend
    End If
End Sub

Private Sub HeaderBS()
    LvRemarksBS.ColumnHeaders.ADD , , "Tanggal Request", 1700
    LvRemarksBS.ColumnHeaders.ADD , , "Custid", 1000
    LvRemarksBS.ColumnHeaders.ADD , , "Agent", 1000
    LvRemarksBS.ColumnHeaders.ADD , , "Nama CH", 1000
    LvRemarksBS.ColumnHeaders.ADD , , "Month", 1000
    LvRemarksBS.ColumnHeaders.ADD , , "Year", 1000
    LvRemarksBS.ColumnHeaders.ADD , , "Remarks BS", 4000
    LvRemarksBS.ColumnHeaders.ADD , , "Remarks By Admin", 4000
End Sub
Private Sub IsiRemarksBS()
    Dim ListItem As ListItem
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    CMDSQL = "select * from tbl_req_bs where custid='"
    CMDSQL = CMDSQL + Trim(TxtCustid.Text) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvRemarksBS.ListItems.ADD(, , Format(M_objrs("tgl_req_bs"), "yyyy-mm-dd"))
                ListItem.SubItems(1) = M_objrs("custid")
                ListItem.SubItems(2) = M_objrs("agent")
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("nama")), "", M_objrs("nama"))
                ListItem.SubItems(4) = IIf(IsNull(M_objrs("month_bs")), "", M_objrs("month_bs"))
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("year_bs")), "", M_objrs("year_bs"))
                ListItem.SubItems(6) = IIf(IsNull(M_objrs("remarks_agent")), "", M_objrs("remarks_agent"))
                ListItem.SubItems(7) = IIf(IsNull(M_objrs("remarks")), "", M_objrs("remarks"))
                
            If M_objrs("status") = "0" Then
                ListItem.ForeColor = vbRed
                ListItem.ListSubItems(1).ForeColor = vbRed
                ListItem.ListSubItems(2).ForeColor = vbRed
                ListItem.ListSubItems(3).ForeColor = vbRed
                ListItem.ListSubItems(4).ForeColor = vbRed
                ListItem.ListSubItems(5).ForeColor = vbRed
                ListItem.ListSubItems(6).ForeColor = vbRed
                ListItem.ListSubItems(7).ForeColor = vbRed
            Else
                ListItem.ForeColor = vbBlue
                ListItem.ListSubItems(1).ForeColor = vbBlue
                ListItem.ListSubItems(2).ForeColor = vbBlue
                ListItem.ListSubItems(3).ForeColor = vbBlue
                ListItem.ListSubItems(4).ForeColor = vbBlue
                ListItem.ListSubItems(5).ForeColor = vbBlue
                ListItem.ListSubItems(6).ForeColor = vbBlue
                ListItem.ListSubItems(7).ForeColor = vbBlue
            End If
            M_objrs.MoveNext
        Wend
    End If
End Sub

Private Sub HeaderRS()
    LvRemarksRS.ColumnHeaders.ADD , , "Tanggal Request", 1700
    LvRemarksRS.ColumnHeaders.ADD , , "Custid", 1000
    LvRemarksRS.ColumnHeaders.ADD , , "Agent", 1000
    LvRemarksRS.ColumnHeaders.ADD , , "Total Payment", 1000
    LvRemarksRS.ColumnHeaders.ADD , , "Installment Period", 1000
    LvRemarksRS.ColumnHeaders.ADD , , "Remarks BS", 4000
    LvRemarksRS.ColumnHeaders.ADD , , "Remarks By Admin", 4000
End Sub
Private Sub IsiRemarksRS()
    Dim ListItem As ListItem
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    CMDSQL = "select * from tbl_req_rs where custid='"
    CMDSQL = CMDSQL + Trim(TxtCustid.Text) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvRemarksRS.ListItems.ADD(, , Format(M_objrs("tgl_req_rs"), "yyyy-mm-dd"))
                ListItem.SubItems(1) = M_objrs("custid")
                ListItem.SubItems(2) = M_objrs("agent")
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("tot_payment")), "0", M_objrs("tot_payment"))
                ListItem.SubItems(4) = IIf(IsNull(M_objrs("installment_period")), "0", M_objrs("installment_period"))
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("remarks_agent")), "", M_objrs("remarks_agent"))
                ListItem.SubItems(6) = IIf(IsNull(M_objrs("remarks")), "", M_objrs("remarks"))
                
            If M_objrs("status") = "0" Then
                ListItem.ForeColor = vbRed
                ListItem.ListSubItems(1).ForeColor = vbRed
                ListItem.ListSubItems(2).ForeColor = vbRed
                ListItem.ListSubItems(3).ForeColor = vbRed
                ListItem.ListSubItems(4).ForeColor = vbRed
                ListItem.ListSubItems(5).ForeColor = vbRed
                ListItem.ListSubItems(6).ForeColor = vbRed
            Else
                ListItem.ForeColor = vbBlue
                ListItem.ListSubItems(1).ForeColor = vbBlue
                ListItem.ListSubItems(2).ForeColor = vbBlue
                ListItem.ListSubItems(3).ForeColor = vbBlue
                ListItem.ListSubItems(4).ForeColor = vbBlue
                ListItem.ListSubItems(5).ForeColor = vbBlue
                ListItem.ListSubItems(6).ForeColor = vbBlue
            End If
            M_objrs.MoveNext
        Wend
    End If
End Sub


Private Sub HeaderOST()
    LvRemarksOST.ColumnHeaders.ADD , , "Tanggal Request", 1700
    LvRemarksOST.ColumnHeaders.ADD , , "Custid", 1000
    LvRemarksOST.ColumnHeaders.ADD , , "Agent", 1000
    LvRemarksOST.ColumnHeaders.ADD , , "Address Request", 1000
    LvRemarksOST.ColumnHeaders.ADD , , "Remarks OST", 4000
    LvRemarksOST.ColumnHeaders.ADD , , "Remarks By Admin", 4000
End Sub
Private Sub IsiRemarksOST()
    Dim ListItem As ListItem
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    CMDSQL = "select * from tbl_req_ost where custid='"
    CMDSQL = CMDSQL + Trim(TxtCustid.Text) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvRemarksOST.ListItems.ADD(, , Format(M_objrs("tgl_req_ost"), "yyyy-mm-dd"))
                ListItem.SubItems(1) = M_objrs("custid")
                ListItem.SubItems(2) = M_objrs("agent")
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("addr")), "0", M_objrs("addr"))
                ListItem.SubItems(4) = IIf(IsNull(M_objrs("remarks_agent")), "", M_objrs("remarks_agent"))
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("remarks")), "", M_objrs("remarks"))
                
            If M_objrs("status") = "0" Then
                ListItem.ForeColor = vbRed
                ListItem.ListSubItems(1).ForeColor = vbRed
                ListItem.ListSubItems(2).ForeColor = vbRed
                ListItem.ListSubItems(3).ForeColor = vbRed
                ListItem.ListSubItems(4).ForeColor = vbRed
                ListItem.ListSubItems(5).ForeColor = vbRed
            Else
                ListItem.ForeColor = vbBlue
                ListItem.ListSubItems(1).ForeColor = vbBlue
                ListItem.ListSubItems(2).ForeColor = vbBlue
                ListItem.ListSubItems(3).ForeColor = vbBlue
                ListItem.ListSubItems(4).ForeColor = vbBlue
                ListItem.ListSubItems(5).ForeColor = vbBlue
            End If
            M_objrs.MoveNext
        Wend
    End If
End Sub

Private Sub HeaderProblem()
    LvRemarksProblem.ColumnHeaders.ADD , , "Tanggal Request", 1700
    LvRemarksProblem.ColumnHeaders.ADD , , "Custid", 1000
    LvRemarksProblem.ColumnHeaders.ADD , , "Agent", 1000
    LvRemarksProblem.ColumnHeaders.ADD , , "Nama Agent", 1000
    LvRemarksProblem.ColumnHeaders.ADD , , "Problem", 4000
    LvRemarksProblem.ColumnHeaders.ADD , , "Solving", 4000
End Sub
Private Sub IsiRemarksProblem()
    Dim ListItem As ListItem
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    CMDSQL = "select * from tbl_req_problem where custid='"
    CMDSQL = CMDSQL + Trim(TxtCustid.Text) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvRemarksProblem.ListItems.ADD(, , Format(M_objrs("tgl"), "yyyy-mm-dd"))
                ListItem.SubItems(1) = M_objrs("custid")
                ListItem.SubItems(2) = M_objrs("agent")
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("nama_agent")), "0", M_objrs("nama_agent"))
                ListItem.SubItems(4) = IIf(IsNull(M_objrs("problem")), "", M_objrs("problem"))
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("solve")), "", M_objrs("solve"))
                
            If M_objrs("status") = "0" Then
                ListItem.ForeColor = vbRed
                ListItem.ListSubItems(1).ForeColor = vbRed
                ListItem.ListSubItems(2).ForeColor = vbRed
                ListItem.ListSubItems(3).ForeColor = vbRed
                ListItem.ListSubItems(4).ForeColor = vbRed
                ListItem.ListSubItems(5).ForeColor = vbRed
            Else
                ListItem.ForeColor = vbBlue
                ListItem.ListSubItems(1).ForeColor = vbBlue
                ListItem.ListSubItems(2).ForeColor = vbBlue
                ListItem.ListSubItems(3).ForeColor = vbBlue
                ListItem.ListSubItems(4).ForeColor = vbBlue
                ListItem.ListSubItems(5).ForeColor = vbBlue
            End If
            M_objrs.MoveNext
        Wend
    End If
End Sub

Private Sub SendRequest(st As String)
    Dim M_objrs As ADODB.Recordset
    Dim STRSQL As String
    On Error Resume Next
    
    'StrSql = "select *,now() as waktu from tbl_ip where "
    'StrSql = StrSql + " tipe='ADMINISTRATOR' or tipe='ADMIN' or tipe='SUPERVISOR'"
    STRSQL = "select now() as waktu "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    i = 0
'    If M_OBJRS.RecordCount > 0 Then
'
'        'Di tutup dulu koneksinya
'        'For J = 1 To M_OBJRS.RecordCount
'            'WinsockSendReq(J).Close
'        'Next J
'
'        While Not M_OBJRS.EOF
'            i = i + 1
'            Load WinsockSendReq(i)
'            WinsockSendReq(i).RemotePort = 3030
'            WinsockSendReq(i).RemoteHost = CStr(Trim(M_OBJRS("ip_addr")))
'            WinsockSendReq(i).Close
'
'            WinsockSendReq(i).Connect
'            'WinsockSendReq_Connect
'
'            WaitSecs (3)
'            DoEvents
'            On Error Resume Next
'            WinsockSendReq(i).SendData CStr(Format(M_OBJRS("waktu"), "yyyy-mm-dd hh:mm:ss")) & " " & st
'            'WinsockSendReq(i).Close
'            'Command1_Click
'            M_OBJRS.MoveNext
'        Wend
'        i = 0
'    End If
'
'    Set M_OBJRS = Nothing
    For i = 1 To jml
        WinsockSendReq(i).SendData "[" & CStr(Format(M_objrs("waktu"), "yyyy-mm-dd hh:mm:ss")) & "] " & " " & st
    Next i
    Set M_objrs = Nothing
End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim STRSQL As String
    On Error Resume Next
    
    'WinsockSendReq(0).RemotePort = 3030
    jml = 0
    
    STRSQL = "select *,now() as waktu from tbl_ip where "
    STRSQL = STRSQL + " tipe='ADMINISTRATOR' or tipe='ADMIN' or tipe='SUPERVISOR'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            jml = jml + 1
            DoEvents
            Load WinsockSendReq(jml)
            WinsockSendReq(jml).RemotePort = 3030
            WinsockSendReq(jml).RemoteHost = CStr(Trim(M_objrs("ip_addr")))
            WinsockSendReq(jml).Close
            
            WinsockSendReq(jml).Connect
            WaitSecs (3)
            M_objrs.MoveNext
        Wend
    End If

    Set M_objrs = Nothing
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    For W = 1 To jml
        WinsockSendReq(W).Close
        Unload WinsockSendReq(W)
    Next W
End Sub

Private Sub WinsockSendReq_Close(Index As Integer)
    WinsockSendReq(Index).Close
End Sub



Private Sub WinsockSendReq_Connect(Index As Integer)
    'MsgBox "OK!"
End Sub

Private Sub WinsockSendReq_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim st As String
    WinsockSendReq(Index).GetData st
End Sub
