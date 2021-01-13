VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form FrmOfferingGuide 
   Caption         =   "Offering Guide..."
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Calculator discon"
      Height          =   1875
      Left            =   180
      TabIndex        =   1
      Top             =   1500
      Width           =   7455
      Begin TDBNumber6Ctl.TDBNumber TdbMaxDisc 
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   780
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   450
         Calculator      =   "FrmOfferingGuide.frx":0000
         Caption         =   "FrmOfferingGuide.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmOfferingGuide.frx":008C
         Keys            =   "FrmOfferingGuide.frx":00AA
         Spin            =   "FrmOfferingGuide.frx":00F4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1834745861
         MinValueVT      =   1970470917
      End
      Begin TDBNumber6Ctl.TDBNumber TdbBalance 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   420
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   556
         Calculator      =   "FrmOfferingGuide.frx":011C
         Caption         =   "FrmOfferingGuide.frx":013C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmOfferingGuide.frx":01A8
         Keys            =   "FrmOfferingGuide.frx":01C6
         Spin            =   "FrmOfferingGuide.frx":0210
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         MaxValue        =   999999999999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TdbDiscon 
         Height          =   255
         Left            =   4440
         TabIndex        =   8
         Top             =   420
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   450
         Calculator      =   "FrmOfferingGuide.frx":0238
         Caption         =   "FrmOfferingGuide.frx":0258
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmOfferingGuide.frx":02C4
         Keys            =   "FrmOfferingGuide.frx":02E2
         Spin            =   "FrmOfferingGuide.frx":032C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1834745861
         MinValueVT      =   1970470917
      End
      Begin TDBNumber6Ctl.TDBNumber TDbBalanceAfter 
         Height          =   375
         Left            =   3780
         TabIndex        =   11
         Top             =   1020
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   661
         Calculator      =   "FrmOfferingGuide.frx":0354
         Caption         =   "FrmOfferingGuide.frx":0374
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmOfferingGuide.frx":03E0
         Keys            =   "FrmOfferingGuide.frx":03FE
         Spin            =   "FrmOfferingGuide.frx":0448
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   32768
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   0
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
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label7 
         Caption         =   "(Balance after discon=balance - (balance*discon)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   1560
         Width           =   5655
      End
      Begin VB.Label Label6 
         Caption         =   "Balance after discon:"
         Height          =   255
         Left            =   3780
         TabIndex        =   10
         Top             =   780
         Width           =   1755
      End
      Begin VB.Label Label5 
         Caption         =   "%"
         Height          =   255
         Left            =   5160
         TabIndex        =   9
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "%"
         Height          =   255
         Left            =   1980
         TabIndex        =   7
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Max.Disc:"
         Height          =   255
         Left            =   300
         TabIndex        =   5
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Discon:"
         Height          =   255
         Left            =   3780
         TabIndex        =   3
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Balance:"
         Height          =   255
         Left            =   300
         TabIndex        =   2
         Top             =   420
         Width           =   735
      End
   End
   Begin VB.Label LblTextGuide 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1275
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "FrmOfferingGuide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TdbDiscon_Change()
    'jika diskon melebihi maksimal diskon maka keluar
    If TdbDiscon.Value > TdbMaxDisc.Value Then
        MsgBox "Maksimal diskon: " & TdbMaxDisc.Value & "%", vbOKOnly + vbInformation, "Informasi"
        TdbDiscon.Value = 0
        TDbBalanceAfter.Value = 0
        Exit Sub
    End If
    TDbBalanceAfter.Value = TdbBalance - (TdbBalance * TdbDiscon.Value / 100)
End Sub
