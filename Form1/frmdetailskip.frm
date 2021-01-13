VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdetailskip 
   BackColor       =   &H0080FF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Detail Skip Tracer"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   16680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   345
      Left            =   14100
      TabIndex        =   57
      Top             =   5310
      Width           =   1215
   End
   Begin VB.TextBox txtaddradd 
      Height          =   675
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txttellppersonal 
      Height          =   315
      Left            =   840
      MaxLength       =   15
      TabIndex        =   23
      Top             =   3090
      Width           =   2385
   End
   Begin VB.TextBox txthppersonal 
      Height          =   315
      Left            =   840
      MaxLength       =   15
      TabIndex        =   22
      Top             =   3450
      Width           =   2385
   End
   Begin VB.TextBox txtofficeaddr 
      Height          =   675
      Left            =   4380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   2370
      Width           =   2745
   End
   Begin VB.TextBox txttelpoffice 
      Height          =   315
      Left            =   4410
      MaxLength       =   15
      TabIndex        =   20
      Top             =   3090
      Width           =   2715
   End
   Begin VB.TextBox txthpoffice 
      Height          =   315
      Left            =   4410
      MaxLength       =   15
      TabIndex        =   19
      Top             =   3420
      Width           =   2715
   End
   Begin VB.TextBox txtnamefamily1 
      Height          =   315
      Left            =   8190
      TabIndex        =   18
      Top             =   2400
      Width           =   1755
   End
   Begin VB.TextBox txtaddrfamiliy1 
      Height          =   495
      Left            =   8160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   2760
      Width           =   2355
   End
   Begin VB.TextBox txttelpfamily 
      Height          =   315
      Left            =   8160
      TabIndex        =   16
      Top             =   3300
      Width           =   1845
   End
   Begin VB.TextBox txthpfamilly 
      Height          =   315
      Left            =   10380
      TabIndex        =   15
      Top             =   3300
      Width           =   1725
   End
   Begin VB.TextBox txtnamefamiliy2 
      Height          =   315
      Left            =   12900
      TabIndex        =   14
      Top             =   2400
      Width           =   1755
   End
   Begin VB.TextBox txtaddrfamiliy2 
      Height          =   495
      Left            =   12900
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   2730
      Width           =   3615
   End
   Begin VB.TextBox txttelpfamilly2 
      Height          =   315
      Left            =   12870
      TabIndex        =   12
      Top             =   3270
      Width           =   1725
   End
   Begin VB.TextBox txthpfamiliy2 
      Height          =   315
      Left            =   14940
      TabIndex        =   11
      Top             =   3300
      Width           =   1515
   End
   Begin VB.TextBox txtnamefriend1 
      Height          =   315
      Left            =   780
      TabIndex        =   10
      Top             =   4290
      Width           =   1755
   End
   Begin VB.TextBox txtaddrfriend1 
      Height          =   495
      Left            =   750
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4650
      Width           =   3885
   End
   Begin VB.TextBox txttelpfriend1 
      Height          =   315
      Left            =   750
      TabIndex        =   8
      Top             =   5190
      Width           =   1845
   End
   Begin VB.TextBox txthpfriend1 
      Height          =   315
      Left            =   2940
      TabIndex        =   7
      Top             =   5190
      Width           =   1725
   End
   Begin VB.TextBox txtnamefriend2 
      Height          =   315
      Left            =   5670
      TabIndex        =   6
      Top             =   4350
      Width           =   1755
   End
   Begin VB.TextBox txtaddrfriend2 
      Height          =   495
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   4710
      Width           =   3885
   End
   Begin VB.TextBox txttelpfriend2 
      Height          =   315
      Left            =   5640
      TabIndex        =   4
      Top             =   5250
      Width           =   1845
   End
   Begin VB.TextBox txthpfriend2 
      Height          =   315
      Left            =   7830
      TabIndex        =   3
      Top             =   5250
      Width           =   1725
   End
   Begin VB.TextBox txtremarks 
      Height          =   915
      Left            =   9780
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4320
      Width           =   6465
   End
   Begin VB.TextBox txthp 
      Height          =   315
      Left            =   10200
      MaxLength       =   15
      TabIndex        =   1
      Top             =   5280
      Width           =   2655
   End
   Begin MSComctlLib.ListView LvSearch 
      Height          =   1320
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   16500
      _ExtentX        =   29104
      _ExtentY        =   2328
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   14640
      TabIndex        =   56
      Top             =   3330
      Width           =   675
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   7
      Left            =   6570
      TabIndex        =   55
      Top             =   120
      Width           =   2325
   End
   Begin VB.Image Image2 
      Height          =   405
      Index           =   7
      Left            =   0
      Picture         =   "frmdetailskip.frx":0000
      Stretch         =   -1  'True
      Top             =   30
      Width           =   16470
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   12150
      TabIndex        =   54
      Top             =   3900
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   14640
      TabIndex        =   53
      Top             =   3030
      Width           =   675
   End
   Begin VB.Image Image2 
      Height          =   405
      Index           =   6
      Left            =   9780
      Picture         =   "frmdetailskip.frx":049C
      Stretch         =   -1  'True
      Top             =   3870
      Width           =   6600
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Friend2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6690
      TabIndex        =   52
      Top             =   3930
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   405
      Index           =   5
      Left            =   4920
      Picture         =   "frmdetailskip.frx":0938
      Stretch         =   -1  'True
      Top             =   3870
      Width           =   4770
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   660
      TabIndex        =   51
      Top             =   1920
      Width           =   2325
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Office Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4620
      TabIndex        =   50
      Top             =   1890
      Width           =   2325
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   49
      Top             =   2430
      Width           =   675
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Tlp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   30
      TabIndex        =   48
      Top             =   3120
      Width           =   675
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Hp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   30
      TabIndex        =   47
      Top             =   3480
      Width           =   675
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3600
      TabIndex        =   46
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Tlp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3540
      TabIndex        =   45
      Top             =   3180
      Width           =   675
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Hp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3570
      TabIndex        =   44
      Top             =   3480
      Width           =   675
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7320
      TabIndex        =   43
      Top             =   2430
      Width           =   675
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Family 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   12930
      TabIndex        =   42
      Top             =   1950
      Width           =   1335
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7410
      TabIndex        =   41
      Top             =   2820
      Width           =   675
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Telp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7680
      TabIndex        =   40
      Top             =   3330
      Width           =   675
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Hp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10020
      TabIndex        =   39
      Top             =   3360
      Width           =   675
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12180
      TabIndex        =   38
      Top             =   2430
      Width           =   675
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12270
      TabIndex        =   37
      Top             =   2820
      Width           =   675
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Telp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12240
      TabIndex        =   36
      Top             =   3330
      Width           =   675
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   4320
      Width           =   675
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   4710
      Width           =   675
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Telp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   270
      TabIndex        =   33
      Top             =   5220
      Width           =   675
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4890
      TabIndex        =   32
      Top             =   4380
      Width           =   675
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4890
      TabIndex        =   31
      Top             =   4770
      Width           =   675
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Telp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5160
      TabIndex        =   30
      Top             =   5280
      Width           =   675
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Hp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7500
      TabIndex        =   29
      Top             =   5310
      Width           =   675
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Family 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   8160
      TabIndex        =   28
      Top             =   1890
      Width           =   2325
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Friend1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   870
      TabIndex        =   27
      Top             =   3870
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10560
      TabIndex        =   26
      Top             =   4050
      Width           =   675
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Hp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9780
      TabIndex        =   25
      Top             =   5310
      Width           =   675
   End
   Begin VB.Image Image2 
      Height          =   405
      Index           =   0
      Left            =   -30
      Picture         =   "frmdetailskip.frx":0DD4
      Stretch         =   -1  'True
      Top             =   1890
      Width           =   3540
   End
   Begin VB.Image Image2 
      Height          =   405
      Index           =   1
      Left            =   3600
      Picture         =   "frmdetailskip.frx":1270
      Stretch         =   -1  'True
      Top             =   1890
      Width           =   3630
   End
   Begin VB.Image Image2 
      Height          =   405
      Index           =   2
      Left            =   7290
      Picture         =   "frmdetailskip.frx":170C
      Stretch         =   -1  'True
      Top             =   1890
      Width           =   3480
   End
   Begin VB.Image Image2 
      Height          =   405
      Index           =   3
      Left            =   10890
      Picture         =   "frmdetailskip.frx":1BA8
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   5700
   End
   Begin VB.Image Image2 
      Height          =   405
      Index           =   4
      Left            =   0
      Picture         =   "frmdetailskip.frx":2044
      Stretch         =   -1  'True
      Top             =   3870
      Width           =   4770
   End
End
Attribute VB_Name = "frmdetailskip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
HeaderList
Dim M_objrs As New ADODB.Recordset
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open "SELECT * FROM mandiri.opening_screen where name like '%" + FrmCC_Colection.lblnama + "%'", M_OBJCONN, adOpenDynamic, adLockOptimistic
LvSearch.ListItems.CLEAR
While Not M_objrs.EOF
   Set ListItem = LvSearch.ListItems.ADD(, , M_objrs("idopening"))
            ListItem.SubItems(1) = IIf(IsNull(M_objrs("name")), "", M_objrs("name"))
M_objrs.MoveNext
 
Wend
End Sub
Private Sub HeaderList()
    'Header 0
    LvSearch.ColumnHeaders.ADD 1, , "Id.", 500
    'Header 1
    LvSearch.ColumnHeaders.ADD 2, , "Name", 3000
    'Header 2
   End Sub
Public Sub showdata()
Dim mobjrs As New ADODB.Recordset
If LvSearch.ListItems.Count <> 0 Then

   Set mobjrs = New ADODB.Recordset
       mobjrs.CursorLocation = adUseClient
       STRSQL = "select * from mandiri.opening_screen where idopening =" + CStr(LvSearch.SelectedItem.Text) + ""
       mobjrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
       If mobjrs.RecordCount > 0 Then
       txtaddradd.Text = IIf(IsNull(mobjrs!personal_alamat), "", mobjrs!personal_alamat)
       txttellppersonal.Text = IIf(IsNull(mobjrs!personal_telp), "", mobjrs!personal_telp)
       txthppersonal.Text = IIf(IsNull(mobjrs!personal_hp), "", mobjrs!personal_hp)
       txtofficeaddr.Text = IIf(IsNull(mobjrs!office_alamat), "", mobjrs!office_alamat)
       txttelpoffice.Text = IIf(IsNull(mobjrs!office_telp), "", mobjrs!office_telp)
       txthpoffice.Text = IIf(IsNull(mobjrs!office_hp), "", mobjrs!office_hp)
       txtnamefamily1.Text = IIf(IsNull(mobjrs!familiy1_name), "", mobjrs!familiy1_name)
       txtaddrfamiliy1.Text = IIf(IsNull(mobjrs!familiy1_alamat), "", mobjrs!familiy1_alamat)
       txttelpfamily.Text = IIf(IsNull(mobjrs!familiy1_telp), "", mobjrs!familiy1_telp)
       txthpfamilly.Text = IIf(IsNull(mobjrs!familiy1_hp), "", mobjrs!familiy1_hp)
       txtnamefamiliy2.Text = IIf(IsNull(mobjrs!familiy2_name), "", mobjrs!familiy2_name)
       txtaddrfamiliy2.Text = IIf(IsNull(mobjrs!familiy2_alamat), "", mobjrs!familiy2_alamat)
       txttelpfamilly2.Text = IIf(IsNull(mobjrs!familiy2_telp), "", mobjrs!familiy2_telp)
       txthpfamiliy2.Text = IIf(IsNull(mobjrs!familiy2_hp), "", mobjrs!familiy2_hp)
       txthp.Text = IIf(IsNull(mobjrs!hp), "", mobjrs!hp)
       'Combo1.Text = IIf(IsNull(mobjrs!stsaccount), "", mobjrs!stsaccount)
       txtnamefriend1.Text = IIf(IsNull(mobjrs!friend1_name), "", mobjrs!friend1_name)
       txtaddrfriend1.Text = IIf(IsNull(mobjrs!friend1_alamat), "", mobjrs!friend1_alamat)
       txttelpfriend1.Text = IIf(IsNull(mobjrs!friend1_telp), "", mobjrs!friend1_telp)
       txthpfriend1.Text = IIf(IsNull(mobjrs!friend1_hp), "", mobjrs!friend1_hp)
       txtnamefriend2.Text = IIf(IsNull(mobjrs!friend2_name), "", mobjrs!friend2_name)
       txtaddrfriend2.Text = IIf(IsNull(mobjrs!friend2_alamat), "", mobjrs!friend2_alamat)
       txttelpfriend2.Text = IIf(IsNull(mobjrs!friend2_telp), "", mobjrs!friend2_telp)
       txthpfriend2.Text = IIf(IsNull(mobjrs!friend2_hp), "", mobjrs!friend2_hp)
       TxtRemarks.Text = IIf(IsNull(mobjrs!Remarks), "", mobjrs!Remarks)
       End If
       End If
End Sub
Private Sub LvSearch_DblClick()
    showdata
End Sub

