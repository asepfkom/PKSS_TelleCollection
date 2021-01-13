VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmCC_Colection 
   BackColor       =   &H00FCFCFC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer"
   ClientHeight    =   10950
   ClientLeft      =   735
   ClientTop       =   435
   ClientWidth     =   19140
   Icon            =   "frmCC_Colection_RITCARD12.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   19140
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   7935
      TabIndex        =   175
      Top             =   45
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.Frame Frame3 
      Caption         =   "0"
      Height          =   285
      Left            =   3225
      TabIndex        =   90
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Frame Frame17 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCFCFC&
         Caption         =   "Other Info"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   9240
         TabIndex        =   275
         Top             =   360
         Width           =   2115
         Begin VB.ComboBox CmbStsKatHome1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmCC_Colection_RITCARD12.frx":000C
            Left            =   7140
            List            =   "frmCC_Colection_RITCARD12.frx":0028
            TabIndex        =   287
            Text            =   "--Pilih Kategori Telepon--"
            Top             =   360
            Width           =   2445
         End
         Begin VB.ComboBox CmbStsKatOffice1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmCC_Colection_RITCARD12.frx":00A6
            Left            =   7140
            List            =   "frmCC_Colection_RITCARD12.frx":00C2
            TabIndex        =   286
            Text            =   "--Pilih Kategori Telepon--"
            Top             =   1080
            Width           =   2445
         End
         Begin VB.ComboBox CmbStsKatOffice2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmCC_Colection_RITCARD12.frx":0140
            Left            =   7140
            List            =   "frmCC_Colection_RITCARD12.frx":015C
            TabIndex        =   285
            Text            =   "--Pilih Kategori Telepon--"
            Top             =   1470
            Width           =   2445
         End
         Begin VB.ComboBox CmbStsKatHP1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmCC_Colection_RITCARD12.frx":01DA
            Left            =   7140
            List            =   "frmCC_Colection_RITCARD12.frx":01F6
            TabIndex        =   284
            Text            =   "--Pilih Kategori Telepon--"
            Top             =   1860
            Width           =   2460
         End
         Begin VB.ComboBox CmbStsKatHP2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmCC_Colection_RITCARD12.frx":0274
            Left            =   7140
            List            =   "frmCC_Colection_RITCARD12.frx":0290
            TabIndex        =   283
            Text            =   "--Pilih Kategori Telepon--"
            Top             =   2220
            Width           =   2460
         End
         Begin VB.ComboBox CmbStsKatHome2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmCC_Colection_RITCARD12.frx":030E
            Left            =   7140
            List            =   "frmCC_Colection_RITCARD12.frx":032A
            TabIndex        =   282
            Text            =   "--Pilih Kategori Telepon--"
            Top             =   720
            Width           =   2445
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FCFCFC&
            Height          =   795
            Left            =   2820
            TabIndex        =   276
            Top             =   3000
            Width           =   3015
            Begin TDBMask6Ctl.TDBMask TxtNoTelpReq 
               Height          =   255
               Left            =   720
               TabIndex        =   277
               Top             =   480
               Width           =   1935
               _Version        =   65536
               _ExtentX        =   3413
               _ExtentY        =   450
               Caption         =   "frmCC_Colection_RITCARD12.frx":03A8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmCC_Colection_RITCARD12.frx":0414
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               AllowSpace      =   -1
               AutoConvert     =   -1
               BackColor       =   12648384
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   -1
               DataProperty    =   0
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
               PromptChar      =   " "
               ReadOnly        =   -1
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "                    "
               Value           =   ""
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00004000&
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H003F9E0C&
               Height          =   255
               Left            =   2700
               TabIndex        =   281
               Top             =   480
               Width           =   195
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "No.Tlp:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   21
               Left            =   60
               TabIndex        =   280
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label TxtKategori 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   720
               TabIndex        =   279
               Top             =   180
               Width           =   1950
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Kat.Tlp:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   15
               Left            =   60
               TabIndex        =   278
               Top             =   180
               Width           =   1575
            End
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeAdd1 
            Height          =   345
            Left            =   4860
            TabIndex        =   288
            Top             =   840
            Visible         =   0   'False
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   609
            Caption         =   "frmCC_Colection_RITCARD12.frx":0456
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD12.frx":04C2
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   16777215
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            PromptChar      =   " "
            ReadOnly        =   1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                  "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeAdd2 
            Height          =   345
            Left            =   5100
            TabIndex        =   289
            Top             =   1230
            Visible         =   0   'False
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   609
            Caption         =   "frmCC_Colection_RITCARD12.frx":0504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD12.frx":0570
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   16777215
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            PromptChar      =   " "
            ReadOnly        =   1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                  "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeAdd1A 
            Height          =   345
            Left            =   1020
            TabIndex        =   290
            Top             =   1080
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   609
            Caption         =   "frmCC_Colection_RITCARD12.frx":05B2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD12.frx":061E
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   12648384
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            PromptChar      =   " "
            ReadOnly        =   1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                  "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeAdd2A 
            Height          =   345
            Left            =   1020
            TabIndex        =   291
            Top             =   1450
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   609
            Caption         =   "frmCC_Colection_RITCARD12.frx":0660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD12.frx":06CC
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   12648384
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            PromptChar      =   " "
            ReadOnly        =   1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                  "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileAdd1 
            Height          =   345
            Left            =   6300
            TabIndex        =   292
            Top             =   1590
            Visible         =   0   'False
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   609
            Caption         =   "frmCC_Colection_RITCARD12.frx":070E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD12.frx":077A
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
            ForeColor       =   0
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                  "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileAdd2 
            Height          =   345
            Left            =   5820
            TabIndex        =   293
            Top             =   1920
            Visible         =   0   'False
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   609
            Caption         =   "frmCC_Colection_RITCARD12.frx":07BC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD12.frx":0828
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
            ForeColor       =   0
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                  "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileAdd1A 
            Height          =   345
            Left            =   1020
            TabIndex        =   294
            Top             =   1830
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   609
            Caption         =   "frmCC_Colection_RITCARD12.frx":086A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD12.frx":08D6
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   12648384
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                  "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileAdd2A 
            Height          =   345
            Left            =   1020
            TabIndex        =   295
            Top             =   2190
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   609
            Caption         =   "frmCC_Colection_RITCARD12.frx":0918
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD12.frx":0984
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   12648384
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                  "
            Value           =   ""
         End
         Begin RichTextLib.RichTextBox AddrNow 
            Height          =   735
            Left            =   120
            TabIndex        =   296
            Top             =   2760
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   1296
            _Version        =   393217
            BackColor       =   12648384
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmCC_Colection_RITCARD12.frx":09C6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin TDBMask6Ctl.TDBMask txtHomeAdd1 
            Height          =   345
            Left            =   4620
            TabIndex        =   297
            Top             =   135
            Visible         =   0   'False
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   609
            Caption         =   "frmCC_Colection_RITCARD12.frx":0A47
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD12.frx":0AB3
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   16777215
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            PromptChar      =   " "
            ReadOnly        =   1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                  "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtHomeAdd2 
            Height          =   345
            Left            =   4740
            TabIndex        =   298
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   609
            Caption         =   "frmCC_Colection_RITCARD12.frx":0AF5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD12.frx":0B61
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   16777215
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            PromptChar      =   " "
            ReadOnly        =   1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                  "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtHomeAdd1A 
            Height          =   345
            Left            =   1020
            TabIndex        =   299
            Top             =   360
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   609
            Caption         =   "frmCC_Colection_RITCARD12.frx":0BA3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD12.frx":0C0F
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   12648384
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            PromptChar      =   " "
            ReadOnly        =   1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                  "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtHomeAdd2A 
            Height          =   345
            Left            =   1020
            TabIndex        =   300
            Top             =   720
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   609
            Caption         =   "frmCC_Colection_RITCARD12.frx":0C51
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD12.frx":0CBD
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   12648384
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            PromptChar      =   " "
            ReadOnly        =   1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                  "
            Value           =   ""
         End
         Begin VB.Label LblBlacklistAddHP2 
            Alignment       =   2  'Center
            BackColor       =   &H00004000&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003F9E0C&
            Height          =   255
            Left            =   7920
            TabIndex        =   313
            Top             =   2220
            Width           =   195
         End
         Begin VB.Label LblBlacklistAddHP1 
            Alignment       =   2  'Center
            BackColor       =   &H00004000&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003F9E0C&
            Height          =   255
            Left            =   7920
            TabIndex        =   312
            Top             =   1860
            Width           =   195
         End
         Begin VB.Label LblBlacklistAddOffice2 
            Alignment       =   2  'Center
            BackColor       =   &H00004000&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003F9E0C&
            Height          =   255
            Left            =   7920
            TabIndex        =   311
            Top             =   1500
            Width           =   195
         End
         Begin VB.Label LblBlacklistAddOffice1 
            Alignment       =   2  'Center
            BackColor       =   &H00004000&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003F9E0C&
            Height          =   255
            Left            =   7920
            TabIndex        =   310
            Top             =   1140
            Width           =   195
         End
         Begin VB.Label LblBlacklistAddHome2 
            Alignment       =   2  'Center
            BackColor       =   &H00004000&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003F9E0C&
            Height          =   255
            Left            =   7920
            TabIndex        =   309
            Top             =   780
            Width           =   195
         End
         Begin VB.Label LblBlacklistAddHome1 
            Alignment       =   2  'Center
            BackColor       =   &H00004000&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003F9E0C&
            Height          =   255
            Left            =   7920
            TabIndex        =   308
            Top             =   420
            Width           =   195
         End
         Begin VB.Label Label19 
            BackColor       =   &H00ABE18E&
            BackStyle       =   0  'Transparent
            Caption         =   "Add  Adress:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   307
            Top             =   3480
            Width           =   795
         End
         Begin VB.Label label1 
            BackColor       =   &H00ABE18E&
            BackStyle       =   0  'Transparent
            Caption         =   "HP I"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   0
            TabIndex        =   306
            Top             =   1830
            Width           =   765
         End
         Begin VB.Label label1 
            BackColor       =   &H00ABE18E&
            BackStyle       =   0  'Transparent
            Caption         =   "HP II"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   305
            Top             =   2190
            Width           =   765
         End
         Begin VB.Label label1 
            BackColor       =   &H00ABE18E&
            BackStyle       =   0  'Transparent
            Caption         =   "P Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   20
            Left            =   0
            TabIndex        =   304
            Top             =   360
            Width           =   885
         End
         Begin VB.Label label1 
            BackColor       =   &H00ABE18E&
            BackStyle       =   0  'Transparent
            Caption         =   "R Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   19
            Left            =   0
            TabIndex        =   303
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label label1 
            BackColor       =   &H00ABE18E&
            BackStyle       =   0  'Transparent
            Caption         =   "P Code Bca"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   17
            Left            =   0
            TabIndex        =   302
            Top             =   1080
            Width           =   1005
         End
         Begin VB.Label label1 
            BackColor       =   &H00ABE18E&
            BackStyle       =   0  'Transparent
            Caption         =   "ROC Bca"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   16
            Left            =   0
            TabIndex        =   301
            Top             =   1500
            Width           =   1125
         End
      End
      Begin VB.ComboBox cbolastcall 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFE4E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmCC_Colection_RITCARD12.frx":0CFF
         Left            =   8190
         List            =   "frmCC_Colection_RITCARD12.frx":0D06
         Locked          =   -1  'True
         TabIndex        =   146
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chk_aktif 
         BackColor       =   &H00ABE18E&
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8940
         TabIndex        =   96
         Top             =   960
         Width           =   1035
      End
      Begin VB.TextBox TXTRUMUS 
         Height          =   315
         Left            =   12960
         TabIndex        =   95
         Top             =   360
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txthasil 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   9960
         TabIndex        =   94
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox CmbBaseOn 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmCC_Colection_RITCARD12.frx":0D15
         Left            =   8280
         List            =   "frmCC_Colection_RITCARD12.frx":0D17
         TabIndex        =   93
         Top             =   720
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton cmd_logcomplaint 
         Caption         =   "Create Complaint"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8400
         TabIndex        =   92
         Top             =   960
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton CmdRequestNumber 
         Caption         =   "Request Number"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8520
         TabIndex        =   91
         Top             =   960
         Width           =   1995
      End
      Begin TDBNumber6Ctl.TDBNumber LblMinPayment 
         Height          =   375
         Left            =   10740
         TabIndex        =   97
         Top             =   1560
         Width           =   1740
         _Version        =   65536
         _ExtentX        =   3069
         _ExtentY        =   661
         Calculator      =   "frmCC_Colection_RITCARD12.frx":0D19
         Caption         =   "frmCC_Colection_RITCARD12.frx":0D39
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":0DA5
         Keys            =   "frmCC_Colection_RITCARD12.frx":0DC3
         Spin            =   "frmCC_Colection_RITCARD12.frx":0E0D
         AlignHorizontal =   2
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   0
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   65280
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
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin MSComctlLib.ListView LstDoubleId 
         Height          =   810
         Left            =   120
         TabIndex        =   98
         Top             =   1320
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   1429
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   10147522
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin TDBNumber6Ctl.TDBNumber tdbmaxad 
         Height          =   255
         Left            =   8760
         TabIndex        =   99
         Top             =   1380
         Visible         =   0   'False
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_RITCARD12.frx":0E35
         Caption         =   "frmCC_Colection_RITCARD12.frx":0E55
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":0EC1
         Keys            =   "frmCC_Colection_RITCARD12.frx":0EDF
         Spin            =   "frmCC_Colection_RITCARD12.frx":0F29
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber tdbminad 
         Height          =   255
         Left            =   9240
         TabIndex        =   100
         Top             =   1560
         Visible         =   0   'False
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_RITCARD12.frx":0F51
         Caption         =   "frmCC_Colection_RITCARD12.frx":0F71
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":0FDD
         Keys            =   "frmCC_Colection_RITCARD12.frx":0FFB
         Spin            =   "frmCC_Colection_RITCARD12.frx":1045
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber Tdbbalance 
         Height          =   255
         Left            =   240
         TabIndex        =   101
         Top             =   1320
         Visible         =   0   'False
         Width           =   105
         _Version        =   65536
         _ExtentX        =   185
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_RITCARD12.frx":106D
         Caption         =   "frmCC_Colection_RITCARD12.frx":108D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":10F9
         Keys            =   "frmCC_Colection_RITCARD12.frx":1117
         Spin            =   "frmCC_Colection_RITCARD12.frx":1161
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   12648384
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber lblLimit 
         Height          =   255
         Left            =   1185
         TabIndex        =   102
         Top             =   2760
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_RITCARD12.frx":1189
         Caption         =   "frmCC_Colection_RITCARD12.frx":11A9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":1215
         Keys            =   "frmCC_Colection_RITCARD12.frx":1233
         Spin            =   "frmCC_Colection_RITCARD12.frx":127D
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15721696
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBDate6Ctl.TDBDate lblOpenDate 
         Height          =   255
         Left            =   9705
         TabIndex        =   103
         Top             =   1215
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   450
         Calendar        =   "frmCC_Colection_RITCARD12.frx":12A5
         Caption         =   "frmCC_Colection_RITCARD12.frx":13BD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":1429
         Keys            =   "frmCC_Colection_RITCARD12.frx":1447
         Spin            =   "frmCC_Colection_RITCARD12.frx":14A5
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15721696
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
      Begin TDBDate6Ctl.TDBDate lblBD 
         Height          =   255
         Left            =   9465
         TabIndex        =   104
         Top             =   1140
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   450
         Calendar        =   "frmCC_Colection_RITCARD12.frx":14CD
         Caption         =   "frmCC_Colection_RITCARD12.frx":15E5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":1651
         Keys            =   "frmCC_Colection_RITCARD12.frx":166F
         Spin            =   "frmCC_Colection_RITCARD12.frx":16CD
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15721696
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
      Begin TDBNumber6Ctl.TDBNumber TxtCurpri 
         Height          =   255
         Left            =   2985
         TabIndex        =   105
         Top             =   3120
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_RITCARD12.frx":16F5
         Caption         =   "frmCC_Colection_RITCARD12.frx":1715
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":1781
         Keys            =   "frmCC_Colection_RITCARD12.frx":179F
         Spin            =   "frmCC_Colection_RITCARD12.frx":17E9
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15721696
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
         ValueVT         =   1769473
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber TDBlate_fee 
         Height          =   255
         Left            =   2985
         TabIndex        =   106
         Top             =   3405
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_RITCARD12.frx":1811
         Caption         =   "frmCC_Colection_RITCARD12.frx":1831
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":189D
         Keys            =   "frmCC_Colection_RITCARD12.frx":18BB
         Spin            =   "frmCC_Colection_RITCARD12.frx":1905
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15721696
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber TxtInterest 
         Height          =   255
         Left            =   9495
         TabIndex        =   107
         Top             =   1080
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_RITCARD12.frx":192D
         Caption         =   "frmCC_Colection_RITCARD12.frx":194D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":19B9
         Keys            =   "frmCC_Colection_RITCARD12.frx":19D7
         Spin            =   "frmCC_Colection_RITCARD12.frx":1A21
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15721696
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber TDB_cur_bal 
         Height          =   255
         Left            =   9015
         TabIndex        =   108
         Top             =   1965
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_RITCARD12.frx":1A49
         Caption         =   "frmCC_Colection_RITCARD12.frx":1A69
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":1AD5
         Keys            =   "frmCC_Colection_RITCARD12.frx":1AF3
         Spin            =   "frmCC_Colection_RITCARD12.frx":1B3D
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15721696
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber LblPrompA 
         Height          =   255
         Left            =   9135
         TabIndex        =   109
         Top             =   1920
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_RITCARD12.frx":1B65
         Caption         =   "frmCC_Colection_RITCARD12.frx":1B85
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":1BF1
         Keys            =   "frmCC_Colection_RITCARD12.frx":1C0F
         Spin            =   "frmCC_Colection_RITCARD12.frx":1C59
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   15721696
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   0
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
         ValueVT         =   1572869
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
         Height          =   255
         Left            =   0
         TabIndex        =   110
         Top             =   1320
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   441
         Calculator      =   "frmCC_Colection_RITCARD12.frx":1C81
         Caption         =   "frmCC_Colection_RITCARD12.frx":1CA1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":1D0D
         Keys            =   "frmCC_Colection_RITCARD12.frx":1D2B
         Spin            =   "frmCC_Colection_RITCARD12.frx":1D75
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
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin Threed.SSCommand CmdKeep 
         Height          =   600
         Left            =   9000
         TabIndex        =   111
         Top             =   1080
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1058
         _Version        =   196610
         Font3D          =   2
         MousePointer    =   16
         ForeColor       =   8388608
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
         Picture         =   "frmCC_Colection_RITCARD12.frx":1D9D
         AutoSize        =   1
         Alignment       =   8
         PictureAlignment=   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   720
         Index           =   5
         Left            =   120
         TabIndex        =   112
         Top             =   4200
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1270
         _Version        =   196610
         Font3D          =   2
         MousePointer    =   16
         ForeColor       =   8388608
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
         Picture         =   "frmCC_Colection_RITCARD12.frx":2A57
         Caption         =   "Offers"
         AutoSize        =   1
         Alignment       =   8
         PictureAlignment=   1
      End
      Begin RichTextLib.RichTextBox lblOfficeAddr1 
         Height          =   795
         Left            =   9195
         TabIndex        =   160
         Top             =   2340
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   1402
         _Version        =   393217
         BackColor       =   -2147483645
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmCC_Colection_RITCARD12.frx":35F3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox lblAddr 
         Height          =   555
         Left            =   8520
         TabIndex        =   162
         Top             =   1155
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   979
         _Version        =   393217
         BackColor       =   -2147483645
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmCC_Colection_RITCARD12.frx":366F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TDBMask6Ctl.TDBMask txtHomeNo2 
         Height          =   255
         Left            =   9060
         TabIndex        =   166
         Top             =   1935
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_RITCARD12.frx":36EB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_RITCARD12.frx":3757
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   -2147483645
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                    "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtOfficeNo2 
         Height          =   255
         Left            =   9780
         TabIndex        =   167
         Top             =   915
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_RITCARD12.frx":3799
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_RITCARD12.frx":3805
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   -2147483645
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                    "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtMobileNo2 
         Height          =   255
         Left            =   1020
         TabIndex        =   168
         Top             =   1575
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_RITCARD12.frx":3847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_RITCARD12.frx":38B3
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   -2147483645
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                    "
         Value           =   ""
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Home II"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   8760
         TabIndex        =   170
         Top             =   855
         Width           =   735
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "HP II"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   169
         Top             =   1545
         Width           =   735
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Office Add"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   8415
         TabIndex        =   161
         Top             =   2340
         Width           =   720
      End
      Begin VB.Label Label31 
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Speak With"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   8640
         TabIndex        =   147
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label lblstop_time 
         Caption         =   "Label9"
         Height          =   255
         Left            =   9120
         TabIndex        =   145
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lbltime_save 
         Caption         =   "Label9"
         Height          =   375
         Left            =   7920
         TabIndex        =   144
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Assg Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   21
         Left            =   9165
         TabIndex        =   142
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   8550
         TabIndex        =   141
         Top             =   930
         Width           =   1545
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H003F9E0C&
         Caption         =   "MIN.PAYMENT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10260
         TabIndex        =   140
         Top             =   1920
         Width           =   1740
      End
      Begin VB.Label LblCycle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   9900
         TabIndex        =   139
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label LblMap 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   8880
         TabIndex        =   138
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H003F9E0C&
         Caption         =   "CYCLE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   9900
         TabIndex        =   137
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H003F9E0C&
         Caption         =   "MAP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8160
         TabIndex        =   136
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         Caption         =   "Max A.d"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   135
         Top             =   2520
         Visible         =   0   'False
         Width           =   840
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   7860
         TabIndex        =   134
         Top             =   420
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         Caption         =   "Min A.d"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   8160
         TabIndex        =   133
         Top             =   480
         Visible         =   0   'False
         Width           =   840
         WordWrap        =   -1  'True
      End
      Begin VB.Label Txtperiod 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFE4E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1185
         TabIndex        =   132
         Top             =   3030
         Width           =   1545
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Limit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   255
         TabIndex        =   131
         Top             =   2805
         Width           =   1245
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   130
         Top             =   3060
         Width           =   885
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Wo Date"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   8760
         TabIndex        =   129
         Top             =   1140
         Width           =   885
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Open Date"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7920
         TabIndex        =   128
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Tag"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   22
         Left            =   2040
         TabIndex        =   127
         Top             =   3990
         Width           =   885
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Segment"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   20
         Left            =   2040
         TabIndex        =   126
         Top             =   3720
         Width           =   885
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2985
         TabIndex        =   125
         Top             =   3990
         Width           =   1545
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   2985
         TabIndex        =   124
         Top             =   3690
         Width           =   1545
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Late Fee"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   123
         Top             =   3435
         Width           =   885
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Cur  Pri"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   2040
         TabIndex        =   122
         Top             =   3150
         Width           =   885
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Curr Bal"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   8400
         TabIndex        =   121
         Top             =   1125
         Width           =   885
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Interest"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   16
         Left            =   8760
         TabIndex        =   120
         Top             =   885
         Width           =   885
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Principle"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   8280
         TabIndex        =   119
         Top             =   240
         Width           =   765
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H003F9E0C&
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   7
         Left            =   8400
         TabIndex        =   118
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblClass 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8760
         TabIndex        =   117
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "ID No"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8280
         TabIndex        =   116
         Top             =   390
         Width           =   720
      End
      Begin VB.Label lblID 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFE4E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8940
         TabIndex        =   115
         Top             =   2040
         Width           =   3030
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H003F9E0C&
         Caption         =   "HOT PR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   8745
         TabIndex        =   114
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H003F9E0C&
         Caption         =   "CPA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   2
         Left            =   8415
         TabIndex        =   113
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Timer TimerBlinkCPA 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8760
      Top             =   6960
   End
   Begin VB.Timer TimerBlink 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   10080
      Top             =   7080
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   11025
      Left            =   -30
      TabIndex        =   28
      Top             =   480
      Width           =   19260
      _ExtentX        =   33973
      _ExtentY        =   19447
      _Version        =   196610
      Font3D          =   1
      ForeColor       =   12583104
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCFCFC&
         BorderStyle     =   0  'None
         Caption         =   "Frame11"
         ForeColor       =   &H80000008&
         Height          =   10875
         Left            =   6855
         TabIndex        =   188
         Top             =   0
         Width           =   13095
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H00B8E2D4&
            ForeColor       =   &H80000008&
            Height          =   1725
            Left            =   6060
            TabIndex        =   259
            Top             =   7200
            Visible         =   0   'False
            Width           =   5805
            Begin VB.Timer Timer_cek_inbox 
               Enabled         =   0   'False
               Interval        =   30000
               Left            =   4020
               Top             =   420
            End
            Begin VB.OptionButton Option9 
               BackColor       =   &H00B8E2D4&
               Caption         =   "Inbox"
               Height          =   255
               Left            =   4710
               TabIndex        =   261
               Top             =   120
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton Option10 
               BackColor       =   &H00B8E2D4&
               Caption         =   "Send"
               Height          =   255
               Left            =   4710
               TabIndex        =   260
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.Frame FrmPayment 
            Appearance      =   0  'Flat
            BackColor       =   &H00FCFCFC&
            Caption         =   "PAYMENT DETAILS"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   4215
            Left            =   4260
            TabIndex        =   244
            Top             =   0
            Width           =   7380
            Begin TDBNumber6Ctl.TDBNumber txtSisaHutang 
               Height          =   255
               Left            =   5835
               TabIndex        =   245
               Top             =   885
               Width           =   1230
               _Version        =   65536
               _ExtentX        =   2170
               _ExtentY        =   450
               Calculator      =   "frmCC_Colection_RITCARD12.frx":38F5
               Caption         =   "frmCC_Colection_RITCARD12.frx":3915
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":3981
               Keys            =   "frmCC_Colection_RITCARD12.frx":399F
               Spin            =   "frmCC_Colection_RITCARD12.frx":39E9
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483645
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###,##0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   0
               Format          =   "###,###,###,##0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999999999999
               MinValue        =   -999999999999
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
               MaxValueVT      =   6750213
               MinValueVT      =   3538949
            End
            Begin TDBNumber6Ctl.TDBNumber TxtAfterPay 
               Height          =   255
               Left            =   5835
               TabIndex        =   246
               Top             =   615
               Width           =   1230
               _Version        =   65536
               _ExtentX        =   2170
               _ExtentY        =   450
               Calculator      =   "frmCC_Colection_RITCARD12.frx":3A11
               Caption         =   "frmCC_Colection_RITCARD12.frx":3A31
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":3A9D
               Keys            =   "frmCC_Colection_RITCARD12.frx":3ABB
               Spin            =   "frmCC_Colection_RITCARD12.frx":3B05
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483645
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###,##0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   0
               Format          =   "###,###,###,##0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999999999999
               MinValue        =   -999999999999
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
               MaxValueVT      =   6750213
               MinValueVT      =   3538949
            End
            Begin TDBNumber6Ctl.TDBNumber TxtPayment2 
               Height          =   255
               Left            =   5835
               TabIndex        =   247
               Top             =   330
               Width           =   1230
               _Version        =   65536
               _ExtentX        =   2170
               _ExtentY        =   450
               Calculator      =   "frmCC_Colection_RITCARD12.frx":3B2D
               Caption         =   "frmCC_Colection_RITCARD12.frx":3B4D
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":3BB9
               Keys            =   "frmCC_Colection_RITCARD12.frx":3BD7
               Spin            =   "frmCC_Colection_RITCARD12.frx":3C21
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483645
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###,##0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   0
               Format          =   "###,###,###,##0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   -99999999999
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
            Begin MSComctlLib.ListView listview1 
               Height          =   3750
               Index           =   0
               Left            =   165
               TabIndex        =   248
               Top             =   345
               Width           =   4590
               _ExtentX        =   8096
               _ExtentY        =   6615
               View            =   3
               LabelEdit       =   1
               SortOrder       =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   0
               BackColor       =   12632256
               BorderStyle     =   1
               Appearance      =   0
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
            Begin TDBNumber6Ctl.TDBNumber TxtLPAPayment 
               Height          =   255
               Left            =   5835
               TabIndex        =   249
               Top             =   1440
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   450
               Calculator      =   "frmCC_Colection_RITCARD12.frx":3C49
               Caption         =   "frmCC_Colection_RITCARD12.frx":3C69
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":3CD5
               Keys            =   "frmCC_Colection_RITCARD12.frx":3CF3
               Spin            =   "frmCC_Colection_RITCARD12.frx":3D3D
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483645
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
               ReadOnly        =   -1
               Separator       =   ","
               ShowContextMenu =   -1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBDate6Ctl.TDBDate TxtLPDPayment 
               Height          =   255
               Left            =   5835
               TabIndex        =   250
               Top             =   1155
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
               _ExtentY        =   450
               Calendar        =   "frmCC_Colection_RITCARD12.frx":3D65
               Caption         =   "frmCC_Colection_RITCARD12.frx":3E7D
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":3EE9
               Keys            =   "frmCC_Colection_RITCARD12.frx":3F07
               Spin            =   "frmCC_Colection_RITCARD12.frx":3F65
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483645
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
            Begin Threed.SSCommand CmddetailPayment 
               Height          =   675
               Left            =   8775
               TabIndex        =   251
               Top             =   1800
               Visible         =   0   'False
               Width           =   2220
               _ExtentX        =   3916
               _ExtentY        =   1191
               _Version        =   196610
               MousePointer    =   16
               ForeColor       =   4210752
               PictureMaskColor=   -2147483644
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCC_Colection_RITCARD12.frx":3F8D
               Caption         =   "&"
               AutoSize        =   1
               ButtonStyle     =   2
               BevelWidth      =   0
            End
            Begin Threed.SSCommand Command2 
               Height          =   675
               Left            =   5490
               TabIndex        =   252
               ToolTipText     =   "SMS"
               Top             =   2400
               Visible         =   0   'False
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   1191
               _Version        =   196610
               MousePointer    =   16
               ForeColor       =   4210752
               PictureMaskColor=   -2147483644
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCC_Colection_RITCARD12.frx":6BE3
               Caption         =   "&"
               AutoSize        =   1
               ButtonStyle     =   2
               BevelWidth      =   0
            End
            Begin Threed.SSCommand SSCommand1 
               Height          =   675
               Index           =   3
               Left            =   5010
               TabIndex        =   253
               Top             =   3270
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   1191
               _Version        =   196610
               Font3D          =   2
               MousePointer    =   16
               ForeColor       =   12582912
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
               Picture         =   "frmCC_Colection_RITCARD12.frx":966C
               Caption         =   "&"
               AutoSize        =   1
               Alignment       =   4
               ButtonStyle     =   2
               BevelWidth      =   0
            End
            Begin VB.Label Label10 
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Jml PTP"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   0
               Left            =   4875
               TabIndex        =   258
               Top             =   330
               Width           =   885
            End
            Begin VB.Label Label13 
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Jml Dibayar:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4875
               TabIndex        =   257
               Top             =   615
               Width           =   1005
            End
            Begin VB.Label Label15 
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Sisa"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   4875
               TabIndex        =   256
               Top             =   885
               Width           =   1005
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "LPA"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   17
               Left            =   4875
               TabIndex        =   255
               Top             =   1440
               Width           =   885
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "LPD"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   18
               Left            =   4875
               TabIndex        =   254
               Top             =   1155
               Width           =   885
            End
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   2
            Left            =   12435
            TabIndex        =   242
            Top             =   0
            Visible         =   0   'False
            Width           =   2895
            Begin VB.Image Image1 
               Height          =   375
               Index           =   2
               Left            =   90
               Picture         =   "frmCC_Colection_RITCARD12.frx":C692
               Stretch         =   -1  'True
               Top             =   60
               Width           =   375
            End
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Emergency Contact"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   2
               Left            =   510
               TabIndex        =   243
               Top             =   120
               Width           =   2175
            End
         End
         Begin VB.Timer TimerBlinkDetailMapping 
            Interval        =   1000
            Left            =   3240
            Top             =   6720
         End
         Begin VB.TextBox TxtTelpKe 
            BackColor       =   &H0000C0C0&
            Height          =   285
            Left            =   540
            TabIndex        =   241
            Text            =   "NoPhone"
            Top             =   6180
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Frame Frame16 
            Appearance      =   0  'Flat
            BackColor       =   &H00FCFCFC&
            Caption         =   "PHONE INFO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   4215
            Left            =   75
            TabIndex        =   199
            Top             =   0
            Width           =   4095
            Begin VB.CommandButton Command1001 
               Caption         =   "Alternatif Icoll"
               Height          =   615
               Left            =   3240
               TabIndex        =   316
               Top             =   1560
               Width           =   855
            End
            Begin VB.TextBox Text11m 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   274
               Top             =   2370
               Width           =   1875
            End
            Begin VB.TextBox txtadd_phone 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Index           =   7
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   273
               Top             =   2040
               Width           =   1875
            End
            Begin VB.TextBox txtOfficeNo1m 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   272
               Top             =   540
               Width           =   1875
            End
            Begin VB.TextBox txtMobileNo1m 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   271
               Top             =   840
               Width           =   1875
            End
            Begin VB.TextBox txtadd_phone 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Index           =   6
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   270
               Top             =   1140
               Width           =   1875
            End
            Begin VB.TextBox txtadd_phone 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Index           =   5
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   269
               Top             =   1440
               Width           =   1875
            End
            Begin VB.TextBox txtadd_phone 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Index           =   4
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   268
               Top             =   1740
               Width           =   1875
            End
            Begin VB.TextBox txtHomeNo1m 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   267
               Top             =   240
               Width           =   1875
            End
            Begin VB.ComboBox CmbPhone 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               ItemData        =   "frmCC_Colection_RITCARD12.frx":DF2C
               Left            =   1140
               List            =   "frmCC_Colection_RITCARD12.frx":DF33
               Locked          =   -1  'True
               TabIndex        =   210
               Text            =   "CmbPhone"
               Top             =   3075
               Width           =   1920
            End
            Begin VB.TextBox stshangup 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5280
               TabIndex        =   209
               Top             =   2520
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox txtHomeNo1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   208
               Top             =   240
               Width           =   1875
            End
            Begin VB.TextBox txtadd_phone 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Index           =   2
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   207
               Top             =   1740
               Width           =   1875
            End
            Begin VB.TextBox txtadd_phone 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Index           =   1
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   206
               Top             =   1440
               Width           =   1875
            End
            Begin VB.TextBox txtadd_phone 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Index           =   0
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   205
               Top             =   1140
               Width           =   1875
            End
            Begin VB.TextBox txtMobileNo1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   204
               Top             =   840
               Width           =   1875
            End
            Begin VB.TextBox txtOfficeNo1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   203
               Top             =   540
               Width           =   1875
            End
            Begin VB.TextBox txtadd_phone 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Index           =   3
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   202
               Top             =   2040
               Width           =   1875
            End
            Begin VB.TextBox ec_desc 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   201
               Top             =   2715
               Width           =   1875
            End
            Begin VB.TextBox Text11 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Height          =   285
               Left            =   1140
               Locked          =   -1  'True
               TabIndex        =   200
               Top             =   2370
               Width           =   1875
            End
            Begin TDBMask6Ctl.TDBMask txtHomeNo2A 
               Height          =   255
               Left            =   7380
               TabIndex        =   211
               Top             =   3075
               Width           =   1755
               _Version        =   65536
               _ExtentX        =   3096
               _ExtentY        =   450
               Caption         =   "frmCC_Colection_RITCARD12.frx":DF3C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmCC_Colection_RITCARD12.frx":DFA8
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               AllowSpace      =   -1
               AutoConvert     =   -1
               BackColor       =   12648384
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   -1
               DataProperty    =   0
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
               PromptChar      =   " "
               ReadOnly        =   -1
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "                    "
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask txtOfficeNo2A 
               Height          =   255
               Left            =   1140
               TabIndex        =   212
               Top             =   4245
               Width           =   1755
               _Version        =   65536
               _ExtentX        =   3096
               _ExtentY        =   450
               Caption         =   "frmCC_Colection_RITCARD12.frx":DFEA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmCC_Colection_RITCARD12.frx":E056
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               AllowSpace      =   -1
               AutoConvert     =   -1
               BackColor       =   12648384
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   -1
               DataProperty    =   0
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
               PromptChar      =   " "
               ReadOnly        =   -1
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "                    "
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask txtMobileNo1A 
               Height          =   255
               Left            =   1380
               TabIndex        =   213
               Top             =   4305
               Width           =   1755
               _Version        =   65536
               _ExtentX        =   3096
               _ExtentY        =   450
               Caption         =   "frmCC_Colection_RITCARD12.frx":E098
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmCC_Colection_RITCARD12.frx":E104
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               AllowSpace      =   -1
               AutoConvert     =   -1
               BackColor       =   12648384
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   -1
               DataProperty    =   0
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
               PromptChar      =   " "
               ReadOnly        =   -1
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "                    "
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask txtMobileNo2A 
               Height          =   255
               Left            =   2100
               TabIndex        =   214
               Top             =   4230
               Width           =   1755
               _Version        =   65536
               _ExtentX        =   3096
               _ExtentY        =   450
               Caption         =   "frmCC_Colection_RITCARD12.frx":E146
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmCC_Colection_RITCARD12.frx":E1B2
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               AllowSpace      =   -1
               AutoConvert     =   -1
               BackColor       =   12648384
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   -1
               DataProperty    =   0
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
               PromptChar      =   " "
               ReadOnly        =   -1
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "                    "
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask txtHomeNo1A 
               Height          =   255
               Left            =   2880
               TabIndex        =   215
               Top             =   4320
               Width           =   555
               _Version        =   65536
               _ExtentX        =   979
               _ExtentY        =   450
               Caption         =   "frmCC_Colection_RITCARD12.frx":E1F4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmCC_Colection_RITCARD12.frx":E260
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               AllowSpace      =   -1
               AutoConvert     =   -1
               BackColor       =   15721696
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   -1
               DataProperty    =   0
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
               PromptChar      =   " "
               ReadOnly        =   -1
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "                    "
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask txtOfficeNo1A 
               Height          =   255
               Left            =   1140
               TabIndex        =   216
               Top             =   4275
               Width           =   1755
               _Version        =   65536
               _ExtentX        =   3096
               _ExtentY        =   450
               Caption         =   "frmCC_Colection_RITCARD12.frx":E2A2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmCC_Colection_RITCARD12.frx":E30E
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               AllowSpace      =   -1
               AutoConvert     =   -1
               BackColor       =   12648384
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   -1
               DataProperty    =   0
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
               PromptChar      =   " "
               ReadOnly        =   -1
               ShowContextMenu =   0
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "                    "
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask tdbtelptrace 
               Height          =   255
               Left            =   1125
               TabIndex        =   217
               Top             =   2385
               Visible         =   0   'False
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   450
               Caption         =   "frmCC_Colection_RITCARD12.frx":E350
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmCC_Colection_RITCARD12.frx":E3BC
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               AllowSpace      =   -1
               AutoConvert     =   -1
               BackColor       =   16777215
               BorderStyle     =   0
               ClipMode        =   0
               CursorPosition  =   -1
               DataProperty    =   0
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
               PromptChar      =   " "
               ReadOnly        =   -1
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "                    "
               Value           =   ""
            End
            Begin Threed.SSCommand SSCommand1 
               Height          =   675
               Index           =   0
               Left            =   2400
               TabIndex        =   218
               Top             =   3480
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   1191
               _Version        =   196610
               MousePointer    =   16
               ForeColor       =   4210752
               PictureMaskColor=   -2147483644
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCC_Colection_RITCARD12.frx":E3FE
               Caption         =   "&"
               AutoSize        =   1
               ButtonStyle     =   2
               BevelWidth      =   0
            End
            Begin Threed.SSCommand SSCommand1 
               Height          =   675
               Index           =   1
               Left            =   3225
               TabIndex        =   219
               Top             =   3480
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   1191
               _Version        =   196610
               MousePointer    =   16
               ForeColor       =   4210752
               PictureMaskColor=   -2147483644
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCC_Colection_RITCARD12.frx":114DD
               Caption         =   "&"
               AutoSize        =   1
               ButtonStyle     =   2
               BevelWidth      =   0
            End
            Begin TDBMask6Ctl.TDBMask TxtAdditional 
               Height          =   255
               Left            =   6420
               TabIndex        =   220
               Top             =   3510
               Visible         =   0   'False
               Width           =   555
               _Version        =   65536
               _ExtentX        =   979
               _ExtentY        =   450
               Caption         =   "frmCC_Colection_RITCARD12.frx":1473A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "frmCC_Colection_RITCARD12.frx":147A6
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               AllowSpace      =   -1
               AutoConvert     =   -1
               BackColor       =   15721696
               BorderStyle     =   1
               ClipMode        =   0
               CursorPosition  =   -1
               DataProperty    =   0
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#############"
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
            Begin Threed.SSCommand cmd_req_telp 
               Height          =   600
               Left            =   3255
               TabIndex        =   221
               ToolTipText     =   "Request Number"
               Top             =   255
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   1058
               _Version        =   196610
               MousePointer    =   16
               ForeColor       =   4210752
               PictureMaskColor=   -2147483644
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCC_Colection_RITCARD12.frx":147E8
               Caption         =   "&"
               AutoSize        =   1
               ButtonStyle     =   2
               BevelWidth      =   0
            End
            Begin VB.Label LblBlacklistHp2 
               Alignment       =   2  'Center
               BackColor       =   &H00004000&
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H003F9E0C&
               Height          =   255
               Left            =   5820
               TabIndex        =   240
               Top             =   2295
               Visible         =   0   'False
               Width           =   195
            End
            Begin VB.Label LblBlacklistHp1 
               Alignment       =   2  'Center
               BackColor       =   &H00004000&
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H003F9E0C&
               Height          =   255
               Left            =   6540
               TabIndex        =   239
               Top             =   2385
               Visible         =   0   'False
               Width           =   195
            End
            Begin VB.Label LblBlacklistOfficeno2 
               Alignment       =   2  'Center
               BackColor       =   &H00004000&
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H003F9E0C&
               Height          =   255
               Left            =   6540
               TabIndex        =   238
               Top             =   2100
               Visible         =   0   'False
               Width           =   195
            End
            Begin VB.Label LblBlacklistOffice1 
               Alignment       =   2  'Center
               BackColor       =   &H00004000&
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H003F9E0C&
               Height          =   255
               Left            =   6540
               TabIndex        =   237
               Top             =   1740
               Visible         =   0   'False
               Width           =   195
            End
            Begin VB.Label LblBlacklistHome2 
               Alignment       =   2  'Center
               BackColor       =   &H00004000&
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H003F9E0C&
               Height          =   255
               Left            =   6540
               TabIndex        =   236
               Top             =   1440
               Visible         =   0   'False
               Width           =   195
            End
            Begin VB.Label LblBlakcListHome1 
               Alignment       =   2  'Center
               BackColor       =   &H00004000&
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H003F9E0C&
               Height          =   255
               Left            =   6540
               TabIndex        =   235
               Top             =   1110
               Width           =   195
            End
            Begin VB.Label Label22 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Add."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   2
               Left            =   6360
               TabIndex        =   234
               Top             =   3660
               Visible         =   0   'False
               Width           =   735
               WordWrap        =   -1  'True
            End
            Begin VB.Label LblMother 
               Appearance      =   0  'Flat
               BackColor       =   &H00EFE4E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2835
               TabIndex        =   233
               Top             =   4260
               Width           =   540
            End
            Begin VB.Label Label22 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Mother Name"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   232
               Top             =   4260
               Width           =   735
               WordWrap        =   -1  'True
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Dest Call"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   9
               Left            =   90
               TabIndex        =   231
               Top             =   3105
               Width           =   1005
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "New Num"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   7
               Left            =   90
               TabIndex        =   230
               Top             =   540
               Width           =   855
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Old Num"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   6
               Left            =   90
               TabIndex        =   229
               Top             =   240
               Width           =   735
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Office Num"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   4
               Left            =   90
               TabIndex        =   228
               Top             =   840
               Width           =   855
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "HP add"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   23
               Left            =   90
               TabIndex        =   227
               Top             =   1755
               Width           =   960
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Home add"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   24
               Left            =   90
               TabIndex        =   226
               Top             =   1140
               Width           =   960
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Office add"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   25
               Left            =   90
               TabIndex        =   225
               Top             =   1440
               Width           =   960
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Other"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   26
               Left            =   240
               TabIndex        =   224
               Top             =   1920
               Width           =   1035
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "EC Info"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   27
               Left            =   90
               TabIndex        =   223
               Top             =   2730
               Width           =   1035
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "EC Num"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   28
               Left            =   105
               TabIndex        =   222
               Top             =   2370
               Width           =   855
            End
         End
         Begin VB.Frame Frame10 
            Appearance      =   0  'Flat
            BackColor       =   &H00FCFCFC&
            Caption         =   "HISTORY REMARKS"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   6015
            Left            =   120
            TabIndex        =   193
            Top             =   4305
            Width           =   12105
            Begin VB.Timer Timer1 
               Interval        =   4000
               Left            =   1920
               Top             =   1320
            End
            Begin VB.TextBox getservertime 
               Height          =   285
               Left            =   840
               TabIndex        =   196
               Text            =   "Text5"
               Top             =   3120
               Visible         =   0   'False
               Width           =   2415
            End
            Begin VB.TextBox txtgetnomor 
               Height          =   285
               Left            =   960
               TabIndex        =   195
               Text            =   "Text5"
               Top             =   2280
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Timer TimerOfferingDiscon 
               Interval        =   1500
               Left            =   3000
               Top             =   1320
            End
            Begin VB.TextBox Text6 
               Height          =   285
               Left            =   600
               TabIndex        =   194
               Text            =   "Text6"
               Top             =   1500
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Timer TimerBlinkSms 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   2400
               Top             =   1260
            End
            Begin MSComctlLib.ListView listview1 
               Height          =   5565
               Index           =   1
               Left            =   240
               TabIndex        =   197
               Top             =   285
               Width           =   11745
               _ExtentX        =   20717
               _ExtentY        =   9816
               View            =   3
               LabelEdit       =   1
               SortOrder       =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   0
               BackColor       =   12632256
               BorderStyle     =   1
               Appearance      =   0
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
            Begin VB.Label Label16 
               Caption         =   "Label16"
               Height          =   495
               Left            =   240
               TabIndex        =   198
               Top             =   480
               Visible         =   0   'False
               Width           =   1215
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FCFCFC&
            Height          =   1110
            Left            =   10260
            TabIndex        =   189
            Top             =   -1065
            Visible         =   0   'False
            Width           =   2050
            Begin VB.TextBox Text5 
               Height          =   285
               Left            =   405
               TabIndex        =   191
               Top             =   1785
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.TextBox Text9 
               Height          =   285
               Left            =   360
               TabIndex        =   190
               Top             =   2460
               Visible         =   0   'False
               Width           =   675
            End
            Begin Threed.SSCommand SSCommand1 
               Height          =   675
               Index           =   4
               Left            =   1020
               TabIndex        =   192
               Top             =   165
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   1191
               _Version        =   196610
               MousePointer    =   16
               ForeColor       =   4210752
               PictureMaskColor=   -2147483644
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "CPA"
               AutoSize        =   1
               ButtonStyle     =   2
               PictureAlignment=   1
               BevelWidth      =   0
            End
         End
         Begin VB.Label LBLEXP 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   7980
            TabIndex        =   264
            Top             =   7080
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   263
            Top             =   3960
            Width           =   735
         End
         Begin VB.Label lbl_agentlama 
            BackStyle       =   0  'Transparent
            Caption         =   "Agent Lama"
            Height          =   375
            Left            =   1200
            TabIndex        =   262
            Top             =   3960
            Width           =   975
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00FCFCFC&
         Caption         =   "CALL RESULT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   135
         TabIndex        =   60
         Top             =   7890
         Width           =   6645
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "frmCC_Colection_RITCARD12.frx":17047
            Left            =   180
            List            =   "frmCC_Colection_RITCARD12.frx":17054
            Locked          =   -1  'True
            TabIndex        =   265
            Text            =   "PHONE"
            Top             =   1200
            Width           =   2145
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   3
            Left            =   0
            TabIndex        =   81
            Top             =   -480
            Width           =   7035
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Call Actvity"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   3
               Left            =   480
               TabIndex        =   82
               Top             =   60
               Width           =   1455
            End
            Begin VB.Image Image1 
               Height          =   375
               Index           =   3
               Left            =   75
               Picture         =   "frmCC_Colection_RITCARD12.frx":17068
               Stretch         =   -1  'True
               Top             =   30
               Width           =   375
            End
         End
         Begin VB.TextBox txtremarks 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1755
            Left            =   2460
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   70
            Top             =   540
            Width           =   3015
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00EFE4E0&
            Height          =   315
            ItemData        =   "frmCC_Colection_RITCARD12.frx":175B0
            Left            =   180
            List            =   "frmCC_Colection_RITCARD12.frx":175B2
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   2700
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.ComboBox cboaccount 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            ItemData        =   "frmCC_Colection_RITCARD12.frx":175B4
            Left            =   180
            List            =   "frmCC_Colection_RITCARD12.frx":175B6
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   540
            Width           =   2145
         End
         Begin TDBDate6Ctl.TDBDate cmbDateSch 
            Height          =   315
            Left            =   180
            TabIndex        =   61
            Top             =   1860
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
            _ExtentY        =   556
            Calendar        =   "frmCC_Colection_RITCARD12.frx":175B8
            Caption         =   "frmCC_Colection_RITCARD12.frx":176D0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD12.frx":1773C
            Keys            =   "frmCC_Colection_RITCARD12.frx":1775A
            Spin            =   "frmCC_Colection_RITCARD12.frx":177B8
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   12648447
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
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   1.12794198814265E-317
            CenturyMode     =   0
         End
         Begin TDBTime6Ctl.TDBTime cmbTimeSch 
            Height          =   315
            Left            =   1500
            TabIndex        =   62
            Top             =   1860
            Width           =   795
            _Version        =   65536
            _ExtentX        =   1402
            _ExtentY        =   556
            Caption         =   "frmCC_Colection_RITCARD12.frx":177E0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD12.frx":1784C
            Spin            =   "frmCC_Colection_RITCARD12.frx":1789C
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   12648447
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
            Text            =   "__:__"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   1.02960316199441E-317
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   2055
            Index           =   2
            Left            =   5640
            TabIndex        =   65
            Top             =   240
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   3625
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   8388608
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
            Picture         =   "frmCC_Colection_RITCARD12.frx":178C4
            Caption         =   "&"
            AutoSize        =   1
            Alignment       =   8
            ButtonStyle     =   2
            BevelWidth      =   0
         End
         Begin VB.Label label1 
            BackColor       =   &H00ABE18E&
            BackStyle       =   0  'Transparent
            Caption         =   "CONTACT MODE"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   29
            Left            =   180
            TabIndex        =   266
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label label1 
            BackColor       =   &H00ABE18E&
            BackStyle       =   0  'Transparent
            Caption         =   "GROUP CALL"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   12
            Left            =   180
            TabIndex        =   69
            Top             =   2460
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label label1 
            BackColor       =   &H00ABE18E&
            BackStyle       =   0  'Transparent
            Caption         =   "CALL STATUS"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   180
            TabIndex        =   68
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label39 
            BackColor       =   &H00ABE18E&
            BackStyle       =   0  'Transparent
            Caption         =   "CALL BACK DATE"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   180
            TabIndex        =   64
            Top             =   1620
            Width           =   1485
         End
         Begin VB.Label Label31 
            BackColor       =   &H00ABE18E&
            BackStyle       =   0  'Transparent
            Caption         =   "REMARKS"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   2460
            TabIndex        =   63
            Top             =   300
            Width           =   1275
         End
      End
      Begin VB.Frame Frame13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCFCFC&
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         ForeColor       =   &H80000008&
         Height          =   10935
         Left            =   15
         TabIndex        =   35
         Top             =   15
         Width           =   6885
         Begin VB.Frame frmPTP 
            Appearance      =   0  'Flat
            BackColor       =   &H00FCFCFC&
            Caption         =   "PTP [ Promise To Pay ]"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   3540
            Left            =   120
            TabIndex        =   49
            Top             =   4290
            Width           =   6615
            Begin VB.CheckBox C_PTP 
               BackColor       =   &H00FCFCFC&
               Caption         =   "PTP"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1515
               TabIndex        =   156
               Top             =   1995
               Width           =   1710
            End
            Begin VB.ComboBox cboPTP 
               Appearance      =   0  'Flat
               BackColor       =   &H00EFE4E0&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmCC_Colection_RITCARD12.frx":1ACEC
               Left            =   720
               List            =   "frmCC_Colection_RITCARD12.frx":1ACEE
               Locked          =   -1  'True
               TabIndex        =   155
               Top             =   3600
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Frame Frame18 
               Appearance      =   0  'Flat
               BackColor       =   &H00FCFCFC&
               Caption         =   "RESERVED"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   1605
               Left            =   3240
               TabIndex        =   152
               Top             =   120
               Width           =   3255
               Begin MSComctlLib.ListView LstReserve 
                  Height          =   1275
                  Left            =   120
                  TabIndex        =   153
                  Top             =   225
                  Width           =   3075
                  _ExtentX        =   5424
                  _ExtentY        =   2249
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   12632256
                  BorderStyle     =   1
                  Appearance      =   0
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
               Begin Threed.SSCommand SSCommand2 
                  Height          =   1035
                  Index           =   3
                  Left            =   2400
                  TabIndex        =   154
                  Top             =   225
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   1826
                  _Version        =   196610
                  MousePointer    =   16
                  ForeColor       =   4210752
                  PictureMaskColor=   -2147483644
                  PictureFrames   =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Picture         =   "frmCC_Colection_RITCARD12.frx":1ACF0
                  Caption         =   "&"
                  AutoSize        =   1
                  ButtonStyle     =   2
                  BevelWidth      =   0
               End
            End
            Begin VB.Frame Frame5 
               Appearance      =   0  'Flat
               BackColor       =   &H00FCFCFC&
               Caption         =   "OVERDUE"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   1605
               Left            =   3240
               TabIndex        =   148
               Top             =   1800
               Width           =   3255
               Begin MSComctlLib.ListView LstPayment 
                  Height          =   1275
                  Left            =   120
                  TabIndex        =   149
                  Top             =   240
                  Width           =   3075
                  _ExtentX        =   5424
                  _ExtentY        =   2249
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  Checkboxes      =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   12632256
                  BorderStyle     =   1
                  Appearance      =   0
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
               Begin Threed.SSCommand SSCommand2 
                  Height          =   735
                  Index           =   1
                  Left            =   3690
                  TabIndex        =   150
                  Top             =   1710
                  Visible         =   0   'False
                  Width           =   750
                  _ExtentX        =   1323
                  _ExtentY        =   1296
                  _Version        =   196610
                  PictureFrames   =   1
                  Picture         =   "frmCC_Colection_RITCARD12.frx":1DEAE
                  Caption         =   "&Ubah"
                  Alignment       =   8
               End
               Begin Threed.SSCommand SSCommand2 
                  Height          =   1035
                  Index           =   2
                  Left            =   2400
                  TabIndex        =   151
                  Top             =   240
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   1826
                  _Version        =   196610
                  MousePointer    =   16
                  ForeColor       =   4210752
                  PictureMaskColor=   -2147483644
                  PictureFrames   =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Picture         =   "frmCC_Colection_RITCARD12.frx":1E437
                  Caption         =   "&"
                  AutoSize        =   1
                  ButtonStyle     =   2
                  BevelWidth      =   0
               End
            End
            Begin VB.ComboBox CmbViaPtp 
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               ItemData        =   "frmCC_Colection_RITCARD12.frx":215F5
               Left            =   180
               List            =   "frmCC_Colection_RITCARD12.frx":21608
               TabIndex        =   89
               Top             =   2400
               Width           =   2895
            End
            Begin VB.ComboBox cmbDiscount 
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               ItemData        =   "frmCC_Colection_RITCARD12.frx":21639
               Left            =   4080
               List            =   "frmCC_Colection_RITCARD12.frx":2163B
               TabIndex        =   72
               Text            =   "0"
               Top             =   3660
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.CheckBox C_Payment 
               Enabled         =   0   'False
               Height          =   255
               Left            =   3180
               TabIndex        =   51
               Top             =   3600
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.CheckBox Chktenor 
               BackColor       =   &H00EFE4E0&
               Caption         =   "Tenor"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1500
               TabIndex        =   50
               Top             =   600
               Width           =   795
            End
            Begin TDBNumber6Ctl.TDBNumber txttenor 
               Height          =   255
               Left            =   2450
               TabIndex        =   52
               Top             =   600
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   450
               Calculator      =   "frmCC_Colection_RITCARD12.frx":2163D
               Caption         =   "frmCC_Colection_RITCARD12.frx":2165D
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":216C9
               Keys            =   "frmCC_Colection_RITCARD12.frx":216E7
               Spin            =   "frmCC_Colection_RITCARD12.frx":21731
               AlignHorizontal =   2
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483645
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
               ValueVT         =   1989214213
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBDate6Ctl.TDBDate TDBDate3 
               Height          =   285
               Left            =   1500
               TabIndex        =   53
               Top             =   1305
               Width           =   1590
               _Version        =   65536
               _ExtentX        =   2805
               _ExtentY        =   503
               Calendar        =   "frmCC_Colection_RITCARD12.frx":21759
               Caption         =   "frmCC_Colection_RITCARD12.frx":21871
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":218DD
               Keys            =   "frmCC_Colection_RITCARD12.frx":218FB
               Spin            =   "frmCC_Colection_RITCARD12.frx":21959
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   12648447
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   6815745
               Value           =   39876
               CenturyMode     =   0
            End
            Begin TDBNumber6Ctl.TDBNumber txtPayment 
               Height          =   255
               Left            =   1500
               TabIndex        =   54
               Top             =   300
               Width           =   1590
               _Version        =   65536
               _ExtentX        =   2805
               _ExtentY        =   450
               Calculator      =   "frmCC_Colection_RITCARD12.frx":21981
               Caption         =   "frmCC_Colection_RITCARD12.frx":219A1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":21A0D
               Keys            =   "frmCC_Colection_RITCARD12.frx":21A2B
               Spin            =   "frmCC_Colection_RITCARD12.frx":21A75
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483645
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
               ValueVT         =   5
               Value           =   88888888
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber Tdabamoint 
               Height          =   255
               Left            =   1740
               TabIndex        =   55
               Top             =   3585
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   450
               Calculator      =   "frmCC_Colection_RITCARD12.frx":21A9D
               Caption         =   "frmCC_Colection_RITCARD12.frx":21ABD
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":21B29
               Keys            =   "frmCC_Colection_RITCARD12.frx":21B47
               Spin            =   "frmCC_Colection_RITCARD12.frx":21B91
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   12648384
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBDate6Ctl.TDBDate tdbptpnew 
               Height          =   285
               Left            =   1500
               TabIndex        =   56
               Top             =   900
               Width           =   1590
               _Version        =   65536
               _ExtentX        =   2805
               _ExtentY        =   503
               Calendar        =   "frmCC_Colection_RITCARD12.frx":21BB9
               Caption         =   "frmCC_Colection_RITCARD12.frx":21CD1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":21D3D
               Keys            =   "frmCC_Colection_RITCARD12.frx":21D5B
               Spin            =   "frmCC_Colection_RITCARD12.frx":21DB9
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   12648447
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   0
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   6815745
               Value           =   39876
               CenturyMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate TdbTglTagih 
               Height          =   285
               Left            =   1500
               TabIndex        =   80
               Top             =   1740
               Width           =   1590
               _Version        =   65536
               _ExtentX        =   2805
               _ExtentY        =   503
               Calendar        =   "frmCC_Colection_RITCARD12.frx":21DE1
               Caption         =   "frmCC_Colection_RITCARD12.frx":21EF9
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":21F65
               Keys            =   "frmCC_Colection_RITCARD12.frx":21F83
               Spin            =   "frmCC_Colection_RITCARD12.frx":21FE1
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   12648447
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   0
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   6815745
               Value           =   39876
               CenturyMode     =   0
            End
            Begin Threed.SSCommand SSCommand2 
               Height          =   795
               Index           =   0
               Left            =   5640
               TabIndex        =   143
               Top             =   720
               Visible         =   0   'False
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   1402
               _Version        =   196610
               MousePointer    =   16
               ForeColor       =   4210752
               PictureMaskColor=   -2147483644
               PictureFrames   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCC_Colection_RITCARD12.frx":22009
               Caption         =   "&"
               AutoSize        =   1
               ButtonStyle     =   2
               BevelWidth      =   0
            End
            Begin Threed.SSCommand CmdSendPTP 
               Height          =   435
               Left            =   180
               TabIndex        =   159
               Top             =   2970
               Width           =   2820
               _ExtentX        =   4974
               _ExtentY        =   767
               _Version        =   196610
               MousePointer    =   16
               ForeColor       =   0
               PictureMaskColor=   -2147483644
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "CREATE"
               AutoSize        =   1
               ButtonStyle     =   2
               BevelWidth      =   0
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00FCFCFC&
               Caption         =   "RESULT PTP"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   2505
               TabIndex        =   158
               Top             =   3660
               Visible         =   0   'False
               Width           =   1245
               WordWrap        =   -1  'True
            End
            Begin VB.Label LblResultPTP 
               Appearance      =   0  'Flat
               BackColor       =   &H00EFE4E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3600
               TabIndex        =   157
               Top             =   3615
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label label1 
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Payment Via"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   180
               TabIndex        =   83
               Top             =   2160
               Width           =   1665
            End
            Begin VB.Label label1 
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl.Tagih"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   11
               Left            =   180
               TabIndex        =   79
               Top             =   1740
               Width           =   1005
            End
            Begin VB.Label label1 
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Date PTPNew"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   18
               Left            =   180
               TabIndex        =   71
               Top             =   900
               Width           =   1245
            End
            Begin VB.Label label1 
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Date Payment Effective"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   435
               Index           =   0
               Left            =   180
               TabIndex        =   59
               Top             =   1260
               Width           =   1605
            End
            Begin VB.Label label1 
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Total Amount Deal"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   495
               Index           =   77
               Left            =   180
               TabIndex        =   58
               Top             =   300
               Width           =   1425
            End
            Begin VB.Label label1 
               BackColor       =   &H00ABE18E&
               Caption         =   "Installment"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   42
               Left            =   1800
               TabIndex        =   57
               Top             =   3705
               Width           =   1005
            End
         End
         Begin VB.Frame Frame14 
            Appearance      =   0  'Flat
            BackColor       =   &H00FCFCFC&
            Caption         =   "PERSONAL INFO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   4230
            Left            =   -120
            TabIndex        =   36
            Top             =   -15
            Width           =   6960
            Begin VB.TextBox Text10 
               Height          =   285
               Left            =   7095
               TabIndex        =   181
               Top             =   2370
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.TextBox lblCustId 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   975
               Locked          =   -1  'True
               TabIndex        =   180
               Top             =   250
               Width           =   1780
            End
            Begin VB.TextBox lblNama 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   975
               Locked          =   -1  'True
               TabIndex        =   179
               Top             =   600
               Width           =   3090
            End
            Begin VB.TextBox txtremarks_old 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   4320
               MultiLine       =   -1  'True
               TabIndex        =   176
               Top             =   3785
               Width           =   2475
            End
            Begin VB.TextBox lblOfficeAddr 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   900
               Left            =   975
               MultiLine       =   -1  'True
               TabIndex        =   173
               Top             =   1935
               Width           =   3075
            End
            Begin VB.TextBox Text8 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   975
               MultiLine       =   -1  'True
               TabIndex        =   163
               Top             =   950
               Width           =   3075
            End
            Begin TDBNumber6Ctl.TDBNumber lblAmount 
               Height          =   255
               Left            =   5070
               TabIndex        =   44
               Top             =   240
               Width           =   1785
               _Version        =   65536
               _ExtentX        =   3149
               _ExtentY        =   450
               Calculator      =   "frmCC_Colection_RITCARD12.frx":24BFA
               Caption         =   "frmCC_Colection_RITCARD12.frx":24C1A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":24C86
               Keys            =   "frmCC_Colection_RITCARD12.frx":24CA4
               Spin            =   "frmCC_Colection_RITCARD12.frx":24CEE
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483645
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
               ReadOnly        =   -1
               Separator       =   ","
               ShowContextMenu =   -1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber lblLastPay 
               Height          =   255
               Left            =   5070
               TabIndex        =   75
               Top             =   2315
               Width           =   1785
               _Version        =   65536
               _ExtentX        =   3149
               _ExtentY        =   450
               Calculator      =   "frmCC_Colection_RITCARD12.frx":24D16
               Caption         =   "frmCC_Colection_RITCARD12.frx":24D36
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":24DA2
               Keys            =   "frmCC_Colection_RITCARD12.frx":24DC0
               Spin            =   "frmCC_Colection_RITCARD12.frx":24E0A
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483645
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
               ReadOnly        =   -1
               Separator       =   ","
               ShowContextMenu =   -1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBDate6Ctl.TDBDate lblPayDt 
               Height          =   255
               Left            =   5070
               TabIndex        =   76
               Top             =   2015
               Width           =   1785
               _Version        =   65536
               _ExtentX        =   3149
               _ExtentY        =   450
               Calendar        =   "frmCC_Colection_RITCARD12.frx":24E32
               Caption         =   "frmCC_Colection_RITCARD12.frx":24F4A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":24FB6
               Keys            =   "frmCC_Colection_RITCARD12.frx":24FD4
               Spin            =   "frmCC_Colection_RITCARD12.frx":25032
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483645
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
            Begin TDBNumber6Ctl.TDBNumber lbl_principal 
               Height          =   255
               Left            =   7575
               TabIndex        =   164
               Top             =   615
               Visible         =   0   'False
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calculator      =   "frmCC_Colection_RITCARD12.frx":2505A
               Caption         =   "frmCC_Colection_RITCARD12.frx":2507A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":250E6
               Keys            =   "frmCC_Colection_RITCARD12.frx":25104
               Spin            =   "frmCC_Colection_RITCARD12.frx":2514E
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483645
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber txtdenda 
               Height          =   255
               Left            =   7575
               TabIndex        =   171
               Top             =   1020
               Visible         =   0   'False
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calculator      =   "frmCC_Colection_RITCARD12.frx":25176
               Caption         =   "frmCC_Colection_RITCARD12.frx":25196
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":25202
               Keys            =   "frmCC_Colection_RITCARD12.frx":25220
               Spin            =   "frmCC_Colection_RITCARD12.frx":2526A
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483645
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber Txtlimit 
               Height          =   255
               Left            =   5070
               TabIndex        =   314
               Top             =   530
               Width           =   1785
               _Version        =   65536
               _ExtentX        =   3149
               _ExtentY        =   450
               Calculator      =   "frmCC_Colection_RITCARD12.frx":25292
               Caption         =   "frmCC_Colection_RITCARD12.frx":252B2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD12.frx":2531E
               Keys            =   "frmCC_Colection_RITCARD12.frx":2533C
               Spin            =   "frmCC_Colection_RITCARD12.frx":25386
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483645
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin VB.Label Txteusex 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   975
               TabIndex        =   326
               Top             =   3840
               Width           =   1785
            End
            Begin VB.Label Label40 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "EU Sex"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               TabIndex        =   325
               Top             =   3840
               Width           =   840
            End
            Begin VB.Label Label36 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "JK"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4200
               TabIndex        =   324
               Top             =   3195
               Width           =   840
            End
            Begin VB.Label Txtjk 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   5070
               TabIndex        =   323
               Top             =   3195
               Width           =   1785
            End
            Begin VB.Label Zip 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Zip Code"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4200
               TabIndex        =   322
               Top             =   1755
               Width           =   960
            End
            Begin VB.Label Txtzip_code 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   5055
               TabIndex        =   321
               Top             =   1715
               Width           =   1785
            End
            Begin VB.Label Txtstatus 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   5055
               TabIndex        =   320
               Top             =   1410
               Width           =   1785
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Status"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4200
               TabIndex        =   319
               Top             =   1440
               Width           =   960
            End
            Begin VB.Label Label35 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Blk Code"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   4200
               TabIndex        =   318
               Top             =   1140
               Width           =   960
            End
            Begin VB.Label Txtblk_code 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   5055
               TabIndex        =   317
               Top             =   1110
               Width           =   1785
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Limit"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   9
               Left            =   4215
               TabIndex        =   315
               Top             =   540
               Width           =   645
            End
            Begin VB.Label Txtcycle 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   5070
               TabIndex        =   187
               Top             =   840
               Width           =   1785
            End
            Begin VB.Label Label33 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Cylce"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4215
               TabIndex        =   186
               Top             =   840
               Width           =   1200
            End
            Begin VB.Label Label25 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   5070
               TabIndex        =   185
               Top             =   2895
               Width           =   1785
            End
            Begin VB.Label Label23 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Card Num"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4215
               TabIndex        =   184
               Top             =   2895
               Width           =   840
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   5070
               TabIndex        =   183
               Top             =   2615
               Width           =   1785
            End
            Begin VB.Label Label20 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "DPD"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   4215
               TabIndex        =   182
               Top             =   2615
               Width           =   720
            End
            Begin VB.Label Txtro 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2730
               TabIndex        =   178
               Top             =   2910
               Width           =   1320
            End
            Begin VB.Label Label17 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Remarks Old :"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   405
               Left            =   5040
               TabIndex        =   177
               Top             =   3480
               Width           =   1260
            End
            Begin VB.Label Label9 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Office Address"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   720
               Left            =   120
               TabIndex        =   174
               Top             =   2145
               Width           =   720
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Denda + Bunga"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   450
               Index           =   23
               Left            =   7080
               TabIndex        =   172
               Top             =   915
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Principal"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   19
               Left            =   6960
               TabIndex        =   165
               Top             =   615
               Visible         =   0   'False
               Width           =   885
            End
            Begin VB.Label lblRecsource 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   975
               TabIndex        =   87
               Top             =   3525
               Width           =   3075
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H003F9E0C&
               BackStyle       =   0  'Transparent
               Caption         =   "Cmpgn"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   80
               Left            =   120
               TabIndex        =   86
               Tag             =   "0"
               Top             =   3525
               Width           =   780
            End
            Begin VB.Label lblaoc 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   975
               TabIndex        =   85
               Top             =   3225
               Width           =   3075
            End
            Begin VB.Label Label32 
               BackColor       =   &H003F9E0C&
               BackStyle       =   0  'Transparent
               Caption         =   "Agent"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   120
               TabIndex        =   84
               Top             =   3240
               Width           =   735
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "LPA"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   4
               Left            =   4200
               TabIndex        =   78
               Top             =   2315
               Width           =   885
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "LPD"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   2
               Left            =   4215
               TabIndex        =   77
               Top             =   2055
               Width           =   885
            End
            Begin VB.Label lblpurge 
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3435
               TabIndex        =   74
               Top             =   330
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label lbltype 
               Appearance      =   0  'Flat
               BackColor       =   &H00808080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2835
               TabIndex        =   73
               Top             =   330
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Balance"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   6
               Left            =   4215
               TabIndex        =   45
               Top             =   195
               Width           =   885
            End
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   630
               Width           =   720
            End
            Begin VB.Label Label8 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Home Address"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   675
               Left            =   120
               TabIndex        =   42
               Top             =   1200
               Width           =   720
            End
            Begin VB.Label lblZIP 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   -1260
               TabIndex        =   41
               Top             =   2520
               Width           =   1050
            End
            Begin VB.Label Label22 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "RO"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   2415
               TabIndex        =   40
               Top             =   2910
               Width           =   480
            End
            Begin VB.Label Label37 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Kanwil"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   2910
               Width           =   720
            End
            Begin VB.Label lblregion 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   975
               TabIndex        =   38
               Top             =   2910
               Width           =   1140
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00ABE18E&
               BackStyle       =   0  'Transparent
               Caption         =   "Cust No"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   65
               Left            =   120
               TabIndex        =   37
               Top             =   300
               Width           =   720
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   930
         Left            =   9690
         TabIndex        =   29
         Top             =   7890
         Width           =   2775
         Begin VB.Label LblStatus 
            Caption         =   "Label42"
            Height          =   255
            Left            =   600
            TabIndex        =   48
            Top             =   360
            Width           =   255
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   72
            Left            =   75
            TabIndex        =   34
            Top             =   315
            Width           =   60
         End
         Begin VB.Label lblCardNo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2460
            TabIndex        =   33
            Top             =   315
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label CustId 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "# Card"
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
            Height          =   195
            Index           =   0
            Left            =   1905
            TabIndex        =   32
            Top             =   285
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Emergency Contact"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   46
            Left            =   15195
            TabIndex        =   31
            Top             =   1590
            Width           =   1890
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Telp Tambahan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   10680
            TabIndex        =   30
            Top             =   135
            Width           =   1500
         End
      End
   End
   Begin VB.Frame FrmPayment1 
      Height          =   1365
      Left            =   1920
      TabIndex        =   22
      Top             =   8295
      Visible         =   0   'False
      Width           =   2085
      Begin VB.CheckBox Check3 
         Caption         =   "Regular to paid Off"
         Height          =   195
         Left            =   75
         TabIndex        =   25
         Top             =   285
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Iregular to Paid Off"
         Height          =   195
         Left            =   60
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Regular Payment"
         Height          =   195
         Left            =   75
         TabIndex        =   23
         Top             =   870
         Visible         =   0   'False
         Width           =   435
      End
      Begin TDBDate6Ctl.TDBDate TdbPTP 
         Height          =   255
         Left            =   60
         TabIndex        =   26
         Top             =   585
         Visible         =   0   'False
         Width           =   1440
         _Version        =   65536
         _ExtentX        =   2540
         _ExtentY        =   450
         Calendar        =   "frmCC_Colection_RITCARD12.frx":253AE
         Caption         =   "frmCC_Colection_RITCARD12.frx":254C6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":25532
         Keys            =   "frmCC_Colection_RITCARD12.frx":25550
         Spin            =   "frmCC_Colection_RITCARD12.frx":255AE
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
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   3.54027066542603E-316
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TdbDatePTP 
         Height          =   225
         Left            =   60
         TabIndex        =   27
         Top             =   1065
         Visible         =   0   'False
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   397
         Calendar        =   "frmCC_Colection_RITCARD12.frx":255D6
         Caption         =   "frmCC_Colection_RITCARD12.frx":256EE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD12.frx":2575A
         Keys            =   "frmCC_Colection_RITCARD12.frx":25778
         Spin            =   "frmCC_Colection_RITCARD12.frx":257D6
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
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
   End
   Begin VB.Frame Frame9 
      Height          =   3405
      Left            =   15
      TabIndex        =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   1755
      Begin VB.OptionButton Option8 
         Caption         =   "Tambah"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   2
         Top             =   2070
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Batal"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   1395
         TabIndex        =   1
         Top             =   2085
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Frame Frame8 
         ForeColor       =   &H000000FF&
         Height          =   1725
         Left            =   60
         TabIndex        =   3
         Top             =   2145
         Visible         =   0   'False
         Width           =   7560
         Begin VB.TextBox TxtName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   915
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   540
            Width           =   3135
         End
         Begin VB.TextBox TxtCustid 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   3375
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   915
            TabIndex        =   7
            Top             =   225
            Width           =   1815
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Alamat Billing"
            Height          =   195
            Index           =   0
            Left            =   4125
            TabIndex        =   6
            Top             =   855
            Width           =   1440
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Rumah"
            Height          =   195
            Index           =   1
            Left            =   5565
            TabIndex        =   5
            Top             =   855
            Width           =   840
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Kantor"
            Height          =   195
            Index           =   2
            Left            =   6525
            TabIndex        =   4
            Top             =   840
            Width           =   840
         End
         Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
            Height          =   315
            Left            =   915
            TabIndex        =   10
            Top             =   870
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   556
            Calculator      =   "frmCC_Colection_RITCARD12.frx":257FE
            Caption         =   "frmCC_Colection_RITCARD12.frx":2581E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD12.frx":2588A
            Keys            =   "frmCC_Colection_RITCARD12.frx":258A8
            Spin            =   "frmCC_Colection_RITCARD12.frx":258F2
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "####0;;Null"
            EditMode        =   0
            Enabled         =   -1
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin RichTextLib.RichTextBox TXtDetails 
            Height          =   570
            Left            =   4080
            TabIndex        =   11
            Top             =   225
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1005
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmCC_Colection_RITCARD12.frx":2591A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin TDBDate6Ctl.TDBDate TDBDate2 
            Height          =   315
            Left            =   915
            TabIndex        =   12
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Calendar        =   "frmCC_Colection_RITCARD12.frx":2599F
            Caption         =   "frmCC_Colection_RITCARD12.frx":25AB7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD12.frx":25B23
            Keys            =   "frmCC_Colection_RITCARD12.frx":25B41
            Spin            =   "frmCC_Colection_RITCARD12.frx":25B9F
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "mm/dd/yyyy"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            FirstMonth      =   4
            ForeColor       =   -2147483640
            Format          =   "mm/dd/yyyy"
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
            ValueVT         =   2010382337
            Value           =   2.12482692446619E-314
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate1 
            Height          =   315
            Left            =   1590
            TabIndex        =   13
            Top             =   870
            Visible         =   0   'False
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Calendar        =   "frmCC_Colection_RITCARD12.frx":25BC7
            Caption         =   "frmCC_Colection_RITCARD12.frx":25CDF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD12.frx":25D4B
            Keys            =   "frmCC_Colection_RITCARD12.frx":25D69
            Spin            =   "frmCC_Colection_RITCARD12.frx":25DC7
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
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
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   2010382337
            Value           =   2.12482692446619E-314
            CenturyMode     =   0
         End
         Begin RichTextLib.RichTextBox TxtAddress 
            Height          =   540
            Left            =   4065
            TabIndex        =   14
            Top             =   1065
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   953
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmCC_Colection_RITCARD12.frx":25DEF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Nomor"
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
            Height          =   255
            Index           =   1
            Left            =   30
            TabIndex        =   21
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Note:"
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
            Height          =   255
            Left            =   2925
            TabIndex        =   20
            Top             =   195
            Width           =   1095
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Visit Date"
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
            Height          =   255
            Left            =   30
            TabIndex        =   19
            Top             =   1245
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Visit Ke"
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
            Height          =   255
            Left            =   30
            TabIndex        =   18
            Top             =   930
            Width           =   810
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Nama"
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
            Height          =   255
            Left            =   30
            TabIndex        =   17
            Top             =   540
            Width           =   810
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Custid"
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
            Index           =   2
            Left            =   420
            TabIndex        =   16
            Top             =   3375
            Width           =   1095
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Visit Ke:"
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
            Height          =   270
            Left            =   3390
            TabIndex        =   15
            Top             =   915
            Width           =   615
         End
      End
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   7695
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.TextBox txtPhoneA 
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   7680
      Width           =   1905
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CALL INFORMATION"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   88
      Top             =   30
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D4B9AF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "FrmCC_Colection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim m_cust As ADODB.Recordset
Dim M_update As ADODB.Recordset
Dim M_objrs As ADODB.Recordset
Dim stscall As Boolean
Dim TYPETELP As String
Dim kontak As Boolean
Dim spend As Boolean
Dim adaSCH As Boolean
Dim adaREG As Boolean
Dim adaPO As Boolean
Dim vrcek As String
Dim vrdateptp As String
Dim vramount As String
Dim vrtdbdateptp As String
Dim vrbaseon As String
Dim vrdiskon As String
Dim vrtenor As String
Dim vrttlptp As String
Dim TglPTPNew As String
Dim vrnewdate As String
Dim KelapKelip As Integer
Dim kelapkelipDetail As Integer
Dim nomortins1, nomortins2 As String
'@@02-05-2012 Tambahan buat Catet Status Kategori
Dim StsKategoriTelepon As String
Dim KelompokKategoriTlp As String
Dim StatusSpeakWith As String
Dim StatusAccount As String
'@@15092012, Catat Apakah Di Sudah Melakukan Call?
Dim AktifitasCall As String
Dim calling As String
'@@221012 Tanggal PaidOff
Dim TanggalPaidOff As String

Dim sudahCall As Boolean
Dim kat_aktif_telp As String


Private Sub C_Contacted_Click()
    
    
If C_Contacted.Value Then
        C_VALID.Value = False
        C_SKIP.Value = False
        C_Payment.Value = False
        C_PTP.Value = False
      '  C_POPSP.Value = False
        FrmContacted.Enabled = True
      '  cboPOPSP.Text = ""
   Else
        cmbContacted.text = ""
        cmbDescCon.text = ""
        FrmContacted.Enabled = False
        If cboPOPSP.text = "" Then
            C_Payment.Value = False
        End If
        CmbBaseOn.text = ""
        cmbDiscount.text = 0
        TdbPTP.Value = ""
        txtPayment.Value = 0
End If
End Sub

Private Sub C_NotContacted_Click()
   If C_NotContacted.Value Then
      FrmUnContacted.Enabled = True
      C_Contacted.Value = False
      C_Payment.Value = False
   Else
      FrmUnContacted.Enabled = False
      cmbDescUn.text = ""
      cmbUncontacted = ""
   End If
End Sub

Private Sub C_Payment_Click()
   If C_Payment.Value Then
     ' Frame54.Enabled = True
   Else
     ' Frame54.Enabled = False
     'If cboPOPSP.Text <> "" Then
     'Exit Sub
     'End If
     
      'cmbDiscount.Text = ""
   End If
End Sub
Private Sub C_PTP_Click()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim m_objrs_payment As ADODB.Recordset
    
    

If C_PTP.Value Then
       '@@ 29 Desember 2011, Cek terlebih dahulu, apakah ada CPA atau tidak, jika tidak ada CPA maka
        'tidak bisa melakukan PTP

       CMDSQL = "select * from tblcpa where vcustid='"
       CMDSQL = CMDSQL + Trim(lblCustId.text) + "' order by nid desc"
       Set M_objrs = New ADODB.Recordset
       M_objrs.CursorLocation = adUseClient
       M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

       If M_objrs.RecordCount = 0 Then
        'C_PTP.Value = vbUnchecked
        'MsgBox "Untuk membuat status account PTP, harus dibuat terlebih dahulu CPA nya!", vbOKOnly + vbInformation, "Informasi"
        'Set M_OBJRS = Nothing
        'Exit Sub
       Else
        'Ambil Nilai Payment di CPA untuk di tempatkan di Total Amount Deal
        txtPayment.Value = IIf(IsNull(M_objrs("nttlpayment")), "", M_objrs("nttlpayment"))
        txttenor.Value = IIf(IsNull(M_objrs("nperiod")), "", M_objrs("nperiod"))
        Chktenor.Value = vbChecked
       End If

       Set M_objrs = Nothing
       
 '@@ 11042012 Dinonaktifkan
'       If Left(cboaccount.Text, 3) <> "ON-" Then
'         cboaccount.Text = ""
'       End If
       
        bcekptp = False
 '       C_VALID.Value = False
'        C_SKIP.Value = False
'        C_Contacted.Value = False
        frmPTP.Enabled = True
        FrmPayment.Enabled = True
        'cboPOPSP.Tag = 0
        Label43(2).Visible = True
       ' cboPOPSP.Text = ""
        C_Payment.Value = 1
        If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
            SSCommand1(4).Visible = False
            Label43(2).Visible = False
            Else
            SSCommand1(4).Visible = True
            Label43(2).Visible = True
        End If
        CmbViaPtp.Enabled = True
        
        '@@22 Desember 2011 Tambahan, jika tidak ada pembayaran maka status PTP= PTP NEW
'        If listview1(0).ListItems.Count = 0 Then
'            cboPTP.Text = "PTP-NEW"
'        End If
'        If listview1(0).ListItems.Count > 0 Then
'            cboPTP.Text = "PTP-POP"
'        End If
        CMDSQL = "SELECT b.custid as custid1, a.CUSTID,a.PayDate,a.Payment,"
        CMDSQL = CMDSQL + "a.Agent,a.FieldName,a.Id from tbllunas a inner join mgm b "
        CMDSQL = CMDSQL + "on a.custid=b.custid WHERE a.custid='"
        CMDSQL = CMDSQL + Trim(lblCustId.text) + "' and date(a.Paydate)+1  > b.tglsource "
        Set m_objrs_payment = New ADODB.Recordset
        m_objrs_payment.CursorLocation = adUseClient
        m_objrs_payment.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If m_objrs_payment.RecordCount = 0 Then
            cboPTP.text = "PTP-NEW"
        Else
            cboPTP.text = "PTP-POP"
        End If
        Set m_objrs_payment = Nothing
        
   Else
       bcekptp = False
       Label43(2).Visible = False
        'C_Payment.Value = 0
       ' CmbBaseOn.Text = ""
       ' cmbDiscount.Text = 0
        'txtPayment.Value = 0
'        TxtPtpAddr.Text = ""
 '       TxtPhonePTP.Text = ""
      '  FrmPayment.Enabled = False
        cboPTP.text = ""
        SSCommand1(4).Visible = False
        frmPTP.Enabled = False
        TdbPTP.Value = ""
        CmbBaseOn.text = ""
        cmbDiscount.text = 0
        TdbPTP.Value = ""
        txtPayment.Value = 0
        'C_Payment = False
        Chktenor.Value = vbUnchecked
        txttenor.Value = 0
        TDBDate3.Value = ""
        CmbViaPtp.text = ""
        tdbptpnew.Value = ""
        TdbTglTagih.Value = ""
        CmbViaPtp.Enabled = False
End If

End Sub

Private Sub C_SKIP_Click()
If C_SKIP.Value Then
        C_VALID.Value = False
        C_Contacted.Value = False
        C_Payment.Value = False
        C_PTP.Value = False
     
        FrmSKIP.Enabled = True
   Else
        cboskip.text = ""
        cbodescskip.text = ""
        FrmSKIP.Enabled = False
End If

End Sub

Private Sub C_VALID_Click()
If C_VALID.Value Then
        C_Contacted.Value = False
        C_SKIP.Value = False
        C_Payment.Value = False
        C_PTP.Value = False
        
        FrMValid.Enabled = True
   Else
        cbovalid.text = ""
        cbodescvalid.text = ""
        FrMValid.Enabled = False
End If

End Sub

Private Sub cbodescskip_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
KeyAscii = 0
End If

End Sub

Private Sub cbodescvalid_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
KeyAscii = 0
End If

End Sub

Private Sub cboaccount_Click()
'        Dim M_COL1 As New ADODB.Recordset
'        cboaccount.Locked = True
'
'    '@@ 11-04-2012, Dinonaktifkan
'    '    If Left(cboaccount, 3) <> "ON-" Then
'    '        C_Payment.Value = vbUnchecked
'    '        C_PTP.Value = vbUnchecked
'    '    End If
'
'     C_Payment.Value = vbUnchecked
'     C_PTP.Value = vbUnchecked
'
'    If UCase(Left(cboaccount.Text, 2)) = "SP" Then
'        C_PTP.Value = 0
'        CmbBaseOn.Text = ""
'        cmbDiscount.Text = ""
'        txtPayment.Value = 0
'        Tdabamoint.Value = 0
'        TDBDate3.Value = ""
'        txttenor.Value = 0
'        C_Payment.Value = 1
'        FrmPayment.Enabled = True
'                Set M_COL1 = New ADODB.Recordset
'                CMDSQL = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon,dateptp,tenor,amountptp from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
'                M_COL1.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'    '            CmbBaseOn.Text = "PRINCIPLE"
'                txtPayment.Value = CStr(IIf(IsNull(M_COL1!ttlptp), "", M_COL1!ttlptp))
'                CmbBaseOn.Text = CStr(IIf(IsNull(M_COL1!CmbBaseOn), "", M_COL1!CmbBaseOn))
'                TdbPTP.Value = CStr(IIf(IsNull(M_COL1!TdbDatePTP), "", M_COL1!TdbDatePTP))
'                cmbDiscount.Text = CStr(IIf(IsNull(M_COL1!discpersen), "", M_COL1!discpersen))
'                TDBDate3.Value = CStr(IIf(IsNull(M_COL1!dateptp), "", M_COL1!dateptp))
'                txttenor.Value = CStr(IIf(IsNull(M_COL1!Tenor), 0, M_COL1!Tenor))
'                Tdabamoint.Value = CStr(IIf(IsNull(M_COL1!amountptp), 0, M_COL1!amountptp))
'    End If
    Combo1.ListIndex = cboaccount.ListIndex
End Sub
Private Sub cboaccount_DropDown()
     cboaccount.Locked = False
     
End Sub

Private Sub chk_aktif_Click()
    If chk_aktif = vbChecked Then
        M_OBJCONN.Execute "UPDATE tbl_cek_framePTP SET status_cek_frame = 1 "
    Else
        M_OBJCONN.Execute "UPDATE tbl_cek_framePTP SET status_cek_frame = 0 "
    End If
End Sub

Private Sub cmd_req_telp_Click()
    FrmReqTelepon.Show 1
End Sub

Private Sub Combo1_DropDown()
     Combo1.Locked = False
End Sub

Private Sub cbolastcall_DropDown()
     cbolastcall.Locked = False
End Sub

Private Sub cboaccount_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbolastcall_Click()
    Select Case UCase(cbolastcall.text)
        Case "CH"
            StatusSpeakWith = "CH"
        Case "RECEPTION/OPERATOR/SEC/OB"
            StatusSpeakWith = "ROSO"
        Case "ATASAN"
            StatusSpeakWith = "BOSS"
        Case "HRD"
            StatusSpeakWith = "HRD"
        Case "TEMAN KANTOR"
            StatusSpeakWith = "FRND"
        Case "ORANG TUA"
            StatusSpeakWith = "PRNT"
        Case "KAKAK/ADIK/ANAK"
            StatusSpeakWith = "BSSD"
        Case "SPOUSE"
            StatusSpeakWith = "SPS"
        Case "KELUARGA DEKAT LAINNYA"
            StatusSpeakWith = "OFAM"
        Case "EX SPOUSE"
            StatusSpeakWith = "ESPS"
        Case "PEMBANTU/SUPIR"
            StatusSpeakWith = "MAID"
        Case "OTHER"
            StatusSpeakWith = "OTH"
        Case "TETANGGA"
            StatusSpeakWith = "NGBR"
        Case "PENGURUS LINGKUNGAN"
            StatusSpeakWith = "RTRW"
        Case "KONTRAKAN"
            StatusSpeakWith = "KNTK"
        Case "LAWYER"
            StatusSpeakWith = "LAW"
        Case "TEMAN"
            StatusSpeakWith = "FRND"
        Case "TEMAN KANTOR"
            StatusSpeakWith = "FRND"
        Case "LAWYER"
            StatusSpeakWith = "LAW"
         Case "UNRECEIVE"
            StatusSpeakWith = "NRCV"
    End Select
End Sub

Private Sub cbolastcall_GotFocus()
'cbolastcall.CLEAR
'Dim M_OBJRS As ADODB.Recordset
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'If UCase(mdiform1.txtlevel.text) = "AGENT" Then
'    If Left(cmbContacted.Text, 2) = "OP" Then
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented not in('SP-SETTLE PAYMENT','PTP-PROMISE TO PAY') "
'    ElseIf Left(cboPTP.Text, 3) = "PTP" Then
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented not in('OP-ON PROGRESS','SP-SETTLE PAYMENT') "
'    Else
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented not in('SP-SETTLE PAYMENT')"
'    End If
' Else
'    If Left(cmbContacted.Text, 2) = "OP" Then
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented <> 'PTP-PROMISE TO PAY' "
'    ElseIf Left(cboPTP.Text, 3) = "PTP" Then
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented <> 'OP-ON PROGRESS' "
'    Else
'    CMDSQL = " Select * from ContactedDesc"
'    End If
' End If
'M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not M_OBJRS.EOF
'    cbolastcall.AddItem M_OBJRS("KdNoProdPresented")
'    M_OBJRS.MoveNext
'Wend
'Set M_OBJRS = Nothing
'
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not M_OBJRS.EOF
'    cbolastcall.AddItem M_OBJRS("KdNoProdPresented")
'    M_OBJRS.MoveNext
'Wend
'Set M_OBJRS = Nothing
End Sub

Private Sub cbolastcall_KeyDown(KeyCode As Integer, Shift As Integer)

cbolastcall.text = ""
Exit Sub
End Sub

Private Sub cboPOPSP_Click()
Dim M_COL1 As New ADODB.Recordset
If Left(cboPOPSP.text, 2) = "SP" Then
    C_Contacted.Value = 0
    C_SKIP.Value = 0
    C_PTP.Value = 0
    C_VALID.Value = 0
    CmbBaseOn.text = ""
    cmbDiscount.text = ""
    txtPayment.Value = 0
    Tdabamoint.Value = 0
    TDBDate3.Value = ""
    txttenor.Value = 0
    cmbDescCon.Enabled = False
    C_Payment.Value = 1
    FrmPayment.Enabled = True
            Set M_COL1 = New ADODB.Recordset
            CMDSQL = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon,dateptp,tenor,amountptp from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
            M_COL1.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'            CmbBaseOn.Text = "PRINCIPLE"
            txtPayment.Value = CStr(IIf(IsNull(M_COL1!ttlptp), "", M_COL1!ttlptp))
            CmbBaseOn.text = CStr(IIf(IsNull(M_COL1!CmbBaseOn), "", M_COL1!CmbBaseOn))
            TdbPTP.Value = CStr(IIf(IsNull(M_COL1!TdbDatePTP), "", M_COL1!TdbDatePTP))
            cmbDiscount.text = CStr(IIf(IsNull(M_COL1!discpersen), "", M_COL1!discpersen))
            TDBDate3.Value = CStr(IIf(IsNull(M_COL1!dateptp), "", M_COL1!dateptp))
            txttenor.Value = CStr(IIf(IsNull(M_COL1!Tenor), 0, M_COL1!Tenor))
            Tdabamoint.Value = CStr(IIf(IsNull(M_COL1!amountptp), 0, M_COL1!amountptp))
End If

'C_Payment.Value = 0



'txtPayment.Value = 0

End Sub

Private Sub cboPOPSP_KeyDown(KeyCode As Integer, Shift As Integer)

cboPOPSP.text = ""
End Sub


Private Sub cboskip_Click()
cbodescskip.CLEAR
If Left(cboskip.text, 2) <> "MV" Then
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
   M_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
         For i = 0 To 3
           cbodescskip.AddItem M_objrs("Description")
           M_objrs.MoveNext
         Next i
   Set M_objrs = Nothing
   C_Payment.Value = 0
Else
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
      M_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
       While Not M_objrs.EOF
           cbodescskip.AddItem M_objrs("Description")
           M_objrs.MoveNext
       Wend
   Set M_objrs = Nothing
   C_Payment.Value = 0
End If

End Sub

Private Sub cbovalid_Click()
Dim i As Integer
cbodescvalid.CLEAR
If Left(cbovalid.text, 2) = "NA" Then
        cbodescvalid.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
          M_objrs.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_objrs.EOF
            cbodescvalid.AddItem M_objrs("Description")
            M_objrs.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_objrs = Nothing
'        FrmPayment.Enabled = False
Else
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
          M_objrs.Open "Select * from DescunContacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_objrs.EOF
            cbodescvalid.AddItem M_objrs("Description")
            M_objrs.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_objrs = Nothing
End If

End Sub

Private Sub cbovalid_KeyDown(KeyCode As Integer, Shift As Integer)

cbovalid.text = ""
Exit Sub
End Sub



Private Sub cbolastcall_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboPTP_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Check1_Click()
regnego = False
Check2.Value = 0
Check3.Value = 0
If CmbBaseOn.text = "PRINCIPLE" Then
    MsgBox "Regular payment only TOTAL AMOUNT"
    CmbBaseOn.SetFocus
    Exit Sub
Else
'Call CEKPTP
'If adaSCH Then
'    MsgBox "Hapus dulu PTP yang ada atau selesaikan paymennya!"
'    Exit Sub
'Else
    Call ISIJMLPAYMENT
    If Check1.Value = 1 Then
        frmregpayment.Show
    End If
End If
End Sub

Sub CEKPTP()
Dim rs As New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "select TYPE from TBLNEGOPTP where custid='" & lblCustId.text & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
If rs.BOF And rs.EOF Then
Else
    While Not rs.EOF
        If rs!Type = "SCH" Then
            adaSCH = True
        ElseIf rs!Type = "REG" Then
            adaREG = True
        ElseIf rs!Type = "PO" Then
            adaPO = True
        End If
        rs.MoveNext
    Wend
End If
Set rs = Nothing
End Sub


Private Sub Check2_Click()
Check1.Value = 0
Check3.Value = 0
If Check2.Value = 1 Then
'    If CmbBaseOn.Text = "PRINCIPLE" Then
'        MsgBox "Regular payment only TOTAL AMOUNT"
'        CmbBaseOn.SetFocus
'        Exit Sub
'    Else
'        Call CEKPTP
'        If adaREG Then
'            MsgBox "Hapus dulu PTP yang ada atau selesaikan paymennya!"
'            Exit Sub
'        Else
            'Call ISIJMLPAYMENT
            regnego = True
            FrmNegoPTP.Show
'        End If
End If
'End If
End Sub

Private Sub Check3_Click()
regnego = False
Check1.Value = 0
Check2.Value = 0

'Call CEKPTP
'If adaPO Then
'    MsgBox "Hapus dulu PTP yang ada atau selesaikan paymennya!"
'    Exit Sub
'Else
    Call ISIJMLPAYMENT
    If Check3.Value = 1 Then
        Frmpaidoff.Show
    End If
'End If
End Sub

Private Sub chkAppv_Click(Index As Integer)
'Select Case Index
'Case 0:
'    chkAppv(1).Value = 0
'Case 1:
'    chkAppv(0).Value = 0
'End Select
End Sub

Private Sub Chktenor_Click()
If Chktenor.Value = 1 Then
    txttenor.Enabled = True
    txttenor.BackColor = vbWhite
Else
    txttenor.Enabled = False
    txttenor.BackColor = &H4000&
    Chktenor.Value = 0
    txttenor.Value = 0
End If


End Sub

Private Sub CmbBaseOn_Click()
If CmbBaseOn.text = "PRINCIPLE" Then
CmbBaseOn.text = ""
End If
    Call cmbDiscount_Click
End Sub

Private Sub CmbBaseOn_LostFocus()
    'Call cmbDiscount_Click
End Sub

Private Sub cmbContacted_Click()
'DESCRIPTION CONTACTED
Dim i As Integer
cmbDescCon.CLEAR

'If Left(vrcek, 2) = "BP" And Left(cmbContacted.Text, 3) = "POP" Then
'    cmbContacted.Text = ""
'End If

If Left(cmbContacted.text, 2) = "RP" Then
    cmbDescCon.Enabled = True
    CmbBaseOn.text = ""
    txtPayment.text = 0
    cmbDiscount.text = ""
    TdbPTP.text = ""
    TdbDatePTP.text = ""
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
     M_objrs.Open "Select * from DescContacted where id <= 12", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_objrs.EOF
        cmbDescCon.AddItem M_objrs("Description")
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    C_Payment.Value = 0
   ' FrmPayment.Enabled = False
    Else
'    If Left(cmbContacted.Text, 2) = "NA" Then
'        cmbDescCon.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
'        Set M_OBJRS = New ADODB.Recordset
'          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'        While Not M_OBJRS.EOF
'            cmbDescCon.AddItem M_OBJRS("Description")
'            M_OBJRS.MoveNext
'        Wend
'        C_Payment.Value = 0
'        FrmPayment.Enabled = False
        
'    Else
         If Left(cmbContacted.text, 2) = "PT" Then
            cmbDescCon.Enabled = False
            C_Payment.Value = 1
            FrmPayment.Enabled = True
            CmbBaseOn.text = "PRINCIPLE"
    Else
        If Left(cmbContacted.text, 2) = "BP" Then
            cmbDescCon.Enabled = False
            CmbBaseOn.text = ""
            txtPayment.text = 0
            cmbDiscount.text = ""
            TdbPTP.text = ""
            TdbDatePTP.text = ""
            C_Payment.Value = 0
           ' FrmPayment.Enabled = False
    Else
    If Left(cmbContacted.text, 2) = "OP" Then
            cmbDescCon.Enabled = False
'            CmbBaseOn.Text = ""
'            txtPayment.Text = 0
'            cmbDiscount.Text = ""
'            TdbPTP.Text = ""
'            TdbDatePTP.Text = ""
          '  C_Payment.Value = 1
             'C_Payment.Value = False
            FrmPayment.Enabled = True
      Else
      
    If Left(cmbContacted.text, 2) = "PO" Or Left(cmbContacted.text, 2) = "SP" Then
            cmbDescCon.Enabled = False
            C_Payment.Value = 1
            FrmPayment.Enabled = True
        Set m_cust = New ADODB.Recordset
        m_cust.CursorLocation = adUseClient
        CMDSQL = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon,dateptp,tenor, amountptp from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
        m_cust.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'            CmbBaseOn.Text = "PRINCIPLE"
            txtPayment.Value = CStr(IIf(IsNull(m_cust!ttlptp), "", m_cust!ttlptp))
           CmbBaseOn.text = CStr(IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn))
            TdbPTP.Value = CStr(IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP))
            cmbDiscount.text = CStr(IIf(IsNull(m_cust!discpersen), "", m_cust!discpersen))
            TDBDate3.Value = CStr(IIf(IsNull(m_cust!dateptp), "", m_cust!dateptp))
            txttenor.Value = CStr(IIf(IsNull(m_cust!Tenor), "0", m_cust!Tenor))
            Tdabamoint.Value = CStr(IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp))
            
      Set m_cust = Nothing
    End If
End If
End If
End If
End If
'End If

Set M_objrs = Nothing
End Sub

Private Sub cmbContacted_KeyDown(KeyCode As Integer, Shift As Integer)

cmbContacted.text = ""
Exit Sub
End Sub

Private Sub cmbDescCon_GotFocus()
'DESCRIPTION CONTACTED
Dim i As Integer
cmbDescCon.CLEAR
If Left(cmbContacted.text, 2) = "RP" Then
    cmbDescCon.Enabled = True
    CmbBaseOn.text = ""
    txtPayment.text = 0
    cmbDiscount.text = ""
    TdbPTP.text = ""
    TdbDatePTP.text = ""
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
     M_objrs.Open "Select * from DescContacted where id <= 12", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_objrs.EOF
        cmbDescCon.AddItem M_objrs("Description")
        M_objrs.MoveNext
    Wend
    C_Payment.Value = 0
   ' FrmPayment.Enabled = False
    Set M_objrs = Nothing
    Else
'    If Left(cmbContacted.Text, 2) = "NA" Then
'        cmbDescCon.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
'        Set M_OBJRS = New ADODB.Recordset
'          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'        While Not M_OBJRS.EOF
'            cmbDescCon.AddItem M_OBJRS("Description")
'            M_OBJRS.MoveNext
'        Wend
'        C_Payment.Value = 0
'        FrmPayment.Enabled = False
        
'    Else
         If Left(cmbContacted.text, 2) = "PT" Then
            cmbDescCon.Enabled = False
            C_Payment.Value = 1
            FrmPayment.Enabled = True
            CmbBaseOn.text = "PRINCIPLE"
    Else
        If Left(cmbContacted.text, 2) = "BP" Then
            cmbDescCon.Enabled = False
            CmbBaseOn.text = ""
            txtPayment.text = 0
            cmbDiscount.text = ""
            TdbPTP.text = ""
            TdbDatePTP.text = ""
            C_Payment.Value = 0
'            FrmPayment.Enabled = False
    Else
    If Left(cmbContacted.text, 2) = "OP" Then
            cmbDescCon.Enabled = False
            CmbBaseOn.text = ""
            txtPayment.text = 0
            cmbDiscount.text = ""
            TdbPTP.text = ""
            TdbDatePTP.text = ""
            C_Payment.Value = 0
           ' FrmPayment.Enabled = False
      Else
      
    If Left(cmbContacted.text, 2) = "PO" Or Left(cmbContacted.text, 2) = "SP" Then
            cmbDescCon.Enabled = False
            C_Payment.Value = 1
            FrmPayment.Enabled = True
Set m_cust = New ADODB.Recordset
m_cust.CursorLocation = adUseClient
CMDSQL = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
    m_cust.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'            CmbBaseOn.Text = "PRINCIPLE"
            txtPayment.Value = CStr(IIf(IsNull(m_cust!ttlptp), "", m_cust!ttlptp))
            CmbBaseOn.text = CStr(IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn))
            TdbPTP.Value = CStr(IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP))
            cmbDiscount.text = CStr(IIf(IsNull(m_cust!discpersen), "", m_cust!discpersen))
            
      Set m_cust = Nothing
    End If
End If
End If
End If
End If
'End If

Set M_objrs = Nothing
End Sub

Private Sub cmbDescCon_KeyDown(KeyCode As Integer, Shift As Integer)

cmbDescCon.text = ""
Exit Sub
End Sub

Private Sub cmbDescCon_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
KeyAscii = 0
End If

End Sub

Private Sub cmbDescUn_GotFocus()
Dim i As Integer
cmbDescUn.CLEAR
If Left(cmbUncontacted.text, 2) = "NA" Then
        cmbDescUn.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
          M_objrs.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_objrs.EOF
            cmbDescUn.AddItem M_objrs("Description")
            M_objrs.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_objrs = Nothing
'        FrmPayment.Enabled = False
Else
If Left(cmbUncontacted.text, 2) <> "MV" Then
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
   M_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
         For i = 0 To 3
           cmbDescUn.AddItem M_objrs("Description")
           M_objrs.MoveNext
         Next i
   Set M_objrs = Nothing
   C_Payment.Value = 0
Else
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
'   If kontak = True Then
'        m_objrs.Open "Select * from DescUncontacted where ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    Else
      M_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    End If
       While Not M_objrs.EOF
           cmbDescUn.AddItem M_objrs("Description")
           M_objrs.MoveNext
       Wend
   Set M_objrs = Nothing
   C_Payment.Value = 0
End If
End If
End Sub

Private Sub cmbDescUn_KeyDown(KeyCode As Integer, Shift As Integer)

cmbDescUn.text = ""
Exit Sub
End Sub

Private Sub cmbDiscount_Change()
Call ISIJMLPAYMENT
End Sub

Private Sub cmbDiscount_Click()
Call ISIJMLPAYMENT
'Check1.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'If Left(cmbContacted.Text, 2) = "OP" Then
'    Check1.Enabled = False
'    Check3.Enabled = False
'End If
End Sub

Sub ISIJMLPAYMENT()
Dim M_objrs As New ADODB.Recordset
'If cmbDiscount.Text = "" Then
'    Exit Sub
'End If

M_objrs.CursorLocation = adUseClient
M_objrs.Open "Select * from tbldiscount where Description = '" + cmbDiscount.text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If M_objrs.RecordCount <> 0 Then
    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + IIf(IsNull(M_objrs!hari), 7, M_objrs!hari)
Else
    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + 7
End If
If cmbDiscount.text = "0" Or cmbDiscount.text = "" Then
    If CmbBaseOn.text = "PRINCIPLE" Then
        txtPayment.Value = LblPrompA.Value
    Else
    
         txtPayment.Value = lblAmount.Value
         Exit Sub
         
'         If CmbBaseOn.Text = "TOTAL AMOUNT" Then
'            If lblAmount.Value = 0 Or lblAmount.ValueIsNull Or cmbDiscount = "" Then
'                txtPayment.Value = 0
'            Else
'                txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'                txtPayment.Value = lblAmount.Value - (CCur(txtDiscount.Text) * lblAmount.Value)
'            End If
'        End If
    End If
End If

        If CmbBaseOn.text = "TOTAL AMOUNT" Then
            If lblAmount.Value = 0 Or lblAmount.ValueIsNull Or cmbDiscount = "" Then
                txtPayment.Value = 0
            Else
                txtdiscount.text = CStr((cmbDiscount.text) / 100)
                txtPayment.Value = lblAmount.Value - (CCur(txtdiscount.text) * lblAmount.Value)
                End If

                
            End If
       ' End If

'    If CmbBaseOn.Text = "PRINCIPLE" Then
'        If lblPromPA.Value = 0 Or lblPromPA.ValueIsNull Then
'            txtPayment.Value = 0
'        Else
'            txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'            txtPayment.Value = lblPromPA.Value - (CCur(txtDiscount.Text) * lblPromPA.Value)
'        End If
'    Else
'        If lblAmount.Value = 0 Or lblAmount.ValueIsNull Then
'            txtPayment.Value = 0
'        Else
'            txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'            txtPayment.Value = lblAmount.Value - (CCur(txtDiscount.Text) * lblAmount.Value)
'        End If
'    End If
'End If
'End If

End Sub

Private Sub cmbDiscount_LostFocus()
'Dim M_OBJRS As New ADODB.Recordset
'If cmbDiscount.Text = "" Then
'    Exit Sub
'End If
'
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tbldiscount where Description = '" + cmbDiscount.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'If M_OBJRS.RecordCount <> 0 Then
'    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + IIf(IsNull(M_OBJRS!hari), 7, M_OBJRS!hari)
'Else
'    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + 7
'End If
'If cmbDiscount.Text = "0" Then
'Else
'
'    If CmbBaseOn.Text = "PRINCIPLE" Then
'        If lblPromPA.Value = 0 Then
'            txtPayment.Value = 0
'        Else
'            txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'            txtPayment.Value = lblPromPA.Value - (CCur(txtDiscount.Text) * lblPromPA.Value)
'        End If
'    Else
'        If lblAmount.Value = 0 Then
'            txtPayment.Value = 0
'        Else
'            txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'            txtPayment.Value = lblAmount.Value - (CCur(txtDiscount.Text) * lblAmount.Value)
'        End If
'    End If
'End If
End Sub

Private Sub cmbNextAct_KeyDown(KeyCode As Integer, Shift As Integer)
cmbNextAct.text = ""
Exit Sub
End Sub

Private Sub CmbPhone_Change()
    If CmbPhone.text <> "" Then
        SSCommand1(0).Enabled = True
    End If
End Sub

Private Sub CmbPhone_Click()
    SSCommand1(0).Enabled = True
    CmbPhone.Locked = True
    FrmCC_Colection.Frame3.Caption = "0"
'    If CmbPhone.Text = "Add" Then
'        Frm_Tambah_Telp.Show vbModal
'    End If
End Sub

Private Sub CmbPhone_DropDown()
    CmbPhone.Locked = False
End Sub

Private Sub CmbPhone_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbUncontacted_Click()
'DESCRIPTION UNCONTACTED
Dim i As Integer
cmbDescUn.CLEAR
If Left(cmbUncontacted.text, 2) = "NA" Then
        cmbDescUn.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
          M_objrs.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_objrs.EOF
            cmbDescUn.AddItem M_objrs("Description")
            M_objrs.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_objrs = Nothing
'        FrmPayment.Enabled = False
Else
If Left(cmbUncontacted.text, 2) <> "MV" Then
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
   M_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
         For i = 0 To 3
           cmbDescUn.AddItem M_objrs("Description")
           M_objrs.MoveNext
         Next i
   Set M_objrs = Nothing
   C_Payment.Value = 0
Else
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
'   If kontak = True Then
'        m_objrs.Open "Select * from DescUncontacted where ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    Else
      M_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    End If
       While Not M_objrs.EOF
           cmbDescUn.AddItem M_objrs("Description")
           M_objrs.MoveNext
       Wend
   Set M_objrs = Nothing
   C_Payment.Value = 0
End If
End If
' Set M_OBJRS = New ADODB.Recordset
'   If kontak = False Then
'          M_OBJRS.Open "Select * from UncontactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'       While Not M_OBJRS.EOF
'           cmbDescUn.AddItem M_OBJRS("NMnoProdpresented")
'           M_OBJRS.MoveNext
'       Wend
'        Set M_OBJRS = Nothing
'   End If
'   C_Payment.Value = 0
'End If

End Sub

Private Sub headerDatePayment()
LstPayment.ColumnHeaders.ADD 1, , "", 0 * TXT
LstPayment.ColumnHeaders.ADD 2, , "ID", 1
LstPayment.ColumnHeaders.ADD 3, , "DATE", 1100
LstPayment.ColumnHeaders.ADD 4, , "PAYMENT", 30 * TXT
LstPayment.ColumnHeaders.ADD 5, , "TYPE", 30 * TXT
LstPayment.ColumnHeaders.ADD 6, , "INPUT DATE", 15 * TXT

LstReserve.ColumnHeaders.ADD 1, , "", 0 * TXT
LstReserve.ColumnHeaders.ADD 2, , "ID", 1
LstReserve.ColumnHeaders.ADD 3, , "DATE", 1100
LstReserve.ColumnHeaders.ADD 4, , "PAYMENT", 30 * TXT
LstReserve.ColumnHeaders.ADD 5, , "TYPE", 30 * TXT
LstReserve.ColumnHeaders.ADD 6, , "INPUT DATE", 15 * TXT

End Sub
Private Sub headerCustid_Double()
    LstDoubleId.ColumnHeaders.ADD 1, , "Id", 10 * TXT
    LstDoubleId.ColumnHeaders.ADD 2, , "Nama", 15 * TXT
    LstDoubleId.ColumnHeaders.ADD 3, , "DescColl", 10 * TXT
    LstDoubleId.ColumnHeaders.ADD 4, , "AmountWo", 10 * TXT
    LstDoubleId.ColumnHeaders.ADD 5, , "Principle", 20 * TXT
End Sub
Private Sub cmbUncontacted_KeyDown(KeyCode As Integer, Shift As Integer)
cmbUncontacted.text = ""
Exit Sub
End Sub
Private Sub Cmbwith_KeyDown(KeyCode As Integer, Shift As Integer)
Cmbwith.text = ""
Exit Sub
End Sub



Private Sub CmbStsKatHome1_Click()
    StsKategoriTelepon = Trim(CmbStsKatHome1.text)
    Call PilihSpeakWith
    Call CariKategoriTlp
End Sub

Private Sub CmbStsKatHome1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub



Private Sub CmbStsKatHome2_Click()
    StsKategoriTelepon = Trim(CmbStsKatHome2.text)
    Call PilihSpeakWith
    Call CariKategoriTlp
End Sub

Private Sub CmbStsKatHome2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmbStsKatHP1_Click()
    StsKategoriTelepon = Trim(CmbStsKatHP1.text)
    Call PilihSpeakWith
    Call CariKategoriTlp
End Sub

Private Sub CmbStsKatHP1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmbStsKatHP2_Click()
    StsKategoriTelepon = Trim(CmbStsKatHP2.text)
    Call PilihSpeakWith
    Call CariKategoriTlp
End Sub

Private Sub CmbStsKatHP2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmbStsKatOffice1_Click()
    StsKategoriTelepon = Trim(CmbStsKatOffice1.text)
    Call PilihSpeakWith
    Call CariKategoriTlp
End Sub

Private Sub CmbStsKatOffice1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmbStsKatOffice2_Click()
    StsKategoriTelepon = Trim(CmbStsKatOffice2.text)
    Call PilihSpeakWith
    Call CariKategoriTlp
End Sub

Private Sub CmbStsKatOffice2_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
End Sub

Private Sub CmbViaPtp_Click()
    If UCase(Trim(CmbViaPtp.text)) = "ATM LAINNYA" Then
        FrmPilihanAtmLainnya.Show vbModal
    End If
     '@@09-04-2012
    CariTanggalTagih
End Sub

Private Sub CmbViaPtp_KeyPress(KeyAscii As Integer)
     KeyAscii = 0
End Sub

Private Sub cmd_logcomplaint_Click()
    With Form_complaint
        .txt_custid.text = lblCustId.text
        .txt_custname.text = Replace(lblNama.text, "'", "")
        .txt_agent.text = lblaoc.Caption
        .Frame2.Enabled = False
        .cb_status.text = "OPEN"
        .lbl_ket = "N"
        .Show 1
    End With
End Sub

Private Sub CmdClaimAcc_Click()
    If UCase(lblaoc.Caption) <> "AKSESALL" Then
        MsgBox "Claim account hanya diperuntukkan bagi account yang di collect secara bersama-sama!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    Else
        'Pindahkan status account ke user claim
        FrmClaimAccount.TxtCustid.text = lblCustId.text
        FrmClaimAccount.txtnama.text = Replace(lblNama.text, "'", "")
        FrmClaimAccount.Show vbModal
    End If
End Sub

Private Sub CmdDataMapping_Click()
    '@@ 30-03-2012 Data Mapping dinonaktifkan, udah jarang dipake
    'FrmDataMapping.Show vbModal
    
    Dim a As String
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim M_Objrs_Cek As ADODB.Recordset
    
    a = MsgBox("Apakah anda akan membuat account ini sebagai Kept account untuk anda?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbYes Then
        'cek data dulu
        CMDSQL = "select * from tbl_keep_acc where date_part('year',tglkeep)="
        CMDSQL = CMDSQL + "date_part('year',now()) and date_part('month',tglkeep)="
        CMDSQL = CMDSQL + "date_part('month',now()) and agent='"
        CMDSQL = CMDSQL + lblaoc.Caption + "'"
        
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_objrs.RecordCount >= 20 Then
           MsgBox "Account keep anda sudah lebih mencapai 20 account. Maksimal account keep 20!", vbOKOnly + vbInformation, "Informasi"
        Else
            
            'Cek apakah custid ini sudah termasuk keep account
            CMDSQL = "select * from tbl_keep_acc where date_part('year',tglkeep)="
            CMDSQL = CMDSQL + "date_part('year',now()) and date_part('month',tglkeep)="
            CMDSQL = CMDSQL + "date_part('month',now()) and agent='"
            CMDSQL = CMDSQL + lblaoc.Caption + "' and custid='"
            CMDSQL = CMDSQL + lblCustId.text + "'"
            Set M_Objrs_Cek = New ADODB.Recordset
            M_Objrs_Cek.CursorLocation = adUseClient
            M_Objrs_Cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_Objrs_Cek.RecordCount > 0 Then
                MsgBox "Account ini sudah di keep sebelumnya!", vbOKOnly + vbInformation, "Informasi"
                Set M_Objrs_Cek = Nothing
                Exit Sub
            End If
            
            Set M_Objrs_Cek = Nothing
            
            CMDSQL = "insert into tbl_keep_acc (custid,agent,tglkeep,nama) values ('"
            CMDSQL = CMDSQL + lblCustId.text + "','"
            CMDSQL = CMDSQL + lblaoc.Caption + "','"
            CMDSQL = CMDSQL + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "','"
            CMDSQL = CMDSQL + Replace(lblNama.text, "'", "") + "')"
            M_OBJCONN.Execute CMDSQL
            
            'Update juga di tabel mgm
            CMDSQL = "update mgm set status_keep='1' where custid='"
            CMDSQL = CMDSQL + Trim(lblCustId.text) + "'"
            M_OBJCONN.Execute CMDSQL
            MsgBox "Keep account anda berhasil!", vbOKOnly + vbInformation, "Informasi"
        End If
        Set M_objrs = Nothing
    End If
End Sub

Private Sub CmddetailPayment_Click()
    FrmDetailPayment.Show 1
End Sub

Private Sub CmdHapusRemarks_Click()
    Dim CMDSQL As String
    Dim a As String
    
    If listview1(1).ListItems.Count = 0 Then
        MsgBox "Tidak ada data remarks yang akan dihapus!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Yakin data: " & listview1(1).SelectedItem.SubItems(1) & " akan dihapus?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbNo Then
        Exit Sub
    End If
    
    CMDSQL = "delete from mgm_hst where id='"
    CMDSQL = CMDSQL + Trim(listview1(1).SelectedItem.SubItems(7)) + "'"
    
    M_OBJCONN.Execute CMDSQL
    
    listview1(1).ListItems.Remove listview1(1).SelectedItem.Index
End Sub

Private Sub CmdKeep_Click()
 Dim CMDSQL As String
 Dim M_objrs As ADODB.Recordset
 Dim a As String
 
 CMDSQL = "select * from mgm where custid='"
 CMDSQL = CMDSQL + Trim(lblCustId.text) + "'"
 Set M_objrs = New ADODB.Recordset
 M_objrs.CursorLocation = adUseClient
 M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
 
 If M_objrs.RecordCount = 0 Then
    Set M_objrs = Nothing
    Exit Sub
 End If
 
 If M_objrs("status_htc") = "1" Then
    a = MsgBox("Apakah anda yakin akan menghilangkan status account ini tidak menjadi Hot Prospect?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbYes Then
        CMDSQL = "update mgm set status_htc=null where custid='"
        CMDSQL = CMDSQL + Trim(lblCustId.text) + "'"
        M_OBJCONN.Execute CMDSQL
        MsgBox "Status Hot Prospect untuk account ini telah dihilangkan!", vbOKOnly + vbInformation, "Informasi"
    End If
    
    '@@ 03-04-2012, Tanyakan ke user, apakah ingin menghapus data ini sebagai keep account juga?
    a = MsgBox("Apakah anda juga akan menghapus Kept Account untuk CH ini?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbYes Then
        CMDSQL = "delete from tbl_keep_acc where date_part('year',tglkeep)="
        CMDSQL = CMDSQL + "date_part('year',now()) and date_part('month',tglkeep)="
        CMDSQL = CMDSQL + "date_part('month',now()) and agent='"
        CMDSQL = CMDSQL + Trim(lblaoc.Caption) + "' and custid='"
        CMDSQL = CMDSQL + Trim(lblCustId.text) + "'"
        M_OBJCONN.Execute CMDSQL
        
        'Update status keep di mgm
        CMDSQL = "update mgm set status_keep=null where custid='"
        CMDSQL = CMDSQL + Trim(lblCustId.text) + "'"
        M_OBJCONN.Execute CMDSQL
        
        MsgBox "Kept Account untuk CH ini sudah dihapus!", vbOKOnly + vbInformation, "Informasi"
    End If
 End If
 
 If IsNull(M_objrs("status_htc")) = True Then
    a = MsgBox("Apakah anda yakin akan  menjadikan account ini  menjadi Hot Prospect?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbYes Then
        CMDSQL = "update mgm set status_htc='1' where custid='"
        CMDSQL = CMDSQL + Trim(lblCustId.text) + "'"
        M_OBJCONN.Execute CMDSQL
        MsgBox "Status Hot Prospect telah ditandai dalam account ini!", vbOKOnly + vbInformation, "Informasi"
    End If
    
    CmdDataMapping_Click
 End If
 
 
End Sub

Private Sub CmdOther_Click()
    FrmOther.Show vbModal
End Sub

Private Sub CmdRequest_Click()
'    '@@ 07-04-2011 Tambahan bikin Form Request
'    With Frm_Request
'        .TxtAgent.Text = lblaoc.Caption
'        .TxtCustid.Text = lblCustId.text
'        .TxtNamaCH.Text = lblNama.text
'
'        .TXtAmountWoPUM.Value = TDB_cur_bal.Value
'        .TxtPaymentDatePUM.Text = Format(lblPayDt.Value, "yyyy-mm-dd")
'        .Show vbModal
'    End With
    
    FrmListKeepAcc.Show vbModal
End Sub

Private Sub CmdRequestNumber_Click()
    With FrmReqTelepon
        .TxtCustid.text = lblCustId.text
        .Show vbModal
    End With


End Sub

Private Sub CmdSendPTP_Click()
    Dim M_Objrs_Cek2 As ADODB.Recordset
    
    If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Or _
       UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        If Trim(Mid(vrcek, 1, 3)) = "PO-" Then
            MsgBox "Untuk account yang statusnya PO-PAID OFF, tidak bisa send PTP! Hubungi SPV anda untuk mengubahnya!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
    End If
    
    CMDSQL = "select now() as tglskrg "
    Set M_Objrs_Cek2 = New ADODB.Recordset
    M_Objrs_Cek2.CursorLocation = adUseClient
    M_Objrs_Cek2.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    'If Format(M_Objrs_Cek2("tglskrg"), "d") > 26 Then
    '    MsgBox "Anda tidak bisa membuat PTP Lebih dari tanggal 25"
    'Else
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Or UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.text) = "MANAGER" Then
        FrmSendPTP.Show vbModal
    Else
        MsgBox "Anda tidak bisa melakukan PTP, Hanya deskcoll yang diperbolehkan!!"
    End If
    'End If
End Sub

Private Sub Combo1_Click()
'    Dim M_objrs As ADODB.Recordset
'    Dim CMDSQL As String
'    Dim M_Objrs_Cek2 As ADODB.Recordset
'
'    If Trim(UCase(Combo1.Text)) = "INCOMING" Then
'        CMDSQL = "select f_cek_new from mgm where custid='"
'        CMDSQL = CMDSQL + CStr(Trim(lblCustId.text)) + "'"
'        Set M_objrs = New ADODB.Recordset
'        M_objrs.CursorLocation = adUseClient
'        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If M_objrs.RecordCount > 0 Then
'            If IsNull(M_objrs("f_cek_new")) = False Then
'                cboaccount.Text = Trim(M_objrs("f_cek_new"))
'            Else
'                cboaccount.Text = "OS-"
'            End If
'        End If
'        cbolastcall.AddItem "CH"
'        cbolastcall.AddItem "SPOUSE"
'        cbolastcall.AddItem "FAMILY"
'        cbolastcall.AddItem "TETANGGA"
'        cbolastcall.AddItem "FRIEND"
'        cbolastcall.AddItem "HRD"
'        cbolastcall.AddItem "ATASAN"
'        cbolastcall.AddItem "OTHER"
'    Else
'        CMDSQL = "select f_cek_new from mgm where custid='"
'        CMDSQL = CMDSQL + CStr(Trim(lblCustId.text)) + "'"
'        Set M_Objrs_Cek2 = New ADODB.Recordset
'        M_Objrs_Cek2.CursorLocation = adUseClient
'        M_Objrs_Cek2.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If M_Objrs_Cek2.RecordCount > 0 Then
'            If IsNull(M_Objrs_Cek2("f_cek_new")) = False Then
'                cboaccount.Text = Trim(M_Objrs_Cek2("f_cek_new"))
'            Else
'                cboaccount.Text = "OS-"
'            End If
'        End If
'        'cboaccount.Text = ""
'    End If
End Sub

'@@ 05-10-2011, Tombol Unlock ditiadakan
'Private Sub CmdUnlock_Click()
'    '@@ 01/02/2011 UnLock Data Oleh agent
'    Dim a As String
'    Dim ID As String
'    Dim M_OBJRS As ADODB.Recordset
'    Dim m_objrs_cekid As ADODB.Recordset
'    Dim CMDSQL As String
'    Dim UpdateDtCloseSession As String
'    Dim m_objrs_ambilTL As ADODB.Recordset
'    Dim cmdsql_ambilTL As String
'
'    Dim pesan As String
'    Dim TglLock As String
'    Dim StartLock As String
'    Dim EndLock As String
'    Dim AccLock As String
'    Dim Status_lock As String
'
'    'Cek dulu apakah yang login agent?
'    If UCase(Trim(mdiform1.txtlevel.text)) <> "AGENT" Then
'        MsgBox "Unlock data ini hanya untuk AGENT!", vbOKOnly + vbExclamation, "Peringatan"
'        Exit Sub
'    End If
'
'    a = MsgBox("Anda yakin akan melakukan Unlock Data?", vbYesNo + vbQuestion, "Konfirmasi")
'    If a = vbNo Then
'        Exit Sub
'    End If
'
'    'Cek apakah ada data yang sedang di lock?
'    Set M_OBJRS = New ADODB.Recordset
'    M_OBJRS.CursorLocation = adUseClient
'    CMDSQL = "select * from usertbl where userid='"
'    CMDSQL = CMDSQL + Trim(mdiform1.txtusername.text) + "'"
'    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If M_OBJRS("lockdarispv") = "" And M_OBJRS("lock_entry_lpd") = "" And M_OBJRS("lockmarkup") = "" Then
'        MsgBox "Tidak ada data yang akan di unlock!", vbOKOnly + vbInformation, "Informasi"
'        Set M_OBJRS = Nothing
'        Exit Sub
'    End If
'    Set M_OBJRS = Nothing
'
'    'Cari id data yang sedang di lock
'    CMDSQL = "select *,now() as tanggal_sekarang from tbltemplockacc_current where id in "
'    CMDSQL = CMDSQL + "(select max(idlock) as idlock from tblperformpersessionlock where agent='"
'    CMDSQL = CMDSQL + Trim(mdiform1.txtusername.text) + "')"
'
'    Set m_objrs_cekid = New ADODB.Recordset
'    m_objrs_cekid.CursorLocation = adUseClient
'    m_objrs_cekid.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    ID = Trim(m_objrs_cekid("id"))
'    TglLock = Format(m_objrs_cekid("date_lock"), "yyyy-mm-dd hh:mm:ss")
'    StartLock = Format(m_objrs_cekid("start_lock"), "yyyy-mm-dd hh:mm:ss")
'    EndLock = Format(m_objrs_cekid("end_lock"), "yyyy-mm-dd hh:mm:ss")
'    AccLock = Trim(IIf(IsNull(m_objrs_cekid("account_lock")), "", m_objrs_cekid("account_lock")))
'    Status_lock = Trim(m_objrs_cekid("status_lock"))
'
'
'    'Catat ke dalam log
'    CMDSQL = "insert into log_unlock_agent (script_lock,date_lock,"
'    CMDSQL = CMDSQL + "start_lock,end_lock,account_lock,lock_by,f_locked,tgl_unlock,agent_unlock,status_lock,id) values ('"
'    CMDSQL = CMDSQL + Trim(IIf(IsNull(m_objrs_cekid("script_lock")), "", m_objrs_cekid("script_lock"))) + "','"
'    CMDSQL = CMDSQL + Format(m_objrs_cekid("date_lock"), "yyyy-mm-dd hh:mm:ss") + "','"
'    CMDSQL = CMDSQL + Format(m_objrs_cekid("start_lock"), "yyyy-mm-dd hh:mm:ss") + "','"
'    CMDSQL = CMDSQL + Format(m_objrs_cekid("end_lock"), "yyyy-mm-dd hh:mm:ss") + "','"
'    CMDSQL = CMDSQL + Trim(IIf(IsNull(m_objrs_cekid("account_lock")), "", m_objrs_cekid("account_lock"))) + "','"
'    CMDSQL = CMDSQL + Trim(IIf(IsNull(m_objrs_cekid("lock_by")), "", m_objrs_cekid("lock_by"))) + "','"
'    CMDSQL = CMDSQL + Trim(IIf(IsNull(m_objrs_cekid("f_locked")), "", m_objrs_cekid("f_locked"))) + "','"
'    CMDSQL = CMDSQL + Format(m_objrs_cekid("tanggal_sekarang"), "yyyy-mm-dd hh:mm:ss") + "','"
'    CMDSQL = CMDSQL + Trim(mdiform1.txtusername.text) + "','"
'    CMDSQL = CMDSQL + Trim(m_objrs_cekid("status_lock")) + "','"
'    CMDSQL = CMDSQL + Trim(ID) + "')"
'
'    M_OBJCONN.Execute CMDSQL
'
'    'Bikin pesan ke TL,jika lock datanya sudah di unlock oleh agent
'    pesan = vbCrLf + "INFORMASI OLEH SISTEM : " + vbCrLf
'    pesan = pesan + "Agent: " + mdiform1.txtusername.text + vbCrLf
'    pesan = pesan + "Melakukan Unlock data untuk accountnya sendiri." + vbCrLf
'    pesan = pesan + "Berikut informasi lock data yang di unlock:" + vbCrLf
'    pesan = pesan + "------------------------------------------------" + vbCrLf
'    pesan = pesan + "Tgl.Lock data :" + StartLock + vbCrLf
'    pesan = pesan + "Start.Lock data:" + EndLock + vbCrLf
'    pesan = pesan + "Account yang di lock:" + AccLock + vbCrLf
'    pesan = pesan + "Status yang di lock:" + Status_lock + vbCrLf
'    pesan = pesan + "------------------------------------------------" + vbCrLf
'    pesan = pesan + "Terima Kasih" + vbCrLf
'    pesan = pesan + "Message Created automatic by system"
'
'    MsgBox "Silahkan tunggu sebentar! Setelah menekan tombol OK ini, sistem akan melakukan unlock data. Harap Tunggu hingga muncul pesan Unlock data berhasil!", vbOKOnly + vbInformation, "Informasi"
'
'    'Pindahkan data ke tabel tblperformpersessionlock
'    DoEvents
'    UpdateDtCloseSession = "update tblperformpersessionlock set f_ceksekrg=a.f_cek_akhir ,"
'    UpdateDtCloseSession = UpdateDtCloseSession + " tgl_f_ceksekrg=a.tgl_akhir,endlock='" + CStr(Format(m_objrs_cekid("tanggal_sekarang"), "yyyy-mm-dd hh:mm:ss")) + "' from "
'    UpdateDtCloseSession = UpdateDtCloseSession + " (select mgm.custid as custid_mgm,mgm.agent as agent_mgm,"
'    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.custid as custid_lock,"
'    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.agent as agent_lock,"
'    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.idlock as id_lock,"
'    UpdateDtCloseSession = UpdateDtCloseSession + " mgm.f_cek_new as f_cek_akhir, mgm.tglcall as tgl_akhir"
'    UpdateDtCloseSession = UpdateDtCloseSession + " from tblperformpersessionlock inner join mgm "
'    UpdateDtCloseSession = UpdateDtCloseSession + " on mgm.custid=tblperformpersessionlock.custid "
'    UpdateDtCloseSession = UpdateDtCloseSession + " and mgm.agent=tblperformpersessionlock.agent) as a "
'    UpdateDtCloseSession = UpdateDtCloseSession + " where tblperformpersessionlock.custid=a.custid_mgm and tblperformpersessionlock.agent=a.agent_mgm and a.id_lock='"
'    UpdateDtCloseSession = UpdateDtCloseSession + Trim(ID) + "' and tblperformpersessionlock.agent='"
'    UpdateDtCloseSession = UpdateDtCloseSession + Trim(mdiform1.txtusername.text) + "'"
'    M_OBJCONN.Execute UpdateDtCloseSession
'
'    Set m_objrs_cekid = Nothing
'
'    cmdsqlserver = "update usertbl set dilockoleh='Release by:" + Trim(mdiform1.txtlevel.text) + "',"
'    cmdsqlserver = cmdsqlserver + "lockdarispv=null,lock_entry_lpd=null,fromaccount=null,"
'    cmdsqlserver = cmdsqlserver + "lockmarkup=null,lockdarispvbuattl=null where userid='"
'    cmdsqlserver = cmdsqlserver + Trim(mdiform1.txtusername.text) + "'"
'    M_OBJCONN.Execute cmdsqlserver
'
'    'Berikan pesan ke TL-nya
'    cmdsql_ambilTL = "select * from usertbl where userid='"
'    cmdsql_ambilTL = cmdsql_ambilTL + Trim(mdiform1.txtusername.text) + "'"
'    Set m_objrs_ambilTL = New ADODB.Recordset
'    m_objrs_ambilTL.CursorLocation = adUseClient
'    m_objrs_ambilTL.Open cmdsql_ambilTL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    CMDSQL = "insert into msgtbl  (recipient, datetime, sender, sentfrom, msg) VALUES ('"
'    CMDSQL = CMDSQL + Trim(m_objrs_ambilTL("team")) + "','"
'    CMDSQL = CMDSQL + CStr(Format(Now, "yyyymmdd")) + "','"
'    CMDSQL = CMDSQL + Trim(mdiform1.txtusername.text) + "','"
'    CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
'    CMDSQL = CMDSQL + Trim(pesan) + "')"
'    M_OBJCONN.Execute CMDSQL
'
'    Set m_objrs_ambilTL = Nothing
'
'    MsgBox "Data anda berhasil di Unlock!", vbOKOnly + vbInformation, "Informasi"
'    VIEW_MGMDATA.listview1.ListItems.CLEAR
'End Sub

Private Sub Command1_Click()
     If Command1.Tag = 0 Then
        Tdbbalance.Visible = True
        
        '@@ 0408201 Dibuang
        'tdbprincipal.Visible = True
        
        Label11(14).Visible = True
        
        '@@ 04082011 dibuang
        'Label11(15).Visible = True
        
        Command1.Tag = 1
        LblPrompA.Visible = True
        Label11(8).Visible = True
        Else
        Tdbbalance.Visible = False
        tdbprincipal.Visible = False
        Label11(14).Visible = False
        
        '@@ 04082011 dibuang
        'Label11(15).Visible = False
        
        Label11(8).Visible = False
        Command1.Tag = 0
        LblPrompA.Visible = False
        End If
        
End Sub

Private Sub Command1001_Click()
    frmaddphone.Show 1
End Sub

Private Sub Command2_Click()
    FrmSendSmsNew.Show vbModal
End Sub

Private Sub Command3_Click()
    If MsgBox("Account ini akan diset set menjadi decease??", vbYesNo + vbQuestion, "Confirm") = vbYes Then
        ' DELETE BEFORE
        M_OBJCONN.Execute "DELETE FROM tblreq_decease WHERE custid='" & CStr(Trim(lblCustId.text)) & "'"
        M_OBJCONN.Execute "INSERT INTO tblreq_decease(custid,agent) VALUES('" & CStr(Trim(lblCustId.text)) & "','" & MDIForm1.TxtUsername.text & "')"
        MsgBox "Account telah diset menjadi Acc Decease, Tunggu approval dari TL", vbOKOnly + vbInformation, "INFO"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Exit Sub
End Sub

Private Sub Form_Load()
    
    DisableCloseBtn Me
    
    Dim dLast_Payment As Date
    Dim dLast_PTP As Date
    Dim status_cek_frame As Integer
    
    MDIForm1.Timer100.Enabled = False
    
    On Error Resume Next
    
    waktu_mulai_ngitung = waktu_server_sekarang
    
    'RANDY : CEK AKTIF / TIDAK CEKBOX UNTUK AGENT CHANGE PTP (REQ DODDY)
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "SELECT status_cek_frame FROM tbl_cek_framePTP", M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    status_cek_frame = IIf(IsNull(M_objrs!status_cek_frame), "", M_objrs!status_cek_frame)
    
    If status_cek_frame = 1 Then
        chk_aktif.Value = 1
    Else
        chk_aktif.Value = 0
    End If
    
    lbltime_save.Caption = waktu_server_sekarang
    lblstop_time.Caption = waktu_server_sekarang
        
    LstPayment.Checkboxes = True
        
    SSCommand1(0).Enabled = False
    
    ' ## Set Status Form Customer Aktif 12 Mei 2013 By Izuddin
    bAktif_form_customer = True
    ' # 08 April 2013 Monitoring Activity By Izuddin
    i_monitoring_activity = 0
    'MDIForm1.Timer2.Enabled = True
    
    '@@15092012 Aktifitas Call di kosongin dulu
    AktifitasCall = ""
    calling = ""
    
    StsKategoriTelepon = ""
    KelompokKategoriTlp = ""
    kat_aktif_telp = ""
    
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        SSCommand1(4).Visible = False
        Command1.Visible = False
        'Jika agent c_ptp didisable 11 Juni 2012
        C_PTP.Enabled = False
        
    ElseIf UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.text) = "ADMIN" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Or UCase(MDIForm1.txtlevel.text) = "MANAGER" Then
            SSCommand1(4).Visible = True
            Command1.Visible = False
            CmdHapusRemarks.Visible = True
            cmd_logcomplaint.Visible = True
    End If
    
    '@@19042012, Tombol Hangup Di nonaktifkan dulu
    SSCommand1(1).Enabled = False
    
    
    FrmCC_Colection.Left = 10
    FrmCC_Colection.Top = 20
    
    'cek list pelunasan
    Dim i, iIndex As Integer
    Dim sKata, cCombo As String
    
    
    '------->>>  setting No Visit  <<<---------------
    Text1.text = Format(Now, "yymmddhhmmss")
    TDBDate1.Value = Now
    'If UCase(Left(mdiform1.txtlevel.text, 5)) = "ADMIN" Or UCase(Left(mdiform1.txtlevel.text, 5)) = "SUPER" Then
    If UCase(Left(MDIForm1.txtlevel.text, 5)) = "ADMIN" Then
        txtHomeNo1.Visible = True
        txtHomeNo1A.Visible = False
        txtHomeNo2.Visible = True
        txtHomeNo2A.Visible = False
        txtOfficeNo1.Visible = True
        txtOfficeNo1A.Visible = False
        txtOfficeNo2.Visible = True
        txtOfficeNo2A.Visible = False
        txtMobileNo1.Visible = True
        txtMobileNo1A.Visible = False
        txtMobileNo2.Visible = True
        txtMobileNo2A.Visible = False
        txtPhone.Visible = True
        txtPhoneA.Visible = False
        txtHomeAdd1.Visible = True
        txtHomeAdd1A.Visible = False
        txtHomeAdd2.Visible = True
        txtHomeAdd2A.Visible = False
        txtOfficeAdd1.Visible = True
        txtOfficeAdd1A.Visible = False
        txtOfficeAdd2.Visible = True
        txtOfficeAdd2A.Visible = False
        txtMobileAdd1.Visible = True
        txtMobileAdd1A.Visible = False
        txtMobileAdd2.Visible = True
        txtMobileAdd2A.Visible = False
        txtECno.Visible = True
        txtECnoA.Visible = False
        Tdbbalance.Visible = False
            '@@ 0408201 Dibuang
            'tdbprincipal.Visible = False
            
            Label11(14).Visible = False
            
            '@@ 04082011 Dibuang
            'Label11(15).Visible = False
            
            'aktifkan recsource @@ 160610
            label1(80).Visible = True
            lblRecsource.Visible = True
    End If
    
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
            C_lunas.Enabled = False
            TdbLunas.Enabled = False
            'chkAppv(0).Enabled = False '@@25/01/2012 Buangin komponen tak terpakai 25012012
            'chkAppv(1).Enabled = False '@@25/01/2012 Buangin komponen tak terpakai 25012012
            TDBTot_payment.Enabled = False
            TxtFieldName.Enabled = False
            
            '@@ 05-10-2011 Tombol Hapus Tabel Lunas ditiadakan terlebih dahulu
            'CmdDeletePelunasan.Enabled = False
             
             'DISABLE AGENT CHANGE PTP
             'TDBDate3.Enabled = False
             
             ' Tampilkan PRincipal
            SSCommand2(3).Enabled = False
            SSCommand2(2).Enabled = False
            
            lblhapus.Enabled = False
            Label41.Enabled = False
            LblPrompA.Visible = True
            Label11(8).Visible = True
            Tdbbalance.Visible = False
            '@@ 0408201 Dibuang
            'tdbprincipal.Visible = False
            
            Label11(14).Visible = False
            
            '@@ 04082011 Dibuang
            'Label11(15).Visible = False
           
    Else ' utk SPV tampilkan no telp
            txtHomeAdd1.ReadOnly = False
            txtHomeAdd2.ReadOnly = False
            txtOfficeAdd1.ReadOnly = False
            txtOfficeAdd2.ReadOnly = False
            txtMobileAdd1.ReadOnly = False
            txtMobileAdd2.ReadOnly = False
            SSCommand2(3).Enabled = True
            SSCommand2(2).Enabled = True
            lblhapus.Enabled = True
            Label41.Enabled = True
            
            txtHomeNo1.Visible = True
            txtHomeNo1A.Visible = False
            txtHomeNo2.Visible = True
            txtHomeNo2A.Visible = False
            
            txtOfficeNo1.Visible = True
            txtOfficeNo1A.Visible = False
            
            txtOfficeNo2.Visible = True
            txtOfficeNo2A.Visible = False
            
            txtMobileNo1.Visible = True
            txtMobileNo1A.Visible = False
            txtMobileNo2.Visible = True
            txtMobileNo2A.Visible = False
            
            txtECno.Visible = True
            txtECnoA.Visible = False
            
            txtHomeAdd1.Visible = True
            txtHomeAdd1A.Visible = False
            txtHomeAdd2.Visible = True
            txtHomeAdd2A.Visible = False
            
            txtOfficeAdd1.Visible = True
            txtOfficeAdd1A.Visible = False
            txtOfficeAdd2.Visible = True
            txtOfficeAdd2A.Visible = False
            
            txtMobileAdd1.Visible = True
            txtMobileAdd1A.Visible = False
            txtMobileAdd2.Visible = True
            txtMobileAdd2A.Visible = False
            ' Tampilkan PRincipal
            LblPrompA.Visible = True
            Label11(8).Visible = True
            'aktifkan recsource @@ 160610
            label1(80).Visible = True
            lblRecsource.Visible = True
            
    End If
     
     '  FrmContacted.Enabled = False
   FrmUnContacted.Enabled = False
   'FrmPayment.Enabled = False
   
    Call headerDatePayment
    Call headerCustid_Double
    Call HEADER_HISTORY
    Call HEADER_HISTORY_PAID
    'Call HEADER_RequestVisit
    Call HEADER_Detail_Customer
    'Call HEADER_SendSMS
    On Error Resume Next
    Call show_cust
    
    '@@ 05-06-2012, Jika Status Complain dan Paid OFF maka kategori telepon tidak dapat dipilih
    If StatusAccount = "CO-" Or StatusAccount = "PO-" Then
        CmbStsKatHome1.Enabled = False
        CmbStsKatHome2.Enabled = False
        CmbStsKatOffice1.Enabled = False
        CmbStsKatOffice2.Enabled = False
        CmbStsKatHP1.Enabled = False
        CmbStsKatHP2.Enabled = False
        CmdRequestNumber.Enabled = False
     End If
    
    Call VisitNo
'    Call isi_lastcall
    
    If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Or UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
        Call aktifphone
    End If
    
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        Call aktifphoneAGENT
    End If
    
    '@@14022011
    Call CekSms
            
      '  SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    SSTab1.Tab = 0
    cmbDateSch.Value = Now
    cmbDateSch.Value = ""
    'CONTACTED
    CmbBaseOn.AddItem "PRINCIPLE"
    CmbBaseOn.AddItem "TOTAL AMOUNT"
    
        
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "Select * from tblptp where KdNoProdPresented not like 'PTP-PAID%' ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_objrs.EOF
        cboPTP.AddItem M_objrs!KdNoProdPresented
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    
        
    '@@ 24 May 2012 Akses 108, untuk agent tertentu saja
    Dim M_objrs_108 As ADODB.Recordset
    CMDSQL = "select sts_108 from usertbl where userid='"
    CMDSQL = CMDSQL + CStr(MDIForm1.TxtUsername.text) + "' " 'and sts_108='1'"
    Set M_objrs_108 = New ADODB.Recordset
    M_objrs_108.CursorLocation = adUseClient
    M_objrs_108.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs_108.RecordCount > 0 Then
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open "Select * from tbllayanantelkom", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        While Not M_objrs.EOF
            CmbPhone.AddItem IIf(IsNull(M_objrs("Nolayanan")), "", M_objrs("Nolayanan"))
            M_objrs.MoveNext
        Wend
        Set M_objrs = Nothing
    End If
    Set M_objrs_108 = Nothing
    
    '@@25052012 Jika yang login Admin,Superviso
    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or _
       UCase(MDIForm1.txtlevel.text) = "ADMIN" Or _
       UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open "Select * from tbllayanantelkom", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        While Not M_objrs.EOF
            CmbPhone.AddItem IIf(IsNull(M_objrs("Nolayanan")), "", M_objrs("Nolayanan"))
            M_objrs.MoveNext
        Wend
        Set M_objrs = Nothing
    End If
    
    'sembunyiin principle kecuali SPV
    If UCase(MDIForm1.txtlevel.text) <> "SUPERVISOR" Then
        LblPrompA.Visible = False
        Label11(8).Visible = False
    Else
        LblPrompA.Visible = True
        Label11(8).Visible = True
    End If
    
    '@@ 15-04-2011 Panggil CekCPA, jika ada data CPA maka kelap-kelip
    Call CekCPA
    
    '@@11 Juni 2012 Jika Yang Login Agent maka form PTP disable
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        frmPTP.Enabled = False
    End If
    
    'RANDY UPDATE
    If chk_aktif = vbChecked Then
        frmPTP.Enabled = True
    End If
    
    'RANDY
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Or UCase(MDIForm1.txtlevel.text) = "ADMIN" Or UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
        chk_aktif.Visible = False
    End If
    sudahCall = False
    
    ' SAMAKAN TGL PEMBAYARAN DENGAN TANGGAL PTP YG TERAKHIR 01 JULI 2014 BY IZUDDIN VIA DODDY
    If (listview1(0).ListItems.Count > 0 And LstPayment.ListItems.Count > 0) Then
        dLast_Payment = Format(listview1(0).ListItems(1).text, "yyyy-mm-dd")
        dLast_PTP = Format(LstPayment.ListItems(1).SubItems(2), "yyyy-mm-dd")
        If (Month(dLast_Payment) = Month(dLast_PTP)) And (Year(dLast_Payment) = Year(dLast_PTP)) Then
            If dLast_PTP > dLast_Payment Then
                ' Cek juga di list Item yang ke-2
                If Month(dLast_PTP) <> Month(Format(LstPayment.ListItems(2).SubItems(2), "yyyy-mm-dd")) Then
'                    M_OBJCONN.Execute "UPDATE tblnegoptp SET promisedate='" & Format(dLast_Payment, "yyyy-mm-dd") & "' WHERE id=" & LstPayment.ListItems(1).SubItems(1)
'                    M_OBJCONN.Execute "UPDATE mgm SET dateptp='" & Format(dLast_Payment, "yyyy-mm-dd") & "' WHERE custid='" & lblCustId.text & "'"
                    Call Show_NEGOPTP
                End If
            End If
        End If
    End If
    ' ---------------------------------------------------------------------------------------
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        'Label1(80).Visible = False
        'lblRecsource.Visible = False
    End If
    
    CMDSQL = "  SELECT tblstatuscall_keterangan,grp_call FROM tblstatuscall WHERE tblstatuscall_kdstatus='1' and tblstatuscall_keterangan not in ('New Data')order by tblstatuscall_id,2 "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    'cboaccount.CLEAR
    While Not M_objrs.EOF
        Combo1.AddItem IIf(IsNull(M_objrs!grp_call), "", M_objrs!grp_call)
        cboaccount.AddItem IIf(IsNull(M_objrs!tblstatuscall_keterangan), "", M_objrs!tblstatuscall_keterangan)
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    
    SSFrame1.Left = (Screen.Width - SSFrame1.Width) / 2
    SSFrame1.Top = ((Screen.Height - SSFrame1.Height) / 2) + 375
    Shape1.Width = Screen.Width
    
    If MDIForm1.txtlevel.text = "Agent" Then
'        If cboaccount.Text <> "" Then
'            Call statusgroup
'        End If
        
        If C_PTP.Value = vbChecked Then
            Text10.text = "1"
        Else
            Text10.text = "0"
        End If
    End If
    
    
    'Call isi_datacustomer
'    txtHomeNo1.Enabled = False
'    txtOfficeNo1.Enabled = False
'    txtMobileNo1.Enabled = False
    If lblRecsource.Caption = "Satukosonglapan" Then
        cmd_req_telp.Visible = False
        SSCommand1(2).Visible = False
    End If
End Sub

'validasi statuscall TIAN (21Dec2016)
Private Sub statusgroup()
   
    statuscall = cboaccount.text
    
    query = "SELECT grp_call from tblstatuscall where tblstatuscall_keterangan = '" + statuscall + "'"
    Set rs_1 = New ADODB.Recordset
    rs_1.CursorLocation = adUseClient
    rs_1.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    cboaccount.CLEAR
    grpcall = rs_1!grp_call
    If grpcall = "CONTACT" Then
        query = "SELECT tblstatuscall_keterangan,grp_call FROM tblstatuscall WHERE tblstatuscall_kdstatus='1' and (grp_call= 'CONTACT' or coalesce(grp_call,'')='') AND tblstatuscall_keterangan not in ('New Data') order by 1 "
        Set rs_2 = New ADODB.Recordset
        rs_2.CursorLocation = adUseClient
        rs_2.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ElseIf grpcall = "CONNECT" Then
        query = "SELECT tblstatuscall_keterangan,grp_call FROM tblstatuscall WHERE tblstatuscall_kdstatus='1' and (grp_call != 'UNCONNECT' or coalesce(grp_call,'')='')  AND tblstatuscall_keterangan not in ('New Data') order by 1 "
        Set rs_2 = New ADODB.Recordset
        rs_2.CursorLocation = adUseClient
        rs_2.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ElseIf grpcall = "UNCONNECT" Then
        query = "SELECT tblstatuscall_keterangan,grp_call FROM tblstatuscall WHERE tblstatuscall_kdstatus='1' AND tblstatuscall_keterangan not in('New Data') order by 2,1 "
        Set rs_2 = New ADODB.Recordset
        rs_2.CursorLocation = adUseClient
        rs_2.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic
    End If
    While Not rs_2.EOF
        Combo1.AddItem IIf(IsNull(rs_2!grp_call), "", rs_2!grp_call)
        cboaccount.AddItem IIf(IsNull(rs_2!tblstatuscall_keterangan), "", rs_2!tblstatuscall_keterangan)
        rs_2.MoveNext
    Wend
    cboaccount.text = statuscall
End Sub

Sub isi_lastcall()
cbolastcall.CLEAR
Dim M_objrs As ADODB.Recordset
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient

If MDIForm1.txtlevel.text = "AGENT" Then
    M_objrs.Open "Select * from ContactedDesc where kdnoprodpresented <> 'SP-SETTLE PAYMENT' ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    Else
    M_objrs.Open "Select * from ContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
End If
While Not M_objrs.EOF
    cbolastcall.AddItem Trim(M_objrs("KdNoProdPresented"))
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing

Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not M_objrs.EOF
    cbolastcall.AddItem Trim(M_objrs("KdNoProdPresented"))
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing
End Sub

Private Sub aktifphone()
'@@03-05-2012 DinonAktifkan
'AHomeAdd1(0).ReadOnly = False
'@@03-05-2012 Dinonaktifkan
'AHomeAdd2(1).ReadOnly = False

txtHomeAdd1.ReadOnly = False
txtHomeAdd1A.ReadOnly = False
txtHomeAdd2.ReadOnly = False
txtHomeAdd2A.ReadOnly = False

'@@03-05-2012 Dinonaktifkan
'AOfficeAdd(2).ReadOnly = False
'AOfficeAdd(3).ReadOnly = False

txtOfficeAdd1.ReadOnly = False
txtOfficeAdd1A.ReadOnly = False
txtOfficeAdd2.ReadOnly = False
txtOfficeAdd2A.ReadOnly = False
txtMobileAdd1.ReadOnly = False
txtMobileAdd1A.ReadOnly = False
txtMobileAdd2.ReadOnly = False
txtMobileAdd2A.ReadOnly = False

'txtECno.ReadOnly = False
'txtECnoA.ReadOnly = False
'@@11052012 EC dinonaktifkan
txtECno.ReadOnly = True
txtECnoA.ReadOnly = True
End Sub

Private Sub aktifphoneAGENT()
If txtHomeAdd1.Value = "" Then
    txtHomeAdd1.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AHomeAdd1(0).ReadOnly = False
End If
If txtHomeAdd1A.Value = "" Then
    txtHomeAdd1A.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AHomeAdd1(0).ReadOnly = False
End If
If txtHomeAdd2.Value = "" Then
    txtHomeAdd2.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AHomeAdd2(1).ReadOnly = False
End If
If txtHomeAdd2A.Value = "" Then
    txtHomeAdd2A.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AHomeAdd2(1).ReadOnly = False
End If
If txtOfficeAdd1.Value = "" Then
    txtOfficeAdd1.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AOfficeAdd(2).ReadOnly = False
End If
If txtOfficeAdd1A.Value = "" Then
    txtOfficeAdd1A.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AOfficeAdd(2).ReadOnly = False
End If
If txtOfficeAdd2.Value = "" Then
    txtOfficeAdd2.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AOfficeAdd(3).ReadOnly = False
End If
If txtOfficeAdd2A.Value = "" Then
    txtOfficeAdd2A.ReadOnly = False
    '@@03-05-2012 Dinonaktifkan
    'AOfficeAdd(3).ReadOnly = False
End If
If txtMobileAdd1.Value = "" Then
    txtMobileAdd1.ReadOnly = False
End If
If txtMobileAdd1A.Value = "" Then
    txtMobileAdd1A.ReadOnly = False
End If
If txtMobileAdd2.Value = "" Then
    txtMobileAdd2.ReadOnly = False
End If
If txtMobileAdd2A.Value = "" Then
    txtMobileAdd2A.ReadOnly = False
End If
If txtECno.Value = "" Then
    txtECno.ReadOnly = True
End If
If txtECnoA.Value = "" Then
    txtECnoA.ReadOnly = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim n%
    For n = 1 To LstPayment.ListItems.Count
            If LstPayment.ListItems(n).SubItems(4) = "UNSCH" And regnego = True Then
                regnego = True
            End If
    Next n
    
    If regnego = False Or LstPayment.ListItems.Count = 0 Then
        kontak = False
        shedulePTP_Show = False
        regnego = False
        ' 'M_OBJCONN.Close
        ' Di cek lagi kenapa ada putus koneksi 17-09-2013 // CEK BY IZUDDIN
'        M_OBJCONN.Close
'        Set M_OBJCONN = Nothing
'        M_OBJCONN.Open CMDSQLOPEN
        VIEW_MGMDATA.WindowState = 2
    Else
        MsgBox "Lakukan PTP yang benar,Jumlah PTP harus >= Deal Payment " & txtPayment.text & " , Atau data simpan dulu!!!"
        Cancel = 1
        i_monitoring_activity = 0
        Exit Sub
    End If
    ' Reset and disable monitoring
    i_monitoring_activity = 0
    'MDIForm1.Timer2.Enabled = False
    ' ####
    ' Reset REMINDER ##############
    bAktif_form_customer = False
    bReminder_agent = False
    bAktif_Cust_Review = False
    ' #############################
    'Call VIEW_MGMDATA.tampil_waktu
End Sub






Private Sub Image1_Click(Index As Integer)
    Select Case Index
       Case 0
'          If Image1(0).Tag = 0 Then
'            Tdbbalance.Visible = True
'            tdbprincipal.Visible = True
'            Label11(14).Visible = True
'            Label11(15).Visible = True
'            Image1(0).Tag = 1
'            LblPrompA.Visible = True
'            Label11(8).Visible = True
'        Else
'            Tdbbalance.Visible = False
'            tdbprincipal.Visible = False
'            Label11(14).Visible = False
'            Label11(15).Visible = False
'            Label11(8).Visible = False
'            Image1(0).Tag = 0
'            LblPrompA.Visible = False
'        End If

    End Select
End Sub

Private Sub Label1_Click(Index As Integer)
  Dim ami As Integer
  
  Select Case Index
        Case 80
        'If label1(80).Tag = 0 Then
          If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.text) = "ADMIN" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
                   Tdbbalance.Visible = True
                   '@@ 0408201 Dibuang
                   'tdbprincipal.Visible = True
                   
                   Label11(14).Visible = True
                   
                   '@@ 04082011 Dibuang
                   'Label11(15).Visible = True
                   
                   label1(80).Tag = 1
                   LblPrompA.Visible = True
                   Label11(8).Visible = True
                   For ami = 1 To LstDoubleId.ListItems.Count
                       LstDoubleId.ListItems(ami).SubItems(4) = ENCRIPY(True, LstDoubleId.ListItems(ami).SubItems(4))
                   Next ami
               Else
                   Tdbbalance.Visible = False
                   
                   '@@ 0408201 Dibuang
                   'tdbprincipal.Visible = False
                   
                   Label11(14).Visible = False
                   
                   '@@ 04082011 Dibuang
                   'Label11(15).Visible = False
                   
                   Label11(8).Visible = False
                   label1(80).Tag = 0
                   LblPrompA.Visible = False
                    For ami = 1 To LstDoubleId.ListItems.Count
                       LstDoubleId.ListItems(ami).SubItems(4) = ENCRIPY(False, LstDoubleId.ListItems(ami).SubItems(4))
                   Next ami
               End If
            Case 8
                frmextensioncc.Show 1
End Select

End Sub


Private Sub Label4_Click()
    Dim CMDSQL, a As String
    
    If TxtNoTelpReq.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = TxtNoTelpReq.Value
            .LblTelp = "Req Telp"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_homeno='1',f_valid_home1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_home1='1', f_sts_valid_home1='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistAddHome1_Click()
    Dim CMDSQL, a As String
    
    If txtHomeAdd1.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = txtHomeAdd1.Value
            .LblTelp.Caption = "AddHome 1"
            If MDIForm1.txtlevel.text = "Agent" Then
                .Show vbModal
            Else
                .Show
            End If
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_homenoadd1='1',f_valid_addhome1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_addhome1='1', f_sts_valid_addhome1='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistAddHome2_Click()
    Dim CMDSQL, a As String
    
    If txtHomeAdd2.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = txtHomeAdd2.Value
            .LblTelp.Caption = "AddHome 2"
            If MDIForm1.txtlevel.text = "Agent" Then
                .Show vbModal
            Else
                .Show
            End If
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_homenoadd2='1',f_valid_addhome2=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_addhome2='1', f_sts_valid_addhome2='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistAddHP1_Click()
      Dim CMDSQL, a As String
    
    If txtMobileAdd1.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = txtMobileAdd1.Value
            .LblTelp.Caption = "AddMobile 1"
            If MDIForm1.txtlevel.text = "Agent" Then
                .Show vbModal
            Else
                .Show
            End If
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_mobilenoadd1='1',f_valid_addmobile1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_addmobile1='1', f_sts_valid_addmobile1='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistAddHP2_Click()
    
    If txtMobileAdd2.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = txtMobileAdd2.Value
            .LblTelp.Caption = "AddMobile 2"
            If MDIForm1.txtlevel.text = "Agent" Then
                .Show vbModal
            Else
                .Show
            End If
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_mobilenoadd2='1',f_valid_addmobile2=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_addmobile2='1', f_sts_valid_addmobile2='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
             MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistAddOffice1_Click()
    Dim CMDSQL, a As String
    
    If txtOfficeAdd1.Value <> Empty Then
        
       a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = txtOfficeAdd1.Value
            .LblTelp.Caption = "AddOffice 1"
            If MDIForm1.txtlevel.text = "Agent" Then
                .Show vbModal
            Else
                .Show
            End If
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_officenoadd1='1',f_valid_addoffice1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_addoffice1='1', f_sts_valid_addoffice1='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
             MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistAddOffice2_Click()
    Dim CMDSQL, a As String
    
    If txtOfficeAdd2.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = txtOfficeAdd2.Value
            .LblTelp.Caption = "AddOffice 2"
            If MDIForm1.txtlevel.text = "Agent" Then
                .Show vbModal
            Else
                .Show
            End If
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_officenoadd2='1',f_valid_addoffice2=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_addoffice2='1', f_sts_valid_addoffice2='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlackliSTEC_Click()
    Dim CMDSQL, a As String
    
    If txtECno.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = txtECno.Value
            .LblTelp.Caption = "EC"
            .Show vbModal
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_ec_telp='1',f_valid_ec=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_ec='1', f_sts_valid_ec='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistHome2_Click()
    Dim CMDSQL, a As String
    
    If txtHomeNo2.Value <> Empty Then
        
       a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = txtHomeNo2.Value
            .LblTelp.Caption = "Home 2"
            If MDIForm1.txtlevel.text = "Agent" Then
                .Show vbModal
            Else
                .Show
            End If
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'             If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_homeno2='1',f_valid_home2=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'             ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_home2='1', f_sts_valid_home2='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
             'End If
             MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistHp1_Click()
    Dim CMDSQL, a As String
    
    If txtMobileNo1.text <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = txtMobileNo1.text
            .LblTelp.Caption = "Mobile 1"
            If MDIForm1.txtlevel.text = "Agent" Then
                .Show vbModal
            Else
                .Show
            End If
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_mobileno='1',f_valid_mobile1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                 'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_mobile1='1', f_sts_valid_mobile1='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistHp2_Click()
    Dim CMDSQL, a As String
    
    If txtMobileNo2.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = txtMobileNo2.Value
            .LblTelp.Caption = "Mobile 2"
            If MDIForm1.txtlevel.text = "Agent" Then
                .Show vbModal
            Else
                .Show
            End If
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_mobileno2='1',f_valid_mobile1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_mobile2='1', f_sts_valid_mobile2='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistOffice1_Click()
    Dim CMDSQL, a As String
    
    If txtOfficeNo1.text <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = txtOfficeNo1.text
            .LblTelp.Caption = "Office 1"
            If MDIForm1.txtlevel.text = "Agent" Then
                .Show vbModal
            Else
                .Show
            End If
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_officeno='1',f_valid_office1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_office1='1', f_sts_valid_office1='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
            'End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlacklistOfficeno2_Click()
    Dim CMDSQL, a As String
    
    If txtOfficeNo2.Value <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = txtOfficeNo2.Value
            .LblTelp.Caption = "Office 2"
            If MDIForm1.txtlevel.text = "Agent" Then
                .Show vbModal
            Else
                .Show
            End If
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'             If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_officeno2='1',f_valid_office2=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_office2='1', f_sts_valid_office2='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblBlakcListHome1_Click()
    Dim CMDSQL, a As String
    
    If txtHomeNo1.text <> Empty Then
        
        a = MsgBox("Apakah anda yakin akan menandai validitas nomor telepon ini?", vbYesNo + vbQuestion, "Konfirmasi")
        
        If a = vbNo Then
            Exit Sub
        End If
        
        With FrmBlackListTelpAgent
            .TxtNoTelp.text = txtHomeNo1.text
            .LblTelp = "Home 1"
            If MDIForm1.txtlevel.text = "Agent" Then
                .Show vbModal
            Else
                .Show
            End If
        End With
        
        If FrmBlackListTelpAgent.ok = True Then
'            If FrmBlackListTelpAgent.STATUS = "Black List Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_homeno='1',f_valid_home1=null where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Data Berhasil di Blacklist!", vbOKOnly + vbInformation, "Informasi"
'            ElseIf FrmBlackListTelpAgent.STATUS = "Valid Number" Then
'                'Update Nomor Telepon
'                Cmdsql = "update mgm set f_valid_home1='1', f_sts_valid_home1='"
'                Cmdsql = Cmdsql + CStr(Trim(IIf(IsNull(FrmBlackListTelpAgent.TxtKeterangan.Text), "", FrmBlackListTelpAgent.TxtKeterangan.Text))) + "' "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql + CStr(lblCustId.text) + "'"
'                M_OBJCONN.Execute Cmdsql
'                MsgBox "Nomor Telepon Berhasil ditandai sebagai valid number!", vbOKOnly + vbInformation, "Informasi"
'            End If
            MsgBox "Data berhasil ditandai!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
End Sub

Private Sub LblMap_Click()
    TimerBlinkDetailMapping.Enabled = False
    FrmDetailMapping.Show vbModal
End Sub

Private Sub ListView1_Click(Index As Integer)
Dim KET As String
Select Case Index
Case 0

Case 1
If listview1(1).ListItems.Count = 0 Then
Exit Sub
Else
   KET = TXtDetails.text
      If Len(TXtDetails) = 0 Then
         TXtDetails.text = " - " + listview1(1).SelectedItem.SubItems(1)
      Else
         TXtDetails.text = KET + " - " + listview1(1).SelectedItem.SubItems(1)
      End If
End If
End Select
End Sub
Private Sub ListView1_DblClick(Index As Integer)
    Dim iret As Long

    Select Case Index
    Case 1
        If MDIForm1.txtlevel.text <> "Agent" Then 'INBOUND
            If listview1(1).ListItems.Count <> 0 Then
               
                If listview1(1).SelectedItem.SubItems(8) <> Empty Then
                    MDIForm1.ActionCTI ("RECORDING|" + listview1(1).SelectedItem.SubItems(8))
                    THandle = FindWindow(vbEmpty, "DCCS Client 1.2.1")
                    iret = BringWindowToTop(THandle)
                Else
                    MsgBox "Unixrecord tidak ada", vbInformation + vbOKOnly, "TINS"
                    Exit Sub
                End If
            End If
        End If
    End Select
End Sub

Private Sub LstDoubleId_DblClick()
     If LstDoubleId.ListItems.Count = 0 Then
        Exit Sub
    End If
    FrmCC_Colection.Hide
    frmCC_Colection2.Show vbModal
End Sub

Private Sub LstPayment_DblClick()
If LstPayment.ListItems.Count = 0 Then
Exit Sub
Else
Call SSCommand2_Click(1)
End If
End Sub
Private Sub Lstscript_DblClick()
  If Lstscript.ListItems.Count > 0 Then
  StartMeUp (Lstscript.SelectedItem.SubItems(2))
  'MsgBox (LstScript.SelectedItem.SubItems(2))
   End If
End Sub
'@@ 11-03-2011 Di remarks, udah tidak diapakai
'Private Sub LstSMS_DblClick()
'If LstSMS.ListItems.Count > 0 Then
'
'no_telp = LstSMS.SelectedItem.Text
'isi_Pesan = LstSMS.SelectedItem.SubItems(3)
'
'MsgBox "No Telepon : " & no_telp & vbCrLf & "Isi Pesan : " & Trim(isi_Pesan)
'
'    Else
'    Exit Sub
' End If
'End Sub

'@@ 11-03-2011 Di remarks, udah tidak diapakai

'Private Sub LstSMS2_DblClick()
'If LstSMS2.ListItems.Count > 0 Then
'
'no_telp = LstSMS2.SelectedItem.Text
'isi_Pesan = LstSMS2.SelectedItem.SubItems(2)
'
'MsgBox "No Telepon : " & no_telp & vbCrLf & "Isi Pesan : " & Trim(isi_Pesan)
'
'    Else
'    Exit Sub
' End If
'End Sub

Private Sub LstVisit_DblClick()
 If LstVisit.ListItems.Count > 0 Then
            
        
           With FRM_UpdateVisit
                .Text1.text = LstVisit.SelectedItem.SubItems(2)
                .Show vbModal
                

'                    M_DATA.UPDATE_NegoPTP M_OBJCONN, .TxtCustid.Text, .TDBDate1.Value, CStr(.TDBNumber1.Value), LstPayment.SelectedItem.SubItems(1)
'
'                    On Error GoTo add_error
'                    If M_DATA.ADD_OK Then
'                        'LstPayment.SelectedItem.SubItems(1) = ""
'                        LstPayment.SelectedItem.SubItems(2) = .TDBDate1.Value
'                        LstPayment.SelectedItem.SubItems(3) = .TDBNumber1.Value
'
'
'                    On Error GoTo 0
'                    End If
'                End If
               End With
Else
Exit Sub
End If

End Sub

Private Sub Option1_Click()
'@@ 03-05-2012, Dinonaktifkan
'If Option1.Value = True Then
'TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(AHome1.Value & txtHomeNo1.Value))
'   If txtHomeNo1.Value <> "" Then
'        txtPhoneA.Text = CStr(AHome1.Value & txtHomeNo1A.Value)
'    Else
'        txtPhoneA.Text = ""
'    End If
'   Option2.Value = False
'   Option3.Value = False
'   Option4.Value = False
'   Option5.Value = False
'End If
End Sub

Private Sub Option2_Click()
'@@ 03-05-2012 Dinonaktifkan
'If Option2.Value = True Then
'TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(AHome2.Value & txtHomeNo2.Value))
'   If txtHomeNo2.Value <> "" Then
'        txtPhoneA.Text = CStr(AHome2.Value & txtHomeNo2A.Value)
'    Else
'        txtPhoneA.Text = ""
'    End If
'   Option1.Value = False
'   Option3.Value = False
'   Option4.Value = False
'   Option5.Value = False
'End If
End Sub

Private Sub Option3_Click()
    '@@ 03-05-2012 DinonAktifkan
'   If Option3.Value = True Then
'   TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(AOffice2.Value & txtOfficeNo2.Value))
'   If txtOfficeNo2.Value <> "" Then
'        txtPhoneA.Text = CStr(AOffice2.Value & txtOfficeNo2A.Value)
'    Else
'        txtPhoneA.Text = ""
'   End If
'   Option2.Value = False
'   Option4.Value = False
'   Option1.Value = False
'   Option5.Value = False
'   End If
End Sub

Private Sub Option4_Click()
'@@DinonAktifkan 03-05-2012
'   If Option4.Value = True Then
'   TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(AOffice1.Value & txtOfficeNo1.Value))
'   If txtOfficeNo1.Value <> "" Then
'        txtPhoneA.Text = CStr(AOffice1.Value & txtOfficeNo1A.Value)
'    Else
'        txtPhoneA.Text = ""
'   End If
'   Option2.Value = False
'   Option3.Value = False
'   Option1.Value = False
'   Option5.Value = False
'End If
End Sub

Private Sub Option5_Click()
 If Option5.Value = True Then
 TYPETELP = ""
   txtPhone.text = GetNumber(CStr(txtMobileNo2.Value))
    If txtMobileNo2.Value <> "" Then
        txtPhoneA.text = CStr(txtMobileNo2A.Value)
    Else
        txtPhoneA.text = ""
   End If
   Option2.Value = False
   Option3.Value = False
   Option1.Value = False
   Option4.Value = False
   Option6.Value = False
   End If
End Sub

Private Sub Option6_Click()
 If Option6.Value = True Then
 TYPETELP = ""
   txtPhone.text = GetNumber(CStr(txtMobileNo1.text))
   If txtMobileNo1.text <> "" Then
        txtPhoneA.text = CStr(txtMobileNo1A.Value)
    Else
        txtPhoneA.text = ""
   End If
   Option2.Value = False
   Option3.Value = False
   Option1.Value = False
   Option4.Value = False
   Option5.Value = False
   End If
End Sub

Private Sub Option7_Click(Index As Integer)
Select Case Index
Case 0
TxtAddress.text = AddrNow.text
Case 1
TxtAddress.text = lblAddr.text
Case 2
TxtAddress.text = lblOfficeAddr.text
End Select

End Sub

Private Sub Option8_Click(Index As Integer)
Select Case Index
Case 0
Frame8.Enabled = True
VisitYES
Case 1
VisitNo
Frame8.Enabled = False
End Select
End Sub

Private Sub Option9_Click()
If Option9.Value = True Then
'@@ 11-03-2011 Di remarks, udah tidak diapakai
'LstSMS.Visible = True
'LstSMS2.Visible = False
End If
End Sub

Private Sub Option10_Click()
If Option10.Value = True Then
'@@ 11-03-2011 Di remarks, udah tidak diapakai
'LstSMS.Visible = False
'LstSMS2.Visible = True
End If

End Sub

Private Sub SSCommand1_Click(Index As Integer)
    Dim rsshut As New ADODB.Recordset
    Dim waktu_server_skrg As Date
    Dim lblagent_review As String
    Dim tian As String
    Dim xpincall As String
    'jejaktian30052016
    Dim query As String
    Dim rs, RcTian As ADODB.Recordset
    Dim n As Integer
    Dim loc As String
    
    Select Case Index
    Case 5
        'FRMSCRIPT.Show 1
        '@@ 09092011 Offering Discon digabung sama offering yang lama
        Call OfferingDiscGuide
    Case 0
        Unload frmaddphone
'        If (cboaccount.Text = "PTP-POP") Or (cboaccount.Text = "PTP-NEW") Or (cboaccount.Text = "PTP-PO") Or (cboaccount.Text = "PTP") Or (cboaccount.Text = "PTP-NE") Or (cboaccount.Text = "PTP-PAIDOFF") Then
'            query = "SELECT * from enabledptp"
'            Set RcTian = New ADODB.Recordset
'            RcTian.CursorLocation = adUseClient
'            RcTian.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If RcTian!Enabled = 1 Then
'                If Format(LstPayment.SelectedItem.ListSubItems(2), "yyyy-mm-dd") > Format(waktu_server_sekarang, "yyyy-mm-dd") Then
'                    MsgBox "Tidak Bisa di Call!! Karena belum masuk tanggal jatuh tempo", vbOKOnly + vbInformation, "Informasi"
'                Exit Sub
'                End If
'            End If
'        End If
        
        StsKategoriTelepon = ""
        KelompokKategoriTlp = ""
        
        Select Case CmbPhone
            '@@02-05-2011 Tambahan Telp Additional
            Case "TelpAdditional"
                txtPhone.text = Trim(TxtAdditional.Value)
                telpno = txtPhone.text
                '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                '@@02052012,Jika telepon additional pindahkan ke kotak additional yang baru
                'untuk memasukkan kategori telepon
                MsgBox "Sebelum anda melakukan call, harap pindahkan terlebih dahulu kategori teleponnya! Terima Kasih!", vbOKOnly + vbInformation, "Informasi"
                FrmReqTelepon.TxtCustid = Trim(lblCustId.text)
                FrmReqTelepon.TxtNoTelp.text = Trim(txtPhone.text)
                FrmReqTelepon.Show vbModal
                'Kosongkan telp_additional
                CMDSQL = "update mgm set telp_additional=null where custid='"
                CMDSQL = CMDSQL + CStr(lblCustId.text) + "'"
                M_OBJCONN.Execute CMDSQL
            Case "Office Num"
                txtPhone.text = Trim(txtMobileNo1.text)
                telpno = txtPhone.text
                '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                '@@11052012, Tambahan Kategori Telepon
                StsKategoriTelepon = "OFFICE"
            Case "Hp2"
                txtPhone.text = txtMobileNo2.Value
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = "HP"
            Case "Old Num"
                txtPhone.text = Trim(txtHomeNo1.text)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = "Old Num"
            Case "HomePhone2"
                txtPhone.text = Trim(txtHomeNo2.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = "Home"
            Case "New Num"
                txtPhone.text = Trim(txtOfficeNo1.text)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = "New Num"
            Case "OfficePhone2"
                txtPhone.text = Trim(txtOfficeNo2.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = "Office"
            Case "EC Num"
'                If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
'                    MsgBox "Simpan Data terlebih dahulu"
'                    Exit Sub
'                End If
                If Left(Text11.text, 3) = "021" Then
                 txtPhone.text = Trim(Mid(Text11.text, 4, 16))
                 Else
                 txtPhone.text = Trim(Text11.text)
                End If
                txtPhone.text = Text11.text
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = "EC"
            Case "AddHome1"
                txtPhone.text = Trim(txtadd_phone(0).text)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = Trim(CmbStsKatHome1.text)
            Case "AddHome2"
                txtPhone.text = Trim(txtHomeAdd2.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = Trim(CmbStsKatHome2.text)
            Case "AddOffice1"
                txtPhone.text = Trim(txtadd_phone(1).text)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = Trim(CmbStsKatOffice1.text)
            Case "AddOffice2"
                txtPhone.text = Trim(txtOfficeAdd2.Value)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = Trim(CmbStsKatOffice2.text)
            Case "AddMobile1"
                txtPhone.text = Trim(txtadd_phone(2).text)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = Trim(CmbStsKatHP1.text)
            Case "AddOtherphone"
                txtPhone.text = Trim(txtadd_phone(3).text)
                telpno = txtPhone.text
                 '@@ 16 Agustus 2011, Tambahan buat remarks, agent nelpon ke mana
                TxtTelpKe.text = Trim(CmbPhone.text)
                StsKategoriTelepon = Trim(CmbStsKatHP2.text)
            Case Else
            
''''                If CmbPhone.text Like "*ICO*" Then
''''                    txtPhone.text = Txtperiod.Caption
''''                    telpno = txtPhone.text
''''                    TxtTelpKe.text = Trim(CmbPhone.text)
''''                    StsKategoriTelepon = "Alternatif ICOLL"
''''                Else
                
                    '@@ 11 Juni 2012, Revisi Tambahan Telepon
                     
                     If FrmCC_Colection.Frame3.Caption = "0" Then
                        txtPhone.text = Replace(CmbPhone.text, " ", "")
                        txtPhone.text = Replace(CmbPhone.text, "'", "")
                        TxtTelpKe.text = Trim(CmbPhone.text)
                     Else
                        telpno = txtPhone.text
                        TxtTelpKe.text = Trim(CmbPhone.text)
                     End If
                     
                     
                     'telpno = txtPhone.text
                     
                    'StsKategoriTelepon = Trim(CmbStsKatHP2.text)
                     
'                     CMDSQL = "select * from tbllayanantelkom where nolayanan='"
'                     CMDSQL = CMDSQL + Trim(txtPhone.text) + "'"
'                     Set M_Objrs_Cek = New ADODB.Recordset
'                     M_Objrs_Cek.CursorLocation = adUseClient
'                     M_Objrs_Cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'                     If M_Objrs_Cek.RecordCount > 0 Then
'                        TxtTelpKe.text = CmbPhone.text
'
'                     Else
'                        If TxtNoTelpReq.Value <> Empty Then
'                            txtPhone.text = Trim(TxtNoTelpReq.Value)
'                            TxtTelpKe.text = Trim(CmbPhone.text)
'                            KelompokKategoriTlp = TxtKategori.Caption
'                            StsKategoriTelepon = TxtTelpKe.text
'                        Else
'                           Set M_Objrs_Cek = Nothing
'                           MsgBox "Maaf, anda tidak dapat menelepon nomor yang tidak terdapat dalam database!", vbOKOnly + vbCritical, "Peringatan"
'                           Exit Sub
'                        End If
'                     End If
'
'''''              End If
               Set M_Objrs_Cek = Nothing
        End Select
        
        txtPhone.text = Replace(txtPhone.text, "'", "")
        
        '@@31-05-2012 Jika Status Account=PO dan CO maka tidak dapat di call
        If StatusAccount = "PO-" Or StatusAccount = "CO-" Then
            MsgBox "Mohon maaf! Status Account PAID OFF atau COMPLAIN tidak dapat di call!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If

        kat_aktif_telp = " [ " & CmbPhone.text & " ] "
        
        'Cek no telepon yang apakah masuk daftar blacklist. Jika masuk maka keluar sub!
        CMDSQL = "select no_telp from tblblacklist where no_telp='"
        CMDSQL = CMDSQL + Replace(Trim(txtPhone.text), " ", "") + "'"
    
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_objrs.RecordCount <> 0 Then
            MsgBox "No.Telepon yang anda hubungi masuk dalam daftar blacklist!. Silahkan hubungi TL  anda!.", vbOKOnly + vbExclamation, "Peringatan"
            Exit Sub
        End If
        Set M_objrs = Nothing
    
        CMDSQL = "select no_telp from tblunvalid_number where no_telp='"
        CMDSQL = CMDSQL + Replace(Trim(txtPhone.text), " ", "") + "' "
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_objrs.RecordCount <> 0 Then
            MsgBox "No.Telepon yang anda hubungi masuk dalam daftar Unvalid number!. Silahkan hubungi TL  anda!.", vbOKOnly + vbExclamation, "Peringatan"
            Exit Sub
        End If
        
        ' ----------- CEK WIT OR WITA 05 FEB 2014 -----------
        If M_objrs.State = 1 Then M_objrs.Close
        M_objrs.Open "SELECT now() as wkt_server"
        If M_objrs.RecordCount > 0 Then
            waktu_server_skrg = M_objrs!wkt_server
            lbltime_save.Caption = Format(M_objrs!wkt_server, "yyyy-mm-dd hh:mm:ss")
        End If
        
        If M_objrs.State = 1 Then M_objrs.Close
        M_objrs.Open "SELECT * FROM tbl_timezone WHERE trim(kode)='" & Left(Replace(Trim(txtHomeNo1A.text), " ", ""), 4) & "'"
        If M_objrs.RecordCount > 0 Then
            If Format(waktu_server_skrg, "hh:mm") >= Format(M_objrs!time_limit, "hh:mm") Then
                MsgBox "Maaf anda tidak diperkenankan Telp pada Pukul atau melebihi " & M_objrs!time_limit & " Pada area " & M_objrs!group_time, vbCritical + vbOKOnly, "INFO"
                Exit Sub
            End If
        End If
    ' ---------------------------------------------------
    Set M_objrs = Nothing
    
    ' 19-04-2013 untuk 5x Blok -------------------------
    sPhone_Agent = Trim(MDIForm1.TxtUsername.text)
    sPhone_CustID = CStr(lblCustId.text)
    sPhone_TelpNo = Replace(Trim(txtPhone.text), " ", "")
    ' ---------------------------------------------------
    
    '@@ 18-04-2012, Cek setiap agent yang menelepon
    'ke nomor yang sama nomor teleponnya tidak bisa dihubungi lagi
    Dim M_Objrs_Cek_Panggilan As ADODB.Recordset
    
    'CEK_SEGMENT_CALL
    If vrcek <> "OS-" Then
        If Label14(0).Caption <> "" Then
            'VHP
            If Label14(0).Caption = "VHP" Then
                If FuncCekSegmen(GetNumber(CStr(Replace(txtPhone.text, " ", "")))) = 6 Then
                    MsgBox "Nomor Tersebut Tidak Bisa Di Call Lebih Dari 6 Kali!", vbOKOnly + vbInformation, "Informasi"
                    Exit Sub
                End If
            ElseIf Label14(0).Caption = "HP" Then
                If FuncCekSegmen(GetNumber(CStr(Replace(txtPhone.text, " ", "")))) = 6 Then
                    MsgBox "Nomor Tersebut Tidak Bisa Di Call Lebih Dari 6 Kali!", vbOKOnly + vbInformation, "Informasi"
                    Exit Sub
                End If
            ElseIf Label14(0).Caption = "MP" Then
                If FuncCekSegmen(GetNumber(CStr(Replace(txtPhone.text, " ", "")))) = 4 Then
                    MsgBox "Nomor Tersebut Tidak Bisa Di Call Lebih Dari 4 Kali!", vbOKOnly + vbInformation, "Informasi"
                    Exit Sub
                End If
            ElseIf Label14(0).Caption = "LP" Then
                If FuncCekSegmen(GetNumber(CStr(Replace(txtPhone.text, " ", "")))) = 4 Then
                    MsgBox "Nomor Tersebut Tidak Bisa Di Call Lebih Dari 4 Kali!", vbOKOnly + vbInformation, "Informasi"
                    Exit Sub
                End If
            ElseIf Label14(0).Caption = "VLP" Then
                If FuncCekSegmen(GetNumber(CStr(Replace(txtPhone.text, " ", "")))) = 2 Then
                    MsgBox "Nomor Tersebut Tidak Bisa Di Call Lebih Dari 2 Kali!", vbOKOnly + vbInformation, "Informasi"
                    Exit Sub
                End If
            End If
        End If
    End If
    
    Set M_Objrs_Cek_Panggilan = Nothing
   
    CMDSQL = "Insert Into tblphonemonitorhst(UserId, CustId, NamaCh,StartDate, TelpNo, Recsource,status_telp,tgl) Values "
    CMDSQL = CMDSQL + " ('" + MDIForm1.TxtUsername.text + "' , '" + FrmCC_Colection.lblCustId.text + "','"
    CMDSQL = CMDSQL + Replace(FrmCC_Colection.lblNama.text, "'", "") + "', '"
    CMDSQL = CMDSQL + Format(CStr(MDIForm1.TDBDate1.Value), "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
    CMDSQL = CMDSQL + "' , '" + Replace(txtPhone.text, " ", "") + "' ,'"
    CMDSQL = CMDSQL + FrmCC_Colection.lblRecsource.Caption + "','"
    CMDSQL = CMDSQL + IIf(IsNull(TxtKategori.Caption), "", TxtKategori.Caption) + "',now())"
    M_OBJCONN.Execute CMDSQL
    
    'JEJAKTIAN08032016
    CMDSQL = "insert into tblrrd(custid,agent,phone,start_time,sstatus_awal) values"
    CMDSQL = CMDSQL + "('" + FrmCC_Colection.lblCustId.text + "','" + MDIForm1.TxtUsername.text + "','"
    CMDSQL = CMDSQL + txtPhone.text + "', '" & waktu_server_sekarang & "','" + cboaccount.text + "')"
    M_OBJCONN.Execute CMDSQL
    
    getservertime.text = waktu_server_sekarang
    SSCommand1(2).Enabled = False
    '=====================================================
    
    CMDSQL = "UPDATE mgm set waktu_mulai_call = '" & waktu_server_sekarang & "' WHERE custid = '" + lblCustId.text + "' "
    M_OBJCONN.Execute CMDSQL
    
    ' RESET UNIQUE ID
    MDIForm1.txt_unique_id.text = ""
    
    '@@19042012 Tombol Exit,Tombol Call di Nonaktifkan dulu
    SSCommand1(3).Enabled = False
    '@@19042012 Tombol Hangup Diaktifkan
    SSCommand1(1).Enabled = True
    '@@19042012 Tombol Call Dinonaktifkan
    SSCommand1(0).Enabled = False
    
    '@@25-05-2012 Tombol Save dinonaktifkan
    '@@17122012 Tombol Save Diaktifkan
    'SSCommand1(2).Enabled = False
    'jejaktian23032016 true jadi false
    SSCommand1(2).Enabled = False
    
    '@@ Filter tanda baca ditelepon
    txtPhone.text = Replace(txtPhone.text, "/", "")
    txtPhone.text = Replace(txtPhone.text, "\", "")
    txtPhone.text = Replace(txtPhone.text, "'", "")
    txtPhone.text = Replace(txtPhone.text, ";", "")
    txtPhone.text = Replace(txtPhone.text, ":", "")
    txtPhone.text = Replace(txtPhone.text, "|", "")
    txtPhone.text = Replace(txtPhone.text, ".", "")
    txtPhone.text = Replace(txtPhone.text, ",", "")
    txtPhone.text = Replace(txtPhone.text, "?", "")
    txtPhone.text = Replace(txtPhone.text, "!", "")
    txtPhone.text = Replace(txtPhone.text, " ", "")
    txtPhone.text = Replace(txtPhone.text, "-", "")
    txtPhone.text = Replace(txtPhone.text, "(", "")
    txtPhone.text = Replace(txtPhone.text, ")", "")
    
    If Left(txtPhone.text, 2) = "62" Then
        txtPhone.text = "0" & Right(txtPhone.text, Len(txtPhone.text) - 2)
    End If
    
    sudahCall = True
    xpincall = ""
    
    qs = "select * from tbl_list_client_indium order by 1"
    Set M_objrsc = New ADODB.Recordset
    M_objrsc.CursorLocation = adUseClient
    M_objrsc.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Folder = ""
    For i = 1 To M_objrsc.RecordCount
        If UCase(lblRecsource.Caption) Like "*" & M_objrsc!client & "*" Then
            Folder = M_objrsc!client
            GoTo bawah:
        End If
        M_objrsc.MoveNext
    Next i
    If Folder = "" Then
        Folder = "LAIN"
    End If
        
    loc = Format(FungsiWaktuServer, "yyyy-mm-dd")
    
    Folder = loc
    
'    If UCase(lblRecsource.Caption) Like "*HCI*" Then
'        Folder = "HCI"
'    End If
'    If UCase(lblRecsource.Caption) Like "*BCA*" Then
'        Folder = "BCA"
'    End If
'    If UCase(lblRecsource.Caption) Like "*MANDIRI*" Then
'        Folder = "MANDIRI"
'    End If
'    If UCase(lblRecsource.Caption) Like "*BRI*" Then
'        Folder = "BRI"
'    End If
'    If UCase(lblRecsource.Caption) Like "*MAYBANK*" Then
'        Folder = "MAYBANK"
'    End If
'    If UCase(lblRecsource.Caption) Like "*PANIN*" Then
'        Folder = "PANIN"
'    End If
'    If UCase(lblRecsource.Caption) Like "*PLUS*" Then
'        Folder = "RPPLUS"
'    End If
'    If UCase(lblRecsource.Caption) Like "*UANG*" Then
'        Folder = "UANGEXP"
'    End If
'    If UCase(lblRecsource.Caption) Like "*GLOBAL*" Then
'        Folder = "GLOBALINDO"
'    End If
'    If UCase(lblRecsource.Caption) Like "*COURT*" Then
'        Folder = "COURT"
'    End If
'    If UCase(lblRecsource.Caption) Like "*DANAMON*" Then
'        Folder = "DANAMON"
'    End If
bawah:
    
    'If Obelisk = False Then
        'UNTUK ORANGE CLIENT
    '    MDIForm1.ActionCTI ("DIAL|" & xpincall & GetNumber(CStr(Replace(txtPhone.text, " ", ""))) & "|" & Trim(FrmCC_Colection.lblCustId.text) & "|" & Trim(FrmCC_Colection.lblCustId.text)) & "-" & MDIForm1.TxtUsername.text
    'Else
        'UNTUK OBELISK
        'MDIForm1.ActionCTI ("DIAL|" & GetNumber(CStr(Replace(txtPhone.Text, " ", ""))) & "|" & Trim(FrmCC_Colection.lblCustId.text) & "|" & Trim(FrmCC_Colection.lblCustId.text)) & "-" & mdiform1.txtusername.text

        MDIForm1.ActionCTI ("DIAL|" & xpincall & GetNumber(CStr(Replace(txtPhone.text, " ", ""))) & "|" & Trim(FrmCC_Colection.lblCustId.text) & "|" & Trim(Folder))
        WaitSecs (0.5)
        lg_call = True
        'Call insertlogcti(MDIForm1.TxtStatus.Text, GetNumber(CStr(Replace(txtPhone.Text, " ", ""))))
    'End If
        
    M_OBJCONN.Execute " INSERT INTO user_phone_log(agent,custid,no_telp) " & _
                      " values('" & MDIForm1.TxtUsername.text & "','" & Trim(FrmCC_Colection.lblCustId.text) & "','" & GetNumber(CStr(Replace(txtPhone.text, " ", ""))) & "')"
    
    'INSERT KE TABLE FLAG TIPE - TIPE / SEGMENT CH (VHP, VH, MP, LP, VLP)
'    If Label14(0).Caption <> "" Then
'        CMDSQL = "SELECT nolayanan FROM tbllayanantelkom WHERE nolayanan = '" & GetNumber(CStr(Replace(txtPhone.Text, " ", ""))) & "'"
'        Set M_objrs = New ADODB.Recordset
'        M_objrs.CursorLocation = adUseClient
'        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'        If M_objrs.RecordCount = 0 Then
'            Call INSERT_TEMP_SEGMENT_CALL
'        End If
'    End If
    
    '---------- NEW LOGIC ACCOUNT REVIEW ------------------
    'INSERT KE TABLE REVIEW YANG BARU (RANDY 10FEB 2016)
    CMDSQL = "SELECT nolayanan FROM tbllayanantelkom WHERE nolayanan = '" & GetNumber(CStr(Replace(txtPhone.text, " ", ""))) & "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount = 0 Then
        Call INSERT_TEMP_TELFON_REVIEW
    End If
    '---------- NEW LOGIC ACCOUNT REVIEW ------------------
    
    '@@ 25-07-2011 Dipindah, jadi di form load
    'Call OfferingDiscGuide
    
    '@@15092012Catat AktifitasCall
    AktifitasCall = "1"
    calling = "1"
    stshangup.text = 0
    'MDIForm1.CmbNo.Text = ""
    stscall = True
    TYPETELP = ""
   Case 2
        Unload frmaddphone
        If MDIForm1.txtlevel.text = "Agent" Then
            If cboaccount.text = "PTP" Then
                If Text10.text <> "1" Then
                    MsgBox "Maaf, anda belum mengisi CPA!", vbOKOnly + vbInformation, "Informasi"
                    Exit Sub
                End If
            End If
        End If
        'untuk 3 kali invalid pada hari yang berbeda akan jadi data retur TIAN (21Dec2016)
'        Call tblretur
        '================================================================
        V_SAVE = CEK_DATA_VALID
        

        
        If V_SAVE = False Then
            Exit Sub
        Else
        End If
        If ADD_CUST Then
        Else
            'Tambahan Randy 11-05-2015 (Untuk mencatat status call sebelum ngesave status call terakhir)
            Call UPDATE_STATUS_CALL_SEBELUM
            'Call UPDATE_MGM_HST_SAVE
            Call CEK_UPDATE_PELANGGAN
            'untuk jadikan data retur apa bila status call Data retur TIAN (21Dec2016)
            Call insertretur
            '=========================================================================
            stscall = False
            Call isi_datapayment
        End If
        AktifitasCall = ""
        Call load_reminder
        Call autoremarks
   Case 3
        Unload frmaddphone
        If bRenderrecord = True Then
          '  VIEW_MGMDATA.renderdonk
        End If
        bRenderrecord = False
        kontak = False
        For n = 1 To LstPayment.ListItems.Count
            If LstPayment.ListItems(n).SubItems(4) = "UNSCH" And regnego = True Then
                regnego = True
            End If
        Next n
        If regnego = True And LstPayment.ListItems.Count <> 0 Then
            MsgBox "Lakukan PTP yang benar, Jumlah PTP harus >= Deal Payment " & txtPayment.text & " ,Atau data simpan dulu!!!"
            Exit Sub
        End If
        
        
'        If Calling = "1" Then
'            '@@15092012 Hitung Durasi Call Dari Icentra
'            Call HitungDurasiDariIcentra
'        End If
        '@@15092012 Cek Aktifitas Call Apakah Agent Telah Melakukan Call?
        'Jika sudah, Agent Harus Melakukan Remarks
        If lblRecsource.Caption <> "Satukosonglapan" Then
    
            If AktifitasCall = "1" Then
                ' 01 JULI 2014 SAVE AFTER CALL
                'If Len(Trim(txtremarks.Text)) = 0 Then
                    MsgBox "Maaf, anda belum menulis remarks! Harap tulis remarks terlebih dahulu!", vbOKOnly + vbInformation, "Informasi"
                    Exit Sub
                'End If
            End If
            
        End If
                
        '@@ Akhir 061110 cek lock account sesuai settingan timer
        Dim M_Objrs_Close As ADODB.Recordset
        CMDSQL = "select sts_close from usertbl where userid='"
        CMDSQL = CMDSQL + CStr(MDIForm1.TxtUsername.text) + "' and sts_close='1'"
        Set M_Objrs_Close = New ADODB.Recordset
        M_Objrs_Close.CursorLocation = adUseClient
        M_Objrs_Close.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Close.RecordCount > 0 Then
            MsgBox "Mohon maaf, ada perubahan system. Aplikasi TINS akan di tutup! Harap Login Ulang!", vbOKOnly + vbInformation, "Informasi"
            Set M_Objrs_Close = Nothing
            CMDSQL = "update usertbl set sts_close=null where userid='"
            CMDSQL = CMDSQL + CStr(MDIForm1.TxtUsername.text) + "' "
            M_OBJCONN.Execute CMDSQL
            End
        End If
        
        ' Matikan monitoring activity
        i_monitoring_activity = 0
        'MDIForm1.Timer2.Enabled = False
        main_timer_activity = 0
        'MDIForm1.Timer7.Enabled = True
        ' #####
        'NGAMBIL WAKTU LOGIN UNTUK BLOCK
        waktu_start = waktu_server_sekarang
            
        Set M_Objrs_Close = Nothing
        
        MDIForm1.Timer100.Enabled = True
        Unload Me
        Exit Sub
'KeluarLockAuto:
        'Unload Me
    Case 1
        Unload frmaddphone
        DoEvents
        sChannel = MDIForm1.txtChannel.text
        MDIForm1.ActionCTI ("HANGUP|" + sChannel)
        stshangup.text = 1
        'MDIForm1.ActionCTI ("HANGUP")
        
        'Call insertlogcti(MDIForm1.TxtStatus.Text, GetNumber(CStr(Replace(txtPhone.Text, " ", ""))))
        '@@ 18 April 2012, Catat ketika agent mengakhiri telepon
        Call hangup_event
    Case 4
        StatusCPA = "CPA Form 1"
        frmcpanew.Show 1
        
End Select
Exit Sub
'ke:
strsql = "update usertbl set stsaplikasi=0  where userid ='" + MDIForm1.TxtUsername.text + "'"
M_OBJCONN.Execute (strsql)
MsgBox err.Description
 Exit Sub
 
End Sub

Public Sub hangup_event()
    If addphone = True Then
        Unload frmaddphone
    End If
    
    SSCommand1(1).Enabled = False
    
    WaitSecs (0.5)
    
    CMDSQL = "update tblphonemonitorhst set enddate=now() from "
    CMDSQL = CMDSQL + " (select id as idnew from "
    CMDSQL = CMDSQL + " tblphonemonitorhst where custid='"
    CMDSQL = CMDSQL + Trim(lblCustId.text) + "' and userid='"
    CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "' order by id desc limit 1) as a "
    CMDSQL = CMDSQL + " where tblphonemonitorhst.id=idnew"
    DoEvents
    M_OBJCONN.Execute CMDSQL
        
    Call HitungDurasiCall
    DoEvents
    
    '@@15092012 Hitung Durasi Call Icentra Dicari dari tombol exit saja
    'Call HitungDurasiDariIcentra
    
    '@@19042012 Tombol Exit,diaktifkan
    SSCommand1(3).Enabled = True
    '@@19042012 Tombol Hangup Dinonaktifkan
    SSCommand1(1).Enabled = False
    '@@19042012 Tombol Call Diaktifkan
    SSCommand1(0).Enabled = False
    '@@25-05-2012 Tombol Save Diaktifkan
    SSCommand1(2).Enabled = True
    txtremarks.SetFocus
    
    ' Berhenti di kasih waktu
    lblstop_time.Caption = waktu_server_sekarang
    
    'Call SimpanRemarksCall
    'JEJAKTIAN08032016
    Call updaterrd
    'Update Randy Req : 10Agustus2015
    Call SimpanTempCall
    ' Reset monitoring activity
    ' i_monitoring_activity = 0
    'MDIForm1.Timer2.Enabled = True
    ' #####
    
    '@@08102012, Buat Hangup Xlite
    On Error Resume Next
    Dim iret As Long
    THandle = FindWindow(vbEmpty, "X-Lite")
    If THandle = 0 Then
        MsgBox "Maaf, X-Lite  tidak ditemukan!"
        If lblRecsource.Caption <> "Satukosonglapan" Then
            MsgBox "simpan terlebih dahulu sebelum melakukan call lagi"
        
        End If
        Exit Sub
    End If
    iret = BringWindowToTop(THandle)
    Sendkeys "^h", 0.7
    WaitSecs 0.2
    Sendkeys "^h", 0.7
    If lblRecsource.Caption <> "Satukosonglapan" Then
        MsgBox "simpan terlebih dahulu sebelum melakukan call lagi"
    End If
    
    lg_call = False
    
End Sub

Private Sub tblretur()
    If cboaccount.text = "Invalid" Then
        query = "select * from tblretur where custid = '" + lblCustId.text + "' and date(tgl) = date(now())"
        Set rs_retur = New ADODB.Recordset
        rs_retur.CursorLocation = adUseClient
        rs_retur.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        query = "select * from tblretur where custid = '" + lblCustId.text + "'"
        Set rs_blok = New ADODB.Recordset
        rs_blok.CursorLocation = adUseClient
        rs_blok.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        If rs_retur.RecordCount = 0 Then
            query = "INSERT INTO tblretur VALUES ('" + lblCustId.text + "',now())"
            M_OBJCONN.Execute query
        End If
        
        If rs_blok.RecordCount = 3 Then
            query = "UPDATE mgm set retur = 1 where custid = '" + lblCustId.text + "'"
            M_OBJCONN.Execute query
        End If
    End If
    If MDIForm1.txtlevel.text <> "Agent" Then
        If cboaccount.text <> "Data Retur" Then
            query = "UPDATE mgm set retur = null where custid = '" + lblCustId.text + "'"
            M_OBJCONN.Execute query
        End If
    End If
End Sub

Private Sub insertretur()
    If cboaccount.text = "Data Retur" Then
        query = "UPDATE mgm SET retur = 1 where custid = '" + lblCustId.text + "'"
        M_OBJCONN.Execute query
    End If
End Sub

Private Sub SimpanTempCall()
    Dim sQuery As String
    Dim iQuery As String
    Dim uQuery As String
    Dim RS_Temp_Call As New ADODB.Recordset
    Dim jumlah_sekarang As Double
    Dim jumlah_baru As Double
    
    'CEK TANGGAL CALL
    sQuery = "SELECT custid, tgl_call FROM tbl_temp_jumlah_call WHERE custid = '" + lblCustId.text + "' "
    sQuery = sQuery + "AND date(tgl_call) = '" & Format(lblstop_time.Caption, "yyyy-mm-dd") & "'"
    Set RS_Temp_Call = New ADODB.Recordset
    RS_Temp_Call.CursorLocation = adUseClient
    RS_Temp_Call.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If RS_Temp_Call.RecordCount = 0 Then
        iQuery = "INSERT INTO tbl_temp_jumlah_call(custid, tgl_call, jumlah)"
        iQuery = iQuery + " VALUES('" + lblCustId.text + "', '" & Format(lblstop_time.Caption, "yyyy-mm-dd") & "', '1')"
        
        M_OBJCONN.Execute iQuery
    Else
        Set RS_Temp_Call = Nothing
        sQuery = "SELECT jumlah FROM tbl_temp_jumlah_call WHERE custid = '" + lblCustId.text + "' "
        sQuery = sQuery + "AND date(tgl_call) = '" & Format(lblstop_time.Caption, "yyyy-mm-dd") & "'"
        Set RS_Temp_Call = New ADODB.Recordset
        RS_Temp_Call.CursorLocation = adUseClient
        RS_Temp_Call.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        jumlah_sekarang = Trim(RS_Temp_Call!jumlah)
        
        jumlah_baru = jumlah_sekarang + 1
        
        uQuery = " UPDATE tbl_temp_jumlah_call set jumlah = '" & jumlah_baru & "' "
        uQuery = uQuery + " WHERE custid = '" + lblCustId.text + "' AND date(tgl_call) = '" & Format(lblstop_time.Caption, "yyyy-mm-dd") & "' "
        
        M_OBJCONN.Execute uQuery
    End If
End Sub

Public Sub Show_NEGOPTP()
Dim showlist As New ADODB.Recordset
Dim ListItem As ListItem
Dim CMDSQL As String
Dim TOTPTP As Currency
Dim ssql As String
ssql = "SELECT CUSTID,sum(PAYMENT) as Jum FROM tbllunas WHERE custid = '" + lblCustId.text + "' GROUP BY CUSTID"
showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If showlist.BOF And showlist.EOF Then
    TOTPTP = 0
Else
    TOTPTP = IIf(IsNull(showlist!jum), 0, showlist!jum)
End If
'If ShowList.BOF And ShowList.EOF Then
'    'CMDSQL = "SELECT * FROM TBLNEGOPTP WHERE custid = '" + lblCustId.text + "'"
'    'AND CUSTID NOT IN (SELECT CUSTID FROM tbllunas)"
'    CMDSQL = "SELECT DISTINCT TBLNEGOPTP.PROMISEDATE,TBLNEGOPTP.ID,TBLNEGOPTP.PROMISEPAY,TBLNEGOPTP.TYPE FROM TBLNEGOPTP,tbllunas WHERE "
'    CMDSQL = CMDSQL + "tbllunas.CUSTID<>TBLNEGOPTP.CUSTID AND TBLNEGOPTP.CUSTID='" + lblCustId.text + "' order by TBLNEGOPTP.promisedate desc"
'Else
'    CMDSQL = "SELECT distinct TBLNEGOPTP.PROMISEDATE,TBLNEGOPTP.PROMISEPAY,TBLNEGOPTP.ID,TBLNEGOPTP.TYPE "
'    CMDSQL = CMDSQL + "FROM VWLISTPTP,TBLNEGOPTP WHERE TBLNEGOPTP.CUSTID=VWLISTPTP.CUSTID AND "
'    CMDSQL = CMDSQL + "VWLISTPTP.PAYDATE<TBLNEGOPTP.PROMISEDATE AND TBLNEGOPTP.CUSTID='" + lblCustId.text + "' order by TBLNEGOPTP.promisedate desc"
'End If

CMDSQL = "SELECT * FROM tblnegoptp where custid = '" + lblCustId.text + "' "
'@@ 08-02-2012 , Tambahan untuk filter tabel negoptp
'@@ 26-03-2012 Filter Bulan dan Tahun dinonaktifkan dulu
'CMDSQL = CMDSQL + " and date_part('month',promisedate)>=date_part('month',now()) and "
'CMDSQL = CMDSQL + " date_part('year',promisedate)>=date_part('year',now()) "
CMDSQL = CMDSQL + " order by promisedate desc"

Set showlist = New ADODB.Recordset
showlist.CursorLocation = adUseClient
showlist.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

LstPayment.ListItems.CLEAR
Dim n As Currency
While Not showlist.EOF
    Set ListItem = LstPayment.ListItems.ADD(, , "")
        ListItem.SubItems(1) = CStr(IIf(IsNull(showlist!ID), "", (showlist!ID)))
        ListItem.SubItems(2) = CStr(IIf(IsNull(showlist!PromiseDate), "", Format(showlist!PromiseDate, "yyyy-mm-dd")))
        ListItem.SubItems(3) = CStr(IIf(IsNull(showlist!PromisePay), "", Round((showlist!PromisePay), 0)))
        n = n + Val(ListItem.SubItems(3))
        If n <= TOTPTP Then
            ListItem.ListSubItems(1).ForeColor = vbRed
            ListItem.ListSubItems(2).ForeColor = vbRed
            ListItem.ListSubItems(3).ForeColor = vbRed
        End If
        
        ListItem.SubItems(4) = IIf(IsNull(showlist!Type), "", showlist!Type)
        ListItem.SubItems(5) = CStr(IIf(IsNull(showlist!inputdate), "", Format(showlist!inputdate, "yyyy-mm-dd")))
     showlist.MoveNext
Wend

Set showlist = Nothing
End Sub

Public Sub show_cust()
    Dim ListItem As ListItem
    Dim m_data As New CLS_FRMCUST_CC
    Dim m_cust1 As ADODB.Recordset
    Dim m_cust2 As ADODB.Recordset
    Dim CMDSQL As String
    Dim CMDSQL2 As String
    Dim sPending As String
    Dim CEKREC As New ADODB.Recordset
    Dim sTime_Hst As String
    Dim tgl_exp As Date

    'On Error GoTo HELL:
    'CMDSQL = "SELECT mgm.*, mgm_DETAIL.* FROM mgm INNER JOIN "
    'CMDSQL = CMDSQL + "mgm_DETAIL ON mgm.CUSTID = dbo.mgm_DETAIL.CUSTID"
    
    CMDSQL = "select * from mgm"
    'CMDSQL2 = "select * from mgm_detail"


    Set m_cust = New ADODB.Recordset
    'Set m_cust2 = New ADODB.Recordset
    m_cust.CursorLocation = adUseClient
    'm_cust2.CursorLocation = adUseClient
    If shedulePTP_Show = True Then
        CMDSQL = CMDSQL + " where custid ='" & MDIForm1.LstGrade.SelectedItem.SubItems(1) & "'"
        m_cust.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ' Tambahan untuk reminder AGENT 12 Mei 2013 By Izuddin
    ElseIf bReminder_agent = True Or bAktif_Cust_Review = True Then
        CMDSQL = CMDSQL + " where custid ='" & sReminder_CUST_ID & "'"
        m_cust.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ' +++++++++++++++++++++++++++++++++++++++++++++++++++++
    Else
        CMDSQL = CMDSQL + " where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
        m_cust.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
        'CMDSQL2 = CMDSQL2 + " where custid ='" & VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(1) & "'"
        'm_cust2.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic
        'm_cust.Open "Select * from mgm where custid='" & VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(1) & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
    End If

'tampilkan data tabel mgm
If Not m_cust.EOF Then
    
    On Error Resume Next
    
    'MENGISI TGL EXPIRED PADA FORM CC
    
    tgl_exp = IIf(IsNull(m_cust("tgl_exp_claim")), "", m_cust("tgl_exp_claim"))
    
    If tgl_exp < Now() Then
        'jejaktian11042016
        'lbl_expdate.Caption = ""
        lbl_expdate.Caption = IIf(IsNull(m_cust("tgl_exp_claim")), "", Format(m_cust("tgl_exp_claim"), "dd-mm-yyyy"))
    Else
        lbl_expdate.Caption = IIf(IsNull(m_cust("tgl_exp_claim")), "", Format(m_cust("tgl_exp_claim"), "dd-mm-yyyy"))
    End If
     
    '@@31052012 Buat Menyimpan Status Account
    StatusAccount = ""
    StatusAccount = IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new"))
    lblgender.Caption = IIf(IsNull(m_cust("sex")), "", m_cust("sex"))
    
    '@@ 07-05-2012, Buat menandakan bahwa nomor tersebut UnValid Number
    If m_cust("f_unvalid_home1") = "1" Then
        txtHomeNo1A.BackColor = &HC0C0&
        txtHomeNo1.BackColor = &HC0C0&
        txtHomeNo1.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_home1")), "(Null)", m_cust("f_sts_unvalid_home1"))
        txtHomeNo1A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_home1")), "(Null)", m_cust("f_sts_unvalid_home1"))
    End If
    If m_cust("f_unvalid_home2") = "1" Then
        txtHomeNo2A.BackColor = &HC0C0&
        txtHomeNo2.BackColor = &HC0C0&
        txtHomeNo2.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_home2")), "(Null)", m_cust("f_sts_unvalid_home2"))
        txtHomeNo2A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_home2")), "(Null)", m_cust("f_sts_unvalid_home2"))
    End If
    If m_cust("f_unvalid_office1") = "1" Then
        txtOfficeNo1A.BackColor = &HC0C0&
        txtOfficeNo1.BackColor = &HC0C0&
        txtOfficeNo1.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_office1")), "(Null)", m_cust("f_sts_unvalid_office1"))
        txtOfficeNo1A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_office1")), "(Null)", m_cust("f_sts_unvalid_office1"))
    End If
    If m_cust("f_unvalid_office2") = "1" Then
        txtOfficeNo2A.BackColor = &HC0C0&
        txtOfficeNo2.BackColor = &HC0C0&
        txtOfficeNo2.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_office2")), "(Null)", m_cust("f_sts_unvalid_office2"))
        txtOfficeNo2A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_office2")), "(Null)", m_cust("f_sts_unvalid_office2"))
    End If
    If m_cust("f_unvalid_mobile1") = "1" Then
        txtMobileNo1A.BackColor = &HC0C0&
        txtMobileNo1.BackColor = &HC0C0&
        txtMobileNo1A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_mobile1")), "(Null)", m_cust("f_sts_unvalid_mobile1"))
        txtMobileNo1.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_mobile1")), "(Null)", m_cust("f_sts_unvalid_mobile1"))
    End If
    If m_cust("f_unvalid_mobile2") = "1" Then
        txtMobileNo2A.BackColor = &HC0C0&
        txtMobileNo2.BackColor = &HC0C0&
        txtMobileNo2A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_mobile2")), "(Null)", m_cust("f_sts_unvalid_mobile2"))
        txtMobileNo2.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_mobile2")), "(Null)", m_cust("f_sts_unvalid_mobile2"))
    End If
    If m_cust("f_unvalid_addhome1") = "1" Then
        txtHomeAdd1.BackColor = &HC0C0&
        txtHomeAdd1A.BackColor = &HC0C0&
        txtHomeAdd1.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addhome1")), "(Null)", m_cust("f_sts_unvalid_addhome1"))
        txtHomeAdd1A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addhome1")), "(Null)", m_cust("f_sts_unvalid_addhome1"))
    End If
    If m_cust("f_unvalid_addhome2") = "1" Then
        txtHomeAdd2.BackColor = &HC0C0&
        txtHomeAdd2A.BackColor = &HC0C0&
        txtHomeAdd2.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addhome2")), "(Null)", m_cust("f_sts_unvalid_addhome2"))
        txtHomeAdd2A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addhome2")), "(Null)", m_cust("f_sts_unvalid_addhome2"))
    End If
    If m_cust("f_unvalid_addoffice1") = "1" Then
        txtOfficeAdd1.BackColor = &HC0C0&
        txtOfficeAdd1A.BackColor = &HC0C0&
        txtOfficeAdd1.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addoffice1")), "(Null)", m_cust("f_sts_unvalid_addoffice1"))
        txtOfficeAdd1A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addoffice1")), "(Null)", m_cust("f_sts_unvalid_addoffice1"))
    End If
    If m_cust("f_unvalid_addoffice2") = "1" Then
        txtOfficeAdd2.BackColor = &HC0C0&
        txtOfficeAdd2A.BackColor = &HC0C0&
        txtOfficeAdd2.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addoffice2")), "(Null)", m_cust("f_sts_unvalid_addoffice2"))
        txtOfficeAdd2A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addoffice2")), "(Null)", m_cust("f_sts_unvalid_addoffice2"))
    End If
    If m_cust("f_unvalid_addmobile1") = "1" Then
        txtMobileAdd1.BackColor = &HC0C0&
        txtMobileAdd1A.BackColor = &HC0C0&
        txtMobileAdd1.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addmobile1")), "(Null)", m_cust("f_sts_unvalid_addmobile1"))
        txtMobileAdd1A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addmobile1")), "(Null)", m_cust("f_sts_unvalid_addmobile1"))
    End If
    If m_cust("f_unvalid_addmobile2") = "1" Then
        txtMobileAdd2.BackColor = &HC0C0&
        txtMobileAdd2A.BackColor = &HC0C0&
        txtMobileAdd2.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addmobile2")), "(Null)", m_cust("f_sts_unvalid_addmobile2"))
        txtMobileAdd2A.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_addmobile2")), "(Null)", m_cust("f_sts_unvalid_addmobile2"))
    End If
    If m_cust("f_unvalid_ec") = "1" Then
        txtECnoA.BackColor = &HC0C0&
        txtECno.BackColor = &HC0C0&
        txtECnoA.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_ec")), "(Null)", m_cust("f_sts_unvalid_ec"))
        txtECno.ToolTipText = "Telepon UnValid, karena: " & IIf(IsNull(m_cust("f_sts_unvalid_ec")), "(Null)", m_cust("f_sts_unvalid_ec"))
    End If
    
    
        
    '@@04-05-2012, Jika kategori call telah terisi, combo box dinonaktifkan
    If m_cust("homenoadd1") <> Empty And m_cust("stskathomeadd1") <> Empty Then
        CmbStsKatHome1.Enabled = False
    End If
    If m_cust("homenoadd2") <> Empty And m_cust("stskathomeadd2") <> Empty Then
        CmbStsKatHome2.Enabled = False
    End If
    If m_cust("officenoadd1") <> Empty And m_cust("stskatofficeadd1") <> Empty Then
        CmbStsKatOffice1.Enabled = False
    End If
    If m_cust("officenoadd2") <> Empty And m_cust("stskatofficeadd2") <> Empty Then
        CmbStsKatOffice2.Enabled = False
    End If
    If m_cust("mobilenoadd1") <> Empty And m_cust("stskathpadd1") <> Empty Then
        CmbStsKatHP1.Enabled = False
    End If
    If m_cust("mobilenoadd2") <> Empty And m_cust("stskathpadd2") <> Empty Then
        CmbStsKatHP2.Enabled = False
    End If
    
    '@@03-05-2012 buat nambahin tooltip dari keterangan nomor yang di black list
    Dim m_objrs_tooltip As ADODB.Recordset
    
    '@@220610 - Memberikan tanda merah pada no telepon yang di blacklist
    If m_cust("f_homeno") = 1 Then
        txtHomeNo1.ForeColor = vbRed
        txtHomeNo1A.ForeColor = vbRed
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("homeno") <> Empty Then
            CMDSQL = "select * from tblblacklist where no_telp='"
            CMDSQL = CMDSQL + CStr(Trim(m_cust("homeno"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtHomeNo1.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtHomeNo1A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    If m_cust("f_homeno2") = 1 Then
        txtHomeNo2.ForeColor = vbRed
        txtHomeNo2A.ForeColor = vbRed
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("homeno2") <> Empty Then
            CMDSQL = "select * from tblblacklist where no_telp='"
            CMDSQL = CMDSQL + CStr(Trim(m_cust("homeno2"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtHomeNo2.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtHomeNo2A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_officeno") = 1 Then
        txtOfficeNo1.ForeColor = vbRed
        txtOfficeNo1A.ForeColor = vbRed
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("officeno") <> Empty Then
            CMDSQL = "select * from tblblacklist where no_telp='"
            CMDSQL = CMDSQL + CStr(Trim(m_cust("officeno"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtOfficeNo1.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtOfficeNo1A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_officeno2") = 1 Then
        txtOfficeNo2.ForeColor = vbRed
        txtOfficeNo2A.ForeColor = vbRed
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("officeno2") <> Empty Then
            CMDSQL = "select * from tblblacklist where no_telp='"
            CMDSQL = CMDSQL + CStr(Trim(m_cust("officeno2"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtOfficeNo2.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtOfficeNo2A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_mobileno") = 1 Then
        txtMobileNo1.ForeColor = vbRed
        txtMobileNo1A.ForeColor = vbRed
        
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("mobileno") <> Empty Then
            CMDSQL = "select * from tblblacklist where no_telp='"
            CMDSQL = CMDSQL + CStr(Trim(m_cust("mobileno"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtMobileNo1.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtMobileNo1A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_mobileno2") = 1 Then
        txtMobileNo2.ForeColor = vbRed
        txtMobileNo2A.ForeColor = vbRed
        
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("mobileno2") <> Empty Then
            CMDSQL = "select * from tblblacklist where no_telp='"
            CMDSQL = CMDSQL + CStr(Trim(m_cust("mobileno2"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtMobileNo2.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtMobileNo2A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_homenoadd1") = 1 Then
        txtHomeAdd1.ForeColor = vbRed
        txtHomeAdd1A.ForeColor = vbRed
        
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("homenoadd1") <> Empty Then
            CMDSQL = "select * from tblblacklist where no_telp='"
            CMDSQL = CMDSQL + CStr(Trim(m_cust("homenoadd1"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtadd_phone(0).ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtHomeAdd1A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_homenoadd2") = 1 Then
        txtHomeAdd2.ForeColor = vbRed
        txtHomeAdd2A.ForeColor = vbRed
        
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("homenoadd2") <> Empty Then
            CMDSQL = "select * from tblblacklist where no_telp='"
            CMDSQL = CMDSQL + CStr(Trim(m_cust("homenoadd2"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtHomeAdd2.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtHomeAdd2A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If

    If m_cust("f_officenoadd1") = 1 Then
         txtOfficeAdd1.ForeColor = vbRed
         txtOfficeAdd1A.ForeColor = vbRed
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("officenoadd1") <> Empty Then
            CMDSQL = "select * from tblblacklist where no_telp='"
            CMDSQL = CMDSQL + CStr(Trim(m_cust("officenoadd1"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtadd_phone(1).ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtOfficeAdd1A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_officenoadd2") = 1 Then
        txtOfficeAdd2.ForeColor = vbRed
        txtOfficeAdd2A.ForeColor = vbRed
        
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("officenoadd1") <> Empty Then
            CMDSQL = "select * from tblblacklist where no_telp='"
            CMDSQL = CMDSQL + CStr(Trim(m_cust("officenoadd2"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtOfficeAdd2.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtOfficeAdd2A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_mobilenoadd1") = 1 Then
         txtMobileAdd1.ForeColor = vbRed
         txtMobileAdd1A.ForeColor = vbRed
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("mobilenoadd1") <> Empty Then
            CMDSQL = "select * from tblblacklist where no_telp='"
            CMDSQL = CMDSQL + CStr(Trim(m_cust("mobilenoadd1"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtadd_phone(2).ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtMobileAdd1A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_mobilenoadd2") = 1 Then
        txtMobileAdd2.ForeColor = vbRed
        txtMobileAdd2A.ForeColor = vbRed
        
        '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("mobilenoadd2") <> Empty Then
            CMDSQL = "select * from tblblacklist where no_telp='"
            CMDSQL = CMDSQL + CStr(Trim(m_cust("mobilenoadd2"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtadd_phone(3).ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtMobileAdd2A.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    If m_cust("f_ec_telp") = 1 Then
         txtECno.ForeColor = vbRed
         txtECnoA.ForeColor = vbRed
         '@@ 03-05-2012 Tambahan buat bikin tooltip
        If m_cust("ec_telp") <> Empty Then
            CMDSQL = "select * from tblblacklist where no_telp='"
            CMDSQL = CMDSQL + CStr(Trim(m_cust("ec_telp"))) + "'"
            Set m_objrs_tooltip = New ADODB.Recordset
            m_objrs_tooltip.CursorLocation = adUseClient
            m_objrs_tooltip.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs_tooltip.RecordCount > 0 Then
                txtECno.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
                txtECnoA.ToolTipText = "Nomor ini di BLACKLIST karena : " & IIf(IsNull(m_objrs_tooltip("keterangan")), "(Maaf, Tidak ada Alasan yang Tersedia)", m_objrs_tooltip("keterangan"))
            End If
            Set m_objrs_tooltip = Nothing
        End If
    End If
    
    
    '@@03-05-2012,,Buat Nandain Valid number -------------------------
    If m_cust("f_valid_home1") = 1 Then
        txtHomeNo1.ForeColor = vbBlue
        txtHomeNo1A.ForeColor = vbBlue
        
        txtHomeNo1.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_home1")), "", m_cust("f_sts_valid_home1"))
        txtHomeNo1A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_home1")), "", m_cust("f_sts_valid_home1"))
    End If
    If m_cust("f_valid_home2") = 1 Then
        txtHomeNo2.ForeColor = vbBlue
        txtHomeNo2A.ForeColor = vbBlue
        
        txtHomeNo2.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_home2")), "", m_cust("f_sts_valid_home2"))
        txtHomeNo2A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_home2")), "", m_cust("f_sts_valid_home2"))
    End If
    If m_cust("f_valid_office1") = 1 Then
        txtOfficeNo1.ForeColor = vbBlue
        txtOfficeNo1A.ForeColor = vbBlue
        
        txtOfficeNo1.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_office1")), "", m_cust("f_sts_valid_office1"))
        txtOfficeNo1A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_office1")), "", m_cust("f_sts_valid_office1"))
    End If
    If m_cust("f_valid_office2") = 1 Then
        txtOfficeNo2.ForeColor = vbBlue
        txtOfficeNo2A.ForeColor = vbBlue
        
        txtOfficeNo2.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_office2")), "", m_cust("f_sts_valid_office2"))
        txtOfficeNo2A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_office2")), "", m_cust("f_sts_valid_office2"))
    End If
    If m_cust("f_valid_mobile1") = 1 Then
        txtMobileNo1.ForeColor = vbBlue
        txtMobileNo1A.ForeColor = vbBlue
        
        txtMobileNo1.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_mobile1")), "", m_cust("f_sts_valid_mobile1"))
        txtMobileNo1A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_mobile1")), "", m_cust("f_sts_valid_mobile1"))
    End If
    If m_cust("f_valid_mobile2") = 1 Then
        txtMobileNo2.ForeColor = vbBlue
        txtMobileNo2A.ForeColor = vbBlue
        
        txtMobileNo2.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_mobile2")), "", m_cust("f_sts_valid_mobile2"))
        txtMobileNo2A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_mobile2")), "", m_cust("f_sts_valid_mobile2"))
    End If
    
    If m_cust("f_valid_addhome1") = 1 Then
        txtadd_phone(0).ForeColor = vbBlue
        txtHomeAdd1A.ForeColor = vbBlue
        
        txtadd_phone(0).ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addhome1")), "", m_cust("f_sts_valid_addhome1"))
        txtHomeAdd1A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addhome1")), "", m_cust("f_sts_valid_addhome1"))
    End If
    If m_cust("f_valid_addhome2") = 1 Then
        txtHomeAdd2.ForeColor = vbBlue
        txtHomeAdd2A.ForeColor = vbBlue
        
        txtHomeAdd2.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addhome2")), "", m_cust("f_sts_valid_addhome2"))
        txtHomeAdd2A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addhome2")), "", m_cust("f_sts_valid_addhome2"))
    End If
    If m_cust("f_valid_addoffice1") = 1 Then
        txtadd_phone(1).ForeColor = vbBlue
        txtOfficeAdd1A.ForeColor = vbBlue
        
        txtadd_phone(1).ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addoffice1")), "", m_cust("f_sts_valid_addoffice1"))
        txtOfficeAdd1A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addoffice1")), "", m_cust("f_sts_valid_addoffice1"))
    End If
    If m_cust("f_valid_addoffice2") = 1 Then
        txtOfficeAdd2.ForeColor = vbBlue
        txtOfficeAdd2A.ForeColor = vbBlue
        
        txtOfficeAdd2.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addoffice2")), "", m_cust("f_sts_valid_addoffice2"))
        txtOfficeAdd2A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addoffice2")), "", m_cust("f_sts_valid_addoffice2"))
    End If
    If m_cust("f_valid_addmobile1") = 1 Then
        txtadd_phone(2).ForeColor = vbBlue
        txtMobileAdd1A.ForeColor = vbBlue
        
        txtadd_phone(2).ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addmobile1")), "", m_cust("f_sts_valid_addmobile1"))
        txtMobileAdd1A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addmobile1")), "", m_cust("f_sts_valid_addmobile1"))
    End If
    If m_cust("f_valid_addmobile2") = 1 Then
        txtadd_phone(3).ForeColor = vbBlue
        txtMobileAdd2A.ForeColor = vbBlue
        
        txtadd_phone(3).ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addmobile2")), "", m_cust("f_sts_valid_addmobile2"))
        txtMobileAdd2A.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_addmobile2")), "", m_cust("f_sts_valid_addmobile2"))
    End If
    If m_cust("f_valid_ec") = 1 Then
        txtECnoA.ForeColor = vbBlue
        txtECno.ForeColor = vbBlue
        
        txtECnoA.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_ec")), "", m_cust("f_sts_valid_ec"))
        txtECno.ToolTipText = IIf(IsNull(m_cust("f_sts_valid_ec")), "", m_cust("f_sts_valid_ec"))
    End If
    '@@03-05-2012,,AKHIR Buat Nandain Valid number -------------------------
    
    
    '@@02-05-2012, Tambahan untuk menampilkan kategori telepon di additional phone
     CmbStsKatHome1.text = IIf(IsNull(m_cust("stskathomeadd1")), "", m_cust("stskathomeadd1"))
     CmbStsKatHome2.text = IIf(IsNull(m_cust("stskathomeadd2")), "", m_cust("stskathomeadd2"))
     CmbStsKatOffice1.text = IIf(IsNull(m_cust("stskatofficeadd1")), "", m_cust("stskatofficeadd1"))
     CmbStsKatOffice2.text = IIf(IsNull(m_cust("stskatofficeadd2")), "", m_cust("stskatofficeadd2"))
     CmbStsKatHP1.text = IIf(IsNull(m_cust("stskathpadd1")), "", m_cust("stskathpadd1"))
     CmbStsKatHP2.text = IIf(IsNull(m_cust("stskathpadd2")), "", m_cust("stskathpadd2"))
    
    
    '@@ 17-04-2012, Tambahan untuk request number
    TxtKategori.Caption = IIf(IsNull(m_cust("status_telp")), "", m_cust("status_telp"))
    TxtNoTelpReq.text = IIf(IsNull(m_cust("req_nomor_telp")), "", Trim(m_cust("req_nomor_telp")))
    
    '@@ 09042012, Tambahan untuk Status Risk Account: POP1 dan PP1
    'LblPop.Caption = IIf(IsNull(m_cust("status_pop1")), "", m_cust("status_pop1"))
    LblPP.Caption = IIf(IsNull(m_cust("status_pp1")), "", m_cust("status_pp1"))

    '01-02-2012, tambahkan status hot tobe collected
    If m_cust("status_htc") = "1" Then
        CmdKeep.BackColor = vbRed
        'CmdKeep.Caption = "Hot..."
    Else
        CmdKeep.BackColor = &H8000000F
        'CmdKeep.Caption = "Not Hot..."
    End If
    
    '@@ 29-03-2012 Tambahan status risk
    If IsNull(m_cust("status_risk")) = True Then
        LblStsRisk.ForeColor = &H80000012
    End If
    If IsNull(m_cust("status_risk")) = "1" Then
        LblStsRisk.ForeColor = &HFF&
    End If
    If IsNull(m_cust("status_risk")) = "2" Then
        LblStsRisk.ForeColor = &HFFFF&
    End If
    If IsNull(m_cust("status_risk")) = "3" Then
        LblStsRisk.ForeColor = &H80FF80
    End If
    
    '@@ 04082011 Tambahan Field
     On Error Resume Next
     'tandain20180807
     TxtInstallment.Value = Format(Round(IIf(IsNull(m_cust("instalment")), "0", m_cust("instalment"))), "##,###")
     
     
     Txtperiod.Caption = IIf(IsNull(m_cust("period")), "", m_cust("period"))
     TxtCurpri.Value = IIf(IsNull(m_cust("curpri")), "", m_cust("curpri"))
     lbltype.Caption = IIf(IsNull(m_cust("acc_type")), "", m_cust("acc_type"))
     lblpurge.Caption = IIf(IsNull(m_cust("sts_purge")), "", m_cust("sts_purge"))
     
     '@@ 04082011 Jika type data card instalment dan period di hide
     If (UCase(lbltype.Caption) = "CARD") Then
        Label11(9).Visible = False
        'TxtInstallment.Visible = False
        
        Label11(10).Visible = False
        Txtperiod.Visible = False
     End If
    
    '@@25/01/2012
    LblResultPTP.Caption = IIf(IsNull(m_cust("result_ptp")), "", m_cust("result_ptp"))
    
    '@@ 02022011
    LblMinPayment.Value = IIf(IsNull(m_cust("minpayment")), "0", m_cust("minpayment"))

    LblStatus.Caption = IIf(IsNull(m_cust("statusprior")), "", "Status : " & m_cust("statusprior"))
    lblCustId.text = IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID"))
    If cnull(m_cust("cmbbaseon")) <> "" Then
        Combo2.text = cnull(m_cust("cmbbaseon"))
    End If
    LblMother.Caption = IIf(IsNull(m_cust("mother")), "", m_cust("mother"))
    'sql = "delete  from tblnegoptp where custid in (select custid from tbllunas where custid ='" + IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID")) + "')"
    TxtCustid.text = IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID"))
    TxtName.text = IIf(IsNull(m_cust("NAME")), "", m_cust("NAME"))
    lblaoc.Caption = IIf(IsNull(m_cust("agent")), "", m_cust("Agent"))
    LblInterest.Caption = Format(IIf(IsNull(m_cust("INTEREST")), "0", m_cust("INTEREST")), "##,###")
    LblFees.Caption = Format(IIf(IsNull(m_cust("FEES")), "0", m_cust("FEES")), "##,###")
    lblregion.Caption = IIf(IsNull(m_cust("region")), "", m_cust("region"))
    txtbulan.Caption = IIf(IsNull(m_cust("bulan")), "", m_cust("bulan"))
    '@@ 04082011 Komponennya dibuang
    'lblaging.Caption = IIf(IsNull(m_cust("Aging")), "            ", m_cust("Aging"))
    
    'lblwilling.Caption = IIf(IsNull(m_cust("Willing_Ness")), "              ", m_cust("Willing_Ness"))
    lblRecsource.Caption = IIf(IsNull(m_cust("RECSOURCE")), "", m_cust("RECSOURCE"))
    
    If lblRecsource.Caption Like "*MANDIRI*" Then
        Label33.Caption = "Min Pay"
        label1(65).Caption = "Apply ID"
        Label20.Caption = "CustNo"
        TxtInstallment.Visible = True
        Label11(9).Visible = True
        
        'TxtInstallment.Value = lblAmount.text - Replace(Label34.Caption, ",", "")
        
    End If
    
    If lblRecsource.Caption Like "*BRI*" Or lblRecsource.Caption Like "*PANIN*" Or lblRecsource.Caption Like "*MAYBANK*" Then
        label1(65).Caption = "Card Num"
        Label23.Caption = "Cust No"
    End If
    
    If lblRecsource.Caption Like "*MAYBANK*" Then
        'Label11(6).Caption = "CWX"
        Label33.Caption = "Cur Bal"
        Command1001.Visible = True
    End If

    If lblRecsource.Caption Like "*PLUS*" Then
        'Label11(6).Caption = "CWX"
        label1(8).Visible = True
    End If
    
    LBLEXP.Caption = IIf(IsNull(m_cust("date_into_clas")), "", "Expire date " & Format(DateAdd("d", 60, m_cust("date_into_clas")), "dd-mm-yyyy"))
    
    '@@ 04082011 Dibuang
    'LblRiskLevel.Caption = IIf(IsNull(m_cust("RiskLevel")), "", m_cust("RiskLevel"))
    
    'lblPriority.Caption = IIf(IsNull(m_cust("Priority")), "", m_cust("Priority"))
    lblNama.text = Replace(IIf(IsNull(m_cust("NAME")), "", m_cust("NAME")), "'", "")
    lblCardNo.Caption = IIf(IsNull(m_cust("NoCard")), "", m_cust("NoCard"))
    lblID.Caption = IIf(IsNull(m_cust("ktpno")), "", m_cust("ktpno"))
    'lblDate.Value = IIf(IsNull(m_cust("BIRTHD")), "", Format(m_cust("BIRTHD"), "dd-mmm-yyyy"))
    LblDOB.Caption = IIf(IsNull(m_cust("DOB")), "", Left(m_cust("DOB"), 10))
    Text8.text = IIf(IsNull(m_cust("ADDRNOW")), "", m_cust("ADDRNOW"))
    TDB_cur_bal = IIf(IsNull(m_cust("CURBAL")), "", m_cust("CURBAL"))
    TXTRUMUS.text = IIf(IsNull(m_cust("typerumus")), "", m_cust("typerumus"))
    Combo1.text = IIf(IsNull(m_cust("stscallcust")), "", m_cust("stscallcust"))
    'tdbmaxad.Value = Format(IIf(IsNull(m_cust("maxad")), "0", m_cust("maxad")), "##,###")
    'tdbminad.Value = Format(IIf(IsNull(m_cust("minad")), "0", m_cust("minad")), "##,###")
    lbl_agentlama.Caption = IIf(IsNull(m_cust("agent_asli")), "", m_cust("agent_asli"))
    
    TxtInterest.Value = IIf(IsNull(m_cust("interest")), "", m_cust("interest"))
    
    ' TAMBAHAN CLASS 28 NOP 2013 ------------
    lblClass.Caption = IIf(IsNull(m_cust("cust_class")), "", m_cust("cust_class"))
    '----------------------------------------
     
    '@@ Tambahan 2 field (map dan cycle)
    LblMap = IIf(IsNull(m_cust("map")), "0", m_cust("map"))
    LblCycle = IIf(IsNull(m_cust("cycle")), "0", m_cust("cycle"))

   Set CEKREC = New ADODB.Recordset
    CEKREC.CursorLocation = adUseClient
    CEKREC.Open "select * from opening_screen where custid='" + lblCustId.text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    '@@ 12-10-2011, Blink OST dinonaktifkan
'    If CEKREC.RecordCount > 0 Then
'        'SSCommand1(7).BackColor = vbRed
'        TimerBlink.Enabled = True
'    Else
'        TimerBlink.Enabled = False
'    End If
    
     If InStr(1, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(3), "DE") > 0 Then
        txthasil.Visible = True
     Else
        txthasil.Visible = False
     End If
     
     Text6.text = IIf(IsNull(m_cust("disapp")), "0", m_cust("disapp"))
     
     '@@03-05-2012 DinonAktifkan
     'tdbhptrace.Value = IIf(IsNull(m_cust("hp1trace")), "", m_cust("hp1trace"))
     
     tdbtelptrace.Value = IIf(IsNull(m_cust("tlp1trace")), "", m_cust("tlp1trace"))
     txtremarkstrace.text = IIf(IsNull(m_cust("addrtrace")), "", m_cust("addrtrace"))
     
     bcekptp = False
    vrcek = IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new"))
    '@@22102012 Catet Tanggal Paid Off
    TanggalPaidOff = IIf(IsNull(m_cust("tgl_paid_off")), "", m_cust("tgl_paid_off"))
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
     
    '@@ 04-03-2011 Ubah status jika TL/SPV/Admin yang buka dapat membuka semua status
    If UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "MANAGER" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
       
        If vrcek <> "BP-" Or Mid(vrcek, 1, 3) = "PTP" Or Mid(vrcek, 1, 3) = "POP" Then
            strsql = "Select * from contacteddesc WHERE status=1"
        ElseIf vrcek = "BP-" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('BP-','PO-','CO-') AND status=1"
        ElseIf Mid(vrcek, 1, 3) = "PTP" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('BP-','PO-','CO-') AND status=1"
        ElseIf Mid(vrcek, 1, 3) = "POP" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('POP') AND status=1"
        End If
        
    Else
    '@@ 04-03-2011 Nah ini jika yang login Agent
        If vrcek = "" Then
            strsql = "Select * from contacteddesc WHERE status=1"
        Else
            '@@02102012 Untuk Agent PO- dinonaktifkan
            If vrcek = "VL-" Then
                strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('VL-','PR-','ON-','CO-') and status=1"
            ElseIf vrcek = "OS-" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('VL-','PR-','ON-','SK-','CO-') AND status=1"
            ElseIf vrcek = "PR-" Then
                 strsql = "Select * from contacteddesc WHERE  substring(KdNoProdPresented,1,3) in('PR-','ON-','CO-') AND status=1"
            ElseIf vrcek = "ON-" Then
                 strsql = "Select * from contacteddesc WHERE  substring(KdNoProdPresented,1,3) in('ON-','CO-') AND status=1"
            ElseIf vrcek = "SK-" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('VL-','PR-','SK-','CO-') AND status=1"
            ElseIf vrcek = "SP-" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('SP-','CO-') AND status=1"
            ElseIf vrcek = "BP-" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('BP-','CO-') AND status=1"
            ElseIf Mid(vrcek, 1, 3) = "PTP" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('BP-','CO-') AND status=1"
            ElseIf Mid(vrcek, 1, 3) = "POP" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('POP') AND status=1"
            '@@31052012Tambahan JIKA PAID OFF (PO-) DAN COMPLAIN (CO-)
            ElseIf Mid(vrcek, 1, 3) = "PO-" Then
                strsql = "select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('PO-') AND status=1"
            ElseIf Mid(vrcek, 1, 3) = "CO-" Then
                strsql = "select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('CO-') AND status=1"
            Else
                strsql = " Select * from contacteddesc WHERE status=1 "
            End If
            
        End If
    End If
    
    '@@Jika Status PO- (PAID OFF) yang login team leader maka accountnya tidak dapat di ubah statusnya
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        If Trim(Mid(vrcek, 1, 3)) = "PO-" Then
            strsql = "select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('PO-') AND status=1"
        End If
        If Trim(Mid(vrcek, 1, 3)) = "CO-" Then
            strsql = "select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('CO-') AND status=1"
        End If
    End If
    
    'STRSQL = " Select * from contacteddesc WHERE status=1 "
'    cboaccount.CLEAR
'    cboaccount.AddItem ""
'    M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_objrs.EOF
'        cboaccount.AddItem M_objrs!KdNoProdPresented
'        M_objrs.MoveNext
'    Wend
'    Set M_objrs = Nothing
    
'    '@@31-05-2012 Tambahan 2 Status Account, PAID OFF dan COMPLAIN
'    cboaccount.AddItem "PAID-OFF"
'    cboaccount.AddItem "COMPLAIN"
    
   If Left(IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new")), 3) <> "PTP" Then
    'cboaccount.Text = IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new"))
    cboaccount.text = IIf(IsNull(m_cust("kethslkerja_new")), "", m_cust("kethslkerja_new"))
    
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Or UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        If cboaccount.text = "Already Paid" Then
            cboaccount.Enabled = False
            CmdSendPTP.Enabled = False
        End If
    End If
   ElseIf Left(IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new")), 3) = "PTP" Then
     cboPTP.text = IIf(IsNull(m_cust("kethslkerja_new")), "", m_cust("kethslkerja_new"))
     cboaccount = IIf(IsNull(m_cust("ptpdesc")), "", m_cust("ptpdesc"))
     
     If UCase(MDIForm1.txtlevel.text) = "AGENT" Or UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        If cboaccount.text = "Already Paid" Then
            cboaccount.Enabled = False
            CmdSendPTP.Enabled = False
        End If
    End If
   End If
  
   CmbViaPtp.Enabled = False
   
   If Left(IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new")), 3) = "PTP" Then
        C_PTP.Value = vbChecked
        '@@ 05-10-2011 Tambahan field PTP VIA
        CmbViaPtp.text = IIf(IsNull(m_cust("ptpvia")), "", m_cust("ptpvia"))
        CmbViaPtp.Enabled = True
   End If
   
   If Trim(Mid(cboaccount, 1, 3)) = "POP" Or Trim(Mid(cboaccount, 1, 2)) = "BP" Then
       '@@ 05-10-2011 Tambahan field PTP VIA
        CmbViaPtp.text = IIf(IsNull(m_cust("ptpvia")), "", m_cust("ptpvia"))
   End If
   
   
   
 TglPTPNew = IIf(IsNull(m_cust("tglptpnew")), "", m_cust("tglptpnew"))
  If TglPTPNew <> "" Then
        'tdbptpnew.Value = Format(tglptpnew, "yyyy-mm-dd")
        tdbptpnew.Value = Format(m_cust("tglptpnew"), "yyyy-mm-dd")
  End If
  
If Left(vrcek, 3) = "PTP" Then
        SSCommand1(4).Visible = True
        Label43(2).Visible = True
Else
        SSCommand1(4).Visible = False
        Label43(2).Visible = False
End If

    If Left(vrcek, 2) = "BP" Then
  '  cboPOPSP.Enabled = False
'        FrmContacted.Enabled = False
'        C_Contacted.Enabled = False
'        cmbContacted.Enabled = False
'        cmbDescCon.Enabled = False
     End If
    
    lblOfficeAddr.text = IIf(IsNull(m_cust("ADDRPT")), "", m_cust("ADDRPT"))
    txtremarks_old.text = IIf(IsNull(m_cust("remarks_old")), "", m_cust("remarks_old"))
    lblZIP.Caption = IIf(IsNull(m_cust("ZIPNOW")), "", m_cust("ZIPNOW"))
    '@@04082011 NoCard dihapus dulu
    'lblNoCard.Caption = IIf(IsNull(m_cust("NoCard")), "", m_cust("NoCard"))
    
    lblNoPay.Caption = IIf(IsNull(m_cust("NoPay")), "", m_cust("NoPay"))
      
       
        
        
        
    'Else
    
     LblPrompA.Value = IIf(IsNull(m_cust("Principal")), "", m_cust("Principal"))
     
        
   If UCase(MDIForm1.txtlevel.text) <> "SUPERVISOR" Then
        If IIf(IsNull(m_cust("flaglead")), 0, m_cust("flaglead")) = 1 Then
            LblPrompA.Visible = False
            Label11(8).Visible = False
        Else
            LblPrompA.Visible = True
            Label11(8).Visible = True
       End If
    Else
          LblPrompA.Visible = False
          Label11(8).Visible = False
    End If
    
   ' End If
    
    '@@ 0408201 Dibuang
    'tdbprincipal.Value = IIf(IsNull(m_cust("Principal")), "", m_cust("Principal"))
    
    lblOpenDate.Value = IIf(IsNull(m_cust("OpenDate")), "", m_cust("OpenDate"))
    lblLastBill.Value = IIf(IsNull(m_cust("LastBill")), "", m_cust("LastBill"))
    lblLcAtm.Value = IIf(IsNull(m_cust("LcATMP")), "", m_cust("LcATMP"))
    txttenor.Value = IIf(IsNull(m_cust("tenor")), 0, m_cust("tenor"))
    vrtenor = IIf(IsNull(m_cust("tenor")), 0, m_cust("tenor"))
    lblBrokenPromised.Caption = IIf(IsNull(m_cust("BrokenPromise")), "", m_cust("BrokenPromise"))
    lblBD.Value = IIf(IsNull(m_cust("B_D")), "", m_cust("B_D"))
    lblLimit.Value = IIf(IsNull(m_cust("Limit")), "", m_cust("Limit"))
    vramount = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
    vrcekamont = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
    If listview1(0).ListItems.Count = 0 Then
    lblPayDt.Value = IIf(IsNull(m_cust("Pay_Dt")), "", m_cust("Pay_Dt"))
    End If
    'lblnocard.Caption = IIf(IsNull(m_cust("custno")), "", m_cust("custno"))
    'txttype.Caption = IIf(IsNull(m_cust("product_description")), "", m_cust("product_description"))
    If listview1(0).ListItems.Count = 0 Then
    lblLastPay.Value = IIf(IsNull(m_cust("LastPay")), "", m_cust("LastPay"))
    End If
    
    Label21.Caption = cnull(m_cust("delq_history"))
    Label25.Caption = cnull(m_cust("nocard"))
    '==asep=='
    Txtlimit.Value = cnull(m_cust("limit"))
    Txtcycle.Caption = cnull(m_cust("cycle"))
    Txtblk_code.Caption = cnull(m_cust("block_code_1"))
    Txtstatus.Caption = cnull(m_cust("stskathomeadd1"))
    Txtro.Caption = cnull(m_cust("batchdiskon"))
    Txtzip_code.Caption = cnull(m_cust("zipnow"))
    Txtjk.Caption = cnull(m_cust("f_sts_valid_office1"))
    Txteusex.Caption = cnull(m_cust("f_sts_valid_home2"))
    '====='
    If Right(cnull(m_cust("batchdiskon")), 5) = ",0000" Or Right(cnull(m_cust("batchdiskon")), 5) = ".0000" Then
        Label34.Caption = Format(Round(Left(cnull(m_cust("batchdiskon")), Len(cnull(m_cust("batchdiskon"))) - 5), 1), "##,###")
    Else
        Dim hit As Double
        Label34.Caption = cnull(m_cust("batchdiskon"))
        hit = Replace(Label34.Caption, ",", ".")
        Label34.Caption = Format(Round(hit), "##,###")
    End If
    Text11.text = cnull(m_cust("afaxno"))
    If Text11.text <> "" Then
        Text11m.text = Left(Text11.text, Len(Text11.text))
    End If
'    If Text11.text <> "" Then
'        Text11m.text = Left(Text11.text, Len(Text11.text) - 3) & "###"
'    End If
    ec_desc.text = cnull(m_cust("product_desc"))
    
    
    lblTtlPay.Value = IIf(IsNull(m_cust("TtlPay")), "", m_cust("TtlPay"))
    'If IIf(IsNull(m_cust("flaglead")), 0, m_cust("flaglead")) = 1 Then
     '   lblAmount.Value = IIf(IsNull(m_cust("AmountWo")), "", Format(m_cust("AmountWo"), "##.##0"))
     '   If lblAmount.ValueIsNull Then
      '      lblAmount.Value = 0
      '  Else
       '     lblAmount.Value = lblAmount.Value + (lblAmount.Value * 18.26) / 100
       ' End If
        
    'Else
    
    
    lblAmount.Value = Round(IIf(IsNull(m_cust("curbal")), "", Format(m_cust("curbal"), "##.##0")))
    Dim ab, ac As Double
    
    ab = Replace(lblAmount.text, ",", "")
    ac = Replace(Label34.Caption, ",", "")
    
    If ab < ac Then
        Dim changea, changeb As String
        changea = Label34.Caption
        changeb = lblAmount.text
        lblAmount.text = changea
        Label34.Caption = changeb
    End If
    
    If lblRecsource.Caption Like "*MANDIRI*" And lblRecsource.Caption Like "*MIGRASI*" Then
        Dim z1, z2, z3, z4 As Double
        
        z1 = Replace(lblAmount.Value, ",", "")
        z2 = Replace(TxtInstallment.Value, ",", "")
        
        Label34.Caption = Format(Round(z2 + ((z1 - z2) * (5 / 100))), "##,###")
    End If
    
        
    lbl_principal.Value = IIf(IsNull(m_cust("principal")), "", Format(m_cust("principal"), "##.##0"))
    txtdenda.Value = lblAmount.Value - lbl_principal.Value
    
    'End If
    
    If lblAmount.ValueIsNull Then
    
            tdbmaxad.Value = 0
        Else
            tdbmaxad.Value = lblAmount.Value - (lblAmount.Value * 24) / 100
        End If
    
    
     If lblAmount.ValueIsNull Then
            tdbminad.Value = tdbminad.Value - (lblAmount.Value * 35) / 100
        Else
            tdbminad.Value = lblAmount.Value - (lblAmount.Value * 31) / 100
        End If
        
    Tdbbalance.Value = IIf(IsNull(m_cust("AmountWo")), "", Format(m_cust("AmountWo"), "##.##0"))
    
    ' ----------- LATE FEE -------------
    TDBlate_fee.Value = IIf(IsNull(m_cust("late_fee")), "", Format(m_cust("late_fee"), "##.##0"))
    ' ----------------------------------
    
    ' ------------ CASE DECEASE -----------
    If lblClass.Caption = "835" Then
        Command3.Enabled = False
        Label11(19).Visible = True
    End If
    
    If IIf(IsNull(m_cust("f_decease")), "", m_cust("f_decease")) = 1 Then
        Command3.Enabled = False
        Label11(19).Visible = True
    End If
    ' -------------------------------------
    
    txtHomeNo1.text = IIf(IsNull(m_cust("HOMENO")), "", m_cust("HOMENO"))
'    If txtHomeNo1.text <> "" Then
'        txtHomeNo1m.text = Left(txtHomeNo1.text, Len(txtHomeNo1.text) - 3) & "###"
'    End If
    '===asep==='
    If txtHomeNo1.text <> "" Then
        txtHomeNo1m.text = Left(txtHomeNo1.text, Len(txtHomeNo1.text))
    End If
    '=========='
    '@@03-05-2012 DinonAktifkan
    'AHome1.Value = IIf(IsNull(m_cust("AHOMENO")), "", m_cust("AHOMENO"))
    
    
    
    If IsNull(m_cust("HOMENO")) = False And m_cust("HOMENO") <> "" Then
        'txtHomeNo1A.Value = Left(m_cust("HOMENO"), Len(m_cust("HOMENO")) - 3) & "XXX"
        txtHomeNo1A.Value = Left(m_cust("HOMENO"), 4) & "BBB" & Mid(m_cust("HOMENO"), 8, 15)
        CmbPhone.AddItem "HomePhone"
    End If
    
    '@@ 03-05-2012 DinonAktifkan
    'AHome2.Value = IIf(IsNull(m_cust("AHOMENO2")), "", m_cust("AHOMENO2"))
    
    txtHomeNo2.Value = IIf(IsNull(m_cust("HOMENO2")), "", m_cust("HOMENO2"))
    If IsNull(m_cust("HOMENO2")) = False And m_cust("HOMENO2") <> "" Then
        'txtHomeNo2A.Value = Left(m_cust("HOMENO2"), Len(m_cust("HOMENO2")) - 3) & "XXX"
        txtHomeNo2A.Value = Left(m_cust("HOMENO2"), 4) & "BBB" & Mid(m_cust("HOMENO2"), 8, 15)
        CmbPhone.AddItem "HomePhone2"
    End If
    
    '@@03-05-2012 DinonAktifkan
    'AOffice1.Value = IIf(IsNull(m_cust("AOFFICENO")), "", m_cust("AOFFICENO"))
    
    txtOfficeNo1.text = IIf(IsNull(m_cust("OFFICENO")), "", m_cust("OFFICENO"))
    '==asep==='
    If txtOfficeNo1.text <> "" Then
        txtOfficeNo1m.text = Left(txtOfficeNo1.text, Len(txtOfficeNo1.text))
    End If
    '========='
'    If txtOfficeNo1.text <> "" Then
'        txtOfficeNo1m.text = Left(txtOfficeNo1.text, Len(txtOfficeNo1.text) - 3) & "###"
'    End If
    If IsNull(m_cust("OFFICENO")) = False And m_cust("OFFICENO") <> "" Then
        'txtOfficeNo1A.Value = Left(m_cust("OFFICENO"), Len(m_cust("OFFICENO")) - 3) & "XXX"
        txtOfficeNo1A.Value = Left(m_cust("OFFICENO"), 4) & "BBB" & Mid(m_cust("OFFICENO"), 8, 15)
        CmbPhone.AddItem "OfficePhone"
    End If
    
    '@@03-05-2012 DinonAktifkan
    'AOffice2.Value = IIf(IsNull(m_cust("AOFFICENO2")), "", m_cust("AOFFICENO2"))
    
    txtOfficeNo2.Value = IIf(IsNull(m_cust("OFFICENO2")), "", m_cust("OFFICENO2"))
    If IsNull(m_cust("OFFICENO2")) = False And m_cust("OFFICENO2") <> "" Then
        'txtOfficeNo2A.Value = Left(m_cust("OFFICENO2"), Len(m_cust("OFFICENO2")) - 3) & "XXX"
        txtOfficeNo2A.Value = Left(m_cust("OFFICENO2"), 4) & "BBB" & Mid(m_cust("OFFICENO2"), 8, 15)
        CmbPhone.AddItem "OfficePhone2"
    End If
    txtMobileNo1.text = IIf(IsNull(m_cust("MOBILENO")), "", m_cust("MOBILENO"))
    '===asep===='
    If txtMobileNo1.text <> "" Then
        txtMobileNo1m.text = Left(txtMobileNo1.text, Len(txtMobileNo1.text))
    End If
    '=========='
    
'    If txtMobileNo1.text <> "" Then
'        txtMobileNo1m.text = Left(txtMobileNo1.text, Len(txtMobileNo1.text) - 3) & "###"
'    End If
    If IsNull(m_cust("MOBILENO")) = False And m_cust("MOBILENO") <> "" Then
        'txtMobileNo1A.Value = Left(m_cust("MOBILENO"), Len(m_cust("MOBILENO")) - 3) & "XXX"
        txtMobileNo1A.Value = Left(m_cust("MOBILENO"), 4) & "BBB" & Mid(m_cust("MOBILENO"), 8, 15)
        CmbPhone.AddItem "Hp"
    End If
    txtMobileNo2.Value = IIf(IsNull(m_cust("MOBILENO2")), "", m_cust("MOBILENO2"))
    If IsNull(m_cust("MOBILENO2")) = False And m_cust("MOBILENO2") <> "" Then
        'txtMobileNo2A.Value = Left(m_cust("MOBILENO2"), Len(m_cust("MOBILENO2")) - 3) & "XXX"
        txtMobileNo2A.Value = Left(m_cust("MOBILENO2"), 4) & "BBB" & Mid(m_cust("MOBILENO2"), 8, 15)
        CmbPhone.AddItem "Hp2"
    End If
    
    '@@ 03-05-2012 Dinonaktifkan
    'AHomeAdd1(0).Value = IIf(IsNull(m_cust("AHOMENOADD1")), "", m_cust("AHOMENOADD1"))
    
    '@@03-05-2012 Dinonaktifkan
    'AHomeAdd2(1).Value = IIf(IsNull(m_cust("AHOMENOADD2")), "", m_cust("AHOMENOADD2"))
    
    '@@03-05-2012 Dinonaktifkan
    'AOfficeAdd(2).Value = IIf(IsNull(m_cust("AOFFICENOADD1")), "", m_cust("AOFFICENOADD1"))
    'AOfficeAdd(3).Value = IIf(IsNull(m_cust("AOFFICENOADD2")), "", m_cust("AOFFICENOADD2"))
   
    txtadd_phone(0).text = IIf(IsNull(m_cust("HOMENOADD1")), "", m_cust("HOMENOADD1"))
    If txtadd_phone(0).text <> "" Then
        txtadd_phone(6).text = Left(txtadd_phone(0).text, Len(txtadd_phone(0).text))
    End If
'    If txtadd_phone(0).text <> "" Then
'        txtadd_phone(6).text = Left(txtadd_phone(0).text, Len(txtadd_phone(0).text) - 3) & "###"
'    End If
    If IsNull(m_cust("HOMENOADD1")) = False And m_cust("HOMENOADD1") <> "" Then
        txtHomeAdd1A.Value = Left(m_cust("HOMENOADD1"), 4) & "BBB" & Mid(m_cust("HOMENOADD1"), 8, 15)
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddHome1"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@26 May 2011 BP- dan Ptp- juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddHome1"
'        End If
        '@@08-06-2011 Semua Telepon dibuka, status apapun
        '@@17-04-2012 Telepon di Non aktifkan
        '@@24-04-2012 Diaktifkan lagi
        CmbPhone.AddItem "AddHome1"
    Else
        txtadd_phone(0).Visible = True
        txtHomeAdd1A.Visible = False
    End If
    txtHomeAdd2.Value = IIf(IsNull(m_cust("HOMENOADD2")), "", m_cust("HOMENOADD2"))
    If IsNull(m_cust("HOMENOADD2")) = False And m_cust("HOMENOADD2") <> "" Then
        txtHomeAdd2A.Value = Left(m_cust("HOMENOADD2"), 4) & "BBB" & Mid(m_cust("HOMENOADD2"), 8, 15)
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddHome2"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@26 May 2011, BP- dan PTP- juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddHome2"
'        End If
        '@@08-06-2011 Telepon dibuka,status apa pun
        '@@17-04-2012 Telepon di Non aktifkan
        '@@24-04-2012 Diaktifkan Lagi
        CmbPhone.AddItem "AddHome2"
    Else
        txtHomeAdd2A.Visible = False
        txtHomeAdd2.Visible = True
    End If
    txtadd_phone(1).text = IIf(IsNull(m_cust("OFFICENOADD1")), "", m_cust("OFFICENOADD1"))
    If txtadd_phone(1).text <> "" Then
        txtadd_phone(5).text = Left(txtadd_phone(1).text, Len(txtadd_phone(1).text))
    End If
'    If txtadd_phone(1).text <> "" Then
'        txtadd_phone(5).text = Left(txtadd_phone(1).text, Len(txtadd_phone(1).text) - 3) & "###"
'    End If
    If IsNull(m_cust("OFFICENOADD1")) = False And m_cust("OFFICENOADD1") <> "" Then
        txtOfficeAdd1A.Value = Left(m_cust("OFFICENOADD1"), 4) & "BBB" & Mid(m_cust("OFFICENOADD1"), 8, 15)
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddOffice1"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@ 26 May 2011, BP- dan PTP- ditampilkan juga
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddOffice1"
'        End If
        '@@08-06-2011 Telepon dibuka, status apapun
        '@@17-04-2012 Telepon di Non aktifkan
        '@@ 24042012 Diaktifkan lagi
        CmbPhone.AddItem "AddOffice1"
    Else
        txtOfficeAdd1A.Visible = False
        txtadd_phone(1).Visible = True
    End If
    txtOfficeAdd2.Value = IIf(IsNull(m_cust("OFFICENOADD2")), "", m_cust("OFFICENOADD2"))
    If IsNull(m_cust("OFFICENOADD2")) = False And m_cust("OFFICENOADD2") <> "" Then
        
        anto = Trim(Left(m_cust("OFFICENOADD2"), 4) + " " + Mid(m_cust("OFFICENOADD2"), 8, 15))
        If Len(anto) = 0 Then
        txtOfficeAdd2A.Value = ""
        
        Else
        
        txtOfficeAdd2A.Value = Left(m_cust("OFFICENOADD2"), 4) & "BBB" & Mid(m_cust("OFFICENOADD2"), 8, 15)
        
        End If
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddOffice2"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@ 26 May 2011 BP- dan PTP- juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddOffice2"
'        End If
        '@@ 08-06-2011 Status telepon dibuka, status apapun
        '@@17-04-2012 Telepon di Non aktifkan
        '@@ 24042012 Diaktifkan lagi
        CmbPhone.AddItem "AddOffice2"
    Else
        txtOfficeAdd2.Visible = True
        txtOfficeAdd2A.Visible = False
    End If
    txtadd_phone(2).text = IIf(IsNull(m_cust("MOBILENOADD1")), "", m_cust("MOBILENOADD1"))
    If txtadd_phone(2).text <> "" Then
        txtadd_phone(4).text = Left(txtadd_phone(2).text, Len(txtadd_phone(2).text))
    End If
'    If txtadd_phone(2).text <> "" Then
'        txtadd_phone(4).text = Left(txtadd_phone(2).text, Len(txtadd_phone(2).text) - 3) & "###"
'    End If
    txtadd_phone(3).text = IIf(IsNull(m_cust("MOBILENOADD2")), "", m_cust("MOBILENOADD2"))
    If txtadd_phone(3).text <> "" Then
        txtadd_phone(7).text = Left(txtadd_phone(3).text, Len(txtadd_phone(3).text))
    End If
    If IsNull(m_cust("MOBILENOADD1")) = False And m_cust("MOBILENOADD1") <> "" Then
        txtMobileAdd1A.text = Left(m_cust("MOBILENOADD1"), 4) & "BBB" & Mid(m_cust("MOBILENOADD1"), 8, 15)
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddMobile1"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@ 26 May 2011 BP- dan PTP- juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddMobile1"
'        End If
        '@@ 08-06-2011 Status Telepon dibuka, status apapun
        '@@17-04-2012 Telepon di Non aktifkan
        '@@ 24042012 Diaktifkan lagi
        CmbPhone.AddItem "AddMobile1"
    Else
        txtadd_phone(2).Visible = True
        txtMobileAdd1A.Visible = False
    End If
    txtMobileAdd2.Value = IIf(IsNull(m_cust("MOBILENOADD2")), "", m_cust("MOBILENOADD2"))
    If IsNull(m_cust("MOBILENOADD2")) = False And m_cust("MOBILENOADD2") <> "" Then
        txtMobileAdd2A.Value = Left(m_cust("MOBILENOADD2"), 4) & "BBB" & Mid(m_cust("MOBILENOADD2"), 8, 15)
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddMobile2"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@ 26 May 2011, BP- dan PTP- ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddMobile2"
'        End If
        '@@ 08-06-2011, status telepon dibuka, status apapun
        '@@17-04-2012 Telepon di Non aktifkan
        '@@ 24042012 Diaktifkan lagi
        CmbPhone.AddItem "AddMobile2"
    Else
        txtMobileAdd2.Visible = True
        txtMobileAdd2A.Visible = False
    End If
   
    AddrNow.text = IIf(IsNull(m_cust("TxtPtpAddr")), "", m_cust("TxtPtpAddr"))
    LblLunas.Caption = IIf(IsNull(m_cust!tgllunas), "", "TELAH LUNAS")
    TxtEC.text = IIf(IsNull(m_cust!ec_name), "", m_cust!ec_name)
    txtECno.Value = IIf(IsNull(m_cust!ec_telp), "", m_cust!ec_telp)
    If IsNull(m_cust("ec_telp")) = False And m_cust("ec_telp") <> "" Then
        txtECnoA.Value = Left(m_cust("ec_telp"), 4) & "BBB" & Mid(m_cust("ec_telp"), 8, 15)
'        '@@ 11-04-2011 EconPhone Di Non Aktifin Dulu, (aktif jika datanya berdasarkan mba wulan)
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "EconPhone"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- dan kosong maka Additional&EC di tampilkan
'        '@@26 May 2011 BP- dan PTP juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "OS-" Or CekStatus = "" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "EconPhone"
'        End If
        '@@ 08-06-2011, Telepon dibuka status apapun
        CmbPhone.AddItem "EconPhone"
    Else
        txtECnoA.Visible = False
        txtECno.Visible = True
    End If
    
    '@@02-05-2011  Tambahan Additional
    TxtAdditional.Value = IIf(IsNull(m_cust("telp_additional")), "", m_cust("telp_additional"))
     If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
            TxtAdditional.Enabled = False
        End If
    If TxtAdditional <> "" Then
        If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
            TxtAdditional.Enabled = False
        End If
        '@@17-04-2012 Telepon di Non aktifkan
        '@@02052012 Diaktifkan Lagi
        CmbPhone.AddItem "TelpAdditional"
    End If
    
    '@@17-04-2012,Tambahan
    If TxtNoTelpReq.Value <> "" Then
        CmbPhone.AddItem TxtKategori.Caption
    End If
    
    txtECAdd.text = IIf(IsNull(m_cust!ECAddr), "", m_cust!ECAddr)
    cbolastcall.text = IIf(IsNull(m_cust!statuscall), "", Trim(m_cust!statuscall))
    cbolastcall.text = IIf(IsNull(m_cust!stscallwith), "", m_cust!stscallwith)
'    If cboaccount.Text <> "" Then
'        Call statusgroup
'    End If
' cari extension
    If InStr(1, txtOfficeNo1.text, "X", vbTextCompare) > 0 Then
        '@@02052012 Extension dinonaktifkan
        'TxtExt1.Text = Right(txtOfficeNo1.Value, Len(txtOfficeNo1.Value) - InStr(1, txtOfficeNo1.Value, "X", vbTextCompare))
    End If
    If InStr(1, txtOfficeNo2.Value, "X", vbTextCompare) > 0 Then
        '@@02052012 Extension dinonaktifkan
        'TxtExt2.Text = Right(txtOfficeNo2.Value, Len(txtOfficeNo2.Value) - InStr(1, txtOfficeNo2.Value, "X", vbTextCompare))
    End If
    If InStr(1, txtOfficeAdd1.Value, "X", vbTextCompare) > 0 Then
        '@@02052012 Extension dinonaktifkan
        'TxtExt3.Text = Right(txtOfficeAdd1.Value, Len(txtOfficeAdd1.Value) - InStr(1, txtOfficeAdd1.Value, "X", vbTextCompare))
    End If
    If InStr(1, txtOfficeAdd2.Value, "X", vbTextCompare) > 0 Then
        '@@02052012 Extension dinonaktifkan
        'TxtExt4.Text = Right(txtOfficeAdd2.Value, Len(txtOfficeAdd2.Value) - InStr(1, txtOfficeAdd2.Value, "X", vbTextCompare))
    End If
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        If Len(txtECno.Value) > 2 Then
            txtECno.ReadOnly = True
        End If
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
        If Len(txtECno.Value) > 2 Then
            txtECno.ReadOnly = True
        End If
    End If
   
    
    sPending = CStr(Trim(IIf(IsNull(m_cust!f_Pending), "", m_cust!f_Pending)))
     If sPending = "Pending" Then
         'chkAppv(0).Value = 0 '@@ 25/01/2012 Komponen Tak Terpakai
    End If
    
'    Select Case m_cust!RECSTATUS
'        Case "V"
'            C_VALID.Value = 1
'            cbovalid.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            cbodescvalid.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'        Case "N"
'            C_NotContacted.Value = 1
'            cmbUncontacted.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            cmbDescUn.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'        Case "C"
'            C_Contacted.Value = 1
'            kontak = True
'            If mdiform1.txtlevel.text = "Agent" Then
'                If Left(vrcek, 3) = "POP" Then
'                    C_SKIP.Enabled = False
'                    C_VALID.Enabled = False
'                    cboPOPSP.Enabled = False
'                    FrmPayment.Enabled = True
'                    C_Payment.Value = 1
'                End If
'            End If
'            cmbContacted.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'      Case "P"
'            C_PTP.Value = 1
'            cboPTP.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            'cmbDescCon.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'            If mdiform1.txtlevel.text = "Agent" Then
'                C_VALID.Enabled = False
'                C_Contacted.Enabled = False
'                FrMValid.Enabled = False
'                C_SKIP.Enabled = False
'                FrmSKIP.Enabled = False
'            End If
'         Case "S"
'            C_SKIP.Value = 1
'            cboskip.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            cbodescskip.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'         Case "O"
'            'C_POPSP.Value = 1
'            cboPOPSP.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            'cmbDescCon.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))      cmbDescCon.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'     End Select
    cboaccount.text = IIf(IsNull(m_cust!statuscall), "", m_cust!statuscall)
    If MDIForm1.txtlevel.text = "Agent" Then
'        If IIf(IsNull(m_cust!RECSTATUS), "", m_cust!RECSTATUS) <> "O" Then
'            frmpopsp.Enabled = False
'           cboPOPSP.Enabled = False
'        End If
    End If
        If IIf(IsNull(m_cust!f_cek_new), "", Left(m_cust!f_cek_new, 3)) = "PTP" Or Left(m_cust!f_cek_new, 3) = "POP" Or Left(m_cust!f_cek_new, 3) = "SP-" Or Left(m_cust!f_cek_new, 3) = "PRE" Then
            C_Payment.Value = 1
            TdbPTP.Value = IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP)
            vrtdbdateptp = IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP)
            vrdateptp = IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP)
            TDBDate3.Value = IIf(IsNull(m_cust!dateptp), "", Format(m_cust!dateptp, "yyyy-mm-dd"))
            vrnewdate = IIf(IsNull(m_cust!dateptp), "", Format(m_cust!dateptp, "yyyy-mm-dd"))
            txtPayment.Value = IIf(IsNull(m_cust!ttlptp), 0, m_cust!ttlptp)
            vrttlptp = IIf(IsNull(m_cust!ttlptp), 0, m_cust!ttlptp)
            Tdabamoint.Value = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
            vramount = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
            TxtPayment2.Value = IIf(IsNull(m_cust!ttlptp), 0, m_cust!ttlptp) 'tampilkan di detail payment
            cmbDiscount.text = IIf(IsNull(m_cust!discpersen), 0, m_cust!discpersen)
            vrdiskon = IIf(IsNull(m_cust!discpersen), 0, m_cust!discpersen)
            CmbBaseOn.text = IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn)
            vrbaseon = IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn)
            'TdbDatePTP.Value = IIf(IsNull(m_cust!TGLINCOMING), "", m_cust!TGLINCOMING)
            
            '@@25/01/2012 Tambahan, tambahkan data tanggal tagih
            TdbTglTagih.Value = IIf(IsNull(m_cust!tgl_tagih), "", Format(m_cust!tgl_tagih, "yyyy-mm-dd"))
        Else
        End If
End If
Call Custid_Double
'Set m_cust1 = M_DATA.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + lblCustId.text + "'", mdiform1.txtlevel.text)
Set m_cust1 = m_data.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + Trim(lblCustId.text) + "'")
While Not m_cust1.EOF
    'Set listitem = ListView1(1).ListItems.ADD(, , Left(m_cust1("TGL"), 4) & "/" & Mid(m_cust1("TGL"), 5, 2) & "/" & IIf(IsNull(m_cust1("TGL")), "", Mid(m_cust1("TGL"), 7, 2)) & " " & IIf(IsNull(m_cust1("TGL")), "", Mid(m_cust1("TGL"), 9, 2)) & ":" & Right(m_cust1("TGL"), 2))
     sTime_Hst = ""
     If IIf(IsNull(m_cust1("TGL")), "", m_cust1!TGL) <> "" Then
         'sTime_Hst = Format(IIf(IsNull(m_cust1("TGL")), "", m_cust1!TGL), "mm-dd-yyyy hh:mm:ss") & Format(IIf(IsNull(m_cust1("stop_time")), "", m_cust1!stop_time), " - hh:mm:ss")
        sTime_Hst = Format(IIf(IsNull(m_cust1("TGL")), "", m_cust1!TGL), "mm-dd-yyyy hh:mm:ss")
     End If
     Set ListItem = listview1(1).ListItems.ADD(, , sTime_Hst)
        ListItem.SubItems(1) = IIf(IsNull(m_cust1("HST")), "", m_cust1("HST"))
        ListItem.SubItems(2) = IIf(IsNull(m_cust1("user_log")), "", m_cust1("user_log"))
        ListItem.SubItems(3) = IIf(IsNull(m_cust1("AGENT")), "", m_cust1("AGENT"))
        ListItem.SubItems(4) = IIf(IsNull(m_cust1("KodeDs")), "", m_cust1("KodeDs"))
        ListItem.SubItems(5) = IIf(IsNull(m_cust1("statuscall")), "", m_cust1("statuscall"))
        ListItem.SubItems(6) = IIf(IsNull(m_cust1("ststelpwith")), "", m_cust1("ststelpwith"))
        ListItem.SubItems(7) = IIf(IsNull(m_cust1("id")), "", m_cust1("id"))
        ListItem.SubItems(8) = IIf(IsNull(m_cust1("unique_id")), "", m_cust1("unique_id"))
        'listitem.SubItems(4) = IIf(IsNull(m_cust1("f_cek")), "", m_cust1("f_cek"))
                
                
                'Data Special 'jejaktian 18032016
                If IIf(IsNull(m_cust1("f_special")), 0, m_cust1("f_special")) = "1" Then
                    For K = 1 To 7
                        ListItem.ListSubItems(K).ForeColor = vbRed
                        ListItem.ListSubItems(K).Bold = True
                    Next K
                End If
        ' ------------------------------------------
m_cust1.MoveNext
Wend


Call isi_datapayment
Call Show_NEGOPTP
Call Show_Reserve
Call Isi_listScript
Call Isi_SendSMS

Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient

'@@ 22-09-2011, penghitungan total payment di tabel lunas juga memperhatikan tgl data masuk
'total payment yang masuk adalah payment yang paydate-nya harus lebih besar dari data yang masuk
'CMDSQL = "Select custid, sum(payment) as jml from tbllunas where custid = '" + lblCustId.text + "' GROUP BY CUSTID"
CMDSQL = "select sum(payment) as jml from "
CMDSQL = CMDSQL + "(SELECT b.custid as custid1, a.CUSTID,a.PayDate, "
CMDSQL = CMDSQL + " a.Payment,a.Agent,a.FieldName,a.Id from tbllunas a "
CMDSQL = CMDSQL + " inner join mgm b on "
CMDSQL = CMDSQL + " a.custid=b.custid  WHERE a.custid='"
CMDSQL = CMDSQL + lblCustId.text + "'  and date(a.Paydate)+1  > b.tglsource  order by a.PayDate asc ) as c"

M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_objrs.EOF
        TxtAfterPay.Value = IIf(IsNull(M_objrs("jml")), 0, M_objrs("jml"))
        M_objrs.MoveNext
Wend
 
 'hitung sisa hutang
 txtSisaHutang.Value = Val(TxtPayment2.Value) - Val(TxtAfterPay.Value)
 
 '---------->> hitung PRINCIPLE & AMOUNTWO  after pay  <<-----------------
 If TxtAfterPay.Value = 0 Then
    '@@04082011 Principle dibuang
    'txtPrinciple_A.Value = 0
    
    txtAmountwo_A.Value = 0
    Else
    If LblPrompA.ValueIsNull Or lblAmount.ValueIsNull Then
    Exit Sub
    End If
  '@@04082011 Principle dibuang
  'txtPrinciple_A.Value = Val(LblPrompA.Value) - Val(TxtAfterPay.Value)
  
  txtAmountwo_A.Value = Val(lblAmount.Value) - Val(TxtAfterPay.Value)
 End If
 
    If lblAmount.ValueIsNull Then
           '@@04082011 Dibuang
           'Woafter.Value = 0
       Else
           '@@04082011 Dibuang
           'Woafter.Value = lblAmount - TxtAfterPay.Value
    End If
  
    If listview1(0).ListItems.Count <> 0 Then
          '@@ 27-07-2011 , dimatiin dulu nih, cznya pay_dtnya jadi ke ambil dari payment disini
          'lblPayDt.Value = listview1(0).ListItems(listview1(0).ListItems.Count).Text
          'lblLastPay.Value = listview1(0).ListItems(listview1(0).ListItems.Count).SubItems(1)
          
'          TxtLPDPayment.Value = ListView1(0).ListItems(ListView1(0).ListItems.Count).Text
'          TxtLPAPayment.Value = ListView1(0).ListItems(ListView1(0).ListItems.Count).SubItems(1)
            
          '@@ 14042012, Karena list payment diubah berdasarkan desc, diubah
          TxtLPDPayment.Value = listview1(0).ListItems(1).text
          TxtLPAPayment.Value = listview1(0).ListItems(1).SubItems(1)
          LBLEXP.Caption = "Expire Date " + glexp
    End If
 
    'jejaktian30052016
'    If m_cust("F_CEK_NEW") = "& %PTP% &" Then
'        CmbPhone.Enabled = False
'        txtHomeNo1A.Enabled = False
'        txtHomeNo1A.Enabled = False
'        txtHomeNo2.Enabled = False
'        txtOfficeNo1.Enabled = False
'        txtOfficeNo2.Enabled = False
'        txtMobileNo1.Enabled = False
'        txtMobileNo2.Enabled = False
'        txtHomeAdd1.Enabled = False
'        txtHomeAdd2.Enabled = False
'        txtOfficeAdd1.Enabled = False
'        txtOfficeAdd2.Enabled = False
'        txtMobileAdd1.Enabled = False
'        txtMobileAdd2.Enabled = False
'    End If
 
    Set m_cust = Nothing
    Set M_objrs = Nothing
Exit Sub
'HELL:
   'MsgBox Err.Description
' Resume
 Set M_objrs = Nothing
Set m_cust = Nothing
End Sub

Private Sub autoremarks()
    Dim m_data As New CLS_FRMCUST_CC
    Set m_cust1 = m_data.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + Trim(lblCustId.text) + "'")
    
    listview1(1).ListItems.CLEAR
    While Not m_cust1.EOF
        'Set listitem = ListView1(1).ListItems.ADD(, , Left(m_cust1("TGL"), 4) & "/" & Mid(m_cust1("TGL"), 5, 2) & "/" & IIf(IsNull(m_cust1("TGL")), "", Mid(m_cust1("TGL"), 7, 2)) & " " & IIf(IsNull(m_cust1("TGL")), "", Mid(m_cust1("TGL"), 9, 2)) & ":" & Right(m_cust1("TGL"), 2))
         sTime_Hst = ""
         If IIf(IsNull(m_cust1("TGL")), "", m_cust1!TGL) <> "" Then
             'sTime_Hst = Format(IIf(IsNull(m_cust1("TGL")), "", m_cust1!TGL), "mm-dd-yyyy hh:mm:ss") & Format(IIf(IsNull(m_cust1("stop_time")), "", m_cust1!stop_time), " - hh:mm:ss")
            sTime_Hst = Format(IIf(IsNull(m_cust1("TGL")), "", m_cust1!TGL), "mm-dd-yyyy hh:mm:ss")
         End If
         Set ListItem = listview1(1).ListItems.ADD(, , sTime_Hst)
            ListItem.SubItems(1) = IIf(IsNull(m_cust1("HST")), "", m_cust1("HST"))
            ListItem.SubItems(2) = IIf(IsNull(m_cust1("user_log")), "", m_cust1("user_log"))
            ListItem.SubItems(3) = IIf(IsNull(m_cust1("AGENT")), "", m_cust1("AGENT"))
            ListItem.SubItems(4) = IIf(IsNull(m_cust1("KodeDs")), "", m_cust1("KodeDs"))
            ListItem.SubItems(5) = IIf(IsNull(m_cust1("statuscall")), "", m_cust1("statuscall"))
            ListItem.SubItems(6) = IIf(IsNull(m_cust1("ststelpwith")), "", m_cust1("ststelpwith"))
            ListItem.SubItems(7) = IIf(IsNull(m_cust1("id")), "", m_cust1("id"))
            ListItem.SubItems(8) = IIf(IsNull(m_cust1("unique_id")), "", m_cust1("unique_id"))
            'listitem.SubItems(4) = IIf(IsNull(m_cust1("f_cek")), "", m_cust1("f_cek"))
                    
                    
                    'Data Special 'jejaktian 18032016
                    If IIf(IsNull(m_cust1("f_special")), 0, m_cust1("f_special")) = "1" Then
                        For K = 1 To 7
                            ListItem.ListSubItems(K).ForeColor = vbRed
                            ListItem.ListSubItems(K).Bold = True
                        Next K
                    End If
            ' ------------------------------------------
    m_cust1.MoveNext
Wend

End Sub


Function ReplaceFirstInstance(SourceString, _
Searchstring, Replacestring)
  'Static StartLoc
  If StartLoc = 0 Then StartLoc = 1
  FoundLoc = InStr(StartLoc, SourceString, Searchstring) '*
  If FoundLoc <> 0 And FoundLoc < 2 Then
     ReplaceFirstInstance = Left(SourceString, FoundLoc - 1) & Replacestring & Right(SourceString, Len(SourceString) - (FoundLoc - 1) - Len(Searchstring))
     StartLoc = FoundLoc + Len(Replacestring)
  ElseIf FoundLoc > 1 Then
  
      ReplaceFirstInstance = Replacestring & "21" & SourceString

  Else
     StartLoc = 1

    ReplaceFirstInstance = SourceString
  End If
End Function

Function FindReplace(SourceString, Searchstring, Replacestring) As String
  tmpString1 = SourceString
 
      tmpString2 = tmpString1
      tmpString1 = ReplaceFirstInstance(tmpString1, _
                   Searchstring, Replacestring)
      
      FindReplace = tmpString1
End Function
Private Sub Isi_SendSMS()
'@@ 11-03-2011 di remarks, cznya udah tidak diapke
'Dim satu As String
'Dim dua As String
'Dim tiga As String
'Dim empat As String
'
'
'Dim RSsms_i As ADODB.Recordset
'Set RSsms_i = New ADODB.Recordset
'
'
'satu = FindReplace(TxtMobileno1.Text, "0", "+62")
'dua = FindReplace(TxtMobileno2.Text, "0", "+62")
'tiga = FindReplace(TxtMobileAdd1.Text, "0", "+62")
'empat = FindReplace(TxtMobileAdd2, "0", "+62")
'
'cmdsql_inbox = "Select receivingdatetime, sendernumber, textdecoded from inbox where (sendernumber='" + Trim$(satu) + "' or sendernumber='" + Trim$(dua) + "' or sendernumber='" + Trim$(tiga) + "' or sendernumber='" + Trim$(empat) + "') and processed='FALSE' "
'RSsms_i.Open cmdsql_inbox, M_OBJCONN1, adOpenDynamic, adLockOptimistic
'While Not RSsms_i.EOF
's = Format(RSsms_i!receivingdatetime, "yyyy-mm-dd hh:mm:ss")
't = Trim(RSsms_i!sendernumber)
'u = Replace(RSsms_i!textdecoded, "'", " ")
'
''u1 = Replace(KATAUBAH, "- -", "-")
'v = FindReplace(t, "+62", "0")
'
'
'
'            CMDSQL = "INSERT INTO receive_sms (tgl_terima, notelp, pesan) VALUES ('" & s & "',"
'            CMDSQL = CMDSQL + " '" + v + "',"
'            CMDSQL = CMDSQL + " '" + u + "')"
'            M_OBJCONN.Execute CMDSQL
'
'            cmdsql_update = "update inbox set processed='TRUE'  where (sendernumber='" + Trim$(satu) + "' or sendernumber='" + Trim$(dua) + "' or sendernumber='" + Trim$(tiga) + "' or sendernumber='" + Trim$(empat) + "')"
'            M_OBJCONN1.Execute cmdsql_update
'
'
'RSsms_i.MoveNext
'Wend
'
''=======================================
'Dim RSsms As ADODB.Recordset
'Set RSsms = New ADODB.Recordset
'Dim lst As listitem
'RSsms.CursorLocation = adUseClient
'If Left(TxtMobileno1, 1) <> "0" And TxtMobileno1 <> "" Then
'satua = "021" & TxtMobileno1
'Else
'satua = TxtMobileno1
'End If
'
'If Left(TxtMobileno2, 1) <> "0" And TxtMobileno2 <> "" Then
'duaa = "021" & TxtMobileno2
'Else
'duaa = TxtMobileno2
'End If
'
'If Left(TxtMobileAdd1, 1) <> "0" And TxtMobileAdd1 <> "" Then
'tigaa = "021" & TxtMobileAdd1
'Else
'tigaa = TxtMobileAdd1
'End If
'
'If Left(TxtMobileAdd2, 1) <> "0" And TxtMobileAdd2 <> "" Then
'empata = "021" & TxtMobileAdd2
'Else
'empata = TxtMobileAdd2
'End If
'
'
'CMDSQL = "Select a.*, b.custid from receive_sms a, mgm b where (a.notelp='" + satua + "' or a.notelp='" + duaa + "' or a.notelp='" + tigaa + "' or a.notelp='" + empata + "') and b.custid='" + lblCustId + "'"
'RSsms.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not RSsms.EOF
'    Set lst = LstSMS.ListItems.ADD(, , IIf(IsNull(RSsms("notelp")), "", RSsms("notelp")))
'         lst.SubItems(1) = lblNama
'         lst.SubItems(2) = IIf(IsNull(RSsms("custid")), "", RSsms("custid"))
'         lst.SubItems(3) = IIf(IsNull(RSsms("pesan")), "", RSsms("pesan"))
'         lst.SubItems(4) = IIf(IsNull(RSsms("tgl_terima")), "", RSsms("tgl_terima"))
'
'RSsms.MoveNext
'Wend
'Set RSsms = Nothing
'Text3 = LstSMS.ListItems.Count
'
''--------------------------------
'If Text4.Text <> "0" Then
'If Int(Text3) > Int(Text2) Then
'
'Dim RSsms_cek As ADODB.Recordset
'Set RSsms_cek = New ADODB.Recordset
'
'RSsms_cek.CursorLocation = adUseClient
'cmdsql_cek = "select * from receive_sms order by tgl_terima desc limit 1"
'RSsms_cek.Open cmdsql_cek, M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not RSsms_cek.EOF
'MsgBox "Anda mendapatkan satu SMS baru" & vbCrLf & "No Telepon : " & RSsms_cek("notelp") & vbCrLf & "Isi Pesan : " & Trim(RSsms_cek("pesan"))
'RSsms_cek.MoveNext
'Wend
'Set RSsms_cek = Nothing
'End If
'End If
'
'Text4.Text = "1"

End Sub
Private Sub Isi_SendSMS2()

Dim RSsms2 As ADODB.Recordset
'@@ 11-03-2011 Di remarks, udah tidak diapakai

'Set RSsms2 = New ADODB.Recordset
'Dim Lst2 As listitem
'RSsms2.CursorLocation = adUseClient
'CMDSQL = "Select * from sentitems where destinationnumber='" + TxtMobileno1 + "' or destinationnumber='" + TxtMobileno2 + "' or destinationnumber='" + TxtMobileAdd1 + "' or destinationnumber='" + TxtMobileAdd2 + "'"
'RSsms2.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic
'While Not RSsms2.EOF
'    Set Lst2 = LstSMS2.ListItems.ADD(, , IIf(IsNull(RSsms2("destinationnumber")), "", RSsms2("destinationnumber")))
'         Lst2.SubItems(1) = lblNama
'         Lst2.SubItems(2) = IIf(IsNull(RSsms2("textdecoded")), "", RSsms2("textdecoded"))
'         Lst2.SubItems(3) = IIf(IsNull(RSsms2("sendingdatetime")), "", RSsms2("sendingdatetime"))
'         Lst2.SubItems(4) = lblCustId
'         'Lst.SubItems(5) = IIf(IsNull(RSsms2("receivingdatetime")), "", RSsms2("receivingdatetime"))
''
'RSsms2.MoveNext
'Wend
'Set RSsms2 = Nothing
End Sub

Private Sub Isi_listScript()
'Mengisi Data di List LstScript
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open "select * from tblinformationlokasi", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_objrs.EOF
  Set ListItem = Lstscript.ListItems.ADD(, , M_objrs.Bookmark)
      ListItem.SubItems(1) = M_objrs("description")
      ListItem.SubItems(2) = M_objrs("direktori")
  M_objrs.MoveNext
Wend
Set M_objrs = Nothing
End Sub

Private Sub isi_datapayment()
Dim m_cust2 As New ADODB.Recordset
Dim NilaiAfterPay As Currency
Dim m_data As New CLS_FRMCUST_CC
Set m_cust2 = m_data.QUERY_HIST_PAID(M_OBJCONN, "a.custid = '" + lblCustId.text + "' ")
listview1(0).ListItems.CLEAR
While Not m_cust2.EOF
    Set ListItem = listview1(0).ListItems.ADD(, , IIf(IsNull(m_cust2("Paydate")), "", Format(m_cust2("Paydate"), "yyyy-mm-dd")))
        ListItem.SubItems(1) = IIf(IsNull(m_cust2("payment")), "0", Format(m_cust2("Payment"), "##,###"))
        ListItem.SubItems(2) = IIf(IsNull(m_cust2("AGENT")), "", m_cust2("AGENT"))
        ListItem.SubItems(3) = IIf(IsNull(m_cust2("FieldName")), "", m_cust2("FieldName"))
        ListItem.SubItems(4) = IIf(IsNull(m_cust2("Id")), "0", m_cust2("Id"))
        NilaiAfterPay = NilaiAfterPay + IIf(IsNull(m_cust2("payment")), "0", m_cust2("Payment"))
    m_cust2.MoveNext
Wend

Set m_cust2 = Nothing
TxtAfterPay.Value = NilaiAfterPay
txtSisaHutang.Value = Format(TxtPayment2.Value - TxtAfterPay.Value, "##,###")
End Sub
'Private Sub isi_datacustomer()
'    Dim M_objrs As ADODB.Recordset
'    Dim CMDSQL As String
'    Dim ListItem As ListItem
'
'    CMDSQL = "select * from tbl_address where custid='"
'    CMDSQL = CMDSQL + lblnocard.Caption + "' order by id"
'
'    Set M_objrs = New ADODB.Recordset
'    M_objrs.CursorLocation = adUseClient
'    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    lvaddress.ListItems.CLEAR
'While Not M_objrs.EOF
'    Set ListItem = lvaddress.ListItems.ADD(, , IIf(IsNull(M_objrs("custid")), "", M_objrs("custid")))
'        ListItem.SubItems(1) = IIf(IsNull(M_objrs("appid")), "", M_objrs("appid"))
'        ListItem.SubItems(2) = IIf(IsNull(M_objrs("adr_type")), "", M_objrs("adr_type"))
'        ListItem.SubItems(3) = IIf(IsNull(M_objrs("contact_address")), "", M_objrs("contact_address"))
'        ListItem.SubItems(4) = IIf(IsNull(M_objrs("address1")), "", M_objrs("address1"))
'        ListItem.SubItems(5) = IIf(IsNull(M_objrs("address2")), "", M_objrs("address2"))
'        ListItem.SubItems(6) = IIf(IsNull(M_objrs("address3")), "", M_objrs("address3"))
'        ListItem.SubItems(7) = IIf(IsNull(M_objrs("address4")), "", M_objrs("address4"))
'        ListItem.SubItems(8) = IIf(IsNull(M_objrs("city")), "", M_objrs("city"))
'        ListItem.SubItems(9) = IIf(IsNull(M_objrs("zipcode")), "", M_objrs("zipcode"))
'        ListItem.SubItems(10) = IIf(IsNull(M_objrs("contact1")), "", M_objrs("contact1"))
'        ListItem.SubItems(11) = IIf(IsNull(M_objrs("contact2")), "", M_objrs("contact2"))
'        ListItem.SubItems(12) = IIf(IsNull(M_objrs("mobileno")), "", M_objrs("mobileno"))
'        ListItem.SubItems(13) = IIf(IsNull(M_objrs("fax")), "", M_objrs("fax"))
'        ListItem.SubItems(14) = IIf(IsNull(M_objrs("email")), "", M_objrs("email"))
'        ListItem.SubItems(15) = IIf(IsNull(M_objrs("relationship_with")), "", M_objrs("relationship_with"))
'
'    M_objrs.MoveNext
'Wend
'
'Set M_objrs = Nothing
'End Sub
Private Sub CEK_UPDATE_PELANGGAN()

    Dim m_data As New CLS_FRMCUST_CC_MGM
    Dim m_Visit As New ClsVisit
    Dim pStatusHstLstCall As String
    Dim StatusPTP As String

    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_waktu As String
    Dim waktu As String
    Dim M_Objrs_Cek_Status As ADODB.Recordset
    Dim cmdsql_cari As String
    
    
    cmdsql_waktu = "select now() as waktu"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql_waktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    waktu = CDate(Format(M_objrs("waktu"), "hh:nn:ss"))
    Set M_objrs = Nothing


    Set M_update = New ADODB.Recordset
    M_update.CursorLocation = adUseServer
    M_update.Open "Select * from mgm where custid='" & lblCustId.text & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            
            
    '@@22102012 Buat nyatet Tanggal Paid Off
    If UCase(Trim(cboaccount.text)) = "PO-PAID OFF" Then
        'Cek apakah tanggal paid off masih kosong, jika ya update tanggal paid offnya
        If TanggalPaidOff = "" Or IsNull(TanggalPaidOff) = True Then
            M_update("tgl_paid_off") = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & waktu
        End If
    End If
    
    '20180727
        M_update("cmbbaseon") = cnull(Combo2.text)
            
    '@@02-05-2012, Buat Simpan kategori telepon
    If txtHomeAdd1.Value <> Empty Then
        M_update("stskathomeadd1") = CmbStsKatHome1.text
    End If
    If txtHomeAdd2.Value <> Empty Then
        M_update("stskathomeadd2") = CmbStsKatHome2.text
    End If
    If txtOfficeAdd1.Value <> Empty Then
        M_update("stskatofficeadd1") = CmbStsKatOffice1.text
    End If
    If txtOfficeAdd2.Value <> Empty Then
        M_update("stskatofficeadd2") = CmbStsKatOffice2.text
    End If
    If txtMobileAdd1.Value <> Empty Then
        M_update("stskathpadd1") = CmbStsKatHP1.text
    End If
    If txtMobileAdd2.Value <> Empty Then
        M_update("stskathpadd2") = CmbStsKatHP2.text
    End If
            
    '@@ 19/08/2011 Untuk telpon additional hanya boleh admin/supervisor (sebelumnya agent bisa, tapi sekrg ngga)
    If UCase(Left(MDIForm1.txtlevel.text, 5)) = "ADMIN" Or _
       UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or _
       UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
        M_update("telp_additional") = IIf(IsNull(TxtAdditional.Value), "", TxtAdditional.Value)
   End If
    
    M_update!maxad = tdbmaxad.Value
    M_update!minad = tdbminad.Value
    vrcekamont = Tdabamoint.Value
    
    '@@ 15 Juni 2011 Tambahkan SPV dan TeamLeader juga bisa save telepon
    If UCase(Left(MDIForm1.txtlevel.text, 5)) = "ADMIN" Or _
       UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or _
       UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
'        M_update("HOMENOADD1") = txtHomeAdd1.Value
'        M_update("HOMENOADD2") = txtHomeAdd2.Value
'        M_update("OFFICENOADD1") = txtOfficeAdd1.Value
'        M_update("OFFICENOADD2") = txtOfficeAdd2.Value
'        M_update("MOBILENOADD1") = txtMobileAdd1.Value
'        M_update("MOBILENOADD2") = txtMobileAdd2.Value
'        M_update!TxtPtpAddr = AddrNow.text
        'M_update!ec_name = TxtEC.Text
        'M_update!ec_telp = txtECno.Value
    Else
        If txtHomeAdd1A.Value = "" And txtHomeAdd1A.Visible = True Then
            M_update("HOMENOADD1") = txtHomeAdd1A.Value
        ElseIf txtHomeAdd1.Value <> "" And txtHomeAdd1.Visible = True Then
            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
            'M_update("HOMENOADD1") = txtHomeAdd1.Value
        End If
            
        If txtHomeAdd2A.Value = "" And txtHomeAdd2A.Visible = True Then
            M_update("HOMENOADD2") = txtHomeAdd2A.Value
        ElseIf txtHomeAdd2.Value <> "" And txtHomeAdd2.Visible = True Then
            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
            'M_update("HOMENOADD2") = txtHomeAdd2.Value
        ElseIf txtHomeAdd2.Value = "" And txtHomeAdd2.Visible = True Then
            M_update("HOMENOADD2") = txtHomeAdd2.Value
        End If
                
        If txtOfficeAdd1A.Value = "" And txtOfficeAdd1A.Visible = True Then
            M_update("OFFICENOADD1") = txtOfficeAdd1A.Value
        ElseIf txtOfficeAdd1.Value <> "" And txtOfficeAdd1.Visible = True Then
            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
            'M_update("OFFICENOADD1") = txtOfficeAdd1.Value
        ElseIf txtOfficeAdd1.Value = "" And txtOfficeAdd1.Visible = True Then
            M_update("OFFICENOADD1") = txtOfficeAdd1.Value
        End If
                
        If txtOfficeAdd2A.Value = "" And txtOfficeAdd2A.Visible = True Then
            M_update("OFFICENOADD2") = txtOfficeAdd2A.Value
        ElseIf txtOfficeAdd2.Value <> "" And txtOfficeAdd2.Visible = True Then
            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
            'M_update("OFFICENOADD2") = txtOfficeAdd2.Value
        ElseIf txtOfficeAdd2.Value = "" And txtOfficeAdd2.Visible = True Then
            M_update("OFFICENOADD2") = txtOfficeAdd2.Value
        End If
            
        If txtMobileAdd1A.Value = "" And txtMobileAdd1A.Visible = True Then
            M_update("MOBILENOADD1") = txtMobileAdd1A.Value
        ElseIf txtMobileAdd1.Value <> "" And txtMobileAdd1.Visible = True Then
            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
            'M_update("MOBILENOADD1") = txtMobileAdd1.Value
        ElseIf txtMobileAdd1.Value = "" And txtMobileAdd1.Visible = True Then
            M_update("MOBILENOADD1") = txtMobileAdd1.Value
        End If
            
        If txtMobileAdd2A.Value = "" And txtMobileAdd2A.Visible = True Then
            M_update("MOBILENOADD2") = txtMobileAdd2A.Value
        ElseIf txtMobileAdd2.Value <> "" And txtMobileAdd2.Visible = True Then
            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
            'M_update("MOBILENOADD2") = txtMobileAdd2.Value
        ElseIf txtMobileAdd2.Value = "" And txtMobileAdd2.Visible = True Then
            M_update("MOBILENOADD2") = txtMobileAdd2.Value
        End If
            
        M_update!TxtPtpAddr = AddrNow.text
        'M_update!ec_name = TxtEC.Text
        'M_update!ECAddr = txtECAdd.Text
                 
'        If txtECnoA.Value = "" And txtECnoA.Visible = True Then
'            M_update("ec_telp") = txtECnoA.Value
'        ElseIf txtECno.Value <> "" And txtECno.Visible = True Then
'            '@@ 15 Juni 2011, Agent tidak boleh update additional sendiri
'            'M_update!ec_telp = txtECno.Value
'        End If
    End If
        
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
'        If Len(txtECno.Value) > 2 Then
'            txtECno.ReadOnly = True
'        End If
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
    
    '@@121110 Tambahan nih buat nyatet history perubahan status account
    If (IsNull(M_update!tglcall)) = True Then
        tglcalllalu = ""
    Else
        tglcalllalu = CStr(M_update("tglcall"))
    End If
        
    '@@ 05-10-2011, Jika status account=PTP or POP maka catat via mana dia bayarnya
    If Trim(Mid(cboaccount, 1, 3)) = "POP" Or Trim(Mid(cboaccount, 1, 2)) = "BP" Then
        M_update!ptpvia = IIf(IsNull(CmbViaPtp.text), "", CmbViaPtp.text)
    End If
        
        
    M_update("TGLCALL") = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & waktu
    'sebelum f_cek diubah statusnya
    StatusPTP = IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new)

    Dim StatusAccCurrent As String  '@@ 121110 tambahan nih buat nyatet history f_cek_new
        
    If C_PTP.Value = vbChecked Then
        GoTo keptp
    End If
        
    If cboaccount.text <> "" Then
        pStatusLstCall = cboaccount.text
        M_update!f_cek_new = cboaccount.text
        'txtResult.Text = pStatusLstCall '@@15/01/2012 KOmponen Tidak Terpakai
        '@@121110 tambahan buat nyatet history f_cek_new
        StatusAccCurrent = cboaccount.text
    Else
keptp:
       
        Dim M_Objrs_PTPNew As New ADODB.Recordset
        Dim Cmdsql_PTPNew As String
        
        If C_PTP.Value Then
            M_update!ptpvia = IIf(IsNull(CmbViaPtp.text), "", CmbViaPtp.text)
            M_update!ptpdesc = cboaccount.text
            
            '//////////////////////// Awal Logika PTP 1 ////////////////////////////////////////////
            If vrcek = "BP-" And Len(TglPTPNew) > 0 And UCase(cboPTP.text) = "PTP-NEW" Then
                M_update!TglPTPNew = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
                                        
                    If TDBDate1.ValueIsNull Then
                        M_update!dateptpnew = Null
                    Else
                        M_update!dateptpnew = Format(TDBDate3.Value, "yyyy-mm-dd")
                        '@@25/01/2012, tambahkan tanggal tagih
                        M_update!tgl_tagih = Format(TdbTglTagih.Value, "yyyy-mm-dd")
                    End If
                                        
                     '@@ 06-01-2012 amountnew yang digunakan untuk amountptp ptp-new
                     'sekarang diambil dari tblnegoptp id terakhir
'                    If Tdabamoint.ValueIsNull Then
'                        M_update!amountnew = 0
'                    Else
'                        M_update!amountnew = Tdabamoint.Value
'                    End If
                   
                    '@@ 16 APRIL 2012, bukan ID terakhir, tetapi inputdate terakhir
                    Cmdsql_PTPNew = "select * from tblnegoptp where custid='"
                    Cmdsql_PTPNew = Cmdsql_PTPNew + lblCustId.text + "' order by inputdate desc limit 1"
                    
                    
                    Set M_Objrs_PTPNew = New ADODB.Recordset
                    M_Objrs_PTPNew.CursorLocation = adUseClient
                    M_Objrs_PTPNew.Open Cmdsql_PTPNew, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    
                    M_update!AmountNew = M_Objrs_PTPNew("promisepay")
                    Set M_Objrs_PTPNew = Nothing
            Else
                If cboPTP.text = "PTP-NEW" Then
                    If vrcek <> "PTP-NE" Then
                    
                        If UCase(cboPTP.text) = "PTP-NEW" And listview1(0).ListItems.Count = 0 Then
                            M_update!TglPTPNew = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
                            If TDBDate1.ValueIsNull Then
                                M_update!dateptpnew = Null
                            Else
                                M_update!dateptpnew = Format(TDBDate3.Value, "yyyy-mm-dd")
                                '@@25/01/2012 , Tambahkan untuk tanggal tagih
                                M_update!tgl_tagih = Format(TdbTglTagih.Value, "yyyy-mm-dd")
                                
                            End If
                                        
                             '@@ 06-01-2012 amountnew yang digunakan untuk amountptp ptp-new
                            'sekarang diambil dari tblnegoptp id terakhir
'                            If Tdabamoint.ValueIsNull Then
'                                M_update!amountnew = 0
'                            Else
'                                M_update!amountnew = Tdabamoint.Value
'                            End If
                            
                            Cmdsql_PTPNew = "select * from tblnegoptp where custid='"
                            Cmdsql_PTPNew = Cmdsql_PTPNew + lblCustId.text + "' order by id desc limit 1"
                
                            Set M_Objrs_PTPNew = New ADODB.Recordset
                            M_Objrs_PTPNew.CursorLocation = adUseClient
                            M_Objrs_PTPNew.Open Cmdsql_PTPNew, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                            
                            If M_Objrs_PTPNew.RecordCount = 0 Then
                                M_update!AmountNew = 0
                            Else
                                M_update!AmountNew = M_Objrs_PTPNew("promisepay")
                            End If
                            
                            'M_update!amountnew = IIf(IsNull(M_Objrs_PTPNew("promisepay")), "0", M_Objrs_PTPNew("promisepay"))
                            Set M_Objrs_PTPNew = Nothing
                            
                        End If
                                                    
                    End If
                End If
            End If
            '//////////////////////// Akhir Logika PTP 1 ////////////////////////////////////////////
            
            '//////////////////////// Awal Logika PTP 2 ////////////////////////////////////////////
            If vrcek = "BP-" And Len(TglPTPNew) > 0 And Left(UCase(cboPTP.text), 3) = "PTP" Then
                M_update!tglallptp = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
            Else
                If Left(cboPTP.text, 3) = "PTP" Then
                    If Left(vrcek, 6) <> Left(cboPTP.text, 6) Then
                        M_update!tglallptp = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
                    ElseIf vrnewdate <> TDBDate3.text Then
                        M_update!tglallptp = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
                    End If
                End If
            End If
            '//////////////////////// Akhir Logika PTP 2 ////////////////////////////////////////////
    
            pStatusLstCall = cboPTP.text
            'txtResult.Text = pStatusLstCall '@@15/01/2012 Komponen Tak Terpakai
            'txtResultDesc.Text = pStatusLstCalldesc '@@15/01/2012 Komponen Tak Terpakai
            M_update("RECSTATUS") = "P"
            M_update!f_cek_new = Left(cboPTP.text, 6)
                                
            '@@121110 tambahan buat nyatet history f_cek_new
            StatusAccCurrent = Left(cboPTP.text, 6)
            
        Else
        End If
    End If
        
    If C_Payment.Value Then
        If StatusPTP <> Empty Then
            If StatusPTP = M_update!f_cek_new Then
            Else
                M_update!TGLINCOMING = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
            End If
        End If
        M_update!ttlptp = txtPayment.Value
        'M_update!amountptp = Tdabamoint.Value
        '@@ 05-01-2012,tdabamoint sudah tidak dipakai, langsung pakai txtpayment
        M_update!amountptp = txtPayment.Value
        M_update!discpersen = cmbDiscount.text
        M_update!Tenor = txttenor.Value
        M_update!dateptp = Format(TDBDate3.Value, "yyyy/mm/dd")
        '@@25/01/2012, Update tanggal tagih
        If TdbTglTagih.ValueIsNull = False Then
         M_update!tgl_tagih = Format(TdbTglTagih.Value, "yyyy-mm-dd")
       End If
    Else
        M_update!ttlptp = 0
        M_update!discpersen = 0
    End If
               
    If Trim(UCase(IIf(IsNull(M_update("kethslkerja_new")), "", M_update("kethslkerja_new")))) = Trim(UCase(pStatusLstCall)) Then
        TGLSTATUS = IIf(IsNull(M_update("TGLSTATUS")), "", Format(M_update("TGLSTATUS"), "yyyy/mm/dd"))
    Else
        M_update("kethslkerja_new") = pStatusLstCall
        M_update("TGLSTATUS") = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
        TGLSTATUS = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")
    End If
        M_update!stscallwith = cbolastcall.text
        M_update("kethslkerja_new") = pStatusLstCall
        pStatusHstLstCall = IIf(IsNull(M_update("kethslkerja_new")), "", M_update("kethslkerja_new"))
        M_update("kethslkerjadesc_new") = cboaccount.text
        M_update("REMARKS") = Replace(txtremarks.text, "'", "`")
    If Not (cmbDateSch.ValueIsNull) Then
        M_update!NEXTACTDATE = Format(cmbDateSch.Value, "yyyy/mm/dd") & " " & Format(cmbTimeSch.Value, "hh:nn")
    End If
        
    M_update("Statuscall") = Trim(cboaccount.text)
    M_update("homeno") = Trim(txtHomeNo1.text)
    M_update("officeno") = Trim(txtOfficeNo1.text)
    M_update("mobileno") = Trim(txtMobileNo1.text)
    M_update("stscallcust") = Trim(Combo1.text)
    M_update.update
        
      
    '@@ 25-Januari-2012 Tulis Result PTPnya
    'RANDY NAMBAHIN CEK_AKTIF
    If chk_aktif.Value = 1 Then
'        If Trim(Left(cboaccount.Text, 3)) = "PTP" And UCase(mdiform1.txtlevel.text) = "AGENT" Then
'            FrmResultPTP.txtStatusAcc.Text = Trim(cboPTP.Text)
'            FrmResultPTP.Show vbModal
'        Else
            'Kalo yang statusnya POP tampilkan juga result ptp nya
            '@@ 29-06-2012
            'If LstPayment.ListItems.Count > 0 Then
        cmdsql_cari = "select f_cek_new from mgm where custid='"
        cmdsql_cari = cmdsql_cari + CStr(lblCustId.text) + "'"
        Set M_Objrs_Cek_Status = New ADODB.Recordset
        M_Objrs_Cek_Status.CursorLocation = adUseClient
        M_Objrs_Cek_Status.Open cmdsql_cari, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
               
        'If UCase(mdiform1.txtlevel.text) = "AGENT" Then
        If Trim(M_Objrs_Cek_Status("f_cek_new")) = "POP" Or _
           Trim(Left(M_Objrs_Cek_Status("f_cek_new"), 3)) = "PTP" Or _
           Trim(Left(M_Objrs_Cek_Status("f_cek_new"), 2)) = "BP" Then
             FrmResultPTP.txtStatusAcc = Trim(M_Objrs_Cek_Status("f_cek_new"))
             'FrmResultPTP.Show vbModal
        End If
        'End If
            'End If
        Set M_Objrs_Cek_Status = Nothing
        'End If
    Else
        cmdsql_cari = "select f_cek_new from mgm where custid='"
        cmdsql_cari = cmdsql_cari + CStr(lblCustId.text) + "'"
        Set M_Objrs_Cek_Status = New ADODB.Recordset
        M_Objrs_Cek_Status.CursorLocation = adUseClient
        M_Objrs_Cek_Status.Open cmdsql_cari, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
               
        If UCase(MDIForm1.txtlevel.text) <> "AGENT" Then
            If Trim(M_Objrs_Cek_Status("f_cek_new")) = "POP" Or _
               Trim(Left(M_Objrs_Cek_Status("f_cek_new"), 3)) = "PTP" Or _
               Trim(Left(M_Objrs_Cek_Status("f_cek_new"), 2)) = "BP" Then
                 'FrmResultPTP.txtStatusAcc = Trim(M_Objrs_Cek_Status("f_cek_new"))
                 'FrmResultPTP.Show vbModal
            End If
        End If
        Set M_Objrs_Cek_Status = Nothing
    End If
    
    
    
'JEJAKTIAN26022016
    '@@21 May 2012,Penulisan Remarks dipecah per 90 Karakter
    Dim BanyakBaris As Integer
    Dim AW As Integer
    Dim AwalRemarks As String
    Dim pesan, Unik As String
    If cboaccount.text <> "" Then
        If txtremarks.text <> Empty Then
'            BanyakBaris = Ceiling(Val(Len(TxtRemarks.Text)) / 87)
'            Unik = Format(Now, "ddmmyyyyhhmmss")
'
'            'Bikin Baris KOsong....
'            M_DATA.ADD_HISTORY lblCustId.text, MDIForm1.TDBDate1.Text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", "------------------------------------------------------------------------------------", CStr(pStatusLstCall), "", "", CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Combo1.Text, Combo1.Text, CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Left(cboaccount.Text, 3), cbolastcall.Text, mdiform1.txtusername.text, Unik, BanyakBaris + 1
'            For AW = 1 To BanyakBaris
'                'AwalRemarks = (87 * AW) - 87
'                AwalRemarks = (87 * ((BanyakBaris + 1) - AW)) - 87
'                pesan = "(" & BanyakBaris + 1 - AW & "/" & BanyakBaris & ") " & Mid(TxtRemarks.Text, AwalRemarks + 1, 87)
'                M_DATA.ADD_HISTORY lblCustId.text, MDIForm1.TDBDate1.Text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", IIf(IsNull(pesan), "", pesan), CStr(pStatusLstCall), "", "", CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Combo1.Text, Combo1.Text, CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Left(cboaccount.Text, 3), cbolastcall.Text, mdiform1.txtusername.text, Unik, BanyakBaris + 1 - AW
'            Next AW

            'M_DATA.ADD_HISTORY lblCustId.text, MDIForm1.TDBDate1.Text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", txtremarks.Text, CStr(pStatusLstCall), "", "", CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Combo1.Text, Combo1.Text, CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Left(cboaccount.Text, 3), cbolastcall.Text, mdiform1.txtusername.text, "", "0"
            ' UPDATE 02 07 2014 'JEJAKTIANREMARK 'openremark
            m_data.ADD_HISTORY lblCustId.text, MDIForm1.TDBDate1.text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", txtremarks.text, CStr(pStatusLstCall), "", "", CStr(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new)), Combo1.text, cboaccount.text, CStr(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new)), cboaccount.text, cbolastcall.text, MDIForm1.TxtUsername.text, "", "0", Format(lbltime_save.Caption, "yyyy-mm-dd hh:mm:ss"), Format(lblstop_time.Caption, "yyyy-mm-dd hh:mm:ss"), lblCustId.text, MDIForm1.txt_unique_id.text, kat_aktif_telp
            End If
    End If
    If C_PTP.Value = vbChecked Then
        GoTo BRO
    End If
BRO:
    'm_data.ADD_HISTORY lblCustId.Text, MDIForm1.TDBDate1.Text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", txtremarks.Text, CStr(pStatusLstCall), "", "", CStr(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new)), Combo1.Text, cboaccount.Text, CStr(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new)), cboaccount.Text, cbolastcall.Text, MDIForm1.txtusername.Text, "", "0", Format(lbltime_save.Caption, "yyyy-mm-dd hh:mm:ss"), Format(lblstop_time.Caption, "yyyy-mm-dd hh:mm:ss"), lblCustId.Text, MDIForm1.txt_unique_id.Text
    If C_PTP.Value = 1 Then
        If txtremarks.text <> Empty Then
             BanyakBaris = Ceiling(Val(Len(txtremarks.text)) / 87)
            Unik = Format(Now, "ddmmyyyyhhmmss")
         End If
    End If

    If Len(TDBTot_payment) > 2 Then
        m_data.ADD_tbllunas M_OBJCONN, lblCustId.text, Format(TdbLunas.Value, "yyyy/mm/dd"), CCur(TDBTot_payment.Value), VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), TxtFieldName.text, ""
    Else
        On Error Resume Next
    End If
    '------------>> simpan ke table Visit <<--------------------
    If Option8(0).Value Then
        m_Visit.ADD_RequestVisit M_OBJCONN, lblCustId.text, M_update!f_cek_new, Text1.text, Format(TDBDate1.Value, "yyyy-mm-dd"), TXtDetails.text, TDBNumber1.Value, TxtAddress.text, Trim(UCase(VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11)))
    Else
        On Error Resume Next
    End If

    MsgBox "Data Sudah Tersimpan", vbInformation + vbOKOnly, "Sukses"
    
    kontak = False
    Set M_update = Nothing

    If shedulePTP_Show = True Then
    Else
        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(7) = txtremarks.text
        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(8) = pStatusLstCall
        If cboaccount <> "" Then
            VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(10) = Left(cboaccount, 3)
        ElseIf cboPTP <> "" Then
            VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(10) = Left(cboPTP, 6)
        End If
    End If
    pStatusLstCall = ""
    pStatusHstLstCall = ""
    txtremarks.text = Empty


    Set m_data = Nothing
    Exit Sub
    Resume
End Sub

Private Sub HEADER_HISTORY()
    listview1(1).ColumnHeaders.ADD 1, , "Tanggal(mm-dd-yyyy)", 10 * TXT
    listview1(1).ColumnHeaders.ADD 2, , "History", 80 * TXT
    listview1(1).ColumnHeaders.ADD 3, , "User Log", 10 * TXT
    listview1(1).ColumnHeaders.ADD 4, , "Handle By", 10 * TXT
    listview1(1).ColumnHeaders.ADD 5, , "Sts Account", 10 * TXT
    listview1(1).ColumnHeaders.ADD 6, , "Sts Call", 10 * TXT
    listview1(1).ColumnHeaders.ADD 7, , "Sts Telp With", 25 * TXT
    listview1(1).ColumnHeaders.ADD 8, , "Id", 25 * TXT
    listview1(1).ColumnHeaders.ADD 9, , "Unique ID", 0 * TXT
End Sub
Private Sub HEADER_RequestVisit()
    LstVisit.ColumnHeaders.ADD 1, , "RequestDate", 10 * TXT
    LstVisit.ColumnHeaders.ADD 2, , "VisitDate", 10 * TXT
    LstVisit.ColumnHeaders.ADD 3, , "VisitNo", 10 * TXT
    LstVisit.ColumnHeaders.ADD 4, , "Details", 20 * TXT
    LstVisit.ColumnHeaders.ADD 5, , "Hasil Visit", 20 * TXT
    LstVisit.ColumnHeaders.ADD 6, , "VisitKe", 2 * TXT
    LstVisit.ColumnHeaders.ADD 7, , "ID", 1 * TXT
    LstVisit.ColumnHeaders.ADD 8, , "Status", 1 * TXT
    End Sub
    
Private Sub HEADER_HISTORY_PAID()
    listview1(0).ColumnHeaders.ADD 1, , "PayDate", 15 * TXT
    listview1(0).ColumnHeaders.ADD 2, , "Payment", 30 * TXT
    listview1(0).ColumnHeaders.ADD 3, , "Agent", 10 * TXT
    listview1(0).ColumnHeaders.ADD 4, , "FieldName", 30 * TXT
    listview1(0).ColumnHeaders.ADD 5, , "Id", 30 * TXT
End Sub
Private Sub HEADER_Detail_Customer()
    lvaddress.ColumnHeaders.ADD 1, , "Customer ID", 15 * TXT
    lvaddress.ColumnHeaders.ADD 2, , "App ID", 30 * TXT
    lvaddress.ColumnHeaders.ADD 3, , "Address Type", 10 * TXT
    lvaddress.ColumnHeaders.ADD 4, , "Contact Address", 10 * TXT
    lvaddress.ColumnHeaders.ADD 5, , "Address 1", 20 * TXT
    lvaddress.ColumnHeaders.ADD 6, , "Address 2", 20 * TXT
    lvaddress.ColumnHeaders.ADD 7, , "Address 3", 20 * TXT
    lvaddress.ColumnHeaders.ADD 8, , "Address 4", 20 * TXT
    lvaddress.ColumnHeaders.ADD 9, , "City", 10 * TXT
    lvaddress.ColumnHeaders.ADD 10, , "Zipcode", 10 * TXT
    lvaddress.ColumnHeaders.ADD 11, , "Contact 1", 15 * TXT
    lvaddress.ColumnHeaders.ADD 12, , "Contact 2", 15 * TXT
    lvaddress.ColumnHeaders.ADD 13, , "Mobile No", 15 * TXT
    lvaddress.ColumnHeaders.ADD 14, , "Fax", 15 * TXT
    lvaddress.ColumnHeaders.ADD 15, , "Email", 15 * TXT
    lvaddress.ColumnHeaders.ADD 16, , "Relationship With", 15 * TXT
End Sub

Private Sub UPDATE_STATUS_CALL_SEBELUM()
    Dim status_call_sebelum As String
    Dim M_Objrs_Cek_Status_Call  As ADODB.Recordset
    Dim sQuery As String
    
    status_call_sebelum = ""
    
    'AMBIL DULU STATUS CALL TERAKHIR
    sQuery = " SELECT f_cek_new from mgm where custid = '" & Trim(lblCustId.text) & "' "
    Set M_Objrs_Cek_Status_Call = New ADODB.Recordset
        M_Objrs_Cek_Status_Call.CursorLocation = adUseClient
        M_Objrs_Cek_Status_Call.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    status_call_sebelum = IIf(IsNull(M_Objrs_Cek_Status_Call("f_cek_new")), "", M_Objrs_Cek_Status_Call("f_cek_new"))
    
    'UPDATE STATUS_CALL_SEBELUM
    sQuery = "update mgm set status_call_sebelum=('"
    sQuery = sQuery & status_call_sebelum & "') where custid='"
    sQuery = sQuery & Trim(lblCustId.text) & "'"
    M_OBJCONN.Execute sQuery
        
End Sub

Private Function CEK_DATA_VALID() As Boolean
    Dim m_msgbox As Variant
    Dim CMDSQL As String
    Dim M_Objrs_Cek_PTP  As ADODB.Recordset
    Dim m_objrs_reserve As ADODB.Recordset
    Dim TotalPtp As Double
    Dim pesan As String
    
    If TDBTot_payment > 2 Then
        CEK_DATA_VALID = True
        Exit Function
    Else


'TIANINDIUM20181212
        '@@04-06-2012 Cek dulu apakah data ptp? kalo iya harus cek cpa
'        If C_PTP.Value Then
'            CMDSQL = "select * from tblcpa where vcustid='"
'            CMDSQL = CMDSQL + Trim(lblCustId.text) + "' order by nid desc limit 1 "
'            Set M_objrs = New ADODB.Recordset
'            M_objrs.CursorLocation = adUseClient
'            M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'            If M_objrs.RecordCount = 0 Then
'
'                MsgBox "Account anda PTP.Untuk membuat status account PTP, harus dibuat terlebih dahulu CPA nya! Tekan Send PTP Untuk membuat CPA dan PTP!", vbOKOnly + vbInformation, "Informasi"
'                MsgBox "Data PTP gagal disimpan!", vbOKOnly + vbExclamation, "Peringatan"
'
'                Set M_objrs = Nothing
'                CEK_DATA_VALID = False
'                Exit Function
'            End If
'
'        End If

        
        '@@ 16 May 2012, Cek jika status PTP-POP atau PTP NEW tapi data di tblnegoptp tidak ada
        'Ubah otomastis ke BP
        Dim M_Objrs_NegoPTP As ADODB.Recordset
        Dim WA As String
        If cboPTP.text = "PTP-POP" Then
            'Cek Apakah data di tabelnegoptp ada?
            CMDSQL = "select * from tblnegoptp where custid='"
            CMDSQL = CMDSQL + CStr(lblCustId.text) + "' order by promisedate desc limit 1 "
            Set M_Objrs_NegoPTP = New ADODB.Recordset
            M_Objrs_NegoPTP.CursorLocation = adUseClient
            M_Objrs_NegoPTP.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            'Ini Jika Tidak ditemukan data di tabel tblnegoptp, maka ubah status account menjadi BP-POP
            'Agar data bisa di dave
            If M_Objrs_NegoPTP.RecordCount = 0 Then
                WA = MsgBox("Benarkah account ini PTP? Jika benar, silahkan sempurnakan datanya, List PTP Jatuh Tempo anda masih kosong!. TEKAN YES jika anda ingin mengisi data PTP atau TEKAN NO jika data ini BUKAN PTP!", vbYesNo + vbQuestion, "Konfirmasi")
                If WA = vbYes Then
                    MsgBox "Sempurnakan terlebih dahulu Form PTP anda. Kemudian lakukan penyimpanan ulang remarks anda!", vbOKOnly + vbInformation, "Informasi"
                    CEK_DATA_VALID = False
                    Exit Function
                End If
                CMDSQL = "update mgm set tglstatus= now() ,F_CEK='BP-',LASTSTATUS='BP-POP',"
                CMDSQL = CMDSQL + "KETHSLKERJA_NEW='BP-POP',F_CEK_NEW='BP-',"
                CMDSQL = CMDSQL + "KETHSLKERJADESC_NEW='BP-BROKEN PROMISE',"
                CMDSQL = CMDSQL + "KETHSLKERJA='BP-PTP POP BROKEN PROMISE',"
                CMDSQL = CMDSQL + "REMARKS = 'BP-POP BROKEN PROMISE @',"
                CMDSQL = CMDSQL + "RECSTATUS='C',OTO='Y' where f_cek_NEW like 'PTP-PO' and custid='"
                CMDSQL = CMDSQL + CStr(lblCustId.text) + "'"
                M_OBJCONN.Execute CMDSQL
                C_PTP.Value = vbUnchecked
                cboaccount.text = "BP-POP"
                C_Payment.Value = vbUnchecked
            End If
            Set M_Objrs_NegoPTP = Nothing
        End If
                
                
        If cboPTP.text = "PTP-NEW" Then
            'Cek Apakah data di tabelnegoptp ada?
            CMDSQL = "select * from tblnegoptp where custid='"
            CMDSQL = CMDSQL + CStr(lblCustId.text) + "' order by promisedate desc limit 1 "
            Set M_Objrs_NegoPTP = New ADODB.Recordset
            M_Objrs_NegoPTP.CursorLocation = adUseClient
            M_Objrs_NegoPTP.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            'Ini Jika Tidak ditemukan data di tabel tblnegoptp, maka ubah status account menjadi BP-POP
            'Agar data bisa di dave
            If M_Objrs_NegoPTP.RecordCount = 0 Then
                WA = MsgBox("Benarkah account ini PTP? Jika benar, tolong sempurnakan datanya, List PTP Jatuh Tempo and masih kosong!. TEKAN YES jika anda ingin mengisi data PTP atau TEKAN NO jika data ini BUKAN PTP!", vbYesNo + vbQuestion, "Konfirmasi")
                If WA = vbYes Then
                    MsgBox "Sempurnakan terlebih dahulu Form PTP anda. Kemudian lakukan penyimpanan ulang remarks anda!", vbOKOnly + vbInformation, "Informasi"
                    CEK_DATA_VALID = False
                    Exit Function
                End If
                CMDSQL = "update mgm set tglstatus= now() ,F_CEK='BP-',LASTSTATUS='BP-NEW',"
                CMDSQL = CMDSQL + "KETHSLKERJA_NEW='BP-NEW',F_CEK_NEW='BP-',"
                CMDSQL = CMDSQL + "KETHSLKERJADESC_NEW='BP-BROKEN PROMISE',"
                CMDSQL = CMDSQL + "KETHSLKERJA='BP-PTP NEW BROKEN PROMISE',"
                CMDSQL = CMDSQL + "REMARKS = 'BP-NEW BROKEN PROMISE @',"
                CMDSQL = CMDSQL + "RECSTATUS='C',OTO='Y' where f_cek_NEW like 'PTP-NE' and custid='"
                CMDSQL = CMDSQL + CStr(lblCustId.text) + "'"
                M_OBJCONN.Execute CMDSQL
                C_PTP.Value = vbUnchecked
                cboaccount.text = "BP-NEW"
                C_Payment.Value = vbUnchecked
            End If
            Set M_Objrs_NegoPTP = Nothing
        End If
                
        
        If Left(cmbContacted, 3) = "PTP" And LstPayment.ListItems.Count = 0 Then
            MsgBox "PTP harus buat Nego PTP di tabel yang hijau !!!", vbInformation + vbOKOnly, "TINS"
            CEK_DATA_VALID = False
            Exit Function
        End If
        
        
        If cboaccount.text = "" And C_PTP.Value = vbUnchecked Then
            MsgBox "Status Account harus diisi!", vbInformation + vbOKOnly, "TINS"
            CEK_DATA_VALID = False
            Exit Function
        End If
        
    
        If C_PTP.Value = vbChecked Then
              '@@ 11 Januari 2012 dinonaktifkan, tidak menggunakan tdabmoint
        '       If Val(vrcekamont) <> Tdabamoint.Value And bcekptp = False Then
        '            MsgBox "anda harus klik tambah di Call Activity untuk Negotiation", vbInformation + vbOKOnly, "TINS"
        '
        '            CEK_DATA_VALID = False
        '            Exit Function
        '        End If
        
            '@@ 05-10-2011, Jika melakukan PTP maka combo via ptp harus diisi
            If CmbViaPtp.text = "" Then
                MsgBox "Combo Via tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
                CEK_DATA_VALID = False
                Exit Function
            End If
            
            'Tambahan, Jika Status data PTP, hitung tanggal tagih
            If TDBDate3.ValueIsNull Then
                MsgBox "Anda belum menentukan tanggal effective pembayaran!", vbOKOnly + vbInformation, "Informasi"
                CEK_DATA_VALID = False
                Exit Function
            End If
            
            Call CariTanggalTagih
            
        End If
    
        If C_Payment.Value = 1 Then
            CmbBaseOn.text = "TOTAL AMOUNT"
            If TDBDate3.ValueIsNull Then
                CEK_DATA_VALID = False
                MsgBox "Tanggal PTP Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
                Exit Function
            End If
        End If
                   
        If C_PTP.Value = 1 Then
            If cboPTP.text = Empty Then
                CEK_DATA_VALID = False
                MsgBox "Description PTP Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
                Exit Function
                SSTab1.Tab = 3
            End If
        End If

       
        If txtremarks.text = "" Then
            CEK_DATA_VALID = False
            MsgBox "Remarks Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
            Exit Function
        End If
 
        If ADD_CUST = True Then
        Else
            If cboaccount.text <> "" Then
                Dim StatusRemarks As String
'                '@@ 16 Agustus 2011, pola remarks diubah
'                StatusRemarks = Combo1.Text & "-"
'                'StatusRemarks = StatusRemarks & cbolastcall.Text & "-"
'                '@@04-05-2012  Cbolastcall disingkat di statusspeakwith
'                StatusRemarks = StatusRemarks & StatusSpeakWith & "-"
'                StatusRemarks = StatusRemarks & "[" & cboaccount.Text & "] - "
'                StatusRemarks = StatusRemarks & TxtTelpKe.Text
'                '@@03-05-2012 Tambahan Status Telepon
'                StatusRemarks = StatusRemarks & "-" & IIf(IsNull(KelompokKategoriTlp), "", KelompokKategoriTlp)
'                TxtRemarks.Text = StatusRemarks & " // " & TxtRemarks.Text
                 
                '@@10052012 Mengubah Pola Remarks
                'StatusRemarks = IIf(IsNull(KelompokKategoriTlp), "", KelompokKategoriTlp) & "/"
                'StatusRemarks = StatusRemarks & IIf(Combo1.Text = "Receive", "RCVD", "NRCV") & "/"
                StatusRemarks = StatusRemarks & cboaccount.text & "/"
                'jejaktian11042016
                'StatusRemarks = StatusRemarks & "Exp Date " & lbl_expdate.Caption & "/ "
                '============================
                StatusRemarks = StatusRemarks & kat_aktif_telp & ": " & txtPhone.text & "/"
                txtremarks.text = StatusRemarks & txtremarks.text
                
                
             ElseIf cboPTP.text <> "" Then
'                '@@ 16 Agustus 2011, pola remarks diubah
'                StatusRemarks = Combo1.Text & "-"
'                'StatusRemarks = StatusRemarks & cbolastcall.Text & "-"
'                '@@04-05-2012  Cbolastcall disingkat di statusspeakwith
'                StatusRemarks = StatusRemarks & StatusSpeakWith & "-"
'                StatusRemarks = StatusRemarks & " PTP Via:" & CmbViaPtp.Text + "-"
'                StatusRemarks = StatusRemarks & "[ " & cboPTP.Text & "-"
'                StatusRemarks = StatusRemarks & "AmountPTP:" & txtPayment.Text & "- "
'                StatusRemarks = StatusRemarks & "DatePTP:" & TDBDate3.Value & " ] -"
'                StatusRemarks = StatusRemarks & TxtTelpKe.Text
'                '@@03-05-2012 Tambahan Status Telepon
'                StatusRemarks = StatusRemarks & "-" & IIf(IsNull(KelompokKategoriTlp), "", KelompokKategoriTlp)
'                TxtRemarks.Text = StatusRemarks & " // " & TxtRemarks.Text
                
                '@@10052012 Menubah Pola Remarks
                StatusRemarks = IIf(IsNull(KelompokKategoriTlp), "", KelompokKategoriTlp) & "/"
                StatusRemarks = StatusRemarks & IIf(Combo1.text = "Receive", "RCVD", "NRCV") & "/"
                StatusRemarks = StatusRemarks & StatusSpeakWith & "/"
                StatusRemarks = StatusRemarks & cboPTP.text & "/"
                StatusRemarks = StatusRemarks & "PTP Via " & CmbViaPtp.text & "/"
                StatusRemarks = StatusRemarks & "Amount PTP " & txtPayment.text & "/"
                'jejaktian11042016
                StatusRemarks = StatusRemarks & "Exp Date " & lbl_expdate.Caption & "/ "
                '============================
                StatusRemarks = StatusRemarks & "Date PTP " & TDBDate3.Value & ": " & kat_aktif_telp
                txtremarks.text = StatusRemarks & txtremarks.text
                
            
            End If
            
            If stscall = True Then
                If C_PTP.Value = vbUnchecked And cboaccount.text = "" Then
                    CEK_DATA_VALID = False
                    MsgBox "Status Account Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
                    SSTab1.Tab = 3
                    Exit Function
                End If
            End If
        End If
    End If

        
'        If cmbDiscount.Text = "" Then
'            MsgBox "Diskon harus diisi", vbInformation + vbOKOnly, "TINS"
'            CEK_DATA_VALID = False
'            Exit Function
'        End If
      
    '@@23022012 Cek dulu apakah status data BP atau POP
    'JIka BP atau POP lewat saja pengecekan PTP
    If Mid(cboaccount.text, 1, 3) = "BP-" Or Mid(cboaccount.text, 1, 3) = "POP" Then
        GoTo Lanjut_1
    End If
      
    pesan = "Informasi: " & vbCrLf
    pesan = pesan & "Anda hanya dapat membuat status PTP " & vbCrLf
    pesan = pesan & "jika CPA untuk account tersebut telah dibuat! " & vbCrLf
    pesan = pesan & "Mintalah kepada TL anda untuk membuat CPA!" & vbCrLf & vbCrLf
    pesan = pesan & "Jika anda mengalami kesulitan untuk menyimpan data remarks anda, kemungkinan adalah: " & vbCrLf
    pesan = pesan & "1. Ada data di list PTP Jatuh Tempo, tetapi Form PTP kosonng. Seperti Total Amount Deal dan Date Payment Effective." & vbCrLf
    pesan = pesan & "2. Ada data di Form PTP, tetapi data di list PTP Jatuh tempo kosong! " & vbCrLf
    pesan = pesan & "3. Jumlah data di list RESERVED PTP tidak sama dengan Tenor di Form PTP!" & vbCrLf
    pesan = pesan & "4. Ada data di list Reserved PTP, tetapi data di Form PTP masih kosong!" & vbCrLf
    pesan = pesan & "5. Date Payment Effective harus sama dengan tanggal di list PTP jatuh tempo!"
      
      
    '@@ 07-02-2012, cek data negoptp jika status data PTP
    If MDIForm1.txtlevel.text = "Agent" Then
        If C_PTP.Value = 1 Then
                    
            'Cek Nilai Payment
            If txtPayment.Value = "0" Or txtPayment.ValueIsNull = True Then
                MsgBox "Anda mencentang data PTP, Total Amount Deal tidak boleh kosong!", vbOKOnly + vbExclamation, "Informasi"
                MsgBox pesan, vbOKOnly + vbInformation, "Informasi"
                CEK_DATA_VALID = False
                Exit Function
            End If
            
            'Cek Nilai Date Payment Effective
            If TDBDate3.ValueIsNull = True Then
                MsgBox "Anda mencentang data PTP, Date Payment Effective tidak boleh kosong!", vbOKOnly + vbExclamation, "Informasi"
                MsgBox pesan, vbOKOnly + vbInformation, "Informasi"
                CEK_DATA_VALID = False
                Exit Function
            End If
            
            'Cek combo via
            If CmbViaPtp.text = "" Then
                MsgBox "Anda mencentang data PTP, Combo VIA tidak boleh kosong!", vbOKOnly + vbExclamation, "Informasi"
                MsgBox pesan, vbOKOnly + vbInformation, "Informasi"
                CEK_DATA_VALID = False
                Exit Function
            End If
            
        End If
    End If
Cek_PTP_Reserved:
        Set M_Objrs_Cek_PTP = Nothing
Lanjut_1:
    
    If C_PTP.Value = 1 Then
        txtremarks.text = txtremarks.text
    End If

    If regnego = True Then
        Dim n%
        Dim jum As Currency
        For n = 1 To FrmCC_Colection.LstPayment.ListItems.Count
            jum = jum + FrmCC_Colection.LstPayment.ListItems(n).SubItems(3)
        Next n
        If jum < FrmCC_Colection.txtPayment.Value Then
            MsgBox "Jumlah PTP Belum sama dengan Jumlah Deal Payment"
            CEK_DATA_VALID = False
            txtremarks.text = ""
            Exit Function
        End If
    End If
    regnego = False
    CEK_DATA_VALID = True
    
    
    
End Function
Public Sub Custid_Double()
Dim ListItem As ListItem
Dim test As String
Dim CMDSQL As String



Set m_cust = New ADODB.Recordset
m_cust.CursorLocation = adUseClient
test = Format(LblDOB.Caption, "yyyy/mm/dd")

'@@ 26-11-10 Ubah logik double custid, harus cek ktpnya dulu
If Trim(lblID.Caption) <> "" Then
    CMDSQL = "Select a.custid, a.name,a.agent, a.amountwo,"
    CMDSQL = CMDSQL + "a.principal,a.flaglead from mgm a where (a.name='"
    CMDSQL = CMDSQL + Trim(TxtName.text) + "' and dob='"
    CMDSQL = CMDSQL + test + "' or ktpno='"
    CMDSQL = CMDSQL + Trim(lblID.Caption) + "')  and a.custid <> '"
    CMDSQL = CMDSQL + Trim(lblCustId.text) + "'"
Else
    CMDSQL = "Select a.custid, a.name,a.agent, a.amountwo,"
    CMDSQL = CMDSQL + "a.principal,a.flaglead from mgm a where a.name='"
    CMDSQL = CMDSQL + Trim(TxtName.text) + "' and dob='"
    CMDSQL = CMDSQL + test + "'"
    CMDSQL = CMDSQL + " and a.custid <> '"
    CMDSQL = CMDSQL + Trim(lblCustId.text) + "'"
End If


'm_cust.Open "Select a.custid, a.name,a.agent, a.amountwo,a.principal,a.flaglead from mgm a where (a.name='" + Trim(txtname.Text) + "' and dob='" + test + "' or ktpno='" & Trim(lblID.Caption) & "') and a.custid <> '" + Trim(lblCustId.text) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
m_cust.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

While Not m_cust.EOF
    Set ListItem = LstDoubleId.ListItems.ADD(, , IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID")))
        ListItem.SubItems(1) = IIf(IsNull(m_cust("NAME")), "", m_cust("NAME"))
        ListItem.SubItems(2) = IIf(IsNull(m_cust("AGENT")), "", m_cust("AGENT")) '
        ListItem.SubItems(3) = Format(IIf(IsNull(m_cust("AmountWo")), 0, m_cust("AmountWo")), "##,###")
        
     If UCase(MDIForm1.txtlevel.text) <> "SUPERVISOR" Then
        If IIf(IsNull(m_cust("flaglead")), 0, m_cust("flaglead")) = 1 Then
            ListItem.SubItems(4) = ""
        Else
           ListItem.SubItems(4) = ENCRIPY(False, CStr(Format(IIf(IsNull(m_cust("principal")), 0, m_cust("principal")), "##,###")))
        End If
    Else
            ListItem.SubItems(4) = ENCRIPY(False, CStr(Format(IIf(IsNull(m_cust("principal")), 0, m_cust("principal")), "##,###")))
    End If
    
    m_cust.MoveNext
Wend
Set m_cust = Nothing
End Sub

Private Sub SSCommand2_Click(Index As Integer)
Dim m_msgbox As Variant
Dim STATUS As String
Dim gaji As Currency
Dim gaji1 As String
Dim ListItem As ListItem
Dim m_data As New ClsNegoPTP
Dim JmlPay As Double
Dim i As Integer
Dim n As Integer
Dim Vrdate As String
Dim jatuhtempo As String
Dim M_Objrs_Cek_PTP As ADODB.Recordset
Dim m_objrs_cek_reserve As ADODB.Recordset

Select Case Index
    Case 0
                    
        If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Or _
            UCase(MDIForm1.txtlevel.text) = "AGENT" Then
            If Trim(Mid(vrcek, 1, 3)) = "PO-" Then
                MsgBox "Untuk account yang statusnya PO-PAID OFF, tidak bisa send PTP! Hubungi SPV anda untuk mengubahnya!", vbOKOnly + vbInformation, "Informasi"
                Exit Sub
            End If
        End If
                    
                    
        If TDBDate3.ValueIsNull Or Tdabamoint.ValueIsNull Or txttenor.ValueIsNull Then
            MsgBox "Pengisian Data Belum Lengkap (installment,tenor,dateptp)!"
            Exit Sub
        End If
        
        '@@ 29 Desember 2011, Cek terlebih dahulu, apakah ada CPA atau tidak, jika tidak ada CPA maka
        'tidak bisa melakukan PTP
       CMDSQL = "select * from tblcpa where vcustid='"
       CMDSQL = CMDSQL + Trim(lblCustId.text) + "' order by nid desc limit 1 "
       Set M_objrs = New ADODB.Recordset
       M_objrs.CursorLocation = adUseClient
       M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
       
       If M_objrs.RecordCount = 0 Then
        'C_PTP.Value = vbUnchecked
        MsgBox "Untuk membuat status account PTP, harus dibuat terlebih dahulu CPA nya!", vbOKOnly + vbInformation, "Informasi"
        MsgBox "Data PTP gagal dibuat!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_objrs = Nothing
        Exit Sub
       End If
       
             
       If txtPayment.Value < Val(M_objrs("nttlpayment")) Then
        MsgBox "Total Amount Deal tidak boleh lebih kecil dari payment di CPA!", vbOKOnly + vbInformation, "Informasi"
        a = MsgBox("Payment di CPA adalah: Rp." + Format(M_objrs("nttlpayment"), "##,###") + ". Anda ingin mengganti Total Amount Deal dengan nilai Payment di CPA?", vbYesNo + vbQuestion, "Konfirmasi")
        If a = vbNo Then
            MsgBox "Data PTP gagal ditambahkan!", vbOKOnly + vbExclamation, "Pemberitahuan"
            Exit Sub
        Else
            'Ambil Nilai Payment di CPA untuk di tempatkan di Total Amount Deal
            txtPayment.Value = IIf(IsNull(M_objrs("nttlpayment")), "0", M_objrs("nttlpayment"))
            GoTo LanjutPtp
        End If
       End If
       
       If txtPayment.Value > Val(M_objrs("nttlpayment")) Then
        MsgBox "Total Amount Deal tidak boleh lebih besar dari payment di CPA!", vbOKOnly + vbInformation, "Informasi"
        a = MsgBox("Payment di CPA adalah: Rp." + Format(M_objrs("nttlpayment"), "##,###") + ". Anda ingin mengganti Total Amount Deal dengan nilai Payment di CPA?", vbYesNo + vbQuestion, "Konfirmasi")
        If a = vbNo Then
            MsgBox "Data PTP gagal ditambahkan!", vbOKOnly + vbExclamation, "Pemberitahuan"
            Exit Sub
        Else
            'Ambil Nilai Payment di CPA untuk di tempatkan di Total Amount Deal
            txtPayment.Value = IIf(IsNull(M_objrs("nttlpayment")), "0", M_objrs("nttlpayment"))
            GoTo LanjutPtp
        End If
       End If
       
       
LanjutPtp:
        
         'Cek apakah Tenor, lebih kecil dari installment period di CPA
             If txttenor.Value < Val(M_objrs("nperiod")) Then
                MsgBox "Tenor tidak boleh lebih kecil dari installment period di CPA!", vbOKOnly + vbInformation, "Informasi"
                a = MsgBox("Installment period di CPA adalah :" + Format(M_objrs("nperiod"), "##,###") + ". Anda ingin mengganti Tenor dengan nilai Installment Period di CPA?", vbYesNo + vbQuestion, "Konfirmasi")
                If a = vbNo Then
                    MsgBox "Data PTP gagal ditambahkan!", vbOKOnly + vbExclamation, "Pemberitahuan"
                    Exit Sub
                Else
                    'Ambil Nilai Tenor dari Installment Period di CPA
                    txttenor.Value = IIf(IsNull(M_objrs("nperiod")), "0", M_objrs("nperiod"))
                    If txttenor > 1 Then
                        Chktenor.Value = vbChecked
                    End If
                    GoTo LanjutPtp2
                End If
            End If
            
            If txttenor.Value > Val(M_objrs("nperiod")) Then
                MsgBox "Tenor tidak boleh lebih besar dari installment period di CPA!", vbOKOnly + vbInformation, "Informasi"
                a = MsgBox("Installment period di CPA adalah :" + Format(M_objrs("nperiod"), "##,###") + ". Anda ingin mengganti Tenor dengan nilai Installment Period di CPA?", vbYesNo + vbQuestion, "Konfirmasi")
                If a = vbNo Then
                    MsgBox "Data PTP gagal ditambahkan!", vbOKOnly + vbExclamation, "Pemberitahuan"
                    Exit Sub
                Else
                    'Ambil Nilai Tenor dari Installment Period di CPA
                    txttenor.Value = IIf(IsNull(M_objrs("nperiod")), "0", M_objrs("nperiod"))
                    If txttenor > 1 Then
                        Chktenor.Value = vbChecked
                    End If
                    GoTo LanjutPtp2
                End If
            End If
            
            Set M_objrs = Nothing

LanjutPtp2:
        
        '@@ 07-02-2012 Cek data dulu, apakah sebelumnya ada data di tblnegoptp? Buat Handle
        'Apakah ada data PTP sebelumnya, kalo ada data ptp sebelumnya dihapus
        '@@ 09-04-2012 filter tanggal dihapus dulu
        CMDSQL = "select * from tblnegoptp where custid='"
        CMDSQL = CMDSQL + Trim(lblCustId.text) + "'  "
        'Cmdsql = Cmdsql + " and date_part('month',promisedate)>=date_part('month',now())  "
        'Cmdsql = Cmdsql + " and date_part('year',promisedate)=date_part('year',now()) "
        '@@13-04-2012 Tambahkan Filter tanggal
        CMDSQL = CMDSQL + " and date(promisedate)='"
        CMDSQL = CMDSQL + CStr(Format(TDBDate3.Value, "yyyy-mm-dd")) + "' "
        CMDSQL = CMDSQL + " order by promisedate,id desc "
        Set M_Objrs_Cek_PTP = New ADODB.Recordset
        M_Objrs_Cek_PTP.CursorLocation = adUseClient
        M_Objrs_Cek_PTP.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Cek_PTP.RecordCount > 0 Then
            Dim KonfirmasiPTP As String
            KonfirmasiPTP = MsgBox("Ada data PTP Sebelumnya dengan TANGGAL YANG SAMA, apakah anda akan menghapus data PTP lama dan menggantinya dengan yang baru?", vbYesNo + vbQuestion, "Konfirmasi")
            If KonfirmasiPTP = vbNo Then
                Set M_Objrs_Cek_PTP = Nothing
                MsgBox "Data PTP gagal ditambahkan!", vbOKOnly + vbInformation, "Informasi"
                Exit Sub
            End If
            
            'Jika memilih Ya, maka cek reservenya
            Dim KonfirmasiReserve As String
            CMDSQL = "select * from tblreserve where custid='"
            CMDSQL = CMDSQL + Trim(lblCustId.text) + "' and stsmove='0'"
            Set m_objrs_cek_reserve = New ADODB.Recordset
            m_objrs_cek_reserve.CursorLocation = adUseClient
            m_objrs_cek_reserve.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If m_objrs_cek_reserve.RecordCount > 0 Then
                
                '@@ 14-04-2012, Cek dulu tenornya jika lebih dari 1 harus hapus data reservenya
                If txttenor.Value > 1 Then
                    KonfirmasiReserve = MsgBox("Tenor lebih dari 1.Apakah anda akan menghapus data reserve yang lama?", vbYesNo + vbQuestion, "Konfirmasi")
                
                    If KonfirmasiReserve = vbNo Then
                        MsgBox "Data PTP gagal ditambahkan!", vbOKOnly + vbExclamation, "Informasi"
                        Set m_objrs_cek_reserve = Nothing
                        Exit Sub
                    End If
                End If
                
                KonfirmasiReserve = vbYes
                
                If KonfirmasiReserve = vbYes Then
                    
                    If M_Objrs_Cek_PTP.RecordCount > 0 Then
                        'Hapus data PTPnya
                        While Not M_Objrs_Cek_PTP.EOF
                            CMDSQL = "delete from tblnegoptp where id='"
                            CMDSQL = CMDSQL + CStr(M_Objrs_Cek_PTP("id")) + "'"
                            M_OBJCONN.Execute CMDSQL
                            M_Objrs_Cek_PTP.MoveNext
                        Wend
                    End If
                    
                    'Hapus data Reservenya
                    If m_objrs_cek_reserve.RecordCount > 0 Then
                        While Not m_objrs_cek_reserve.EOF
                            CMDSQL = "delete from tblreserve where id='"
                            CMDSQL = CMDSQL + CStr(m_objrs_cek_reserve("id")) + "'"
                            M_OBJCONN.Execute CMDSQL
                            m_objrs_cek_reserve.MoveNext
                        Wend
                    End If
                    
                End If
                
            Else
                    'Jika tidak ada data reserve maka langsung hapus saja data ptp nya
                    If M_Objrs_Cek_PTP.RecordCount > 0 Then
                        While Not M_Objrs_Cek_PTP.EOF
                            CMDSQL = "delete from tblnegoptp where id='"
                            CMDSQL = CMDSQL + CStr(M_Objrs_Cek_PTP("id")) + "'"
                            M_OBJCONN.Execute CMDSQL
                            M_Objrs_Cek_PTP.MoveNext
                        Wend
                    End If
            End If
            LstPayment.ListItems.CLEAR
            LstReserve.ListItems.CLEAR
            Set m_objrs_cek_reserve = Nothing
        Else
            'Ini jika PTP Jatuh Temponya kosong!
            'Konfirmasi lagi untuk penghapusan reserve data
            If txttenor.Value > 1 Then
                KonfirmasiReserve = MsgBox("Tenor lebih dari 1. Apakah anda akan membersihkan data reserve PTP?", vbYesNo + vbQuestion, "Konfirmasi")
                If KonfirmasiReserve = vbNo Then
                    MsgBox "Data PTP Gagal ditambahkan!", vbOKOnly + vbInformation, "Informasi"
                    Exit Sub
                End If
                CMDSQL = "delete from tblreserve where custid='"
                CMDSQL = CMDSQL + Trim(lblCustId.text) + "' and stsmove='0'"
                M_OBJCONN.Execute CMDSQL
             End If
        End If
        
        Call CariTanggalTagih
        
        '@@ 22-12-2011 Menentukan nilai awal payment
        If Val(txttenor.Value) > 1 Then
            FrmDealPtp.Show vbModal
            Exit Sub
        End If
        
        bcekptp = True
        '@@ 14 April 2012, Cek tanggal negoptp jika ada yang sama dengan yang diinputkan,
        'yang lama dihapus dan diganti dengan yang baru
        Dim M_Objrs_Cek_Tgl As ADODB.Recordset
           If Chktenor.Value = 0 Then
                
                jatuhtempo = Format(TDBDate3.Value, "yyyy-mm-dd")
                
                '@@14-04-2012 Cek Data
                CMDSQL = "select * from tblnegoptp where custid='"
                CMDSQL = CMDSQL + lblCustId.text + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(TDBDate3.Value, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from tblnegoptp where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                 
                CMDSQL = "INSERT INTO TblNegoPTP "
                CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
                CMDSQL = CMDSQL + "VALUES "
                CMDSQL = CMDSQL + "('" + lblCustId + "', "
                CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
                CMDSQL = CMDSQL + "" + CStr(Tdabamoint.Value) + " , "
                'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
                CMDSQL = CMDSQL + "now(),"
                CMDSQL = CMDSQL + "'IPO')"
                M_OBJCONN.Execute CMDSQL
                
                '@@14042012, tblnegoptp_log di cek aja
                 '@@14-04-2012 Cek Data
                CMDSQL = "select * from tblnegoptp_log where custid='"
                CMDSQL = CMDSQL + lblCustId.text + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(TDBDate3.Value, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from tblnegoptp_log where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                
                ' isi ke tbl log_ptp
                CMDSQL = "INSERT INTO tblnegoptp_log "
                CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
                CMDSQL = CMDSQL + "VALUES "
                CMDSQL = CMDSQL + "('" + lblCustId + "', "
                CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
                CMDSQL = CMDSQL + "" + CStr(Tdabamoint.Value) + " , "
                'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
                CMDSQL = CMDSQL + "now(),"
                CMDSQL = CMDSQL + "'" + lblaoc.Caption + "','P')"
                M_OBJCONN.Execute CMDSQL
                
                Set ListItem = LstPayment.ListItems.ADD(, , "")
                ListItem.SubItems(1) = ""
                ListItem.SubItems(2) = Format(TDBDate3.Value, "yyyy-mm-dd")
                ListItem.SubItems(3) = CStr(Tdabamoint.Value)
                ListItem.SubItems(4) = "IPO"
                ListItem.SubItems(5) = MDIForm1.TDBDate1.Value
            
            Else
            
                jatuhtempo = Format(TDBDate3.Value, "yyyy-mm-dd")
                
                 '@@14-04-2012 Cek Data
                CMDSQL = "select * from tblnegoptp where custid='"
                CMDSQL = CMDSQL + lblCustId.text + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(TDBDate3.Value, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from tblnegoptp where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                
                CMDSQL = "INSERT INTO TblNegoPTP "
                CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
                CMDSQL = CMDSQL + "VALUES "
                CMDSQL = CMDSQL + "('" + lblCustId + "', "
                CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
                CMDSQL = CMDSQL + "" + CStr(Tdabamoint.Value) + " , "
                'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
                CMDSQL = CMDSQL + "now(),"
                CMDSQL = CMDSQL + "'IPO')"
                M_OBJCONN.Execute CMDSQL
                
                 '@@14-04-2012 Cek Data
                CMDSQL = "select * from tblnegoptp_log where custid='"
                CMDSQL = CMDSQL + lblCustId.text + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(TDBDate3.Value, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from tblnegoptp_log where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                        
                ' isi ke tbl log_ptp
                CMDSQL = "INSERT INTO tblnegoptp_log "
                CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
                CMDSQL = CMDSQL + "VALUES "
                CMDSQL = CMDSQL + "('" + lblCustId + "', "
                CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
                CMDSQL = CMDSQL + "" + CStr(Tdabamoint.Value) + " , "
                'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
                CMDSQL = CMDSQL + "now(),"
                CMDSQL = CMDSQL + "'" + lblaoc.Caption + "','P')"
                M_OBJCONN.Execute CMDSQL
                
                Set ListItem = LstPayment.ListItems.ADD(, , "")
                ListItem.SubItems(1) = ""
                ListItem.SubItems(2) = Format(TDBDate3.Value, "yyyy-mm-dd")
                ListItem.SubItems(3) = CStr(Tdabamoint.Value)
                ListItem.SubItems(4) = "IPO"
                ListItem.SubItems(5) = MDIForm1.TDBDate1.Value
    

        n = 0
        For i = 1 To Val(txttenor - 1)
            n = n + 1
            JmlPay = (txtPayment - Tdabamoint) / (txttenor.Value - 1)
            'VRDATE = Format(DateAdd("m", n, TDBDate3.Value), "yyyy-mm-dd")
            Vrdate = DateAdd("m", n, Format(TDBDate3.Value, "yyyy-mm-dd"))
            
                '@@14-04-2012 Cek Data
                CMDSQL = "select * from tblreserve where custid='"
                CMDSQL = CMDSQL + lblCustId.text + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(Vrdate, "yyyy-mm-dd")) + "' and stsmove='0'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from tblreserve where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
            
            CMDSQL = "INSERT INTO tblreserve "
            CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + lblCustId + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(JmlPay) + " , "
            'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "now(),"
            CMDSQL = CMDSQL + "'IPO')"
            M_OBJCONN.Execute CMDSQL
            
            
            '@@14-04-2012 Cek Data
                CMDSQL = "select * from tblnegoptp_log where custid='"
                CMDSQL = CMDSQL + lblCustId.text + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(Vrdate, "yyyy-mm-dd")) + "' "
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from tblnegoptp_log where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
            
            
            CMDSQL = "INSERT INTO TblNegoptp_log "
            CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglinput,agent,stsacc) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + lblCustId + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(JmlPay) + " , "
            'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "now(),"
            CMDSQL = CMDSQL + "'" + lblaoc.Caption + "','R')"
            M_OBJCONN.Execute CMDSQL

        Set ListItem = LstReserve.ListItems.ADD(, , "")
            ListItem.SubItems(1) = ""
                               'listitem.SubItems(2) = .TDBDate1.Value
            ListItem.SubItems(2) = Format(Vrdate, "yyyy-mm-dd")
            ListItem.SubItems(3) = JmlPay
            ListItem.SubItems(4) = "IPO"
            ListItem.SubItems(5) = MDIForm1.TDBDate1.Value
    Next i
   End If
    
    Case 1
        Dim M_Cek_Status As ADODB.Recordset
        Dim Cmdsql_Cek As String
        
        If LstPayment.ListItems.Count = 0 Then
            Exit Sub
        End If
        
        '@@ 11-04-2012 Cek status account terlebih dahulu, data bisa diedit jika status account PTP
        Cmdsql_Cek = "select f_cek_new from mgm where custid='"
        Cmdsql_Cek = Cmdsql_Cek + lblCustId.text + "'"
        Set M_Cek_Status = New ADODB.Recordset
        M_Cek_Status.CursorLocation = adUseClient
        M_Cek_Status.Open Cmdsql_Cek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If IsNull(M_Cek_Status("f_cek_new")) = True Then
            MsgBox "Data hanya dapat diedit jika status account=PTP!", vbOKOnly + vbExclamation, "Peringatan!"
            Set M_Cek_Status = Nothing
            Exit Sub
        End If
        
        If Mid(M_Cek_Status("f_cek_new"), 1, 3) <> "PTP" Then
            MsgBox "Data hanya dapat diedit jika status account=PTP!", vbOKOnly + vbExclamation, "Peringatan!"
            Set M_Cek_Status = Nothing
            Exit Sub
        End If
        
        '@@27062012, Jika agent maka tidak dapat diedit!
        If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
            MsgBox "Mohon maaf anda tidak dapat mengedit PTP!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        
           With FrmNegoPTP
                .Caption = "Ubah Data"
                .SSCommand1(0).Caption = "Update"
                .TDBDate1.Value = Format(LstPayment.SelectedItem.SubItems(2), "yyyy-mm-dd")
                .TDBNumber1.Value = LstPayment.SelectedItem.SubItems(3)
                .Show vbModal
                If .ok Then
                    
                    '@@ Buat Update Tanggal Tagih
                    If C_PTP.Value = vbChecked Then
                                
                        '@@ 05-10-2011, Jika melakukan PTP maka combo via ptp harus diisi
                        If CmbViaPtp.text = "" Then
                            MsgBox "Combo Via tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
                            MsgBox "Data gagal diupdate!", vbOKOnly + vbInformation, "Informasi"
                            Unload Me
                            Exit Sub
                        End If
            
                        'Tambahan, Jika Status data PTP, hitung tanggal tagih
                        If TDBDate3.ValueIsNull Then
                            MsgBox "Anda belum menentukan tanggal effective pembayaran!", vbOKOnly + vbInformation, "Informasi"
                            MsgBox "Data gagal diupdate!", vbOKOnly + vbInformation, "Informasi"
                            Unload Me
                            Exit Sub
                        End If
            
                    End If
                    
                    
                    
                    m_data.UPDATE_NegoPTP M_OBJCONN, .TxtCustid.text, Format(.TDBDate1.Value, "yyyy-mm-dd"), CStr(.TDBNumber1.Value), LstPayment.SelectedItem.SubItems(1)

                    On Error GoTo add_error
                    If m_data.ADD_OK Then
                        'LstPayment.SelectedItem.SubItems(1) = ""
                        LstPayment.SelectedItem.SubItems(2) = Format(.TDBDate1.Value, "yyyy-mm-dd")
                        LstPayment.SelectedItem.SubItems(3) = .TDBNumber1.Value
                        
                        Call CariTanggalTagih
                        
                        CMDSQL = "update mgm set tgl_tagih='"
                        CMDSQL = CMDSQL + Format(TdbTglTagih.Value, "yyyy-mm-dd") + "',dateptp='"
                        CMDSQL = CMDSQL + Format(TDBDate3.Value, "yyyy-mm-dd") + "' "
                        CMDSQL = CMDSQL + " where custid='"
                        CMDSQL = CMDSQL + Trim(lblCustId.text) + "'"
                        M_OBJCONN.Execute CMDSQL
                        
                    On Error GoTo 0
                    End If
                End If
                Unload FrmNegoPTP
            End With
        Exit Sub
    Case 2
         Frmdelete.Show vbModal
    Case 3
        MsgBox "Tidak dapat hapus reserved PTP!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
        frmdeletereserve.Show vbModal
End Select
add_error:
End Sub
Private Sub VisitYES()
Text1.BackColor = &HFF00&
TxtCustid.BackColor = &H80000005
TxtName.BackColor = &H80000005
TDBNumber1.BackColor = &H80000005
TXtDetails.BackColor = &H80000005
'LstVisit.BackColor = &HFF00&
TxtAddress.BackColor = &H80000005
TxtAddress.Enabled = True
TXtDetails.Enabled = True
Option7(0).Enabled = True
Option7(1).Enabled = True
Option7(2).Enabled = True
End Sub
Private Sub VisitNo()
Text1.BackColor = &H8000000F
TxtCustid.BackColor = &H8000000F
TxtName.BackColor = &H8000000F
TDBNumber1.BackColor = &H8000000F
TXtDetails.BackColor = &H8000000F
TxtAddress.BackColor = &H8000000F
'LstVisit.BackColor = &H8000000F
Option8(1).Value = True
Option7(0).Enabled = False
Option7(1).Enabled = False
Option7(2).Enabled = False

TxtAddress.Enabled = False
TXtDetails.Enabled = False
End Sub
'
'Private Sub SSPanel1_Click()
'    Call isi_datacustomer
'End Sub

Private Sub Tdabamoint_Change()
bcekptp = False
End Sub

Private Sub TDBDate3_Change()
   Dim CMDSQL As String
   Dim M_objrs As ADODB.Recordset
   Dim TglPtp As String
   
   
   If C_PTP.Value Then
        '@@ 09-04-2012
        Call CariTanggalTagih
        'Update tanggal negoptp
        CMDSQL = "select * from tblnegoptp where custid='"
        CMDSQL = CMDSQL + lblCustId.text + "'"
        CMDSQL = CMDSQL + " order by promisedate desc limit 1"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_objrs.RecordCount = 0 Then
             Set M_objrs = Nothing
             Exit Sub
        End If
        
        If TDBDate3.Value = Empty Then
             TglPtp = "null"
        Else
             TglPtp = "'" + Format(TDBDate3.Value, "yyyy-mm-dd") + "'"
        End If
        
        On Error GoTo Salah
        CMDSQL = "update tblnegoptp set promisedate="
        CMDSQL = CMDSQL + TglPtp + " where id='"
        CMDSQL = CMDSQL + CStr(M_objrs("id")) + "'"
        M_OBJCONN.Execute CMDSQL
        Call Show_NEGOPTP
        
        '@@27-06-2012 Update juga di negoptp
        CMDSQL = "update mgm set dateptp="
        CMDSQL = CMDSQL + TglPtp + ",tgl_tagih='"
        CMDSQL = CMDSQL + Format(TdbTglTagih.Value, "yyyy-mm-dd") + "' "
        CMDSQL = CMDSQL + " where custid='"
        CMDSQL = CMDSQL + CStr(lblCustId.text) + "'"
        M_OBJCONN.Execute CMDSQL
        
   End If
   Exit Sub
Salah:
   MsgBox "Ada error: " & err.Description
End Sub

Private Sub TdbPTP_Change()
TdbPTP.Value = TDBDate1.Value
End Sub

Private Sub Text11_Click()
    If Len(Text11.text) > 3 Then
        If Text11.text <> Empty Then
            CmbPhone.text = "EC Num"
            txtgetnomor.text = Text11.text
        End If
    Else
        CmbPhone.text = ""
    End If
End Sub

Private Sub Text11m_Click()
    If Len(Text11.text) > 3 Then
        If Text11.text <> Empty Then
            CmbPhone.text = "EC Num"
            txtgetnomor.text = Text11.text
            FrmCC_Colection.Frame3.Caption = "0"
        End If
    Else
        CmbPhone.text = ""
        FrmCC_Colection.Frame3.Caption = "0"
    End If

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Timer1_Timer()
    Text5.text = MDIForm1.Txtstatus.text
    Text7.text = MDIForm1.txt_unique_id.text
End Sub

'Private Sub Timer_cek_inbox_Timer()
''@@ 11-03-2011 Di remarks, udah tidak diapakai
''    Text2 = LstSMS.ListItems.Count
''
''    LstSMS.ListItems.CLEAR
''    LstSMS2.ListItems.CLEAR
''    Isi_SendSMS
''    Isi_SendSMS2
'End Sub

Private Sub blink(Seconds As Single)
 Dim a As Long
 Seconds = Seconds + Timer
 While Seconds > Timer
  a = DoEvents
 Wend
End Sub

Private Sub TimerBlink_Timer()
'@@ 05-10-2011 tombol OST ditiadakan
   
'               If SSCommand1(7).BackColor = vbRed Then
'                 SSCommand1(7).BackColor = vbGreen
'                 KelapKelip = KelapKelip + 1
'               Else
'                 SSCommand1(7).BackColor = vbRed
'                 KelapKelip = KelapKelip + 1
'               End If
'
'           If KelapKelip = 7 Then
'            KelapKelip = 0
'            WaitSecs (3)
'            'TimerBlink.Enabled = False
'           End If
    
End Sub

Private Sub BlinkCPA_Timer()
    Dim kelapkelipCpa As Integer
    
    If SSCommand1(4).BackColor = vbBlack Then
        SSCommand1(4).BackColor = vbRed
        kelapkelipCpa = kelapkelipCpa + 1
    Else
        SSCommand1(4).BackColor = vbBlack
        kelapkelipCpa = kelapkelipCpa + 1
    End If
           
    If kelapkelipCpa = 7 Then
            kelapkelipCpa = 0
            WaitSecs (3)
            SSCommand1(4).BackColor = vbBlack
            TimerBlinkCPA.Enabled = False
    End If
End Sub

Private Sub TimerBlinkDetailMapping_Timer()
    'Dim kelapkelipDetail As Integer
    
    If Val(LblMap.Caption) > 0 Then
        If LblMap.BackColor = vbBlack Then
            LblMap.BackColor = vbRed
            kelapkelipDetail = kelapkelipDetail + 1
        Else
            LblMap.BackColor = vbBlack
            kelapkelipDetail = kelapkelipDetail + 1
        End If

    Else
        TimerBlinkDetailMapping.Enabled = False
    End If
End Sub

Private Sub TimerBlinkSms_Timer()
    If LabelSms.ForeColor = vbBlack Then
        LabelSms.ForeColor = vbRed
        Command2.BackColor = vbRed
        KelapKelip = KelapKelip + 1
    Else
        LabelSms.ForeColor = vbBlack
        Command2.BackColor = vbYellow
        KelapKelip = KelapKelip + 1
    End If
           
    If KelapKelip = 7 Then
            KelapKelip = 0
            WaitSecs (3)
            'TimerBlink.Enabled = False
    End If
End Sub

Private Sub TimerCekMapping_Timer()
'     If CmdDataMapping.BackColor = vbGreen Then
'        CmdDataMapping.BackColor = vbRed
'        KelapKelip = KelapKelip + 1
'    Else
'        CmdDataMapping.BackColor = vbYellow
'        KelapKelip = KelapKelip + 1
'    End If
'
'    If KelapKelip = 7 Then
'            KelapKelip = 0
'            WaitSecs (3)
'            'TimerBlink.Enabled = False
'    End If
    
    If label1(8).BackColor = &HABE18E Then
        label1(8).BackColor = &HFCFCFC
    Else
        label1(8).BackColor = &HABE18E
    End If
End Sub

Private Sub TimerOfferingDiscon_Timer()
    ' Last Update #12042013 by Izuddin
    If Not (listview1(0).ListItems.Count > 0) Then
        'OfferingDiscGuide
    End If
    TimerOfferingDiscon.Enabled = False
End Sub

Private Sub txtadd_phone_Click(Index As Integer)
Dim query As String
Dim M_objrs As ADODB.Recordset
Dim hasil As String
Select Case Index
    Case 0 Or 6
        query = " select count(*) from tblrequestadditionalphone where request_number = '" & txtadd_phone(0).text & "' and agent = '" & MDIForm1.TxtUsername.text & "'"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If Not M_objrs.BOF And Not M_objrs.EOF Then
            Label16.Caption = M_objrs!Count
        End If
        
        If Label16.Caption >= 1 Then
            SSCommand1(0).Enabled = False
            MsgBox "Nomor Belum di Approve", vbInformation
            Exit Sub
        End If
        
        If txtadd_phone(0).text = "" Then
            SSCommand1(0).Enabled = False
            Exit Sub
        End If
        TYPETELP = "HOME1"
            CmbPhone.text = "AddHome1"
            txtgetnomor.text = txtadd_phone(0).text
            FrmCC_Colection.Frame3.Caption = "0"
    Case 1 Or 5
        query = " select count(*) from tblrequestadditionalphone where request_number = '" & txtadd_phone(1).text & "' and agent = '" & MDIForm1.TxtUsername.text & "'"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If Not M_objrs.BOF And Not M_objrs.EOF Then
            Label16.Caption = M_objrs!Count
        End If
        
        If Label16.Caption >= 1 Then
            SSCommand1(0).Enabled = False
            MsgBox "Nomor Belum di Approve", vbInformation
            Exit Sub
        End If
        
        If txtadd_phone(1).text = "" Then
            SSCommand1(0).Enabled = False
            Exit Sub
        End If
        
        TYPETELP = "OFFICE1"
        CmbPhone.text = "AddOffice1"
        txtgetnomor.text = txtadd_phone(1).text
        FrmCC_Colection.Frame3.Caption = "0"
    Case 4
        query = " select count(*) from tblrequestadditionalphone where request_number = '" & txtadd_phone(2).text & "' and agent = '" & MDIForm1.TxtUsername.text & "'"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If Not M_objrs.BOF And Not M_objrs.EOF Then
            Label16.Caption = M_objrs!Count
        End If
        
        If Label16.Caption >= 1 Then
            SSCommand1(0).Enabled = False
            MsgBox "Nomor Belum di Approve", vbInformation
            Exit Sub
        End If
        
        If txtadd_phone(2).text = "" Then
            SSCommand1(0).Enabled = False
            Exit Sub
        End If
        
        TYPETELP = "MOBILE1"
            txtPhone.text = txtadd_phone(2).text
            txtPhoneA.text = txtadd_phone(2).text
            txtgetnomor.text = txtadd_phone(2).text
        If Len(txtadd_phone(2).text) > 3 Then
            CmbPhone.text = "AddMobile1"
            FrmCC_Colection.Frame3.Caption = "0"
            Else
            CmbPhone.text = ""
            FrmCC_Colection.Frame3.Caption = "0"
        End If
    Case 3 Or 7
        query = " select count(*) from tblrequestadditionalphone where request_number = '" & txtadd_phone(3).text & "' and agent = '" & MDIForm1.TxtUsername.text & "'"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If Not M_objrs.BOF And Not M_objrs.EOF Then
            Label16.Caption = M_objrs!Count
        End If
        
        If Label16.Caption >= 1 Then
            SSCommand1(0).Enabled = False
            MsgBox "Nomor Belum di Approve", vbInformation
            Exit Sub
        End If
        
         If txtadd_phone(3).text = "" Then
            SSCommand1(0).Enabled = False
            Exit Sub
        End If
        
        TYPETELP = "MOBILE2"
            txtPhone.text = txtadd_phone(3).text
            txtPhoneA.text = txtadd_phone(3).text
            txtgetnomor.text = txtadd_phone(3).text
        If Len(txtadd_phone(3).text) > 3 Then
            CmbPhone.text = "AddOtherphone"
            FrmCC_Colection.Frame3.Caption = "0"
            Else
            CmbPhone.text = ""
            FrmCC_Colection.Frame3.Caption = "0"
        End If
End Select
End Sub

Private Sub txthasil_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtHomeAdd1_Click()
Dim query As String
Dim M_objrs As ADODB.Recordset
Dim hasil As String

query = " select count(*) from tblrequestadditionalphone where request_number = '" & txtHomeAdd1.Value & "' and agent = '" & MDIForm1.TxtUsername.text & "'"
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

If Not M_objrs.BOF And Not M_objrs.EOF Then
    Label16.Caption = M_objrs!Count
End If

If Label16.Caption >= 1 Then
    SSCommand1(0).Enabled = False
    MsgBox "Nomor Belum di Approve", vbInformation
    Exit Sub
End If

If txtHomeAdd1.Value = "" Then
    SSCommand1(0).Enabled = False
    Exit Sub
End If

TYPETELP = "HOME1"
    CmbPhone.text = "AddHome1"
    txtgetnomor.text = txtHomeAdd1.text
End Sub


Private Sub txtHomeNo1_Click()
    If Len(txtHomeNo1.text) > 3 Then
        If txtHomeNo1.text <> Empty Then
            CmbPhone.text = "Old Num"
            txtgetnomor.text = txtHomeNo1.text
        End If
    Else
        CmbPhone.text = ""
    End If
End Sub

Private Sub txtHomeNo1A_Click()
If Len(txtHomeNo1A.text) > 3 Then
    CmbPhone.text = "HomePhone"
    txtgetnomor.text = txtHomeNo1A.text

    Else
    CmbPhone.text = ""
    End If
End Sub
Private Sub txtHomeNo1A_DblClick()
txthasil.text = txtHomeNo1.text
End Sub

Private Sub txtHomeNo1m_Click()
    If Len(txtHomeNo1.text) > 3 Then
        If txtHomeNo1.text <> Empty Then
            CmbPhone.text = "Old Num"
            txtgetnomor.text = txtHomeNo1.text
            FrmCC_Colection.Frame3.Caption = "0"
        End If
    Else
        CmbPhone.text = ""
        FrmCC_Colection.Frame3.Caption = "0"
    End If
End Sub

Private Sub txtHomeNo2_Click()
    If Len(txtHomeNo2.text) > 3 Then
    CmbPhone.text = "HomePhone2"
    txtgetnomor.text = txtHomeNo2.text
    Else
    CmbPhone.text = ""
    End If
End Sub

Private Sub txtHomeNo2A_Click()
  If Len(txtHomeNo2A.text) > 3 Then
    CmbPhone.text = "HomePhone2"
    txtgetnomor.text = txtHomeNo2A.text
    Else
    CmbPhone.text = ""
    End If
End Sub
Private Sub txtHomeNo2A_DblClick()
txthasil.text = txtHomeNo2.text
End Sub

Private Sub txtMobileAdd1A_Click()
TYPETELP = "MOBILE1"
    txtPhone.text = txtMobileAdd1.Value
    txtPhoneA.text = txtMobileAdd1A.Value
    CmbPhone.text = "AddMobile1"
    txtgetnomor.text = txtMobileAdd1A.text
End Sub

Private Sub txtMobileAdd1A_DblClick()
txthasil.text = txtMobileAdd1.text
End Sub

Private Sub txtMobileAdd2A_Change()
'    txtMobileAdd2.Text = txtMobileAdd2A.Text
End Sub
Private Sub txtMobileAdd2A_Click()
TYPETELP = "MOBILE2"
    txtPhone.text = txtMobileAdd2.Value
    txtPhoneA.text = txtMobileAdd2A.Value
    If Len(txtMobileAdd2A.text) > 3 Then
    CmbPhone.text = "AddMobile2"
    txtgetnomor.text = txtMobileAdd2A.text
    Else
    CmbPhone.text = ""
    End If
End Sub

Private Sub txtMobileAdd2A_DblClick()
txthasil.text = txtMobileAdd2.text
End Sub

Private Sub txtMobileNo1_Click()
If Len(txtMobileNo1.text) > 3 Then
CmbPhone.text = "Office Num"
txtgetnomor.text = txtMobileNo1A.text
Else
CmbPhone.text = ""
End If
End Sub

Private Sub txtMobileNo1A_Click()
If Len(txtMobileNo1A.text) > 3 Then
CmbPhone.text = "Hp"
txtgetnomor.text = txtMobileNo1.text
Else
CmbPhone.text = ""
End If
End Sub

Private Sub txtMobileNo1A_DblClick()
txthasil.text = txtMobileNo1.text
End Sub

Private Sub txtMobileNo1m_Click()
If Len(txtMobileNo1.text) > 3 Then
CmbPhone.text = "Office Num"
txtgetnomor.text = txtMobileNo1A.text
FrmCC_Colection.Frame3.Caption = "0"
Else
CmbPhone.text = ""
FrmCC_Colection.Frame3.Caption = "0"
End If

End Sub

Private Sub txtMobileNo2_Click()
If Len(txtMobileNo2.text) > 3 Then
CmbPhone.text = "Hp2"
txtgetnomor.text = txtMobileNo2.text
Else
CmbPhone.text = ""
End If
End Sub
Private Sub txtMobileNo2A_Click()
If Len(txtMobileNo2A.text) > 3 Then
CmbPhone.text = "Hp2"
txtgetnomor.text = txtMobileNo2A.text
Else
CmbPhone.text = ""
End If
End Sub
Private Sub txtMobileNo2A_DblClick()
    txthasil.text = txtMobileNo2.text
End Sub

Private Sub txtOfficeAdd1_Click()
Dim query As String
Dim M_objrs As ADODB.Recordset
Dim hasil As String

query = " select count(*) from tblrequestadditionalphone where request_number = '" & txtOfficeAdd1.Value & "' and agent = '" & MDIForm1.TxtUsername.text & "'"
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

If Not M_objrs.BOF And Not M_objrs.EOF Then
    Label16.Caption = M_objrs!Count
End If

If Label16.Caption >= 1 Then
    SSCommand1(0).Enabled = False
    MsgBox "Nomor Belum di Approve", vbInformation
    Exit Sub
End If

If txtOfficeAdd1.Value = "" Then
    SSCommand1(0).Enabled = False
    Exit Sub
End If

TYPETELP = "OFFICE1"
'@@03-05-2012 Dinonaktifkan
'If Trim(AOfficeAdd(2).Value) = "021" Or AOfficeAdd(2).Value = "" Then
'    txtPhone.Text = txtOfficeAdd1.Value
'    txtPhoneA.Text = txtOfficeAdd1.Value
'Else
'    txtPhone.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1.Value)
'    txtPhoneA.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1.Value)
'End If
CmbPhone.text = "AddOffice1"
txtgetnomor.text = txtOfficeAdd1.text
End Sub

Private Sub txtOfficeAdd1A_Change()
'    txtOfficeAdd1.Text = txtOfficeAdd1A.Text
End Sub

Private Sub txtOfficeAdd1A_Click()
TYPETELP = "OFFICE1"
'@@03-05-2012 Dinonaktifkan
'If Trim(AOfficeAdd(2).Value) = "021" Or AOfficeAdd(2).Value = "" Then
'    txtPhone.Text = txtOfficeAdd1.Value
'    txtPhoneA.Text = txtOfficeAdd1A.Value
'Else
'    txtPhone.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1.Value)
'    txtPhoneA.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1A.Value)
'End If
CmbPhone.text = "AddOffice1"
txtgetnomor.text = txtOfficeAdd1A.text
End Sub
Private Sub txtOfficeAdd1A_DblClick()
    txthasil.text = txtOfficeAdd1.text
End Sub

Private Sub txtOfficeAdd2_Click()
Dim query As String
Dim M_objrs As ADODB.Recordset
Dim hasil As String

query = " select count(*) from tblrequestadditionalphone where request_number = '" & txtOfficeAdd2.Value & "' and agent = '" & MDIForm1.TxtUsername.text & "'"
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

If Not M_objrs.BOF And Not M_objrs.EOF Then
    Label16.Caption = M_objrs!Count
End If

If Label16.Caption >= 1 Then
    SSCommand1(0).Enabled = False
    MsgBox "Nomor Belum di Approve", vbInformation
    Exit Sub
End If

If txtOfficeAdd2.Value = "" Then
    SSCommand1(0).Enabled = False
    Exit Sub
End If

TYPETELP = "OFFICE2"
'@@03-05-2012 Dinonaktifkan
'If Trim(AOfficeAdd(3).Value) = "021" Or AOfficeAdd(3).Value = "" Then
'    txtPhone.Text = txtOfficeAdd2.Value
'    txtPhoneA.Text = txtOfficeAdd2.Value
'Else
'    txtPhone.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
'    txtPhoneA.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
'End If
CmbPhone.text = "AddOffice2"
txtgetnomor.text = txtOfficeAdd2.text
End Sub

Private Sub txtMobileAdd1_Click()
Dim query As String
Dim M_objrs As ADODB.Recordset
Dim hasil As String

query = " select count(*) from tblrequestadditionalphone where request_number = '" & txtMobileAdd1.Value & "' and agent = '" & MDIForm1.TxtUsername.text & "'"
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

If Not M_objrs.BOF And Not M_objrs.EOF Then
    Label16.Caption = M_objrs!Count
End If

If Label16.Caption >= 1 Then
    SSCommand1(0).Enabled = False
    MsgBox "Nomor Belum di Approve", vbInformation
    Exit Sub
End If

If txtMobileAdd1.Value = "" Then
    SSCommand1(0).Enabled = False
    Exit Sub
End If

TYPETELP = "MOBILE1"
    txtPhone.text = txtMobileAdd1.Value
    txtPhoneA.text = txtMobileAdd1.Value
    txtgetnomor.text = txtMobileAdd1.text
If Len(txtMobileAdd1.text) > 3 Then
    CmbPhone.text = "AddMobile1"
    Else
    CmbPhone.text = ""
End If
End Sub

Private Sub txtMobileAdd2_Click()
Dim query As String
Dim M_objrs As ADODB.Recordset
Dim hasil As String

query = " select count(*) from tblrequestadditionalphone where request_number = '" & txtMobileAdd2.Value & "' and agent = '" & MDIForm1.TxtUsername.text & "'"
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

If Not M_objrs.BOF And Not M_objrs.EOF Then
    Label16.Caption = M_objrs!Count
End If

If Label16.Caption >= 1 Then
    SSCommand1(0).Enabled = False
    MsgBox "Nomor Belum di Approve", vbInformation
    Exit Sub
End If

If txtMobileAdd2.Value = "" Then
    SSCommand1(0).Enabled = False
    Exit Sub
End If

TYPETELP = "MOBILE2"
    txtPhone.text = txtMobileAdd2.Value
    txtPhoneA.text = txtMobileAdd2.Value
    txtgetnomor.text = txtMobileAdd2.text
If Len(txtMobileAdd2.text) > 3 Then
    CmbPhone.text = "AddMobile2"
    Else
    CmbPhone.text = ""
End If
    
End Sub
Public Sub UpdateAppv()
'If chkAppv(0).Value Then
'    x = MsgBox("Pindahkan data ke Agent DA ?", vbYesNo + vbExclamation, "Info !")
'    If x = vbYes Then
'        CMDSQL = "update mgm set F_pending='Pending',Agent='DA',PO_Agent='" & lblaoc.Caption & "' where custid='" + lblCustId.text + "'"
'        M_OBJCONN.Execute CMDSQL
'        spend = True
'        MsgBox "Data berhasil dipindah ke agent DA", vbInformation
'        VIEW_MGMDATA.LstVwSearchMgm.ListItems.CLEAR
'        MDIForm1.LstGrade.ListItems.CLEAR
'    End If
'Else
'    If chkAppv(1).Value Then
'        Dim spo As ADODB.Recordset
'        Set spo = New ADODB.Recordset
'        spo.CursorLocation = adUseClient
'        spo.Open "select PO_Agent from mgm where custid='" + lblCustId.text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'        If spo!PO_AGENT <> "" And IsNull(spo!PO_AGENT) = False Then
'            CMDSQL = "update mgm set F_pending='',AGENT=PO_Agent where custid='" + lblCustId.text + "'"
'            M_OBJCONN.Execute CMDSQL
'            CMDSQL = "update mgm set PO_Agent='' where custid='" + lblCustId.text + "'"
'            M_OBJCONN.Execute CMDSQL
'            MsgBox "Data berhasil dikembalikan", vbInformation
'            VIEW_MGMDATA.LstVwSearchMgm.ListItems.CLEAR
'            MDIForm1.LstGrade.ListItems.CLEAR
'        Else
'            MsgBox "Silahkan Pilih Status !," & vbCrLf & "untuk menyimpan hilangkan ceklist NO !", vbInformation
'            Exit Sub
'        End If
'    End If
'End If
End Sub

Private Sub txtOfficeAdd2A_Change()
'    txtOfficeAdd2.Text = txtOfficeAdd2A.Text
End Sub

Private Sub txtOfficeAdd2A_Click()
TYPETELP = "OFFICE2"
'@@03-05-2012 Dinonaktifkan
'If Trim(AOfficeAdd(3).Value) = "021" Or AOfficeAdd(3).Value = "" Then
'    txtPhone.Text = txtOfficeAdd2.Value
'    txtPhoneA.Text = txtOfficeAdd2A.Value
'Else
'    txtPhone.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
'    txtPhoneA.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2A.Value)
'End If

CmbPhone.text = "AddOffice2"
txtgetnomor.text = txtOfficeAdd2A.text
End Sub

Private Sub txtOfficeAdd2A_DblClick()
txthasil.text = txtOfficeAdd2.text
End Sub

Private Sub txtOfficeNo1_Click()
If Len(txtOfficeNo1.text) > 2 Then
CmbPhone.text = "New Num"
txtgetnomor.text = txtOfficeNo1.text
Else
CmbPhone.text = ""
End If
End Sub
Private Sub txtOfficeNo1A_DblClick()
 txthasil.text = txtOfficeNo1.text
End Sub

Private Sub txtOfficeNo1A_Click()
If Len(txtOfficeNo1A.text) > 3 Then
CmbPhone.text = "OfficePhone"
txtgetnomor.text = txtOfficeNo1A.text
Else
CmbPhone.text = ""
End If

End Sub

Private Sub txtOfficeNo1m_Click()
If Len(txtOfficeNo1.text) > 2 Then
CmbPhone.text = "New Num"
txtgetnomor.text = txtOfficeNo1.text
FrmCC_Colection.Frame3.Caption = "0"
Else
CmbPhone.text = ""
FrmCC_Colection.Frame3.Caption = "0"
End If
End Sub

Private Sub txtOfficeNo2_Click()
If Len(txtOfficeNo2.text) > 3 Then
CmbPhone.text = "OfficePhone2"
txtgetnomor.text = txtOfficeNo2.text
Else
CmbPhone.text = ""
End If

End Sub
Private Sub txtOfficeNo2A_Click()
If Len(txtOfficeNo2A.text) > 3 Then
CmbPhone.text = "OfficePhone2"
Else
CmbPhone.text = ""
End If

End Sub
Public Sub Show_Reserve()
Dim showlist As New ADODB.Recordset
Dim ListItem As ListItem
Dim CMDSQL As String
Dim TOTPTP As Currency
Dim ssql As String
ssql = "SELECT CUSTID,sum(PAYMENT) as Jum FROM tbllunas WHERE custid = '" + lblCustId.text + "' GROUP BY CUSTID"
showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If showlist.BOF And showlist.EOF Then
    TOTPTP = 0
Else
    TOTPTP = IIf(IsNull(showlist!jum), 0, showlist!jum)
End If
'If ShowList.BOF And ShowList.EOF Then
'    'CMDSQL = "SELECT * FROM TBLNEGOPTP WHERE custid = '" + lblCustId.text + "'"
'    'AND CUSTID NOT IN (SELECT CUSTID FROM tbllunas)"
'    CMDSQL = "SELECT DISTINCT TBLNEGOPTP.PROMISEDATE,TBLNEGOPTP.ID,TBLNEGOPTP.PROMISEPAY,TBLNEGOPTP.TYPE FROM TBLNEGOPTP,tbllunas WHERE "
'    CMDSQL = CMDSQL + "tbllunas.CUSTID<>TBLNEGOPTP.CUSTID AND TBLNEGOPTP.CUSTID='" + lblCustId.text + "' order by TBLNEGOPTP.promisedate desc"
'Else
'    CMDSQL = "SELECT distinct TBLNEGOPTP.PROMISEDATE,TBLNEGOPTP.PROMISEPAY,TBLNEGOPTP.ID,TBLNEGOPTP.TYPE "
'    CMDSQL = CMDSQL + "FROM VWLISTPTP,TBLNEGOPTP WHERE TBLNEGOPTP.CUSTID=VWLISTPTP.CUSTID AND "
'    CMDSQL = CMDSQL + "VWLISTPTP.PAYDATE<TBLNEGOPTP.PROMISEDATE AND TBLNEGOPTP.CUSTID='" + lblCustId.text + "' order by TBLNEGOPTP.promisedate desc"
'End If
If MDIForm1.txtlevel.text = "SUPERVISOR" Then
    CMDSQL = "SELECT * FROM tblreserve where custid = '" + lblCustId.text + "' order by promisedate"
Else
    CMDSQL = "SELECT * FROM tblreserve where custid = '" + lblCustId.text + "' and stsmove=0 order by promisedate"
End If

Set showlist = New ADODB.Recordset
showlist.CursorLocation = adUseClient
showlist.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

LstReserve.ListItems.CLEAR
Dim n As Currency
While Not showlist.EOF
    Set ListItem = LstReserve.ListItems.ADD(, , "")
        ListItem.SubItems(1) = CStr(IIf(IsNull(showlist!ID), "", (showlist!ID)))
        ListItem.SubItems(2) = CStr(IIf(IsNull(showlist!PromiseDate), "", Format(showlist!PromiseDate, "yyyy-mm-dd")))
        ListItem.SubItems(3) = CStr(IIf(IsNull(showlist!PromisePay), "", (Round(showlist!PromisePay, 0))))
        n = n + Val(ListItem.SubItems(3))
        If n <= TOTPTP Then
            ListItem.ListSubItems(1).ForeColor = vbRed
            ListItem.ListSubItems(2).ForeColor = vbRed
            ListItem.ListSubItems(3).ForeColor = vbRed
        End If
        
        ListItem.SubItems(4) = IIf(IsNull(showlist!Type), "", showlist!Type)
        ListItem.SubItems(5) = CStr(IIf(IsNull(showlist!inputdate), "", Format(showlist!inputdate, "yyyy-mm-dd")))
     showlist.MoveNext
Wend

Set showlist = Nothing
End Sub

Private Sub txtOfficeNo2A_DblClick()
txthasil.text = txtOfficeNo2.text
End Sub

Public Sub PesanLockAuto()
    Dim m_objrsPesanReset As ADODB.Recordset
    Dim m_objrsPesanLock As ADODB.Recordset
    Dim M_ObjWktServer As ADODB.Recordset
    Dim WaktuServer As Date
    Dim CMDSQL As String
    
    'Ambil Waktu Server Sekarang
    Set M_ObjWktServer = New ADODB.Recordset
    M_ObjWktServer.CursorLocation = adUseClient
    M_ObjWktServer.Open "Select now() as WktSrv ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    WaktuServer = Format(M_ObjWktServer(0), "yyyy-mm-dd hh:mm")
    Set M_ObjWktServer = Nothing
    
    'Cek pesan reset
    CMDSQL = "select f_pesanresetauto,f_idsessend from usertbl where userid='"
    CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.text) + "'"
    Set m_objrsPesanReset = New ADODB.Recordset
    m_objrsPesanReset.CursorLocation = adUseClient
    m_objrsPesanReset.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    
    If m_objrsPesanReset.RecordCount <> 0 Then
        If m_objrsPesanReset("f_pesanresetauto") = "1" Then
            MsgBox "Reset Data! Ini adalah lock data automatic, data anda akan segera diperbaharui!", vbOKOnly + vbInformation, "Informasi"
           
            VIEW_MGMDATA.LstVwSearchMgm.ListItems.CLEAR
            '@@20-11-10 akhiri session dengan mencatat hasil akhir perubahan status data yang dikerjain agent
                If m_objrsPesanReset("f_idsessend") <> "" Or IsNull(m_objrsPesanReset("f_idsessend")) = False Or m_objrsPesanReset("f_idsessend") <> Empty Then
                    Dim UpdateDtCloseSession As String
'                    UpdateDtCloseSession = "update tblperformpersessionlock set f_ceksekrg=a.f_cek_akhir ,"
'                    UpdateDtCloseSession = UpdateDtCloseSession + " tgl_f_ceksekrg=a.tgl_akhir,endlock='" + CStr(Format(WaktuServer, "yyyy-mm-dd hh:mm:ss")) + "' from "
'                    UpdateDtCloseSession = UpdateDtCloseSession + " (select mgm.custid as custid_mgm,mgm.agent as agent_mgm,"
'                    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.custid as custid_lock,"
'                    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.agent as agent_lock,"
'                    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.idlock as id_lock,"
'                    UpdateDtCloseSession = UpdateDtCloseSession + " mgm.f_cek_new as f_cek_akhir, mgm.tglcall as tgl_akhir"
'                    UpdateDtCloseSession = UpdateDtCloseSession + " from tblperformpersessionlock inner join mgm "
'                    UpdateDtCloseSession = UpdateDtCloseSession + " on mgm.custid=tblperformpersessionlock.custid "
'                    UpdateDtCloseSession = UpdateDtCloseSession + " and mgm.agent=tblperformpersessionlock.agent) as a "
'                    UpdateDtCloseSession = UpdateDtCloseSession + " where tblperformpersessionlock.custid=a.custid_mgm and tblperformpersessionlock.agent=a.agent_mgm and a.id_lock='"
'                    UpdateDtCloseSession = UpdateDtCloseSession + Trim(m_objrsPesanReset("f_idsessend")) + "' and tblperformpersessionlock.agent='"
'                    UpdateDtCloseSession = UpdateDtCloseSession + Trim(mdiform1.txtusername.text) + "'"
'                    M_OBJCONN.Execute UpdateDtCloseSession
                    'bikin null lagi nilai f_idsessend
                    UpdateDtCloseSession = "update usertbl set f_idsessend=null where userid='"
                    UpdateDtCloseSession = UpdateDtCloseSession + Trim(MDIForm1.TxtUsername.text) + "'"
                    M_OBJCONN.Execute UpdateDtCloseSession
                End If
            '@@20-11-10 akhiri session dengan mencatat hasil akhir perubahan status data yang dikerjain agent
             
            CMDSQL = "update usertbl set f_pesanresetauto=null where userid='"
            CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.text) + "'"
            M_OBJCONN.Execute CMDSQL
        End If
    End If
    
    Set m_objrsPesanReset = Nothing
    
    'Cek pesan Lock
    CMDSQL = "select f_pesanlockauto from usertbl where userid='"
    CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.text) + "'"
    Set m_objrsPesanLock = New ADODB.Recordset
    m_objrsPesanLock.CursorLocation = adUseClient
    m_objrsPesanLock.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If m_objrsPesanLock.RecordCount <> 0 Then
        If m_objrsPesanLock("f_pesanlockauto") = "1" Then
            MsgBox "Lock Data! Ini adalah lock data automatic, data anda akan segera diperbaharui!", vbOKOnly + vbInformation, "Informasi"
            CMDSQL = "update usertbl set f_pesanlockauto=null where userid='"
            CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.text) + "'"
            M_OBJCONN.Execute CMDSQL
            VIEW_MGMDATA.LstVwSearchMgm.ListItems.CLEAR
        End If
     End If
    
    Set m_objrsPesanLock = Nothing
End Sub

'@@ 14022011
Private Sub CekSms()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    '@@ 14/02/2010,, Cek smsnya melalui field blink di usertbl aja, jadinya lebih ringan
    If UCase(Trim(MDIForm1.txtlevel.text)) = "AGENT" Then
        CMDSQL = "select status_sms from usertbl where userid='"
        CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.text) + "'"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_objrs("status_sms") <> "" Then
            TimerBlinkSms.Enabled = True
            LabelSms.Caption = "Ada SMS Baru!"
        Else
            LabelSms.Caption = "Tidak ada SMS baru!"
            LabelSms.ForeColor = vbBlack
            Command2.BackColor = vbGreen
            TimerBlinkSms.Enabled = False
        End If
        
        Set M_objrs = Nothing
    End If
End Sub

'@@ 15-04-2011, Cek CPA , jika ada data cpa  maka kelap-kelip
Private Sub CekCPA()
    Dim strsql As String
    Dim M_objrs As ADODB.Recordset
    
    strsql = "select * from tblcpa where vcustid='" + Trim(lblCustId.text) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        TimerBlinkCPA.Enabled = True
    Else
         TimerBlinkCPA.Enabled = False
    End If
    Set M_objrs = Nothing
End Sub

'@@ 06-May 2011 Tambahan Offering Discon Guide
Private Sub OfferingDiscGuide()
    '@@06 May 2011 Tambahan Offering
        Dim K As Integer
        Dim W As String
        Dim l As Integer
        Dim diskon As Integer
        
        Dim M_objrs As ADODB.Recordset
        Dim m_objrs_waktu As ADODB.Recordset
        Dim CMDSQL As String
              
        
        'Cek dulu ada pembayaran apa ngga di tabel lunas
        CMDSQL = "select * from tbllunas where custid='"
        CMDSQL = CMDSQL + Trim(lblCustId.text) + "'"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        
        'Ambil waktu sekarang
        CMDSQL = "select now() as waktu "
        Set m_objrs_waktu = New ADODB.Recordset
        m_objrs_waktu.CursorLocation = adUseClient
        m_objrs_waktu.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        '@@ 08-06-2011, Jika lblpaydt=kosong on error goto salah
        On Error GoTo Salah
        l = 0
        If Not lblPayDt.ValueIsNull Then
            l = DateDiff("M", Format(lblPayDt.Value, "yyyy-mm-dd"), Format(CDate(m_objrs_waktu("waktu")), "yyyy-mm-dd"))
        End If
        '@@ 09-05-2011 Jika tidak ada nopay atau lpd > 4 bulan dari current date maka
        'tampilkan offering
        
        
        If M_objrs.RecordCount = 0 Or _
            l > 4 Then
            On Error GoTo Salah
            K = DateDiff("M", Format(lblOpenDate.Value, "yyyy-mm-dd"), Format(lblBD.Value, "yyyy-mm-dd"))
            If K < 12 Then
                W = "Penawaran Diskon Maximal 60%"
                diskon = 60
            ElseIf K >= 12 And K <= 17 Then
                W = "Penawaran Diskon Maximal 50%"
                diskon = 50
            ElseIf K >= 18 And K <= 36 Then
                W = "Penawaran Diskon Maximal 40%"
                diskon = 40
            ElseIf K > 37 Then
                W = "Cicilan panjang " & " dan diskon 30%"
                diskon = 30
            End If
        
            'MsgBox "Pemandu Offering: " & w, vbOKOnly + vbInformation, "Offering Disc.Guide..."
            'With FrmOfferingGuide
            With FRMSCRIPT
                On Error Resume Next
                '.LblTextGuide.Caption = "Pemandu Offering: " & W
                .LblTextGuide.Caption = "Pemandu Offering: Cicilan"
                .Tdbbalance.Value = lblAmount.Value
                ' Fixed 40 #12042013 - Joko
                diskon = 40
                .TdbMaxDisc.Value = diskon
                .Show vbModal
            End With
        End If
        
        Set M_objrs = Nothing
        Set m_objrs_waktu = Nothing
        Exit Sub
Salah:
    Set M_objrs = Nothing
    Set m_objrs_waktu = Nothing
    MsgBox "Ada error: " & err.Description
End Sub


'@@ 09092011, Skrip Ofering yang awalnya di FormOfferingGuide, Sekarang Dipindah ke FormScript
Private Sub OfferingDiscGuideNew()
    '@@06 May 2011 Tambahan Offering
        Dim K As Integer
        Dim W As String
        Dim l As Integer
        Dim diskon As Integer
        
        Dim M_objrs As ADODB.Recordset
        Dim m_objrs_waktu As ADODB.Recordset
        Dim CMDSQL As String
              
        
        'Cek dulu ada pembayaran apa ngga di tabel lunas
        CMDSQL = "select * from tbllunas where custid='"
        CMDSQL = CMDSQL + Trim(lblCustId.text) + "'"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        
        'Ambil waktu sekarang
        CMDSQL = "select now() as waktu "
        Set m_objrs_waktu = New ADODB.Recordset
        m_objrs_waktu.CursorLocation = adUseClient
        m_objrs_waktu.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        '@@ 08-06-2011, Jika lblpaydt=kosong on error goto salah
        On Error GoTo Salah
        l = DateDiff("M", Format(lblPayDt.Value, "yyyy-mm-dd"), Format(CDate(m_objrs_waktu("waktu")), "yyyy-mm-dd"))
        
        '@@ 09-05-2011 Jika tidak ada nopay atau lpd > 4 bulan dari current date maka
        'tampilkan offering
        
        
        If M_objrs.RecordCount = 0 Or _
            l > 4 Then
            On Error GoTo Salah
            K = DateDiff("M", Format(lblOpenDate.Value, "yyyy-mm-dd"), Format(lblBD.Value, "yyyy-mm-dd"))
            If K < 12 Then
                W = "Penawaran Diskon Maximal 60%"
                diskon = 60
            ElseIf K >= 12 And K <= 17 Then
                W = "Penawaran Diskon Maximal 50%"
                diskon = 50
            ElseIf K >= 18 And K <= 36 Then
                W = "Penawaran Diskon Maximal 40%"
                diskon = 40
            ElseIf K > 37 Then
                W = "Cicilan panjang " & " dan diskon 30%"
                diskon = 30
            End If
        
            'MsgBox "Pemandu Offering: " & w, vbOKOnly + vbInformation, "Offering Disc.Guide..."
            With FRMSCRIPT
                '.LblTextGuide.Caption = "Pemandu Offering: " & W
                ' Last Update #12042013 Joko by Izuddin
                .LblTextGuide.Caption = "Pemandu Offering: Cicilan"
                .Tdbbalance.Value = lblAmount.Value
                ' Fixed 30 #12042013 - Joko
                diskon = 40
                .TdbMaxDisc.Value = diskon
                '.Show vbModal
            End With
        End If
        
        Set M_objrs = Nothing
        Set m_objrs_waktu = Nothing
        Exit Sub
Salah:
    Set M_objrs = Nothing
    Set m_objrs_waktu = Nothing
End Sub

'@@22-09-2011 Hitung InstallmentPtp
Private Sub HitungInstallmentPtp()
    Dim installment As Double
    
'    If txttenor.Value = 0 Then
'        installment = txtPayment.Value / 1
'    Else
'        installment = txtPayment.Value / txttenor.Value
'    End If
'    Tdabamoint.Value = installment
End Sub

Private Sub txtPayment_Change()
    HitungInstallmentPtp
End Sub

Private Sub txttenor_Change()
    HitungInstallmentPtp
End Sub

Private Sub CariTanggalTagih()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim TglPaymentEffective As String
    
    If IsNull(TDBDate3.Value) = True Then
        MsgBox "Payment effective tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    TglPaymentEffective = Format(TDBDate3.Value, "yyyy-mm-dd")
    
    CMDSQL = "Select  date('" + TglPaymentEffective + "')-"
    If UCase(Trim(CmbViaPtp.text)) = "HSBC" Then
        CMDSQL = CMDSQL + "1"
    ElseIf UCase(Trim(CmbViaPtp.text)) = "BERSAMA" Then
        CMDSQL = CMDSQL + "1"
    ElseIf UCase(Trim(CmbViaPtp.text)) = "KANTOR POS" Then
        CMDSQL = CMDSQL + "3"
    ElseIf UCase(Trim(CmbViaPtp.text)) = "PUM" Then
        CMDSQL = CMDSQL + "1"
    Else
        CMDSQL = CMDSQL + "3"
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    On Error GoTo Salah
    TdbTglTagih.Value = Format(M_objrs(0), "yyyy-mm-dd")
    
    Set M_objrs = Nothing
    Exit Sub
Salah:
    MsgBox "Ada Error: " & err.Description
End Sub

'@@ 17-04-2012, Ini buat hitung durasi call
Private Sub HitungDurasiCall()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim JAM, Menit, Detik As Long
     
    CMDSQL = "select id,enddate-tgl as durasi from tblphonemonitorhst where custid='"
    CMDSQL = CMDSQL + Trim(lblCustId.text) + "' and userid='"
    CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "' order by id desc limit 1"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    DoEvents
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    JAM = Val(Mid(M_objrs("durasi"), 1, 2)) * 3600
    Menit = Val(Mid(M_objrs("durasi"), 4, 2)) * 60
    Detik = Val(Mid(M_objrs("durasi"), 7, 2)) + JAM + Menit
    
    If Detik >= 40 Then
        CMDSQL = "update tblphonemonitorhst set durasi='"
        CMDSQL = CMDSQL + CStr(Detik) + "', flag_review='1' where id='"
        CMDSQL = CMDSQL + CStr(M_objrs("id")) + "'"
    Else
        CMDSQL = "update tblphonemonitorhst set durasi='"
        CMDSQL = CMDSQL + CStr(Detik) + "' where id='"
        CMDSQL = CMDSQL + CStr(M_objrs("id")) + "'"
    End If
    DoEvents
    M_OBJCONN.Execute CMDSQL
    Set M_objrs = Nothing
End Sub

'@@ 19042012,, Buat Hitung Durasi Call dari Icentra
Private Sub HitungDurasiDariIcentra()
    Dim connIcentra As ADODB.Connection
    Dim StrKoneksi As String
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim Initiate, Start, Finish As String
    Dim JAM, Menit, Detik As Long
    
    
    Set connIcentra = New ADODB.Connection
    If Trim(MDIForm1.TxtIPIcentra.text) = "192.168.10.4" Then
       '-- Lokal --
       'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=localhost; PORT=5432; Database=icentra_4; UID=admin; PWD=admin321"
       '-- Database --
       StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.10.4; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    ElseIf Trim(MDIForm1.TxtIPIcentra.text) = "192.168.10.5" Then
       '-- Lokal --
       'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=localhost; PORT=5432; Database=icentra_5; UID=admin; PWD=admin321"
       '-- Database --
       StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.10.5; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    Else
        '@@ 02052012, Jika IP Kosong,, coba dicari dulu di database
        Dim M_Objrs_IP_Icentra As ADODB.Recordset
        
        CMDSQL = "select * from tbl_ip_icentra where ip='"
        CMDSQL = CMDSQL + CStr(MDIForm1.WskCTI.LocalIP) + "'"
        Set M_Objrs_IP_Icentra = New ADODB.Recordset
        M_Objrs_IP_Icentra.CursorLocation = adUseClient
        M_Objrs_IP_Icentra.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_IP_Icentra.RecordCount = 0 Then
            MDIForm1.TxtIPIcentra.text = ""
            Set M_Objrs_IP_Icentra = Nothing
            '@@ Jika IP tidak ditemukan langsung exit, Tapi Cek dulu manual dengan
            'menelusuri server 4 dan 5
            'Call CariIPIcentra
            '@@ 24 May 2012, Cari Berdasarkan Waktu Login aja
            Call CariIPIcentraByWaktuLogin
            Exit Sub
        Else
            MDIForm1.TxtIPIcentra.text = IIf(IsNull(M_Objrs_IP_Icentra("ip_icentra")), "", Trim(M_Objrs_IP_Icentra("ip_icentra")))
            StrKoneksi = "Driver={PostgreSQL ANSI}; Server=" & MDIForm1.TxtIPIcentra.text & "; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
            Set M_Objrs_IP_Icentra = Nothing
        End If
    End If
    '------------ LOKAL ICENTRA --------------------
    'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=localhost; PORT=5432; Database=icentra_4; UID=admin; PWD=admin321"
    '------------ ICENTRA BANDUNG ---------------------
    'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.11.1; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    '------------ ICENTRA SURABAYA ----------------------
    'StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.11.1; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    On Error GoTo Salah
    connIcentra.Open StrKoneksi
    
    '@@15092012 Cek Nomor Telepon yang dicall, jika kosong keluar dari sistem
    If IsNull(txtPhone.text) = True Or txtPhone.text = "" Then
        Exit Sub
    End If
    
    CMDSQL = "select *,finish-start as durasi from acd_log_outgoing_session where destination='"
    CMDSQL = CMDSQL + Trim(Replace(txtPhone.text, " ", "")) + "' and campaign='"
    CMDSQL = CMDSQL + Trim(lblCustId.text) + "' and date(initiate)=date(now()) "
    CMDSQL = CMDSQL + " and start is not null and finish is not null  "
    CMDSQL = CMDSQL + " order by acd_log_outgoing_session_id desc limit 1 "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, connIcentra, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        'Pindahin data dari icentra ke database card
        Initiate = IIf(IsNull(M_objrs("initiate")), "null", "'" & Format(M_objrs("initiate"), "yyyy-mm-dd hh:mm:ss") + "'")
        Start = IIf(IsNull(M_objrs("start")), "null", "'" & Format(M_objrs("start"), "yyyy-mm-dd hh:mm:ss") + "'")
        Finish = IIf(IsNull(M_objrs("finish")), "null", "'" & Format(M_objrs("finish"), "yyyy-mm-dd hh:mm:ss") + "'")
        
        'Hitung Konevrsi Selisih ke detik
        JAM = Val(Mid(M_objrs("durasi"), 1, 2)) * 3600
        Menit = Val(Mid(M_objrs("durasi"), 4, 2)) * 60
        Detik = Val(Mid(M_objrs("durasi"), 7, 2)) + JAM + Menit
        
        CMDSQL = "insert into outgoing_icentra (destination,"
        CMDSQL = CMDSQL + "initiate,start,finish,recording_filename,"
        CMDSQL = CMDSQL + "custid,durasi,agent,acd_log_outgoing_session_id) values ('"
        CMDSQL = CMDSQL + IIf(IsNull(M_objrs("destination")), "", CStr(M_objrs("destination"))) + "',"
        CMDSQL = CMDSQL + Initiate + "," + Start + "," + Finish + ",'"
        CMDSQL = CMDSQL + IIf(IsNull(M_objrs("recording_filename")), "", CStr(M_objrs("recording_filename"))) + "','"
        CMDSQL = CMDSQL + IIf(IsNull(M_objrs("campaign")), "", CStr(M_objrs("campaign"))) + "','"
        CMDSQL = CMDSQL + CStr(Detik) + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
        CMDSQL = CMDSQL + CStr(M_objrs("acd_log_outgoing_session_id")) + "')"
        M_OBJCONN.Execute CMDSQL
    End If
    
    Set M_objrs = Nothing
    Set connIcentra = Nothing
    Exit Sub
Salah:
    Exit Sub
    'MsgBox "Anda tidak terhubung ke Icentra!", vbOKOnly + vbInformation, "Informasi"
    
End Sub

'@@ 02052012, Tambahkan Pilihan Speak With
Private Sub PilihSpeakWith()
    cbolastcall.CLEAR
    If UCase(Trim(TxtTelpKe.text)) = "OTHER CH OFFICE" Or _
       StsKategoriTelepon = "OTHER CH OFFICE" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "Reception/Operator/Sec/OB"
        cbolastcall.AddItem "Atasan"
        cbolastcall.AddItem "HRD"
        cbolastcall.AddItem "Teman kantor"
    End If
    If UCase(Trim(TxtTelpKe.text)) = "OTHER CH HOME" Or _
       StsKategoriTelepon = "OTHER CH HOME" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "Orang Tua"
        cbolastcall.AddItem "Kakak/Adik/Anak"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Keluarga Dekat Lainnya"
        cbolastcall.AddItem "Ex Spouse"
        cbolastcall.AddItem "Pembantu/Supir"
        cbolastcall.AddItem "Kontrakan"
        cbolastcall.AddItem "Other"
    End If
    If UCase(Trim(TxtTelpKe.text)) = "FAMILY" Or _
       StsKategoriTelepon = "FAMILY" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "Orang Tua"
        cbolastcall.AddItem "Kakak/Adik/Anak"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Keluarga Dekat Lainnya"
        cbolastcall.AddItem "Ex Spouse"
        cbolastcall.AddItem "Pembantu/Supir"
    End If
    If UCase(Trim(TxtTelpKe.text)) = "NEIGHBOUR" Or _
       StsKategoriTelepon = "NEIGHBOUR" Then
        cbolastcall.AddItem "Tetangga"
        cbolastcall.AddItem "Pengurus Lingkungan"
        cbolastcall.AddItem "Pembantu/Supir"
    End If
    If UCase(Trim(TxtTelpKe.text)) = "RELATED PERSON" Or _
       StsKategoriTelepon = "RELATED PERSON" Then
        cbolastcall.AddItem "Lawyer"
        cbolastcall.AddItem "Teman"
        cbolastcall.AddItem "Other"
        cbolastcall.AddItem "Reception/Operator/Sec/OB"
        cbolastcall.AddItem "Atasan"
        cbolastcall.AddItem "HRD"
        cbolastcall.AddItem "Teman kantor"
        cbolastcall.AddItem "Orang Tua"
        cbolastcall.AddItem "Kakak/Adik/Anak"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Keluarga Dekat Lainnya"
        cbolastcall.AddItem "Ex Spouse"
        cbolastcall.AddItem "Tetangga"
        cbolastcall.AddItem "Pengurus Lingkungan"
        cbolastcall.AddItem "Pembantu/Supir"
    End If
    
        
    If UCase(Trim(TxtTelpKe.text)) = "OTHER CH MOBILE" Or _
        StsKategoriTelepon = "OTHER CH MOBILE" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "SPOUSE"
        cbolastcall.AddItem "OTHER"
    End If
    
    If UCase(Trim(TxtTelpKe.text)) = "HOMEPHONE" Or _
       UCase(Trim(TxtTelpKe.text)) = "HOMEPHONE2" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "Orang Tua"
        cbolastcall.AddItem "Kakak/Adik/Anak"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Keluarga Dekat Lainnya"
        cbolastcall.AddItem "Ex Spouse"
        cbolastcall.AddItem "Pembantu/Supir"
        cbolastcall.AddItem "Kontrakan"
        cbolastcall.AddItem "Other"
    End If
    
    If UCase(Trim(TxtTelpKe.text)) = "OFFICEPHONE" Or _
       UCase(Trim(TxtTelpKe.text)) = "OFFICEPHONE2" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "Reception/Operator/Sec/OB"
        cbolastcall.AddItem "Atasan"
        cbolastcall.AddItem "HRD"
        cbolastcall.AddItem "Teman Kantor"
    End If
    If UCase(Trim(TxtTelpKe.text)) = "ECONPHONE" Or _
       UCase(Trim(TxtTelpKe.text)) = "ECONPHONE" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "EC"
        cbolastcall.AddItem "LAWYER"
        cbolastcall.AddItem "Teman"
        cbolastcall.AddItem "OTHER"
        cbolastcall.AddItem "Reception/Operator/Sec/OB"
        cbolastcall.AddItem "Atasan"
        cbolastcall.AddItem "HRD"
        cbolastcall.AddItem "Teman Kantor"
        cbolastcall.AddItem "Orang Tua"
        cbolastcall.AddItem "Kakak/Adik/Anak"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Keluarga Dekat Lainnya"
        cbolastcall.AddItem "Ex Spouse"
        cbolastcall.AddItem "Tetangga"
        cbolastcall.AddItem "Pengurus Lingkungan"
        cbolastcall.AddItem "Pembantu/Supir"
    End If
    
    If UCase(Trim(TxtTelpKe.text)) = "HP" Or _
       UCase(Trim(TxtTelpKe.text)) = "HP2" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Other"
    End If
    
    
    If UCase(Trim(TxtTelpKe.text)) = "OTHER EC" Or _
       StsKategoriTelepon = "OTHER EC" Then
        cbolastcall.AddItem "CH"
        cbolastcall.AddItem "EC"
        cbolastcall.AddItem "LAWYER"
        cbolastcall.AddItem "Teman"
        cbolastcall.AddItem "OTHER"
        cbolastcall.AddItem "Reception/Operator/Sec/OB"
        cbolastcall.AddItem "Atasan"
        cbolastcall.AddItem "HRD"
        cbolastcall.AddItem "Teman Kantor"
        cbolastcall.AddItem "Orang Tua"
        cbolastcall.AddItem "Kakak/Adik/Anak"
        cbolastcall.AddItem "Spouse"
        cbolastcall.AddItem "Keluarga Dekat Lainnya"
        cbolastcall.AddItem "Ex Spouse"
        cbolastcall.AddItem "Tetangga"
        cbolastcall.AddItem "Pengurus Lingkungan"
        cbolastcall.AddItem "Pembantu/Supir"
    End If
    
    cbolastcall.AddItem "UnReceive"
    
End Sub

Private Sub CariKategoriTlp()
    If StsKategoriTelepon = "OTHER CH OFFICE" Then
        KelompokKategoriTlp = "OCO"
    ElseIf StsKategoriTelepon = "OTHER CH HOME" Then
        KelompokKategoriTlp = "OCH"
    ElseIf StsKategoriTelepon = "FAMILY" Then
        KelompokKategoriTlp = "FAM"
    ElseIf StsKategoriTelepon = "NEIGHBOUR" Then
        KelompokKategoriTlp = "NEB"
    ElseIf StsKategoriTelepon = "RELATED PERSON" Then
        KelompokKategoriTlp = "RLP"
    ElseIf StsKategoriTelepon = "OTHER EC" Then
        KelompokKategoriTlp = "OEC"
    ElseIf StsKategoriTelepon = "OTHER CH MOBILE" Then
        KelompokKategoriTlp = "OCM"
    ElseIf StsKategoriTelepon = "HP" Then
        KelompokKategoriTlp = "HP"
    ElseIf StsKategoriTelepon = "Home" Then
        KelompokKategoriTlp = "HOME"
    ElseIf StsKategoriTelepon = "Office" Then
        KelompokKategoriTlp = "OFF"
    ElseIf StsKategoriTelepon = "EC" Then
        KelompokKategoriTlp = "EC"
    End If
End Sub

''@@ 16 May 2012, Khusus HSBC JAKARTA
Private Sub CariIPIcentra()
    Dim connIcentra As ADODB.Connection
    Dim StrKoneksi As String
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim Initiate, Start, Finish As String
    Dim JAM, Menit, Detik As Long
    
    '@@ Cek Ke server 4 dulu ---------------------------------------------------------------------------
    StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.10.4; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    On Error GoTo Salah
    connIcentra.Open StrKoneksi
    
    CMDSQL = "select *,finish-start as durasi from acd_log_outgoing_session where destination='"
    CMDSQL = CMDSQL + Trim(Replace(txtPhone.text, " ", "")) + "' and campaign='"
    CMDSQL = CMDSQL + Trim(lblCustId.text) + "' and date(initiate)=date(now()) "
    CMDSQL = CMDSQL + " and start is not null and finish is not null  "
    CMDSQL = CMDSQL + " order by acd_log_outgoing_session_id desc limit 1 "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, connIcentra, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        'Pindahin data dari icentra ke database card
        Initiate = IIf(IsNull(M_objrs("initiate")), "null", "'" & Format(M_objrs("initiate"), "yyyy-mm-dd hh:mm:ss") + "'")
        Start = IIf(IsNull(M_objrs("start")), "null", "'" & Format(M_objrs("start"), "yyyy-mm-dd hh:mm:ss") + "'")
        Finish = IIf(IsNull(M_objrs("finish")), "null", "'" & Format(M_objrs("finish"), "yyyy-mm-dd hh:mm:ss") + "'")
        
        'Hitung Konevrsi Selisih ke detik
        JAM = Val(Mid(M_objrs("durasi"), 1, 2)) * 3600
        Menit = Val(Mid(M_objrs("durasi"), 4, 2)) * 60
        Detik = Val(Mid(M_objrs("durasi"), 7, 2)) + JAM + Menit
        
        CMDSQL = "insert into outgoing_icentra (destination,"
        CMDSQL = CMDSQL + "initiate,start,finish,recording_filename,"
        CMDSQL = CMDSQL + "custid,durasi,agent,acd_log_outgoing_session_id) values ('"
        CMDSQL = CMDSQL + IIf(IsNull(M_objrs("destination")), "", CStr(M_objrs("destination"))) + "',"
        CMDSQL = CMDSQL + Initiate + "," + Start + "," + Finish + ",'"
        CMDSQL = CMDSQL + IIf(IsNull(M_objrs("recording_filename")), "", CStr(M_objrs("recording_filename"))) + "','"
        CMDSQL = CMDSQL + IIf(IsNull(M_objrs("campaign")), "", CStr(M_objrs("campaign"))) + "','"
        CMDSQL = CMDSQL + CStr(Detik) + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
        CMDSQL = CMDSQL + CStr(M_objrs("acd_log_outgoing_session_id")) + "')"
        M_OBJCONN.Execute CMDSQL
        
        MDIForm1.TxtIPIcentra.text = "192.168.10.4"
        
        Set M_objrs = Nothing
        Set connIcentra = Nothing
        Exit Sub
    End If
    Set M_objrs = Nothing
    Set connIcentra = Nothing
    
    '-------------------------------------------------------------------------------------
    
    '---- Cek Server 5 -------------------------------------------------------------------
    StrKoneksi = "Driver={PostgreSQL ANSI}; Server=192.168.10.5; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    On Error GoTo Salah
    connIcentra.Open StrKoneksi
    
    CMDSQL = "select *,finish-start as durasi from acd_log_outgoing_session where destination='"
    CMDSQL = CMDSQL + Trim(Replace(txtPhone.text, " ", "")) + "' and campaign='"
    CMDSQL = CMDSQL + Trim(lblCustId.text) + "' and date(initiate)=date(now()) "
    CMDSQL = CMDSQL + " and start is not null and finish is not null  "
    CMDSQL = CMDSQL + " order by acd_log_outgoing_session_id desc limit 1 "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, connIcentra, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        'Pindahin data dari icentra ke database card
        Initiate = IIf(IsNull(M_objrs("initiate")), "null", "'" & Format(M_objrs("initiate"), "yyyy-mm-dd hh:mm:ss") + "'")
        Start = IIf(IsNull(M_objrs("start")), "null", "'" & Format(M_objrs("start"), "yyyy-mm-dd hh:mm:ss") + "'")
        Finish = IIf(IsNull(M_objrs("finish")), "null", "'" & Format(M_objrs("finish"), "yyyy-mm-dd hh:mm:ss") + "'")
        
        'Hitung Konevrsi Selisih ke detik
        JAM = Val(Mid(M_objrs("durasi"), 1, 2)) * 3600
        Menit = Val(Mid(M_objrs("durasi"), 4, 2)) * 60
        Detik = Val(Mid(M_objrs("durasi"), 7, 2)) + JAM + Menit
        
        CMDSQL = "insert into outgoing_icentra (destination,"
        CMDSQL = CMDSQL + "initiate,start,finish,recording_filename,"
        CMDSQL = CMDSQL + "custid,durasi,agent,acd_log_outgoing_session_id) values ('"
        CMDSQL = CMDSQL + IIf(IsNull(M_objrs("destination")), "", CStr(M_objrs("destination"))) + "',"
        CMDSQL = CMDSQL + Initiate + "," + Start + "," + Finish + ",'"
        CMDSQL = CMDSQL + IIf(IsNull(M_objrs("recording_filename")), "", CStr(M_objrs("recording_filename"))) + "','"
        CMDSQL = CMDSQL + IIf(IsNull(M_objrs("campaign")), "", CStr(M_objrs("campaign"))) + "','"
        CMDSQL = CMDSQL + CStr(Detik) + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
        CMDSQL = CMDSQL + CStr(M_objrs("acd_log_outgoing_session_id")) + "')"
        M_OBJCONN.Execute CMDSQL
        
        MDIForm1.TxtIPIcentra.text = "192.168.10.5"
    End If
    Set M_objrs = Nothing
    Set connIcentra = Nothing
    Exit Sub
Salah:
    Exit Sub
    'MsgBox "Maaf anda tidak terhubung ke Icentra!", vbOKOnly + vbInformation, "Informasi"
End Sub

'@@ 21 May 2012, Tambahan Buat bikin beberapa baris  dari remarks
Private Function Ceiling(number As Double) As Long
    Ceiling = -Int(-number)
End Function

'@@ 24 May 2012, Mencari IP Centra Berdasarkan Waktu Login
Private Sub CariIPIcentraByWaktuLogin()
    Dim KoneksiIcentra As ADODB.Connection
    Dim StrKoneksiIcentra As String
    Dim M_Objrs_Icentra As ADODB.Recordset
    Dim M_Objrs_Telp As ADODB.Recordset
    Dim Initiate, Start, Finish As String
    Dim JAM, Menit, Detik As Long
    
    Set KoneksiIcentra = New ADODB.Connection
    
    'Cek di Server4 Dulu
    StrKoneksiIcentra = "Driver={PostgreSQL ANSI}; Server=192.168.10.4; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    On Error GoTo Salah
    KoneksiIcentra.Open StrKoneksiIcentra
    CMDSQL = "select * from acd_log_agent_session,acd_agent where "
    CMDSQL = CMDSQL + " acd_log_agent_session.acd_agent_id=acd_agent.acd_agent_id "
    CMDSQL = CMDSQL + " and acd_agent.name='"
    CMDSQL = CMDSQL + Trim(Replace(MDIForm1.TxtUsername.text, "TL", "TLCARD")) + "' "
    CMDSQL = CMDSQL + " and date(login_time)=date(now()) limit 1 "
    Set M_Objrs_Icentra = New ADODB.Recordset
    M_Objrs_Icentra.CursorLocation = adUseClient
    DoEvents
    M_Objrs_Icentra.Open CMDSQL, KoneksiIcentra, adOpenDynamic, adLockOptimistic, adCmdText
        
    If M_Objrs_Icentra.RecordCount > 0 Then
        MDIForm1.TxtIPIcentra.text = "192.168.10.4"
        
        '@@15092012 Cek Nomor Telepon yang dicall, jika kosong keluar dari sistem
        If IsNull(txtPhone.text) = True Or txtPhone.text = "" Then
            Exit Sub
        End If
        
        'Cari No Telepon yang terakhir
        CMDSQL = "select *,finish-start as durasi from acd_log_outgoing_session where destination='"
        CMDSQL = CMDSQL + Trim(Replace(txtPhone.text, " ", "")) + "' and campaign='"
        CMDSQL = CMDSQL + Trim(lblCustId.text) + "' and date(initiate)=date(now()) "
        CMDSQL = CMDSQL + " and start is not null and finish is not null  "
        CMDSQL = CMDSQL + " order by acd_log_outgoing_session_id desc limit 1 "
        Set M_Objrs_Telp = New ADODB.Recordset
        M_Objrs_Telp.CursorLocation = adUseClient
        DoEvents
        M_Objrs_Telp.Open CMDSQL, KoneksiIcentra, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Telp.RecordCount > 0 Then
            'Pindahin data dari icentra ke database card
            Initiate = IIf(IsNull(M_Objrs_Telp("initiate")), "null", "'" & Format(M_Objrs_Telp("initiate"), "yyyy-mm-dd hh:mm:ss") + "'")
            Start = IIf(IsNull(M_Objrs_Telp("start")), "null", "'" & Format(M_Objrs_Telp("start"), "yyyy-mm-dd hh:mm:ss") + "'")
            Finish = IIf(IsNull(M_Objrs_Telp("finish")), "null", "'" & Format(M_Objrs_Telp("finish"), "yyyy-mm-dd hh:mm:ss") + "'")
            
            'Hitung Konevrsi Selisih ke detik
            JAM = Val(Mid(M_Objrs_Telp("durasi"), 1, 2)) * 3600
            Menit = Val(Mid(M_Objrs_Telp("durasi"), 4, 2)) * 60
            Detik = Val(Mid(M_Objrs_Telp("durasi"), 7, 2)) + JAM + Menit
            
            CMDSQL = "insert into outgoing_icentra (destination,"
            CMDSQL = CMDSQL + "initiate,start,finish,recording_filename,"
            CMDSQL = CMDSQL + "custid,durasi,agent,acd_log_outgoing_session_id) values ('"
            CMDSQL = CMDSQL + IIf(IsNull(M_Objrs_Telp("destination")), "", CStr(M_Objrs_Telp("destination"))) + "',"
            CMDSQL = CMDSQL + Initiate + "," + Start + "," + Finish + ",'"
            CMDSQL = CMDSQL + IIf(IsNull(M_Objrs_Telp("recording_filename")), "", CStr(M_Objrs_Telp("recording_filename"))) + "','"
            CMDSQL = CMDSQL + IIf(IsNull(M_Objrs_Telp("campaign")), "", CStr(M_Objrs_Telp("campaign"))) + "','"
            CMDSQL = CMDSQL + CStr(Detik) + "','"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
            CMDSQL = CMDSQL + CStr(M_Objrs_Telp("acd_log_outgoing_session_id")) + "')"
            M_OBJCONN.Execute CMDSQL
            
            Set M_Objrs_Telp = Nothing
            Set M_Objrs_Icentra = Nothing
            Set KoneksiIcentra = Nothing
            Exit Sub
        End If
    End If
    Set M_Objrs_Icentra = Nothing
    Set KoneksiIcentra = Nothing
    
    '/////////////////////////////----------- Server 5 ----------------------------------------
    Set KoneksiIcentra = New ADODB.Connection
    StrKoneksiIcentra = "Driver={PostgreSQL ANSI}; Server=192.168.10.5; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
    On Error GoTo Salah
    KoneksiIcentra.Open StrKoneksiIcentra
    CMDSQL = "select * from acd_log_agent_session,acd_agent where "
    CMDSQL = CMDSQL + " acd_log_agent_session.acd_agent_id=acd_agent.acd_agent_id "
    CMDSQL = CMDSQL + " and acd_agent.name='"
    CMDSQL = CMDSQL + Trim(Replace(MDIForm1.TxtUsername.text, "TL", "TLCARD")) + "' "
    CMDSQL = CMDSQL + " and date(login_time)=date(now()) limit 1 "
    Set M_Objrs_Icentra = New ADODB.Recordset
    M_Objrs_Icentra.CursorLocation = adUseClient
    DoEvents
    M_Objrs_Icentra.Open CMDSQL, KoneksiIcentra, adOpenDynamic, adLockOptimistic, adCmdText
        
CariDiServer5:
    If M_Objrs_Icentra.RecordCount > 0 Then
        MDIForm1.TxtIPIcentra.text = "192.168.10.5"
        
        'Cari No Telepon yang terakhir
        CMDSQL = "select *,finish-start as durasi from acd_log_outgoing_session where destination='"
        CMDSQL = CMDSQL + Trim(Replace(txtPhone.text, " ", "")) + "' and campaign='"
        CMDSQL = CMDSQL + Trim(lblCustId.text) + "' and date(initiate)=date(now()) "
        CMDSQL = CMDSQL + " and start is not null and finish is not null  "
        CMDSQL = CMDSQL + " order by acd_log_outgoing_session_id desc limit 1 "
        Set M_Objrs_Telp = New ADODB.Recordset
        M_Objrs_Telp.CursorLocation = adUseClient
        DoEvents
        M_Objrs_Telp.Open CMDSQL, KoneksiIcentra, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Telp.RecordCount > 0 Then
            'Pindahin data dari icentra ke database card
            Initiate = IIf(IsNull(M_Objrs_Telp("initiate")), "null", "'" & Format(M_Objrs_Telp("initiate"), "yyyy-mm-dd hh:mm:ss") + "'")
            Start = IIf(IsNull(M_Objrs_Telp("start")), "null", "'" & Format(M_Objrs_Telp("start"), "yyyy-mm-dd hh:mm:ss") + "'")
            Finish = IIf(IsNull(M_Objrs_Telp("finish")), "null", "'" & Format(M_Objrs_Telp("finish"), "yyyy-mm-dd hh:mm:ss") + "'")
            
            'Hitung Konevrsi Selisih ke detik
            JAM = Val(Mid(M_Objrs_Telp("durasi"), 1, 2)) * 3600
            Menit = Val(Mid(M_Objrs_Telp("durasi"), 4, 2)) * 60
            Detik = Val(Mid(M_Objrs_Telp("durasi"), 7, 2)) + JAM + Menit
            
            CMDSQL = "insert into outgoing_icentra (destination,"
            CMDSQL = CMDSQL + "initiate,start,finish,recording_filename,"
            CMDSQL = CMDSQL + "custid,durasi,agent,acd_log_outgoing_session_id) values ('"
            CMDSQL = CMDSQL + IIf(IsNull(M_Objrs_Telp("destination")), "", CStr(M_Objrs_Telp("destination"))) + "',"
            CMDSQL = CMDSQL + Initiate + "," + Start + "," + Finish + ",'"
            CMDSQL = CMDSQL + IIf(IsNull(M_Objrs_Telp("recording_filename")), "", CStr(M_Objrs_Telp("recording_filename"))) + "','"
            CMDSQL = CMDSQL + IIf(IsNull(M_Objrs_Telp("campaign")), "", CStr(M_Objrs_Telp("campaign"))) + "','"
            CMDSQL = CMDSQL + CStr(Detik) + "','"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
            CMDSQL = CMDSQL + CStr(M_Objrs_Telp("acd_log_outgoing_session_id")) + "')"
            M_OBJCONN.Execute CMDSQL
            
            Set M_Objrs_Telp = Nothing
            Set M_Objrs_Icentra = Nothing
            Set KoneksiIcentra = Nothing
            Exit Sub
        End If
    End If
    Set M_Objrs_Icentra = Nothing
    Set KoneksiIcentra = Nothing
    Exit Sub
Salah:
    Exit Sub
    
End Sub

'@@09072012 Cari Buat Masukin data Contact LPD /Contact Rate
Private Sub CariContactRate()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim SelisihTanggal As String
    
'    'Cari data terakhir pembayaran
'    CMDSQL = "select (date(now())-paydate) as hari from tbllunas where custid='"
'    CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.text) + "' order by paydate desc limit 1"
'    Set M_OBJRS = New ADODB.Recordset
'    M_OBJRS.CursorLocation = adUseClient
'    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If M_OBJRS.RecordCount > 0 Then
'        If Val(M_OBJRS("hari")) < 30 Then
'            'Update f_contact_rate
'            CMDSQL = "update mgm set f_contact_rate='CONTACT LPD 1' where custid='"
'            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.text) + "'"
'            M_OBJCONN.Execute CMDSQL
'        End If
'        If Val(M_OBJRS("hari")) > 30 Then
'            'Update f_contact_rate
'            CMDSQL = "update mgm set f_contact_rate='CONTACT > LPD 1 ' where custid='"
'            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.text) + "'"
'            M_OBJCONN.Execute CMDSQL
'        End If
'    End If
'    Set M_OBJRS = Nothing

    '@@10-07-2012 Ambil Dari Listview1(0)
    If listview1(0).ListItems.Count > 0 Then
        CMDSQL = "select (date(now())- '"
        CMDSQL = CMDSQL + Format(listview1(0).ListItems(1).text, "yyyy-mm-dd") + "') as hari   "
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If Val(M_objrs("hari")) < 30 Then
            'Update f_contact_rate
            CMDSQL = "update mgm set f_contact_rate='CONTACT LPD 1' where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.text) + "'"
            M_OBJCONN.Execute CMDSQL
        End If
        If Val(M_objrs("hari")) > 30 Then
            'Update f_contact_rate
            CMDSQL = "update mgm set f_contact_rate='CONTACT > LPD 1' where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.text) + "'"
            M_OBJCONN.Execute CMDSQL
        End If
        Set M_objrs = Nothing
    Else
        'Update f_contact_rate
        CMDSQL = "update mgm set f_contact_rate='No LPD' where custid='"
        CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.text) + "'"
        M_OBJCONN.Execute CMDSQL
    End If
End Sub
Private Sub updaterrd()
    Dim CMDSQL As String
    
    CMDSQL = "update tblrrd set sstatus_akhir = '" + cboaccount.text + "' where custid = '" + lblCustId.text + "' and agent = '" + MDIForm1.TxtUsername.text + "' and start_time = '" + getservertime.text + "'"
    'M_OBJCONN.Execute CMDSQL
    CMDSQL = "update tblrrd set stop_time = '" + waktu_server_sekarang + "' where custid = '" + lblCustId.text + "' and agent = '" + MDIForm1.TxtUsername.text + "'and start_time = '" + getservertime.text + "'"
    'M_OBJCONN.Execute CMDSQL
    
End Sub

'@@27092012 Setiap Menekan tombol Call, disimpan ke dalam remarks
Private Sub SimpanRemarksCall()
    Dim StatusRemarks As String
    Dim CMDSQL As String
    
    StatusRemarks = IIf(IsNull(KelompokKategoriTlp), "", KelompokKategoriTlp) & "/"
    StatusRemarks = StatusRemarks & StatusSpeakWith & "/"
    StatusRemarks = StatusRemarks & IIf(IsNull(StatusAccount), "", StatusAccount) & ": " & kat_aktif_telp
    StatusRemarks = StatusRemarks & "[Auto by System] -> No Answer / NBPU"
    
    'M_DATA.ADD_HISTORY lblCustId.text, MDIForm1.TDBDate1.Text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", txtremarks.Text, CStr(pStatusLstCall), "", "", IIf(IsNull(StatusAccount), "", StatusAccount), Combo1.Text, Combo1.Text, IIf(IsNull(StatusAccount), "", StatusAccount), IIf(IsNull(StatusAccount), "", StatusAccount), cbolastcall.Text, mdiform1.txtusername.text, "", "0"
    
'    cmdsql = "insert into mgm_hst "
'    cmdsql = cmdsql + " (custid,agent,hst,tgl,kodeds,phoneno,user_log,stop_time) values ('"
'    cmdsql = cmdsql + CStr(Trim(lblCustId.text)) + "','"
'    cmdsql = cmdsql + CStr(Trim(lblaoc.Caption)) + "','" + CStr(StatusRemarks) + "','" & Format(lbltime_save.Caption, "yyyy-mm-dd hh:mm:ss") & "','"
'    cmdsql = cmdsql + IIf(IsNull(StatusAccount), "", StatusAccount) + "','"
'    cmdsql = cmdsql + CStr(txtPhone.Text) + "','"
'    cmdsql = cmdsql + CStr(Trim(mdiform1.txtusername.text)) + "','" & Format(lblstop_time.Caption, "yyyy-mm-dd hh:mm:ss") & "')"

     'jejaktian 02022016
    CMDSQL = "insert into mgm_hst "
    CMDSQL = CMDSQL + " (custid,agent,hst,tgl,kodeds,phoneno,user_log,start_time,stop_time,callwith) values ('"
    CMDSQL = CMDSQL + CStr(Trim(lblCustId.text)) + "','"
    CMDSQL = CMDSQL + CStr(Trim(lblaoc.Caption)) + "','" + CStr(StatusRemarks) + "','" & Format(lbltime_save.Caption, "yyyy-mm-dd hh:mm:ss") & "','"
    CMDSQL = CMDSQL + IIf(IsNull(StatusAccount), "", StatusAccount) + "','"
    CMDSQL = CMDSQL + CStr(txtPhone.text) + "','"
    CMDSQL = CMDSQL + CStr(Trim(MDIForm1.TxtUsername.text)) + "','" & Format(lbltime_save.Caption, "yyyy-mm-dd hh:mm:ss") & "','" & Format(lblstop_time.Caption, "yyyy-mm-dd hh:mm:ss") & "')"

'    cmdsql = "update mgm_hst set stop_time = '" & Format(lblstop_time.Caption, "yyyy-mm-dd hh:mm:ss") & "' where custid = '" & lblCustId.text & "' AND start_time = (select max(start_time) from mgm_hst)"
    M_OBJCONN.Execute CMDSQL
End Sub


'@@19022013 Ini buat pemberian pesan kepada agent kalo dia mempunyai akses untuk semua account
Private Sub CekAksessAllAcc()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    On Error GoTo Salah
    If UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Or _
       UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or _
       UCase(MDIForm1.txtlevel.text) = "ADMIN" Then
        Exit Sub
    End If
    
    DoEvents
    
    ' # Unset account monitor_akses
'    Cmdsql = "update mgm set monitor_akses=null"
'    Cmdsql = Cmdsql + ",waktu_akses=null where custid='" & Trim(lblCustId.text) & "'"
'    M_OBJCONN.Execute Cmdsql
    
    CMDSQL = "select * from tbl_cust_aksesall WHERE kd_profile in " & _
            "(SELECT a.kd_profile FROM tbl_profile_aksesall a, usertbl b WHERE a.kd_profile=b.profile_akses_all " & _
            " AND b.userid='"
    CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "' AND a.waktu_awal < now() and "
    CMDSQL = CMDSQL + " a.waktu_akhir > now() )"
    
    'cek di tabel distribusi
'    Cmdsql = "select * from tbl_distribusi_account where agent='"
'    Cmdsql = Cmdsql + mdiform1.txtusername.text + "' and waktu_awal < now() and "
'    Cmdsql = Cmdsql + " waktu_akhir > now() "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount > 0 Then
        'cek akses allnya
        If AksesAllAcc <> "1" Then
            'update di f_pesanresetauto nya
            CMDSQL = "update usertbl set f_akses_all_acc='1',f_pesanresetauto='1' where "
            CMDSQL = CMDSQL + " userid='"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "'"
            M_OBJCONN.Execute CMDSQL
            AksesAllAcc = "1"
        End If
    Else
'        Cmdsql = "DELETE FROM tbl_cust_aksesall WHERE kd_profile in " & _
'                "(SELECT a.kd_profile FROM tbl_profile_aksesall a, usertbl b WHERE a.kd_profile=b.profile_akses_all " & _
'                " AND b.userid='"
'        Cmdsql = Cmdsql + mdiform1.txtusername.text + "' AND "
'        Cmdsql = Cmdsql + " a.waktu_akhir < now() )"
'        M_OBJCONN.Execute Cmdsql
        ' Balikkin ke agent asli -------
        ' UPDATE 26 JULI 2013 BY IZUDDIN
'        cmdsql = "update mgm set agent=agent_asli,agent_asli=null WHERE monitor_akses is null" & _
'                " AND agent='AKSESALL' AND custid in(SELECT custid FROM tbl_cust_aksesall a,tbl_profile_aksesall b WHERE " & _
'                " a.kd_profile=b.kd_profile AND b.waktu_akhir < now()) AND agent_asli is not null"
        ' UPDATE 30 OKT 2013 BY IZUDDIN
        ' 19 AGUSTUS 2014 dihilangkan agent_asli=null
        CMDSQL = "UPDATE mgm SET agent=agent_asli WHERE " & _
                " agent='AKSESALL' AND custid in(SELECT custid FROM tbl_cust_aksesall a,tbl_profile_aksesall b WHERE " & _
                " a.kd_profile=b.kd_profile AND b.waktu_akhir < now()) AND agent_asli is not null"
        M_OBJCONN.Execute CMDSQL
        
        CMDSQL = "DELETE FROM tbl_cust_aksesall "
        CMDSQL = CMDSQL & " WHERE kd_profile in (SELECT kd_profile FROM tbl_profile_aksesall WHERE waktu_akhir < now()) "
        M_OBJCONN.Execute CMDSQL
        ' -----------------------------
        AksesAllAcc = ""
    End If
        
    Set M_objrs = Nothing
    Exit Sub
Salah:
    MsgBox "Mohon maaf ada error! " & err.Description, vbOKOnly + vbExclamation, "Pesan error"
    
End Sub

Private Sub INSERT_TEMP_SEGMENT_CALL()
    Dim sQuery As String
    Dim iQuery As String
    Dim Rs_Cek_Segment As ADODB.Recordset
    Dim Rs_Temp_Jumlah_Call As ADODB.Recordset
    Dim nomor_telpon As String
    Dim jumlah_call As Double
    
    nomor_telpon = GetNumber(CStr(Replace(txtPhone.text, " ", "")))
    
    sQuery = "SELECT no_telpon, tgl_call FROM tbl_temp_segment_call WHERE date(tgl_call) = '" & Format(waktu_server_sekarang, "yyyy-mm-dd") & "'"
    sQuery = sQuery + " AND no_telpon = '" & nomor_telpon & "' "
    Set Rs_Cek_Segment = New ADODB.Recordset
    Rs_Cek_Segment.CursorLocation = adUseClient
    Rs_Cek_Segment.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Rs_Cek_Segment.RecordCount > 0 Then
        sQuery = "SELECT id, jumlah_call FROM tbl_temp_segment_call WHERE date(tgl_call) = '" & Format(waktu_server_sekarang, "yyyy-mm-dd") & "' "
        sQuery = sQuery + " AND no_telpon = '" & nomor_telpon & "' "
        Set Rs_Temp_Jumlah_Call = New ADODB.Recordset
        Rs_Temp_Jumlah_Call.CursorLocation = adUseClient
        Rs_Temp_Jumlah_Call.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
         
        jumlah_call = Rs_Temp_Jumlah_Call!jumlah_call
        
        jumlah_call = jumlah_call + 1
        
        M_OBJCONN.Execute "UPDATE tbl_temp_segment_call SET jumlah_call = '" & jumlah_call & "' WHERE id = '" & Rs_Temp_Jumlah_Call!ID & "' "
    Else
        M_OBJCONN.Execute "INSERT INTO tbl_temp_segment_call(no_telpon, tgl_call, tipe_segment, jumlah_call) " & _
                          " VALUES ('" & nomor_telpon & "','" & waktu_server_sekarang & "', " & _
                          " '" & Label14(0).Caption & "', '1')"
    End If
End Sub

Private Sub INSERT_TEMP_TELFON_REVIEW()
    Dim sQuery, iQuery, nomor_telpon, CustId, tanggal_telfon, AGENT As String
    Dim Rs_Cek_Tanggal As ADODB.Recordset
    Dim jumlah_call As Double
    
    nomor_telpon = GetNumber(CStr(Replace(txtPhone.text, " ", "")))
    CustId = Trim(FrmCC_Colection.lblCustId.text)
    tanggal_telfon = Format(waktu_server_sekarang, "YYYY-MM-DD")
    AGENT = MDIForm1.TxtUsername.text
    
    sQuery = "SELECT * FROM tbl_temp_telfon_review WHERE no_telfon = '" & nomor_telpon & "'"
    sQuery = sQuery + " AND date(tanggal_telfon) = '" & tanggal_telfon & "'"
    'updatetian30032016
    sQuery = sQuery + " AND custid = '" & CustId & "'"
    '=================================================
    Set Rs_Cek_Tanggal = New ADODB.Recordset
    Rs_Cek_Tanggal.CursorLocation = adUseClient
    Rs_Cek_Tanggal.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Rs_Cek_Tanggal.RecordCount > 0 Then
        jumlah_call = Rs_Cek_Tanggal!jumlah_call
        jumlah_call = jumlah_call + 1
        
        M_OBJCONN.Execute "UPDATE tbl_temp_telfon_review SET jumlah_call = '" & jumlah_call & "' WHERE id = '" & Rs_Cek_Tanggal!ID & "'"
    Else
        M_OBJCONN.Execute "INSERT INTO tbl_temp_telfon_review(custid, no_telfon, tanggal_telfon, jumlah_call, agent) " & _
                          " VALUES ('" & CustId & "','" & nomor_telpon & "', " & _
                          " '" & waktu_server_sekarang & "', '1', '" & AGENT & "')"
        'jejaktian28032016listphonereview
        M_OBJCONN.Execute "INSERT INTO tblloglistreview(custid, no_telfon, tanggal_telfon, agent) " & _
                          " VALUES ('" & CustId & "','" & nomor_telpon & "', " & _
                          " '" & waktu_server_sekarang & "','" & AGENT & "')"
        '===========================================================================
    End If
End Sub

Public Function FungsiWaktuServer()
 'Fungsi Untuk mengambil waktu dan tanggal di server database
 Dim CMDSQL As String
 Dim M_objrs As ADODB.Recordset
 
 CMDSQL = "select now() as waktu"
 
 Set M_objrs = New ADODB.Recordset
 M_objrs.CursorLocation = adUseClient
 
 M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
 
 WaktuServer = Format(M_objrs(0), "yyyy-mm-dd hh:nn:ss")
 FungsiWaktuServer = WaktuServer
 Set M_objrs = Nothing
End Function
