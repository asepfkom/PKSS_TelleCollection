VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "TIDATE6.OCX"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "TIMASK6.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_VER_SEARCH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Data"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "FRM_VER_SEARCH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7455
   Begin VB.CheckBox Check2 
      Caption         =   "Data Upload"
      Height          =   255
      Left            =   855
      TabIndex        =   22
      Top             =   345
      Width           =   1470
   End
   Begin VB.CheckBox Check1 
      Caption         =   "MGM Data"
      Height          =   255
      Left            =   870
      TabIndex        =   21
      Top             =   45
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proses....!!"
      Height          =   675
      Left            =   45
      TabIndex        =   19
      Top             =   4020
      Visible         =   0   'False
      Width           =   3990
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   390
         Left            =   60
         TabIndex        =   20
         Top             =   210
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Hapus"
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
      Index           =   2
      Left            =   2865
      TabIndex        =   10
      Top             =   2730
      Width           =   1545
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   4785
      MaxLength       =   20
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2325
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Left            =   1725
      TabIndex        =   2
      Top             =   1005
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   556
      Calendar        =   "FRM_VER_SEARCH.frx":0442
      Caption         =   "FRM_VER_SEARCH.frx":055A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FRM_VER_SEARCH.frx":05C6
      Keys            =   "FRM_VER_SEARCH.frx":05E4
      Spin            =   "FRM_VER_SEARCH.frx":0642
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
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
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37475
      CenturyMode     =   0
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   3630
      TabIndex        =   6
      Top             =   1635
      Width           =   3825
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   1725
      TabIndex        =   5
      Top             =   1635
      Width           =   1905
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   3630
      TabIndex        =   4
      Top             =   1320
      Width           =   3825
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   1725
      TabIndex        =   3
      Top             =   1320
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Keluar"
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
      Index           =   1
      Left            =   4530
      TabIndex        =   11
      Top             =   2730
      Width           =   1545
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cari"
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
      Index           =   0
      Left            =   1200
      TabIndex        =   9
      Top             =   2730
      Width           =   1545
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   1725
      TabIndex        =   1
      Top             =   690
      Width           =   4590
   End
   Begin TDBMask6Ctl.TDBMask TDBMask1 
      Height          =   315
      Left            =   1725
      TabIndex        =   7
      Top             =   1950
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Caption         =   "FRM_VER_SEARCH.frx":066A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FRM_VER_SEARCH.frx":06D6
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
      ForeColor       =   0
      Format          =   "999-99999"
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
      Text            =   "___-_____"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TDBMask2 
      Height          =   315
      Left            =   1725
      TabIndex        =   8
      Top             =   2265
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Caption         =   "FRM_VER_SEARCH.frx":0718
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FRM_VER_SEARCH.frx":0784
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
      ForeColor       =   0
      Format          =   "9999-99999999"
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
      Text            =   "____-________"
      Value           =   ""
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Lap"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   3075
      TabIndex        =   18
      Top             =   135
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Selular"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   240
      TabIndex        =   17
      Top             =   2340
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Telephone"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   2010
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sumber Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Lahir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   225
      TabIndex        =   14
      Top             =   1050
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Sales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   1380
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   165
      TabIndex        =   12
      Top             =   720
      Width           =   1485
   End
End
Attribute VB_Name = "FRM_VER_SEARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Check2.Value = 0
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Check1.Value = 0
    End If
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim M_DATA As New CLS_FRMSEARCH
Dim m_objrs As ADODB.Recordset
Select Case Index
Case 0
    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(0).Text = m_objrs("USERID")
        Combo1(1).Text = m_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 1
    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(0).Text = m_objrs("USERID")
        Combo1(1).Text = m_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 2
Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(2).Text = m_objrs("KODEDS")
        Combo1(3).Text = m_objrs("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
Case 3
Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(2).Text = m_objrs("KODEDS")
        Combo1(3).Text = m_objrs("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
End Select
Set M_DATA = Nothing
Set m_objrs = Nothing
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim sSearchText As String
Dim lReturn As Long
Select Case Index
Case 0, 1, 2, 3
If KeyAscii = 13 Then
   Combo1_Click (Index)
   KeyAscii = 0
Else
   sSearchText = Left$(Combo1(Index).Text, Combo1(Index).SelStart) & Chr$(KeyAscii)
   lReturn = SendMessage(Combo1(Index).hWnd, CB_FINDSTRING, -1, ByVal sSearchText)
   If lReturn <> CB_ERR Then
      mbIgnoreListClick = True
      Combo1(Index).ListIndex = lReturn
      mbIgnoreListClick = False
      Combo1(Index).Text = Combo1(Index).List(lReturn)
      Combo1(Index).SelStart = Len(sSearchText)
      Combo1(Index).SelLength = Len(Combo1(Index).Text)
      KeyAscii = 0
   End If
End If
End Select
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim M_DATA As New CLS_FRMSEARCH
Dim m_objrs As ADODB.Recordset
Select Case Index
Case 0
    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(0).Text = m_objrs("USERID")
        Combo1(1).Text = m_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 1
    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(0).Text = m_objrs("USERID")
        Combo1(1).Text = m_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 2
Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(2).Text = m_objrs("KODEDS")
        Combo1(3).Text = m_objrs("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
Case 3
Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(2).Text = m_objrs("KODEDS")
        Combo1(3).Text = m_objrs("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
End Select
Set M_DATA = Nothing
Set m_objrs = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim NAMACUST As String
Dim NAMAAGENT As String
Dim DATASOURCE As String
Dim TGLLAHIR As String
Dim OFFPHONE As String
Dim OFFPHONE2 As String
Dim HOMEPHONE As String
Dim HOMEPHONE2 As String
Dim MOBILEPHONE As String
Dim MOBILEPHONE2 As String
Dim FAXPHONE As String
Dim FAXPHONE2 As String
Dim M_DATA As New CLS_FRMSEARCH
Dim m_objrs As ADODB.Recordset
Dim PANJANG As Integer
Select Case Index
    Case 0
        If Text1(0).Text = Empty And Combo1(0).Text = Empty And Combo1(2).Text = Empty And Len(TDBMask2.Value) < 1 And Len(TDBMask1.Value) < 1 And TDBDate1.ValueIsNull And Len(Text1(2).Text) < 3 Then
            MsgBox "Masukan Kriteria Customer Yang Akan Dicari...!!!", vbCritical + vbOKOnly, "Peringatan"
            Text1(0).SetFocus
            Set M_DATA = Nothing
            Set m_objrs = Nothing
            Exit Sub
        Else
            If Len(Text1(2).Text) < 3 Then
                    If Text1(0).Text <> Empty Then
                        NAMACUST = "NAME LIKE " + "'%" + UBAH_QUOTE(Text1(0).Text) + "%'"
                    End If
                    If Combo1(0).Text <> Empty Then
                        NAMAAGENT = "AGENT = '" + Combo1(0).Text + "'"
                    End If
                    If Combo1(2).Text <> Empty Then
                        DATASOURCE = "RECSOURCE = '" + Combo1(2).Text + "'"
                    End If
                    If TDBDate1.ValueIsNull Then
                    Else
                        TGLLAHIR = "BIRTHD = '" + Format(TDBDate1.Text, "mm/dd/yyyy") + "'"
                    End If
                    If Len(TDBMask1.Value) > 1 Then
                        OFFPHONE = "OFFICENO like '%" + TDBMask1.Value + "%'"
                        OFFPHONE2 = "OFFICENO2 like '%" + TDBMask1.Value + "%'"
                        HOMEPHONE = "HOMENO like '%" + TDBMask1.Value + "%'"
                        HOMEPHONE2 = "HOMENO2 like '%" + TDBMask1.Value + "%'"
                        FAXPHONE = "FAXNO like '%" + TDBMask1.Value + "%'"
                        FAXPHONE2 = "FAXNO2 like '%" + TDBMask1.Value + "%'"
                        
                    End If
                    If Len(TDBMask2.Value) > 1 Then
                        MOBILEPHONE = "MOBILENO like '%" + TDBMask2.Value + "%'"
                        MOBILEPHONE2 = "MOBILENO2 like '%" + TDBMask2.Value + "%'"
                    End If
        
                        
                'Unload FRM_SEARCH
                If Check1.Value = 1 Then
                    Set m_objrs = M_DATA.QUERY_SEARCH_CONDITION_mgm(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
                                                            OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text)
                    
                Else
                    If Check2.Value = 1 Then
                        Set m_objrs = M_DATA.QUERY_SEARCH_UPLOAD(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
                                                            OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text)
                    Else
                        Set m_objrs = M_DATA.QUERY_SEARCH_CONDITION(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
                                                            OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text)
                    End If
                End If
        Else
            Set m_objrs = M_DATA.QUERY_SEARCH(M_OBJCONN, "NOLAP = '" + Text1(2).Text + "'", MDIForm1.Text3.Text)
        End If
        
If m_objrs.RecordCount = 0 Then
    MsgBox "Data Tidak Ditemukan", vbInformation + vbOKOnly, "TeleGrandi"
    Set m_objrs = Nothing
    Set M_DATA = Nothing
    Exit Sub
End If
            search_ok = True
            If Check1.Value = 1 Then
                FRM_VER_MGMDATA.Caption = "Search MGM Data"
                FRM_VER_MGMDATA.Show
            Else
                If Check2.Value = 1 Then
                
                Else
                    FRM_VER_PRESCREEN.Caption = "Search Non Mgm Data"
                    FRM_VER_PRESCREEN.Show
                End If
                
            End If
            'FRM_PRESCREEN.Show vbModal

        End If
    Case 1
        Unload Me
    Case 2
        Text1(2).Text = Empty
        Text1(0).Text = Empty
        TDBDate1.Text = Empty
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
        TDBMask1.Text = Empty
        TDBMask2.Text = Empty
End Select
 
Set M_DATA = Nothing
Set m_objrs = Nothing
End Sub




Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New CLS_FRMSEARCH

    Me.Top = 1000
    Me.Left = 1500
    
Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "")
    While Not m_objrs.EOF
        Combo1(0).AddItem m_objrs("USERID")
        Combo1(0).DataField = m_objrs("USERID")
        Combo1(1).AddItem m_objrs("AGENT")
        Combo1(1).DataField = m_objrs("AGENT")
        m_objrs.MoveNext
    Wend
Set m_objrs = Nothing
Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "")
    While Not m_objrs.EOF
        Combo1(2).AddItem m_objrs("KODEDS")
        Combo1(2).DataField = m_objrs("KODEDS")
        Combo1(3).AddItem m_objrs("KETERANGAN")
        Combo1(3).DataField = m_objrs("KETERANGAN")
        m_objrs.MoveNext
    Wend
    

If UCase(MDIForm1.Text3.Text) = "ADMIN" Then
    Label1(5).Visible = True
    Text1(2).Visible = True
End If

Set m_objrs = Nothing
Set M_DATA = Nothing
End Sub



Private Sub Form_Unload(Cancel As Integer)
Frame1.Visible = False
ProgressBar1.Value = 0
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click(0)
End If

Select Case Index
Case 1
    Select Case KeyAscii
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
            Exit Sub
        Case Else
            KeyAscii = 0
    End Select
End Select
End Sub

