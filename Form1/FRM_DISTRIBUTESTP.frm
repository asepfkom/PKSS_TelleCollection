VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FRM_DISTRIBUTESTP 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   Icon            =   "FRM_DISTRIBUTESTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "X-SELL"
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
      Height          =   300
      Index           =   1
      Left            =   2475
      TabIndex        =   17
      Top             =   780
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MGM"
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
      Height          =   300
      Index           =   0
      Left            =   1530
      TabIndex        =   16
      Top             =   765
      Width           =   945
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5865
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   6765
      Width           =   810
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1095
      Width           =   810
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7530
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1095
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4860
      Left            =   75
      TabIndex        =   8
      Top             =   1635
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   8573
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Data Tersedia"
      TabPicture(0)   =   "FRM_DISTRIBUTESTP.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Distribusi"
      TabPicture(1)   =   "FRM_DISTRIBUTESTP.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H80000004&
         Height          =   4425
         Left            =   -74910
         TabIndex        =   11
         Top             =   375
         Width           =   9015
         Begin MSComctlLib.ListView ListView2 
            Height          =   4260
            Left            =   30
            TabIndex        =   12
            Top             =   135
            Width           =   8955
            _ExtentX        =   15796
            _ExtentY        =   7514
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   -2147483643
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
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000004&
         Height          =   4470
         Left            =   105
         TabIndex        =   9
         Top             =   330
         Width           =   9000
         Begin MSComctlLib.ListView ListView1 
            Height          =   4305
            Left            =   30
            TabIndex        =   10
            Top             =   135
            Width           =   8925
            _ExtentX        =   15743
            _ExtentY        =   7594
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   -2147483643
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
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   630
      Left            =   150
      TabIndex        =   6
      Top             =   6660
      Visible         =   0   'False
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   1111
      _Version        =   196610
      ForeColor       =   192
      BackColor       =   -2147483644
      Caption         =   "Proses"
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   360
         Left            =   30
         TabIndex        =   7
         Top             =   225
         Visible         =   0   'False
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   0
      Left            =   1530
      TabIndex        =   4
      Top             =   1110
      Width           =   2040
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   1
      Left            =   3570
      TabIndex        =   3
      Top             =   1110
      Width           =   3060
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Proses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   6855
      TabIndex        =   2
      Top             =   6780
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   7845
      TabIndex        =   1
      Top             =   6780
      Width           =   960
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   1138
      _Version        =   196610
      Font3D          =   5
      ForeColor       =   0
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Distribusi Data"
      BevelWidth      =   2
      BorderWidth     =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Distribute Ke :"
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
      Index           =   1
      Left            =   255
      TabIndex        =   18
      Top             =   795
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "Sumber Data :"
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
      Index           =   0
      Left            =   285
      TabIndex        =   5
      Top             =   1125
      Width           =   1260
   End
End
Attribute VB_Name = "FRM_DISTRIBUTESTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Combo1_Click(Index As Integer)
Dim m_data As New CLS_DISTRIBUSISTP
Dim m_objrs As ADODB.Recordset
Dim JUMLAH As Currency
Dim listitem1 As LISTITEM
Select Case Index
    Case 0
        Set m_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("KODEDS")
            Combo1(1).Text = m_objrs("KETERANGAN")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 1
        Set m_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("KODEDS")
            Combo1(1).Text = m_objrs("KETERANGAN")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    End Select
Set m_objrs = Nothing
ListView2.ListItems.Clear
Text2.Text = 0
JUMLAH = m_data.HITUNG_TEMPCUST_CC(M_OBJCONN, "RECSOURCE = '" + Combo1(0).Text + "'")
Set listitem1 = ListView2.ListItems.ADD(, , Combo1(0).Text)
     listitem1.SubItems(1) = Format(JUMLAH, "##,##0")
     listitem1.SubItems(2) = Text1.Text
     Text2.Text = Format(JUMLAH, "##,##0")
Set m_data = Nothing
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim sSearchText As String
Dim lReturn As Long
Select Case Index
Case 0, 1
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
Dim m_data As New CLS_DISTRIBUSISTP
Dim m_objrs As ADODB.Recordset
Dim listitem1 As LISTITEM
Dim JUMLAH As Currency
Select Case Index
    Case 0
        Set m_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("KODEDS")
            Combo1(1).Text = m_objrs("KETERANGAN")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 1
        Set m_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("KODEDS")
            Combo1(1).Text = m_objrs("KETERANGAN")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    End Select
Set m_objrs = Nothing
ListView2.ListItems.Clear
Text2.Text = 0
JUMLAH = m_data.HITUNG_TEMPCUST_CC(M_OBJCONN, "RECSOURCE = '" + Combo1(0).Text + "'")
Set listitem1 = ListView2.ListItems.ADD(, , Combo1(0).Text)
     listitem1.SubItems(1) = Format(JUMLAH, "##,##0")
     listitem1.SubItems(2) = Text1.Text
     Text2.Text = Format(JUMLAH, "##,##0")
Set m_data = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim m_data As New CLS_DISTRIBUSISTP
Dim i As Integer
Select Case Index
    Case 0
        If CCur(Text3.Text) > CCur(Text2.Text) Then
            MsgBox "Data Tidak Yang Tersedia Tidak Cukup.. Kurangi Jumlah Distribusi", vbInformation + vbOKOnly, "Informasi"
            Exit Sub
        End If
        If Combo1(0).Text = Empty Then
            MsgBox "Data Source Harus Diisi", vbInformation + vbOKOnly, "Informasi"
            Exit Sub
        End If
        For i = 1 To ListView1.ListItems.count
        ListView1.ListItems(i).Selected = True
            If ListView1.ListItems(i).SubItems(2) <> "0" Then
                SSFrame1.Visible = True
                m_data.PROSES M_OBJCONN, M_RPTCONN, Combo1(0).Text, ListView1.ListItems(i).Text, ListView1.ListItems(i).SubItems(2), ListView1.ListItems(i).SubItems(3), Text1.Text, ListView1.ListItems(i).SubItems(1)
            End If
        Next i
            If m_data.ADD_OK Then
                MsgBox "Distribusi Selesai", vbInformation + vbOKOnly, "Informasi"
            Else
                m_data.ADD_OK = True
                MsgBox "Ulangi Lagi", vbInformation + vbOKOnly, "Informasi"
            End If
            Unload Me
    Case 1
        Unload Me
End Select
Set m_data = Nothing
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Dim m_data As New CLS_DISTRIBUSISTP
Dim LISTITEM As LISTITEM
    SSTab1.Tab = 0
    Text2.Text = 0
    Text3.Text = 0
    Option1(0).Value = True
    
    Call header
    Call header1
    Set LISTITEM = ListView1.ListItems.ADD(, , FRM_SETUSERSTP.Combo1(0).Text)
        LISTITEM.SubItems(1) = FRM_SETUSERSTP.Combo1(1).Text
        LISTITEM.SubItems(2) = 0
        LISTITEM.SubItems(3) = Format(MDIForm1.TDBDate1.Text, "yyyymmdd") & Format(Now, "hhmm")
        
    Set m_objrs = m_data.QUERY_USER_ACC(M_RPTCONN, FRM_SETUSERSTP.Combo1(0).Text)
    While Not m_objrs.EOF
         Set LISTITEM = ListView1.ListItems.ADD(, , IIf(IsNull(m_objrs("USERID")), "", m_objrs("USERID")))
             LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("NAMA")), "", m_objrs("NAMA"))
             LISTITEM.SubItems(2) = 0
             LISTITEM.SubItems(3) = Format(MDIForm1.TDBDate1.Text, "yyyymmdd") & Format(Now, "hhmm")
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
Set m_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "")
    While Not m_objrs.EOF
        Combo1(0).AddItem m_objrs("KODEDS")
        Combo1(0).DataField = m_objrs("KODEDS")
        Combo1(1).AddItem m_objrs("KETERANGAN")
        Combo1(1).DataField = m_objrs("KETERANGAN")
        m_objrs.MoveNext
    Wend
Set m_objrs = Nothing
Set m_objrs = m_data.QUERY_SPV(M_OBJCONN, " SPVCODE = '" + FRM_SETUSERSTP.Combo1(0).Text + "'")
If m_objrs.RecordCount <> 0 Then
    Text1.Text = IIf(IsNull(m_objrs("UNIT")), "", m_objrs("UNIT"))
Else
    Text1.Text = Empty
End If
Set m_objrs = Nothing
Set m_data = Nothing
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "User Id", 15 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Nama", 31 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Jumlah", 7 * TXT
    ListView1.ColumnHeaders.ADD 4, , "TglJam", 15 * TXT
End Sub

Private Sub header1()
    ListView2.ColumnHeaders.ADD 1, , "Sumber Data", 15 * TXT
    ListView2.ColumnHeaders.ADD 2, , "Jumlah Data", 15 * TXT, 1
    ListView2.ColumnHeaders.ADD 3, , "Jenis Produk", 15 * TXT
End Sub


Private Sub ListView1_DblClick()
Dim VOLD As Double
Dim VNEW As Double
Dim TGL As String
If Text2.Text < 1 Then
    MsgBox "Tidak Ada Data Untuk Di Distribusikan", vbInformation + vbOKOnly, "TeleGrandi"
    Exit Sub
End If
    With Form1
                .Text1.Text = ListView1.SelectedItem.Text
                .Text2.Text = ListView1.SelectedItem.SubItems(2)
                TGL = Mid(ListView1.SelectedItem.SubItems(3), 7, 2) & "/" & Mid(ListView1.SelectedItem.SubItems(3), 5, 2) & "/" & Left(ListView1.SelectedItem.SubItems(3), 4)
                .TDBDate1.Value = Format(TGL, "dd-mmm-yyyy")
                .TDBTime1.Value = Mid(ListView1.SelectedItem.SubItems(3), 9, 2) & ":" & Right(ListView1.SelectedItem.SubItems(3), 2)
                
                VOLD = ListView1.SelectedItem.SubItems(2)
                .Text1.Locked = True
                .Text1.TabStop = False
                .Text1.BackColor = &H8000000F
                .Text1.Appearance = 0
                .Show vbModal
                If .ok Then
                        VNEW = CCur(.Text2.Text)
                        Text3.Text = (CCur(Text3.Text) - VOLD) + VNEW
                        ListView1.SelectedItem.SubItems(2) = .Text2.Text
                        ListView1.SelectedItem.SubItems(3) = Format(.TDBDate1.Text, "yyyymmdd") & Format(.TDBTime1.Text, "hhmm")
                End If
                Unload Form1
            End With
End Sub
