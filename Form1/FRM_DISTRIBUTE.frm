VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form FRM_DISTRIBUTE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   9345
   ControlBox      =   0   'False
   Icon            =   "FRM_DISTRIBUTE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdUpdateList 
      Caption         =   "Update List"
      Height          =   375
      Left            =   7800
      TabIndex        =   24
      Top             =   60
      Width           =   945
   End
   Begin TDBNumber6Ctl.TDBNumber TdbNDistribusi 
      Height          =   330
      Left            =   5370
      TabIndex        =   22
      Top             =   90
      Width           =   2235
      _Version        =   65536
      _ExtentX        =   3942
      _ExtentY        =   582
      Calculator      =   "FRM_DISTRIBUTE.frx":0442
      Caption         =   "FRM_DISTRIBUTE.frx":0462
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FRM_DISTRIBUTE.frx":04CE
      Keys            =   "FRM_DISTRIBUTE.frx":04EC
      Spin            =   "FRM_DISTRIBUTE.frx":0536
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
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -99999
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
      MaxValueVT      =   5505029
      MinValueVT      =   1028849669
   End
   Begin VB.TextBox TxtZIPCOde 
      Height          =   315
      Left            =   1470
      TabIndex        =   21
      Top             =   45
      Width           =   1935
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
      Left            =   4230
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   6030
      Width           =   2370
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
      Left            =   6615
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   690
      Width           =   2670
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
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   690
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4860
      Left            =   30
      TabIndex        =   7
      Top             =   1080
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   8573
      _Version        =   393216
      Style           =   1
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
      TabPicture(0)   =   "FRM_DISTRIBUTE.frx":055E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Distribusi"
      TabPicture(1)   =   "FRM_DISTRIBUTE.frx":057A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   4425
         Left            =   -74910
         TabIndex        =   10
         Top             =   375
         Width           =   9015
         Begin MSComctlLib.ListView ListView2 
            Height          =   4260
            Left            =   30
            TabIndex        =   11
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
         Height          =   4470
         Left            =   105
         TabIndex        =   8
         Top             =   330
         Width           =   9000
         Begin MSComctlLib.ListView ListView1 
            Height          =   4305
            Left            =   30
            TabIndex        =   9
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
      Height          =   615
      Left            =   30
      TabIndex        =   5
      Top             =   5910
      Visible         =   0   'False
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   1085
      _Version        =   196610
      ForeColor       =   192
      Caption         =   "Proses"
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   360
         Left            =   30
         TabIndex        =   6
         Top             =   225
         Visible         =   0   'False
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   0
      Left            =   1470
      TabIndex        =   3
      Top             =   705
      Width           =   2040
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   1
      Left            =   3510
      TabIndex        =   2
      Top             =   705
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
      Height          =   375
      Index           =   0
      Left            =   6795
      TabIndex        =   1
      Top             =   6000
      Width           =   885
   End
   Begin VB.CommandButton Command1 
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
      Left            =   7785
      TabIndex        =   0
      Top             =   6000
      Width           =   885
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
      Left            =   1710
      TabIndex        =   15
      Top             =   1050
      Visible         =   0   'False
      Width           =   945
   End
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
      Left            =   2655
      TabIndex        =   16
      Top             =   1065
      Visible         =   0   'False
      Width           =   1095
   End
   Begin TDBNumber6Ctl.TDBNumber TdbNDeviasi 
      Height          =   300
      Left            =   1470
      TabIndex        =   23
      Top             =   375
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   529
      Calculator      =   "FRM_DISTRIBUTE.frx":0596
      Caption         =   "FRM_DISTRIBUTE.frx":05B6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FRM_DISTRIBUTE.frx":0622
      Keys            =   "FRM_DISTRIBUTE.frx":0640
      Spin            =   "FRM_DISTRIBUTE.frx":068A
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,##0.0000"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,##0.0000"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999
      MinValue        =   -99
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
      MaxValueVT      =   5505029
      MinValueVT      =   1028849669
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Jumlah Distribusi :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   4
      Left            =   4065
      TabIndex        =   20
      Top             =   60
      Width           =   1260
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Deviasi :"
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
      Index           =   3
      Left            =   195
      TabIndex        =   19
      Top             =   375
      Width           =   1260
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Kode Pos :"
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
      Index           =   2
      Left            =   195
      TabIndex        =   18
      Top             =   90
      Width           =   1260
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Campaign :"
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
      Left            =   195
      TabIndex        =   4
      Top             =   735
      Width           =   1260
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
      Left            =   435
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   1260
   End
End
Attribute VB_Name = "FRM_DISTRIBUTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdUpdateList_Click()
Dim i As Integer
Text3.Text = 0
If ListView1.ListItems.Count = 0 Then
    Exit Sub
End If
For i = 1 To ListView1.ListItems.Count
    ListView1.ListItems(i).SubItems(2) = TdbNDistribusi.Value
    Text3.Text = CCur(Text3.Text) + TdbNDistribusi.Value
Next i
Text3.Text = Format(Text3.Text, "##,###")
MsgBox "Done"
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim m_data As New CLS_DISTRIBUSI
Dim M_objrs As ADODB.Recordset
Dim JUMLAH As Currency
Dim listitem1 As ListItem
Select Case Index
    Case 0
        Set M_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("KODEDS")
            Combo1(1).Text = M_objrs("KETERANGAN")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 1
        Set M_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("KODEDS")
            Combo1(1).Text = M_objrs("KETERANGAN")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    End Select
Set M_objrs = Nothing
ListView2.ListItems.CLEAR
Text2.Text = 0
'JUMLAH = m_data.HITUNG_TEMPCUST_CC(M_OBJCONN, "RECSOURCE = '" + Combo1(0).Text + "' and AGENT IS NULL ")
JUMLAH = m_data.HITUNG_TEMPCUST_CC(M_OBJCONN, "RECSOURCE = '" + Combo1(0).Text + "' and AGENT='swap' ")
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
   lReturn = SendMessage(Combo1(Index).hwnd, CB_FINDSTRING, -1, ByVal sSearchText)
   If lReturn <> CB_ERR Then
      mbIgnoreListClick = True
      Combo1(Index).ListIndex = lReturn
      mbIgnoreListClick = False
      Combo1(Index).Text = Combo1(Index).list(lReturn)
      Combo1(Index).SelStart = Len(sSearchText)
      Combo1(Index).SelLength = Len(Combo1(Index).Text)
      KeyAscii = 0
   End If
End If
End Select
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim m_data As New CLS_DISTRIBUSI
Dim M_objrs As ADODB.Recordset
Dim listitem1 As ListItem
Dim JUMLAH As Currency
Select Case Index
    Case 0
        Set M_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("KODEDS")
            Combo1(1).Text = M_objrs("KETERANGAN")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 1
        Set M_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("KODEDS")
            Combo1(1).Text = M_objrs("KETERANGAN")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    End Select
Set M_objrs = Nothing
ListView2.ListItems.CLEAR
Text2.Text = 0
'JUMLAH = m_data.HITUNG_TEMPCUST_CC(M_OBJCONN, "RECSOURCE = '" + Combo1(0).Text + "' and agent is null")
JUMLAH = m_data.HITUNG_TEMPCUST_CC(M_OBJCONN, "RECSOURCE = '" + Combo1(0).Text + "' and agent='swap'")
Set listitem1 = ListView2.ListItems.ADD(, , Combo1(0).Text)
     listitem1.SubItems(1) = Format(JUMLAH, "##,##0")
     listitem1.SubItems(2) = Text1.Text
     Text2.Text = Format(JUMLAH, "##,##0")
Set m_data = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim m_data As New CLS_DISTRIBUSI
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
        Call distribusi_data("", "", "0")
'        For i = 1 To ListView1.ListItems.count
'        ListView1.ListItems(i).Selected = True
'            If ListView1.ListItems(i).SubItems(2) <> "0" Then
'                SSFrame1.Visible = True
'                'm_data.PROSES M_OBJCONN, M_RPTCONN, Combo1(0).Text, ListView1.ListItems(i).Text, ListView1.ListItems(i).SubItems(2), ListView1.ListItems(i).SubItems(3), Text1.Text, ListView1.ListItems(i).SubItems(1)
'                distribusi_data Combo1(0).Text, ListView1.ListItems(i).Text, ListView1.ListItems(i).SubItems(2)
'            End If
'        Next i
           ' If m_data.ADD_OK Then
                MsgBox "Distribusi Selesai", vbInformation + vbOKOnly, "Informasi"
           ' Else
            '    m_data.ADD_OK = True
                'MsgBox "Ulangi Lagi", vbInformation + vbOKOnly, "Informasi"
           ' End If
'            Unload Me
    Case 1
        Unload Me
End Select
Set m_data = Nothing
End Sub

Private Sub distribusi_data(DATASOURCE As String, USERID As String, JUMLAH As String)
Dim m_rs As ADODB.Recordset
Dim i As Integer
Dim LAgent As String
Dim habisAgent As Boolean
Dim TOTALAMOUNT As Currency
On Error GoTo adderr:
Set m_rs = New ADODB.Recordset
m_rs.CursorLocation = adUseClient
'm_rs.Open "Select CUSTID,AGENT,AmountWo,TGLDISTRIBUSI from mgm where RECSOURCE ='" + Combo1(0).Text + "' AND AGENT IS NULL order by AmountWo Desc", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
m_rs.Open "Select CUSTID,AGENT,AmountWo,TGLDISTRIBUSI,f_agent from mandiri.mgm where RECSOURCE ='" + Combo1(0).Text + "' AND AGENT='swap' order by AmountWo Desc", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_rs.RecordCount <> 0 Then
    m_rs.MoveFirst
End If
While Not m_rs.EOF
    For i = 1 To ListView1.ListItems.Count
        'TOTALAMOUNT = TOTALAMOUNT + m_rs!amountwo
        'm_rs!agent = USERID
        ' ---------- cek agent lamanya -------------- '
'    LAgent = Trim(m_rs!f_agent)
'        If ListView1.ListItems(i).Text = m_rs!f_agent Then
'        Else
            '--------- cek limit distribusi --------------'
            If (CCur(ListView1.ListItems(i).SubItems(3)) + IIf(IsNull(m_rs!amountwo), 0, m_rs!amountwo)) <= (CCur(ListView1.ListItems(i).SubItems(2)) + (CCur(ListView1.ListItems(i).SubItems(2)) * TdbNDeviasi.Value)) Then
                If (CCur(ListView1.ListItems(i).SubItems(2)) + (CCur(ListView1.ListItems(i).SubItems(2)) * TdbNDeviasi.Value)) >= CCur(ListView1.ListItems(i).SubItems(3)) Then
                    habisAgent = False
                    m_rs!AGENT = ListView1.ListItems(i).Text
                    m_rs!TGLDISTRIBUSI = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
                    m_rs.update
                    m_rs.MoveNext
                    ListView1.ListItems(i).SubItems(3) = Format(CCur(ListView1.ListItems(i).SubItems(3)) + IIf(IsNull(m_rs!amountwo), 0, m_rs!amountwo), "###,###")
                    DoEvents
                End If
            End If
'        End If
        If i = ListView1.ListItems.Count Then
        
            If habisAgent = True Then
                m_rs.MoveNext
                habisAgent = False
            End If
            
            i = 0
            habisAgent = True
        End If
        If m_rs.RecordCount = m_rs.Bookmark Then
            Exit Sub
        End If
    Next i
Wend
Set m_rs = Nothing

Set m_rs = New ADODB.Recordset
m_rs.CursorLocation = adUseClient
'm_rs.Open "Select CUSTID,AGENT,AmountWo,TGLDISTRIBUSI from mgm where RECSOURCE ='" + Combo1(0).Text + "' AND AGENT IS NULL order by AmountWo Desc", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
m_rs.Open "Select CUSTID,AGENT,AmountWo,TGLDISTRIBUSI from mandiri.mgm where RECSOURCE ='" + Combo1(0).Text + "' AND AGENT='swap' order by AmountWo Desc", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_rs.RecordCount <> 0 Then
    m_rs.MoveFirst
End If
While Not m_rs.EOF
    For i = 1 To ListView1.ListItems.Count
        'TOTALAMOUNT = TOTALAMOUNT + m_rs!amountwo
        'm_rs!agent = USERID
        If (CCur(ListView1.ListItems(i).SubItems(3)) + IIf(IsNull(m_rs!amountwo), 0, m_rs!amountwo)) <= (CCur(ListView1.ListItems(i).SubItems(2)) + (CCur(ListView1.ListItems(i).SubItems(2)) * TdbNDeviasi.Value)) Then
            If (CCur(ListView1.ListItems(i).SubItems(2)) + (CCur(ListView1.ListItems(i).SubItems(2)) * TdbNDeviasi.Value)) >= CCur(ListView1.ListItems(i).SubItems(3)) Then
                habisAgent = False
                m_rs!AGENT = ListView1.ListItems(i).Text
                m_rs!TGLDISTRIBUSI = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
                m_rs.update
                m_rs.MoveNext
                ListView1.ListItems(i).SubItems(3) = Format(CCur(ListView1.ListItems(i).SubItems(3)) + IIf(IsNull(m_rs!amountwo), 0, m_rs!amountwo), "###,###")
                DoEvents
            End If
        End If
        If i = ListView1.ListItems.Count Then
            If habisAgent = True Then
                m_rs.MoveNext
                habisAgent = False
            End If
            i = 0
            habisAgent = True
        End If
        If m_rs.RecordCount = m_rs.Bookmark Then
            Exit Sub
        End If
    Next i
Wend
Set m_rs = Nothing
Exit Sub
adderr:
    MsgBox err.Description
'    Resume
End Sub

Private Sub Form_Load()
Dim M_objrs As ADODB.Recordset
Dim m_data As New CLS_DISTRIBUSI
Dim ListItem As ListItem
    SSTab1.Tab = 0
    Text2.Text = 0
    Text3.Text = 0
    Option1(0).Value = True
    
    Call header
    Call header1
'    Set listitem = ListView1.ListItems.ADD(, , FRM_SETUSER.Combo1(0).Text)
'        listitem.SubItems(1) = FRM_SETUSER.Combo1(1).Text
'        listitem.SubItems(2) = 0
'        listitem.SubItems(3) = Format(MDIForm1.TDBDate1.Text, "yyyymmdd") & Format(Now, "hhmm")
        
    Set M_objrs = m_data.QUERY_USER_ACC(M_RPTCONN, FRM_SETUSER.Combo1(0).Text)
    While Not M_objrs.EOF
         Set ListItem = ListView1.ListItems.ADD(, , IIf(IsNull(M_objrs("USERID")), "", M_objrs("USERID")))
             ListItem.SubItems(1) = IIf(IsNull(M_objrs("NAMA")), "", M_objrs("NAMA"))
             ListItem.SubItems(2) = 0
             ListItem.SubItems(3) = 0
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
Set M_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "")
    While Not M_objrs.EOF
        Combo1(0).AddItem M_objrs("KODEDS")
        Combo1(0).DataField = M_objrs("KODEDS")
        Combo1(1).AddItem M_objrs("KETERANGAN")
        Combo1(1).DataField = M_objrs("KETERANGAN")
        M_objrs.MoveNext
    Wend
Set M_objrs = Nothing
Set M_objrs = m_data.QUERY_SPV(M_OBJCONN, " SPVCODE = '" + FRM_SETUSER.Combo1(0).Text + "'")
If M_objrs.RecordCount <> 0 Then
    Text1.Text = IIf(IsNull(M_objrs("UNIT")), "", M_objrs("UNIT"))
Else
    Text1.Text = Empty
End If
Set M_objrs = Nothing
Set m_data = Nothing
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Kode Agent", 15 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Nama Nama", 31 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Jumlah", 7 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Jumlah Dibagi", 15 * TXT
End Sub

Private Sub header1()
    ListView2.ColumnHeaders.ADD 1, , "Campaign", 15 * TXT
    ListView2.ColumnHeaders.ADD 2, , "Jumlah", 15 * TXT, 1
    ListView2.ColumnHeaders.ADD 3, , "Produk", 15 * TXT
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
Dim VOLD As Double
Dim VNEW As Double
Dim TGL As String
If Text2.Text < 1 Then
    MsgBox "Tidak Ada Data Untuk Di Distribusikan", vbInformation + vbOKOnly, "Aplikasi"
    Exit Sub
End If
    With Form1
                .Text1.Text = ListView1.SelectedItem.Text
                .Text2.Text = ListView1.SelectedItem.SubItems(2)
                TGL = Mid(ListView1.SelectedItem.SubItems(3), 7, 2) & "/" & Mid(ListView1.SelectedItem.SubItems(3), 5, 2) & "/" & Left(ListView1.SelectedItem.SubItems(3), 4)
                '.TDBDate1.Value = Format(TGL, "dd-mmm-yyyy")
                '.TDBTime1.Value = Mid(ListView1.SelectedItem.SubItems(3), 9, 2) & ":" & Right(ListView1.SelectedItem.SubItems(3), 2)
                
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
                        ListView1.SelectedItem.SubItems(3) = 0
                End If
                Unload Form1
            End With
End Sub


