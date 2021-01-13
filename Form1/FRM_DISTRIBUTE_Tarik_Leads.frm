VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FRM_distribute_Tarik_Leads 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   Icon            =   "FRM_DISTRIBUTE_Tarik_Leads.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "X-Sell"
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
      Left            =   2520
      TabIndex        =   17
      Top             =   765
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MGM"
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
      Left            =   1545
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
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1080
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
      Left            =   7095
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1080
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
      TabPicture(0)   =   "FRM_DISTRIBUTE_Tarik_Leads.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Distribusi"
      TabPicture(1)   =   "FRM_DISTRIBUTE_Tarik_Leads.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H80000004&
         Height          =   4425
         Left            =   90
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
         Left            =   -74895
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
      Top             =   1095
      Width           =   1605
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   1
      Left            =   3135
      TabIndex        =   3
      Top             =   1095
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
      Left            =   6840
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
      Caption         =   "Upload Ke :"
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
      Left            =   270
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
Attribute VB_Name = "FRM_distribute_Tarik_Leads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ADD_OK As Boolean
Private Sub Combo1_Click(Index As Integer)
'Dim M_DATA As New CLS_DISTRIBUSI
Dim m_objrs As ADODB.Recordset
Select Case Index
    Case 0
        Set m_objrs = QUERY_COMBO_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("KODEDS")
            Combo1(1).Text = m_objrs("KETERANGAN")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 1
        Set m_objrs = QUERY_COMBO_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("KODEDS")
            Combo1(1).Text = m_objrs("KETERANGAN")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    End Select
Set m_objrs = Nothing
'Set M_DATA = Nothing
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
'Dim M_DATA As New CLS_DISTRIBUSI
Dim m_objrs As ADODB.Recordset
Dim listitem1 As listitem
Dim JUMLAH As Currency

Select Case Index
    Case 0
        Set m_objrs = QUERY_COMBO_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("KODEDS")
            Combo1(1).Text = m_objrs("KETERANGAN")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 1
        Set m_objrs = QUERY_COMBO_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
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
        JUMLAH = HITUNG_TEMPCUST_CC(M_OBJCONN, "AGENT ='NELLY' and RECSOURCEREF = '" + Combo1(0).Text + "' ")
        Set listitem1 = ListView2.ListItems.ADD(, , Combo1(0).Text)
             listitem1.SubItems(1) = Format(JUMLAH, "##,##0")
             listitem1.SubItems(2) = Text1.Text
             Text2.Text = Format(JUMLAH, "##,##0")
'Set M_DATA = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
'Dim M_DATA As New CLS_DISTRIBUSI
Dim CMDSQL As String
Dim i As Integer

On Error GoTo ERR1
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
        For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Selected = True
            If ListView1.ListItems(i).SubItems(2) <> "0" Then
                SSFrame1.Visible = True
                CMDSQL = "Update cc_custtbl SET AGENT = " + "'" + ListView1.ListItems(i).Text + "'"
                CMDSQL = CMDSQL + " ,TGLSTATUS =" + "'" + Format(MDIForm1.TDBDate1.Value, "YYYY/MM/DD HH:NN") + "'"
                'cmdsql = cmdsql + " ,F_TARIK =" + "'YA'"
                CMDSQL = CMDSQL + " ,NAMAAGENT =" + "'" + ListView1.ListItems(i).SubItems(1) + "'"
                CMDSQL = CMDSQL + " WHERE CUSTID IN (SELECT TOP " + ListView1.ListItems(i).SubItems(2)
                CMDSQL = CMDSQL + " CUSTID FROM CC_CUSTTBL where AGENT ='NELLY' and recsource ='" + Combo1(0).Text + "') "
                M_OBJCONN.Execute CMDSQL
            End If
        Next i
        MsgBox "Distribusi Selesai", vbInformation + vbOKOnly, "Informasi"
        Unload Me
        
    Case 1
        Unload Me
End Select

Exit Sub
ERR1:

MsgBox "Ulangi Lagi", vbInformation + vbOKOnly, "Informasi"

End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
'Dim M_DATA As New CLS_DISTRIBUSI
Dim listitem As listitem
    SSTab1.Tab = 0
    Text2.Text = 0
    Text3.Text = 0
    Option1(1).Enabled = False
    Option1(0).Enabled = False
    Option1(1).Value = True
    Call header
    Call header1
    Set listitem = ListView1.ListItems.ADD(, , FRM_SETUSER_Tarik_Leads.Combo1(0).Text)
        listitem.SubItems(1) = FRM_SETUSER_Tarik_Leads.Combo1(1).Text
        listitem.SubItems(2) = 0
        listitem.SubItems(3) = Format(MDIForm1.TDBDate1.Text, "yyyymmdd") & Format(Now, "hhmm")
      
    Set m_objrs = QUERY_USER_ACC(M_RPTCONN, FRM_SETUSER_Tarik_Leads.Combo1(0).Text)
    While Not m_objrs.EOF
         Set listitem = ListView1.ListItems.ADD(, , IIf(IsNull(m_objrs("USERID")), "", m_objrs("USERID")))
             listitem.SubItems(1) = IIf(IsNull(m_objrs("NAMA")), "", m_objrs("NAMA"))
             listitem.SubItems(2) = 0
             listitem.SubItems(3) = Format(MDIForm1.TDBDate1.Text, "yyyymmdd") & Format(Now, "hhmm")
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    'Call QUERY_COMBO_DATASOURCE
    Set m_objrs = QUERY_COMBO_DATASOURCE(M_OBJCONN, "")
    While Not m_objrs.EOF
        Combo1(0).AddItem m_objrs("KODEDS")
        Combo1(0).DataField = m_objrs("KODEDS")
        Combo1(1).AddItem m_objrs("KETERANGAN")
        Combo1(1).DataField = m_objrs("KETERANGAN")
        m_objrs.MoveNext
    Wend
Set m_objrs = Nothing
'Call QUERY_SPV
Set m_objrs = QUERY_SPV(M_OBJCONN, " SPVCODE = '" + FRM_SETUSER_Tarik_Leads.Combo1(0).Text + "'")
If m_objrs.RecordCount <> 0 Then
    Text1.Text = IIf(IsNull(m_objrs("UNIT")), "", m_objrs("UNIT"))
Else
    Text1.Text = Empty
End If
Set m_objrs = Nothing
'Set M_DATA = Nothing
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
                '.Text3.Text = ListView1.SelectedItem.SubItems(3)
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
                      '  ListView1.SelectedItem.SubItems(3) = .Text3.Text
                        ListView1.SelectedItem.SubItems(3) = Format(.TDBDate1.Text, "yyyymmdd") & Format(.TDBTime1.Text, "hhmm")
                End If
                Unload Form1
            End With
End Sub

Public Function QUERY_COMBO_DATASOURCE(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim m_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM DATASOURCETBL"
'CMDSQL = CMDSQL + " WHERE STATUS ='A'"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
    Else
     CMDSQL = CMDSQL + " where left(kodeds,3)<>'INF'"
    'cmdsql = cmdsql + " and left(kodeds,3)<>'pre'"
 End If
 
    
CMDSQL = CMDSQL + " ORDER BY KODEDS"
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_COMBO_DATASOURCE = m_objrs
Set m_objrs = Nothing
End Function

Public Function QUERY_USER_ACC(M_RPTCONN As ADODB.Connection, SPVCODE As String) As Object
    Dim CMDSQL As String
    Dim m_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM DISTRIBUSI"
CMDSQL = CMDSQL + " ORDER BY USERID"
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_USER_ACC = m_objrs
Set m_objrs = Nothing
End Function

Public Function QUERY_SPV(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim m_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM SPVTBL"
CMDSQL = CMDSQL + " WHERE UNIT <> 'Admin'"
'cmdsql = cmdsql + " WHERE  spvcode='" + MDIForm1.Text1.Text + "'  "
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " AND " + M_WHERE
 End If
CMDSQL = CMDSQL + " ORDER BY SPVCODE"
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SPV = m_objrs
Set m_objrs = Nothing
End Function

Public Function PROSES(M_OBJCONN As ADODB.Connection, M_RPTCONN As ADODB.Connection, DATASOURCE As String, USERID As String, JUMLAH As String, tgljam As String, TIPE_PRODUK As String, NAMAAGENT As String)
Dim TGL As String
Dim JAM As String
Dim tgl1 As String
Dim m_objrs As ADODB.Recordset
On Error GoTo HELL
TIPE_PRODUK = True

M_OBJCONN.BeginTrans
    If Len(tgljam) < 11 Then
        TGL = Format(MDIForm1.TDBDate1.Text, "mm/dd/yy")
        JAM = Format(Now, "hh:mm")
        tgl1 = TGL + " " + JAM
    Else
        TGL = Mid(tgljam, 5, 2) + "/" + Mid(tgljam, 7, 2) + "/" + Left(tgljam, 4)
        JAM = Mid(tgljam, 9, 2) + ":" + Right(tgljam, 2)
        tgl1 = TGL + " " + JAM
    End If
    If TIPE_PRODUK = False Then
    Exit Function
    Else
'        WaitSecs (1)
        Call UPDATE_TEMPCUSTTBL_KTA(M_OBJCONN, USERID, JUMLAH, tgl1, TIPE_PRODUK, NAMAAGENT)
        'WaitSecs (1)
        Call QUERY_TEMPCUSTTBL_KTA(M_OBJCONN, TIPE_PRODUK, USERID, NAMAAGENT)
    End If
    ProgressBar1.Value = ProgressBar1.Max
    ProgressBar1.Visible = False
    ProgressBar1.Value = 0
    
Set m_objrs = Nothing
M_OBJCONN.CommitTrans
ADD_OK = True
Exit Function
HELL:
    ADD_OK = False
    M_OBJCONN.RollbackTrans
    MsgBox Err.Description
End Function

Private Function UPDATE_TEMPCUSTTBL_KTA(M_OBJCONN As ADODB.Connection, USERID As String, JUMLAH As String, tgljam As String, TIPE_PRODUK As String, NAMAAGENT As String)
Dim CMDSQL As String
Dim CustId As String
Dim m_objrs As ADODB.Recordset
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient

    CMDSQL = "SELECT TOP " + JUMLAH
    CMDSQL = CMDSQL + " CUSTID FROM MGM"
    CMDSQL = CMDSQL + " WHERE RECSOURCE ='" + Combo1(0).Text + "' AND AGENT ='" + MDIForm1.Text1.Text + "'  ORDER BY CUSTID"
    
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
CMDSQL = Empty
While Not m_objrs.EOF
    CustId = IIf(IsNull(m_objrs("CUSTID")), " ", m_objrs("CUSTID"))
    If CustId <> " " Then
        CMDSQL = "UPDATE MGM"
        CMDSQL = CMDSQL + " SET AGENT = '" + Trim(USERID) + "',"
        CMDSQL = CMDSQL + " NamaAgent = '" + NAMAAGENT + "',"
        CMDSQL = CMDSQL + " NEXTACTDATE = '" + tgljam + "'"
       ' cmdsql = cmdsql + " TGLSTATUS = '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn") + "'"
        CMDSQL = CMDSQL + " WHERE CUSTID = '" + CustId + "'"
        M_OBJCONN.Execute CMDSQL
    End If
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
End Function




Public Function HITUNG_TEMPCUST_CC(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Currency
Dim CMDSQL As String
Dim m_objrs As ADODB.Recordset

CMDSQL = "SELECT COUNT(CUSTID) AS JML FROM cc_custtbl"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_objrs.RecordCount <> 0 Then
    HITUNG_TEMPCUST_CC = m_objrs("JML")
End If
Set m_objrs = Nothing
End Function

Public Function INSERT_DISTRIBUSI(M_RPTCONN As ADODB.Connection, M_OBJCONN As ADODB.Connection, SPVCODE As String, TANGGAL As String)
Dim CMDSQL As String
Dim USERID As String
Dim Nama As String
Dim TGLJAM1 As String
Dim JAM As String
Dim TGLJAM2 As String
Dim i As Integer
Dim m_objrs As ADODB.Recordset

Call DELETE_DISTRIBUSI(M_RPTCONN)

Set m_objrs = QUERY_USER(M_OBJCONN, SPVCODE)
If m_objrs.RecordCount = 0 Then
    ProgressBar1.Max = 100
Else
    ProgressBar1.Max = 100 * (m_objrs.RecordCount + 1)
    
End If
    ProgressBar1.Visible = True
    ProgressBar1.Value = 100
i = 100

TGLJAM2 = Format(TANGGAL, "mm/dd/yy")
JAM = Format(TGLJAM2, "mm/dd/yy") + " " + Format(Now, "hh:mm")
TGLJAM1 = Format(TGLJAM2, "yyyymmdd") + Format(Now, "hhmm")
While Not m_objrs.EOF
    ProgressBar1.Value = i
    USERID = IIf(IsNull(m_objrs("USERID")), "", m_objrs("USERID"))
    Nama = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
    CMDSQL = "INSERT INTO DISTRIBUSI"
    CMDSQL = CMDSQL + " (USERID,"
    CMDSQL = CMDSQL + " TGLJAM,"
    CMDSQL = CMDSQL + " NAMA)"
    CMDSQL = CMDSQL + " VALUES"
    CMDSQL = CMDSQL + " ('" + Trim(USERID) + "',"
    CMDSQL = CMDSQL + " '" + LTrim(TGLJAM1) + "',"
    CMDSQL = CMDSQL + " '" + Trim(Nama) + "')"
    M_RPTCONN.Execute CMDSQL
    m_objrs.MoveNext
    i = i + 100
Wend
    ProgressBar1.Value = ProgressBar1.Max
    ProgressBar1.Visible = False
End Function

Private Function DELETE_DISTRIBUSI(M_RPTCONN As ADODB.Connection)
Dim CMDSQL As String
    CMDSQL = "DELETE * FROM DISTRIBUSI"
    M_RPTCONN.Execute CMDSQL
End Function


Public Function QUERY_USER(M_OBJCONN As ADODB.Connection, SPVCODE As String) As Object
Dim CMDSQL As String
Dim m_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM USERTBL"
CMDSQL = CMDSQL + " WHERE USERTYPE ='1'"
 If Len(SPVCODE) <> 0 Then
    CMDSQL = CMDSQL + " AND SPVCODE = '" + SPVCODE + "'"
 End If
CMDSQL = CMDSQL + " ORDER BY USERID"
    
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_USER = m_objrs
Set m_objrs = Nothing
End Function

Private Function QUERY_TEMPCUSTTBL_KTA(M_OBJCONN As ADODB.Connection, TIPE_PRODUK As String, USERID As String, NAMAAGENT As String)
Dim CMDSQL As String
Dim i As Integer
Dim CustId As String, NAME1 As String, TITLE1 As String, _
                            BIRTHD As String, AddrNow As String, ZIPNOW As String, CITYNOW As String, AHOMENO As String, _
                            HOMENO As String, AHOMENO2 As String, HOMENO2 As String, MOBILENO As String, MOBILENO2 As String, _
                            CAT As String, JENISUSAHA As String, NAMAPT As String, ADDRPT As String, AFAXNO As String, FAXNO As String, _
                            AFAXNO2 As String, FAXNO2 As String, AOFFICENO As String, OFFICENO As String, EXTOFFICENO As String, _
                            AOFFICENO2 As String, OFFICENO2 As String, EXTOFFICENO2 As String, agent As String, NEXTACT As String, _
                            NEXTACTDATE As String, PRODUCTOFFERED As String, VOLOFFERED As String, PRODUCTAPPROVED As String, _
                            VOLAPPROVED As String, RECSOURCE As String, TGLSOURCE As String, RECSTATUS As String, KETHSLKERJA As String, _
                            TGLSTATUS As String, OTHERS As String, NOLAP As String
Dim m_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM MGM"
CMDSQL = CMDSQL + " WHERE AGENT = '" + USERID + "'"
CMDSQL = CMDSQL + " ORDER BY NAME"

Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_objrs.RecordCount = 0 Then
    'FRM_ALIHBLOKDATA .ProgressBar1.Max = 100
Else
   ProgressBar1.Max = 100 * (m_objrs.RecordCount + 1)
    
End If
    SSFrame1.Visible = True
    ProgressBar1.Visible = True
    ProgressBar1.Value = 100
i = 100
While Not m_objrs.EOF
    ProgressBar1.Value = i
    CustId = IIf(IsNull(m_objrs("CUSTID")), "", m_objrs("CUSTID"))
    NOLAP = IIf(IsNull(m_objrs("NOLAP")), "", m_objrs("NOLAP"))
    NAMAAGENT = IIf(IsNull(m_objrs("NamaAgent")), "", m_objrs("NamaAgent"))
    NAME1 = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
    TITLE1 = IIf(IsNull(m_objrs("TITLE")), "", m_objrs("TITLE"))
    BIRTHD = IIf(IsNull(m_objrs("BIRTHD")), "", m_objrs("BIRTHD"))
    AddrNow = IIf(IsNull(m_objrs("ADDRNOW")), "", m_objrs("ADDRNOW"))
    ZIPNOW = IIf(IsNull(m_objrs("ZIPNOW")), "", m_objrs("ZIPNOW"))
    CITYNOW = IIf(IsNull(m_objrs("CITYNOW")), "", m_objrs("CITYNOW"))
    AHOMENO = IIf(IsNull(m_objrs("AHOMENO")), "", m_objrs("AHOMENO"))
    HOMENO = IIf(IsNull(m_objrs("HOMENO")), "", m_objrs("HOMENO"))
    AHOMENO2 = IIf(IsNull(m_objrs("AHOMENO2")), "", m_objrs("AHOMENO2"))
    HOMENO2 = IIf(IsNull(m_objrs("HOMENO2")), "", m_objrs("HOMENO2"))
    MOBILENO = IIf(IsNull(m_objrs("MOBILENO")), "", m_objrs("MOBILENO"))
    MOBILENO2 = IIf(IsNull(m_objrs("MOBILENO2")), "", m_objrs("MOBILENO2"))
    CAT = IIf(IsNull(m_objrs("CAT")), "", m_objrs("CAT"))
    JENISUSAHA = IIf(IsNull(m_objrs("JENISUSAHA")), "", m_objrs("JENISUSAHA"))
    NAMAPT = IIf(IsNull(m_objrs("NAMAPT")), "", m_objrs("NAMAPT"))
    ADDRPT = IIf(IsNull(m_objrs("ADDRPT")), "", m_objrs("ADDRPT"))
    AFAXNO = IIf(IsNull(m_objrs("AFAXNO")), "", m_objrs("AFAXNO"))
    FAXNO = IIf(IsNull(m_objrs("FAXNO")), "", m_objrs("FAXNO"))
    AFAXNO2 = IIf(IsNull(m_objrs("AFAXNO2")), "", m_objrs("AFAXNO2"))
    FAXNO2 = IIf(IsNull(m_objrs("FAXNO2")), "", m_objrs("FAXNO2"))
    AOFFICENO = IIf(IsNull(m_objrs("AOFFICENO")), "", m_objrs("AOFFICENO"))
    OFFICENO = IIf(IsNull(m_objrs("OFFICENO")), "", m_objrs("OFFICENO"))
    EXTOFFICENO = IIf(IsNull(m_objrs("EXTOFFICE")), "", m_objrs("EXTOFFICE"))
    AOFFICENO2 = IIf(IsNull(m_objrs("AOFFICENO2")), "", m_objrs("AOFFICENO2"))
    OFFICENO2 = IIf(IsNull(m_objrs("OFFICENO2")), "", m_objrs("OFFICENO2"))
    EXTOFFICENO2 = IIf(IsNull(m_objrs("EXTOFFICE2")), "", m_objrs("EXTOFFICE2"))
    agent = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
    NEXTACT = IIf(IsNull(m_objrs("NEXTACT")), "", m_objrs("NEXTACT"))
    NEXTACTDATE = IIf(IsNull(m_objrs("NEXTACTDATE")), "", m_objrs("NEXTACTDATE"))
    RECSOURCE = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
    TGLSOURCE = IIf(IsNull(m_objrs("TGLSOURCE")), "", m_objrs("TGLSOURCE"))
    RECSTATUS = IIf(IsNull(m_objrs("RECSTATUS")), "", m_objrs("RECSTATUS"))
    KETHSLKERJA = IIf(IsNull(m_objrs("KETHSLKERJA")), "", m_objrs("KETHSLKERJA"))
    TGLSTATUS = IIf(IsNull(m_objrs("TGLSTATUS")), "", m_objrs("TGLSTATUS"))
    OTHERS = IIf(IsNull(m_objrs("OTHERS")), "", m_objrs("OTHERS"))
    
    'Call ADD_CUSTTBL(M_OBJCONN, CUSTID, NAME1, TITLE1, _
     '                       BIRTHD, ADDRNOW, ZIPNOW, CITYNOW, AHOMENO, _
      '                      HOMENO, AHOMENO2, HOMENO2, MOBILENO, MOBILENO2, _
       '                     CAT, JENISUSAHA, NAMAPT, ADDRPT, AFAXNO, FAXNO, _
        '                    AFAXNO2, FAXNO2, AOFFICENO, OFFICENO, EXTOFFICENO, _
         '                   AOFFICENO2, OFFICENO2, EXTOFFICENO2, agent, NEXTACT, _
          '                  NEXTACTDATE, RECSOURCE, TGLSOURCE, RECSTATUS, KETHSLKERJA, _
           '                 TGLSTATUS, OTHERS, TIPE_PRODUK, NOLAP, NAMAAGENT)
    On Error GoTo add_error
  '  WaitSecs (1)
        'M_OBJCONN.Execute cmdsql
        ADD_OK = True
        Exit Function
add_error:
            ADD_OK = False

    m_objrs.MoveNext
    i = i + 100
Wend
    ProgressBar1.Value = ProgressBar1.Max
    'WaitSecs (2)
    'Call DELETE_TEMPCUSTTBL(M_OBJCONN, USERID, TIPE_PRODUK)
Set m_objrs = Nothing
End Function

