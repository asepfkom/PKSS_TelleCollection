VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmVisit 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9105
   ClientLeft      =   -3360
   ClientTop       =   45
   ClientWidth     =   12030
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame FrmVisit 
      Height          =   915
      Left            =   90
      TabIndex        =   27
      Top             =   375
      Width           =   11895
      Begin VB.CheckBox Check2 
         Caption         =   "Telah Visit ?"
         Height          =   255
         Left            =   8055
         TabIndex        =   44
         Top             =   150
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   1
         Left            =   5775
         TabIndex        =   35
         Top             =   465
         Width           =   2175
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   0
         Left            =   4695
         TabIndex        =   34
         Top             =   465
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Close"
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
         Left            =   10440
         TabIndex        =   33
         Top             =   135
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Search"
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
         Left            =   9480
         TabIndex        =   32
         Top             =   135
         Width           =   975
      End
      Begin VB.TextBox TxtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1005
         TabIndex        =   31
         Top             =   435
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   5775
         TabIndex        =   30
         Top             =   105
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   4695
         TabIndex        =   29
         Top             =   105
         Width           =   1095
      End
      Begin VB.TextBox TxtCustid 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1005
         TabIndex        =   28
         Top             =   135
         Width           =   1815
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   9480
         TabIndex        =   36
         Top             =   555
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Field Name"
         Height          =   255
         Left            =   2730
         TabIndex        =   40
         Top             =   495
         Width           =   1935
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Desc Caller :"
         Height          =   255
         Left            =   3525
         TabIndex        =   39
         Top             =   135
         Width           =   1125
      End
      Begin VB.Label Label9 
         Caption         =   "Cust Name"
         Height          =   255
         Left            =   60
         TabIndex        =   38
         Top             =   420
         Width           =   945
      End
      Begin VB.Label Label8 
         Caption         =   "Cust Id"
         Height          =   255
         Left            =   375
         TabIndex        =   37
         Top             =   165
         Width           =   615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7770
      Left            =   60
      TabIndex        =   0
      Top             =   1305
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   13705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "UnVisit"
      TabPicture(0)   =   "FrmVisit.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Has Visit"
      TabPicture(1)   =   "FrmVisit.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame Frame1 
         Height          =   7380
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   11655
         Begin VB.Frame Frame2 
            Height          =   7230
            Left            =   6360
            TabIndex        =   3
            Top             =   120
            Width           =   5175
            Begin VB.ComboBox CboField 
               Height          =   315
               Index           =   1
               Left            =   2400
               TabIndex        =   26
               Top             =   3360
               Width           =   2535
            End
            Begin VB.ComboBox CboField 
               Height          =   315
               Index           =   0
               Left            =   1440
               TabIndex        =   25
               Top             =   3360
               Width           =   975
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   1440
               TabIndex        =   23
               Top             =   3720
               Width           =   1815
            End
            Begin VB.TextBox TxtF_CEK 
               Appearance      =   0  'Flat
               Height          =   320
               Left            =   3600
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   600
               Width           =   615
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Sudah Di Visit"
               Height          =   195
               Left            =   2160
               TabIndex        =   18
               Top             =   6960
               Width           =   1335
            End
            Begin Threed.SSCommand SSCommand1 
               Height          =   375
               Left            =   3720
               TabIndex        =   17
               Top             =   6840
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   196610
               Caption         =   "&Update Visit"
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               Height          =   320
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   960
               Width           =   2055
            End
            Begin RichTextLib.RichTextBox TxtDetails1 
               Height          =   1215
               Left            =   1440
               TabIndex        =   15
               Top             =   2040
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   2143
               _Version        =   393217
               Enabled         =   -1  'True
               ReadOnly        =   -1  'True
               Appearance      =   0
               TextRTF         =   $"FrmVisit.frx":0038
            End
            Begin TDBDate6Ctl.TDBDate TDBDate2 
               Height          =   315
               Left            =   1440
               TabIndex        =   14
               Top             =   1680
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   556
               Calendar        =   "FrmVisit.frx":00BA
               Caption         =   "FrmVisit.frx":01D2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmVisit.frx":023E
               Keys            =   "FrmVisit.frx":025C
               Spin            =   "FrmVisit.frx":02BA
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
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
               Left            =   1440
               TabIndex        =   13
               Top             =   1320
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   556
               Calendar        =   "FrmVisit.frx":02E2
               Caption         =   "FrmVisit.frx":03FA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmVisit.frx":0466
               Keys            =   "FrmVisit.frx":0484
               Spin            =   "FrmVisit.frx":04E2
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
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
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               Height          =   320
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   12
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C00000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   320
               Left            =   1440
               TabIndex        =   11
               Top             =   240
               Width           =   1470
            End
            Begin RichTextLib.RichTextBox Txtdetails2 
               Height          =   1215
               Left            =   1440
               TabIndex        =   19
               Top             =   5520
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   2143
               _Version        =   393217
               Enabled         =   -1  'True
               Appearance      =   0
               TextRTF         =   $"FrmVisit.frx":050A
            End
            Begin RichTextLib.RichTextBox TxtAddress 
               Height          =   1215
               Left            =   1440
               TabIndex        =   42
               Top             =   4200
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   2143
               _Version        =   393217
               Enabled         =   -1  'True
               Appearance      =   0
               TextRTF         =   $"FrmVisit.frx":058C
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "Address To Visit "
               Height          =   375
               Left            =   240
               TabIndex        =   41
               Top             =   4200
               Width           =   855
            End
            Begin VB.Label Label14 
               Caption         =   "Field Collector"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   3360
               Width           =   1095
            End
            Begin VB.Label Label12 
               Caption         =   "Status Visit"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   3720
               Width           =   855
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               Caption         =   "Status"
               Height          =   255
               Left            =   3000
               TabIndex        =   20
               Top             =   630
               Width           =   495
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "Visit No"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "Details Visit "
               Height          =   255
               Left            =   240
               TabIndex        =   9
               Top             =   5520
               Width           =   855
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Details Request"
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Visit Date"
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Request Date"
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   1350
               Width           =   1095
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Name"
               Height          =   255
               Left            =   120
               TabIndex        =   5
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Custid"
               Height          =   315
               Left            =   120
               TabIndex        =   4
               Top             =   600
               Width           =   495
            End
         End
         Begin MSComctlLib.ListView LstVisit 
            Height          =   6990
            Left            =   0
            TabIndex        =   2
            Top             =   270
            Width           =   6240
            _ExtentX        =   11007
            _ExtentY        =   12330
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483643
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
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   688
      _Version        =   196610
      Font3D          =   4
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "List Visit Customer"
      BevelWidth      =   2
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
End
Attribute VB_Name = "FrmVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_cari As ADODB.Recordset
Dim m_data As New ClsVisit
Private Sub HEADER_VIEW_visit()
    LstVisit.ColumnHeaders.ADD 1, , "No", 3 * TXT
    LstVisit.ColumnHeaders.ADD 2, , "Visit No", 10 * TXT
    LstVisit.ColumnHeaders.ADD 3, , "Custid", 16 * TXT
    LstVisit.ColumnHeaders.ADD 4, , "Name", 16 * TXT
    LstVisit.ColumnHeaders.ADD 5, , "Request Date", 10 * TXT
    LstVisit.ColumnHeaders.ADD 6, , "VisitDate", 5 * TXT
    LstVisit.ColumnHeaders.ADD 7, , "Details Request", 30 * TXT
    LstVisit.ColumnHeaders.ADD 8, , "Details Visit", 30 * TXT
    LstVisit.ColumnHeaders.ADD 9, , "Status", 3 * TXT
    LstVisit.ColumnHeaders.ADD 10, , "id", 1 * TXT
    LstVisit.ColumnHeaders.ADD 11, , "Sts", 3 * TXT
    LstVisit.ColumnHeaders.ADD 12, , "FFC", 4 * TXT
    LstVisit.ColumnHeaders.ADD 13, , "Agent", 4 * TXT
    LstVisit.ColumnHeaders.ADD 14, , "StatusVisit", 4 * TXT
    LstVisit.ColumnHeaders.ADD 15, , "AddressToVisit", 10 * TXT
    LstVisit.ColumnHeaders.ADD 16, , "VisitKe", 2 * TXT
    
LstVisit.SortKey = 2
LstVisit.Sorted = True
MousePointer = vbNormal
    End Sub
Private Sub ShowVisit()

Dim ListItem As ListItem
Dim Lcustid1 As String
Dim Lcustid2 As String
Dim LCall As String
Dim i As Integer
Dim CMDSQL As String
Dim sPending As String
Dim M_objrs As ADODB.Recordset
Dim VOLUMEAMOUNT As Double
i = 1
On Error GoTo hell
    
    
    LstVisit.ListItems.CLEAR
    Me.MousePointer = vbHourglass
    ProgressBar1.Max = m_cari.RecordCount + 1
    While Not m_cari.EOF
    ProgressBar1.Value = m_cari.Bookmark
        Lcustid1 = CStr(IIf(IsNull(m_cari!CustId), "", m_cari!CustId))
'        sPending = CStr(Trim(IIf(IsNull(m_cari!f_Pending), "", m_cari!f_Pending)))
'        If sPending = "OK" Then sPending = ""
        
        Set ListItem = LstVisit.ListItems.ADD(, , m_cari.Bookmark)
        ListItem.SubItems(1) = IIf(IsNull(m_cari("VisitNo")), "", m_cari("VisitNo"))
        ListItem.SubItems(2) = IIf(IsNull(m_cari("custid")), "", m_cari("Custid"))
        ListItem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
        ListItem.SubItems(4) = IIf(IsNull(m_cari("REquestdate")), "", Format(m_cari("Requestdate"), "yyyy/mm/dd hh:nn"))
        ListItem.SubItems(5) = IIf(IsNull(m_cari("visitdate")), "", Format(m_cari("visitdate"), "yyyy/mm/dd hh:nn"))
        ListItem.SubItems(6) = IIf(IsNull(m_cari("detailsr")), "", m_cari("detailsr"))
        ListItem.SubItems(7) = IIf(IsNull(m_cari("detailsV")), "", m_cari("detailsV"))
        ListItem.SubItems(8) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
        ListItem.SubItems(9) = IIf(IsNull(m_cari("id")), "", m_cari("id"))
        ListItem.SubItems(10) = IIf(IsNull(m_cari("Sts")), "", m_cari("Sts"))
        ListItem.SubItems(11) = IIf(IsNull(m_cari("FFC")), "", m_cari("FFC"))
        ListItem.SubItems(12) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        ListItem.SubItems(13) = IIf(IsNull(m_cari("StatusVisit")), "", m_cari("StatusVisit"))
        ListItem.SubItems(14) = IIf(IsNull(m_cari("AddressToVisit")), "", m_cari("AddressToVisit"))
        ListItem.SubItems(15) = IIf(IsNull(m_cari("VisitKe")), "", m_cari("VisitKe"))
     
        
        m_cari.MoveNext
    Wend
Set m_cari = Nothing
    
        If LstVisit.ListItems.Count = 0 Then
'            TxtJmlDtmgm.Text = "Tidak Ada Data"
'            TxtJmlVolmgm.Text = "0"
'        Else
'            TxtJmlDtmgm.Text = "Total " + CStr(m_cari.RecordCount) + " Records"
'            TxtJmlVolmgm.Text = "Total " + CStr(Format(VOLUMEAMOUNT, "##,###"))
        End If
LstVisit.SortKey = 2
LstVisit.Sorted = True
ProgressBar1.Value = 0
ProgressBar1.Visible = False
MousePointer = vbNormal

Exit Sub
hell:
    Me.MousePointer = vbNormal
    MsgBox err.Description
  ''  Resume
End Sub


Private Sub CboField_Click(Index As Integer)
Dim m_data As New CLS_FRMSEARCH
Dim M_objrs As ADODB.Recordset
Select Case Index
Case 0
    Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + CboField(Index).Text + "'")
    If M_objrs.RecordCount <> 0 Then
        CboField(0).Text = M_objrs("USERID")
        CboField(1).Text = M_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 1
    Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + CboField(Index).Text + "'")
    If M_objrs.RecordCount <> 0 Then
        CboField(0).Text = M_objrs("USERID")
        CboField(1).Text = M_objrs("AGENT")
    Else
        CboField(0).Text = Empty
        CboField(1).Text = Empty
    End If
    
 End Select
 
 Set m_data = Nothing
Set M_objrs = Nothing
End Sub
Private Sub Update_Visit()
Dim M_update As New ADODB.Recordset
Dim CMDSQL As String

If Len(Text1.Text) > 0 Then
Set M_update = New ADODB.Recordset
M_update.CursorLocation = adUseClient
CMDSQL = "SELECT * FROM mandiri.TblVisit where id = '" + LstVisit.SelectedItem.SubItems(9) + "'"
M_update.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

M_update!VisitDate = IIf(IsNull(TDBDate2.Value), Null, Format(TDBDate2.Value, "yyyy-mm-dd"))
M_update!DetailsV = IIf(IsNull(Txtdetails2.Text), "", Trim(Txtdetails2.Text))
M_update!StatusVisit = IIf(IsNull(Combo2.Text), "", Trim(Combo2.Text))
M_update!FFC = Trim(IIf(IsNull(CboField(0).Text), "", CboField(0).Text))
 If Check1.Value Then
M_update!STS = "1"
  End If
M_update.update
MsgBox "Update Done...!"
Set M_update = Nothing
Else
MsgBox "Pilih Data yang akan Di Update...!"
Exit Sub
End If

End Sub
Private Sub Combo1_Click(Index As Integer)
Dim m_data As New CLS_FRMSEARCH
Dim M_objrs As ADODB.Recordset
Select Case Index
Case 0
    Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(0).Text = M_objrs("USERID")
        Combo1(1).Text = M_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 1
    Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(0).Text = M_objrs("USERID")
        Combo1(1).Text = M_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
    
 End Select
 
 Set m_data = Nothing
Set M_objrs = Nothing
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim m_data As New CLS_FRMSEARCH
Dim M_objrs As ADODB.Recordset
Select Case Index
Case 0
    Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(0).Text = M_objrs("USERID")
        Combo1(1).Text = M_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 1
    Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(0).Text = M_objrs("USERID")
        Combo1(1).Text = M_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
 End Select
End Sub

Private Sub Command1_Click(Index As Integer)
Dim NAMACUST As String
Dim NamaAgent As String
Dim DATASOURCE As String
Dim TGLLAHIR As String
Dim OFFPHONE As String
Dim OFFPHONE2 As String
Dim HOMEPHONE As String
Dim HOMEPHONE2 As String
Dim MOBILEPHONE As String
Dim MOBILEPHONE2 As String
Dim FAXPHONE As String
Dim Lcustid As String
Dim FAXPHONE2 As String
Dim KETHSLKERJA As String
Dim lLastCallDate As String
Dim STS As String



Dim m_data As New CLS_FRMSEARCH
Dim M_objrs As New ADODB.Recordset
Dim PANJANG As Integer

Select Case Index
    Case 0
  
'        If Trim(TxtCustid.Text) = Empty And Trim(Combo1(0).Text) = Empty And Combo1(1).Text = Empty And Trim(TxtName.Text) = Empty And Check2.Value = False Then
'            MsgBox "Masukan Kriteria Customer Yang Akan Dicari...!!!", vbCritical + vbOKOnly, "Peringatan"
'            TxtName.SetFocus
'            Set M_DATA = Nothing
'            Exit Sub
'        Else
        
        LstVisit.ListItems.CLEAR
         
            If TxtCustid.Text <> Empty Then
                Lcustid = "tblvisit.CUSTID LIKE " + "'%" + UBAH_QUOTE(TxtCustid.Text) + "%'"
            End If
                If txtname.Text <> Empty Then
                    NAMACUST = "mgm.NAME LIKE " + "'%" + UBAH_QUOTE(txtname.Text) + "%'"
                End If
                If Combo1(0).Text <> Empty Then
                    NamaAgent = "mgm.AGENT = '" + Trim(Combo1(0).Text) + "'"
                 End If
                  If Check2.Value Then
                     STS = "tblvisit.sts = '1' "
                     Else
                     If Check2.Value = False Then
                        STS = "tblvisit.sts = '0' "
                     End If
                   End If
          
                
                    Set m_cari = m_data.QUERY_SEARCH_VISIT(M_OBJCONN, NAMACUST, NamaAgent, Lcustid, STS)
            If m_cari.RecordCount = 0 Then
                MsgBox "Data Tidak Ditemukan", vbInformation + vbOKOnly, "Aplikasi"
                Set m_data = Nothing
                Exit Sub
            Else
               
                search_ok = True
                    SSTab1.Tab = 0
                    Call ShowVisit
        End If
'End If
      Set m_data = Nothing
      ' Frame3.Visible = False

Case 1
Unload Me

End Select

End Sub

Private Sub Form_Load()
Dim M_objrs As ADODB.Recordset
Dim m_objrs2 As ADODB.Recordset
Dim m_data As New CLS_FRMSEARCH
Dim CMDSQL As String
SSTab1.TabVisible(1) = False

Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "")
    While Not M_objrs.EOF
        Combo1(0).AddItem M_objrs("USERID")
        Combo1(1).AddItem M_objrs("AGENT")
        M_objrs.MoveNext
    Wend
Set M_objrs = Nothing


Set m_objrs2 = New ADODB.Recordset
m_objrs2.CursorLocation = adUseClient
CMDSQL = "select DISTINCT F_CEK  from mandiri.TBLVISIT"
m_objrs2.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

While Not m_objrs2.EOF
         Combo2.AddItem Trim(m_objrs2!F_CEK)
         m_objrs2.MoveNext
Wend
Set m_objrs2 = Nothing

'-------->> Tampil Field Collector <<------------
Set m_objrs2 = New ADODB.Recordset
m_objrs2.CursorLocation = adUseClient
CMDSQL = "SELECT USERID,AGENT FROM mandiri.usertbl WHERE USERTYPE='2' ORDER BY USERID"
m_objrs2.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

While Not m_objrs2.EOF
         CboField(0).AddItem Trim(m_objrs2!USERID)
         CboField(1).AddItem Trim(m_objrs2!AGENT)
         m_objrs2.MoveNext
Wend
Set m_objrs2 = Nothing

Call HEADER_VIEW_visit
End Sub
Private Sub Tampil_Data()
If LstVisit.ListItems.Count = 0 Then
Exit Sub
Else
Text1.Text = LstVisit.SelectedItem.ListSubItems(1)
Text2.Text = LstVisit.SelectedItem.ListSubItems(2)
Text3.Text = LstVisit.SelectedItem.ListSubItems(3)
TDBDate1.Value = LstVisit.SelectedItem.ListSubItems(4)
TDBDate2.Value = LstVisit.SelectedItem.ListSubItems(5)
TxtDetails1.Text = LstVisit.SelectedItem.ListSubItems(6)
TxtF_CEK.Text = LstVisit.SelectedItem.ListSubItems(8)
CboField(0).Text = LstVisit.SelectedItem.ListSubItems(11)
Combo2.Text = LstVisit.SelectedItem.ListSubItems(13)
Txtdetails2.Text = LstVisit.SelectedItem.ListSubItems(7)
TxtAddress.Text = LstVisit.SelectedItem.ListSubItems(14)
If LstVisit.SelectedItem.ListSubItems(10).Text = "1" Then
   Check1.Value = 1
   Else
      If LstVisit.SelectedItem.ListSubItems(10).Text = "0" Then
   Check1.Value = 0
      End If
End If
       
End If
End Sub
Private Sub LstVisit_Click()
Call Tampil_Data
End Sub



Private Sub LstVisit_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

LstVisit.SortKey = ColumnHeader.Index - 1
LstVisit.Sorted = True

End Sub

Private Sub LstVisit_KeyDown(KeyCode As Integer, Shift As Integer)
Call Tampil_Data
End Sub

Private Sub LstVisit_KeyUp(KeyCode As Integer, Shift As Integer)
Call Tampil_Data
End Sub

Private Sub SSCommand1_Click()
If LstVisit.ListItems.Count = 0 Then
Exit Sub
Else
Call Update_Visit
End If
'If Check1.Value = 1 Then
'   If TDBDate2.Value = Empty Then
'   MsgBox "tanggal Visit harus di isi..!"
'      TDBDate2.SetFocus
'    End If
'      If Txtdetails2.Text = Empty Then
'      MsgBox "Details Visit harus di isi..!"
'      Txtdetails2.SetFocus
'      End If
'      M_DATA.UPDATE_RequestVisit M_OBJCONN, TDBDate1.Value, "2007-07-07", Txtdetails2.Text, Combo2.Text, LstVisit.SelectedItem.SubItems(9)
'
'                    On Error GoTo add_error
'                    If M_DATA.ADD_OK Then
'                        'LstPayment.SelectedItem.SubItems(1) = ""
''                        LstPayment.SelectedItem.SubItems(2) = .TDBDate1.Value
''                        LstPayment.SelectedItem.SubItems(3) = .TDBNumber1.Value
''
'
'                    On Error GoTo 0
'                    End If
''                End If
'add_error:

End Sub

