VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmPreembosReport 
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11400
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   3405
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbCek 
      Height          =   315
      Left            =   8925
      TabIndex        =   24
      Top             =   2370
      Width           =   2235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   360
      Index           =   0
      Left            =   8940
      TabIndex        =   12
      Top             =   2730
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Index           =   1
      Left            =   10230
      TabIndex        =   13
      Top             =   2730
      Width           =   1125
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   9060
      TabIndex        =   7
      Top             =   1695
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   6570
      TabIndex        =   6
      Top             =   1695
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose One..."
      Height          =   1035
      Left            =   5550
      TabIndex        =   14
      Top             =   570
      Width           =   5805
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   1
         Left            =   3630
         TabIndex        =   2
         Top             =   195
         Width           =   2130
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   1215
         TabIndex        =   1
         Top             =   195
         Width           =   2085
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   1
         Left            =   3630
         TabIndex        =   5
         Top             =   540
         Width           =   2130
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   0
         Left            =   1230
         TabIndex        =   4
         Top             =   540
         Width           =   2085
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Agent        :"
         Height          =   270
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   1125
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Supervisor :"
         Height          =   270
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   555
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   300
         Index           =   2
         Left            =   3375
         TabIndex        =   18
         Top             =   225
         Width           =   270
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   300
         Index           =   6
         Left            =   3405
         TabIndex        =   17
         Top             =   570
         Width           =   270
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3405
      Left            =   -30
      TabIndex        =   16
      Top             =   -15
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   6006
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
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
   Begin Crystal.CrystalReport RPT 
      Left            =   5820
      Top             =   2685
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   5535
      TabIndex        =   15
      Top             =   3165
      Visible         =   0   'False
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Index           =   1
      Left            =   9060
      TabIndex        =   10
      Top             =   2010
      Width           =   1440
      _Version        =   65536
      _ExtentX        =   2540
      _ExtentY        =   556
      Calendar        =   "FrmPreEmbosReport.frx":0000
      Caption         =   "FrmPreEmbosReport.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmPreEmbosReport.frx":0184
      Keys            =   "FrmPreEmbosReport.frx":01A2
      Spin            =   "FrmPreEmbosReport.frx":0200
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
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Index           =   0
      Left            =   6570
      TabIndex        =   8
      Top             =   2010
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmPreEmbosReport.frx":0228
      Caption         =   "FrmPreEmbosReport.frx":0340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmPreEmbosReport.frx":03AC
      Keys            =   "FrmPreEmbosReport.frx":03CA
      Spin            =   "FrmPreEmbosReport.frx":0428
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
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBTime6Ctl.TDBTime DTimeLastCall 
      Height          =   300
      Index           =   0
      Left            =   7950
      TabIndex        =   9
      Top             =   2025
      Width           =   885
      _Version        =   65536
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "FrmPreEmbosReport.frx":0450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmPreEmbosReport.frx":04BC
      Spin            =   "FrmPreEmbosReport.frx":050C
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
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   0.870289351851852
   End
   Begin TDBTime6Ctl.TDBTime DTimeLastCall 
      Height          =   300
      Index           =   1
      Left            =   10485
      TabIndex        =   11
      Top             =   2010
      Width           =   885
      _Version        =   65536
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "FrmPreEmbosReport.frx":0534
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmPreEmbosReport.frx":05A0
      Spin            =   "FrmPreEmbosReport.frx":05F0
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
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   0.870289351851852
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Status Cek :"
      Height          =   315
      Left            =   7635
      TabIndex        =   25
      Top             =   2385
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "to"
      Height          =   255
      Index           =   1
      Left            =   8850
      TabIndex        =   23
      Top             =   1710
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "From Batch :"
      Height          =   300
      Index           =   0
      Left            =   5595
      TabIndex        =   22
      Top             =   1725
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Date :"
      Height          =   300
      Index           =   5
      Left            =   5595
      TabIndex        =   21
      Top             =   2055
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "to"
      Height          =   300
      Index           =   4
      Left            =   8850
      TabIndex        =   20
      Top             =   2040
      Width           =   270
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5565
      TabIndex        =   19
      Top             =   75
      Width           =   5745
   End
End
Attribute VB_Name = "FrmPreembosReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AmbilDtYgDiFU_PerBatch_Leads()
Dim m_hst As New ADODB.Recordset
Dim tglawal As String
Dim tglakhir As String
Dim LAgent As String
Dim Lbatch As String
Dim cmdsql As String
Dim m_msgbox As Variant
DoEvents
On Error GoTo addd1
tglawal = Format(TDBDate1(0).Value, "yyyymmdd") & Format(DTimeLastCall(0).Value, "hhnn")
tglakhir = Format(TDBDate1(1).Value, "yyyymmdd") & Format(DTimeLastCall(1).Value, "hhnn")
m_hst.CursorLocation = adUseClient
cmdsql = " SELECT RECSOURCEREF, KETHSLKERJA, COUNT(KETHSLKERJA) AS JUMLAH FROM"
cmdsql = cmdsql + " (select custid, recsourceREF,kethslkerja, agent from CC_CUSTTBL"
cmdsql = cmdsql + " where custid in (Select distinct custid from CC_CUSTHSTTBL "
'where [datetime] between '" + tglawal + "' and '" + tglakhir + "'))"
cmdsql = cmdsql + " where TGL Between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'))"
cmdsql = cmdsql + " A GROUP BY RECSOURCEREF, KETHSLKERJA"
m_hst.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        Lbatch = CStr(IIf(IsNull(m_hst!RECSOURCEREF), "", m_hst!RECSOURCEREF))
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + "[" + Trim(CStr(IIf(IsNull(m_hst!KETHSLKERJA), "", m_hst!KETHSLKERJA))) + "]"
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_hst!JUMLAH), 0, m_hst!JUMLAH)) + ""
        cmdsql = cmdsql + " Where BATCH = '" + Lbatch + "'"
        If IsNull(m_hst!KETHSLKERJA) Then
        Else
            If m_hst!KETHSLKERJA = "" Then
            Else
                If m_hst!JUMLAH = 0 Then
                Else
                    M_RPTCONN.Execute cmdsql
                End If
            End If
        End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
Exit Sub
addd1:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub AmbilDtYgDiFU_PerAgent_Leads()
Dim m_hst As New ADODB.Recordset
Dim tglawal As String
Dim tglakhir As String
Dim LAgent As String
Dim cmdsql As String
Dim m_msgbox As Variant
DoEvents
On Error GoTo add2
tglawal = Format(TDBDate1(0).Value, "yyyymmdd") & Format(DTimeLastCall(0).Value, "hhnn")
tglakhir = Format(TDBDate1(1).Value, "yyyymmdd") & Format(DTimeLastCall(1).Value, "hhnn")
m_hst.CursorLocation = adUseClient
'cmdsql = "SELECT AGENT, KETHSLKERJA, COUNT(KETHSLKERJA) AS JUMLAH FROM (select custid, kethslkerja, recsource from cc_custtbl where custid in (Select distinct custid from where [datetime] between '" + tglawal + "' and '" + tglakhir + "')) A GROUP BY AGENT, KETHSLKERJA"
cmdsql = " SELECT AGENT, KETHSLKERJA, COUNT(KETHSLKERJA) AS JUMLAH FROM"
cmdsql = cmdsql + " (select custid, recsource,kethslkerja, agent from CC_CUSTTBL"
cmdsql = cmdsql + " where custid in (Select distinct custid from CC_CUSTHSTTBL "
'where [datetime] between '" + tglawal + "' and '" + tglakhir + "'))"
cmdsql = cmdsql + " where TGL Between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'))"
cmdsql = cmdsql + " A GROUP BY AGENT, KETHSLKERJA"
m_hst.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent))
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + "[" + Trim(CStr(IIf(IsNull(m_hst!KETHSLKERJA), "", m_hst!KETHSLKERJA))) + "]"
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_hst!JUMLAH), 0, m_hst!JUMLAH)) + ""
        cmdsql = cmdsql + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!KETHSLKERJA) Then
        Else
            If m_hst!KETHSLKERJA = "" Then
            Else
                If m_hst!JUMLAH = 0 Then
                Else
                    M_RPTCONN.Execute cmdsql
                End If
            End If
        End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
Exit Sub
add2:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub AmbilDtYgDiFU_PerBatch()
Dim m_hst As New ADODB.Recordset
Dim tglawal As String
Dim tglakhir As String
Dim LAgent As String
Dim Lbatch As String
Dim cmdsql As String
Dim m_msgbox As Variant
DoEvents
On Error GoTo add3
tglawal = Format(TDBDate1(0).Value, "yyyymmdd") & Format(DTimeLastCall(0).Value, "hhnn")
tglakhir = Format(TDBDate1(1).Value, "yyyymmdd") & Format(DTimeLastCall(1).Value, "hhnn")
m_hst.CursorLocation = adUseClient
cmdsql = " SELECT RECSOURCE, KETHSLKERJA, COUNT(KETHSLKERJA) AS JUMLAH FROM"
cmdsql = cmdsql + " (select custid, recsource,kethslkerja, agent from MGM"
cmdsql = cmdsql + " where custid in (Select distinct custid from MGM_HST "
'where [datetime] between '" + tglawal + "' and '" + tglakhir + "'))"
cmdsql = cmdsql + " where TGL Between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'))"
cmdsql = cmdsql + " A GROUP BY RECSOURCE, KETHSLKERJA"
m_hst.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        Lbatch = CStr(IIf(IsNull(m_hst!RECSOURCE), "", m_hst!RECSOURCE))
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + "[" + Trim(CStr(IIf(IsNull(m_hst!KETHSLKERJA), "", m_hst!KETHSLKERJA))) + "]"
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_hst!JUMLAH), 0, m_hst!JUMLAH)) + ""
        cmdsql = cmdsql + " Where BATCH = '" + Lbatch + "'"
        If IsNull(m_hst!KETHSLKERJA) Then
        Else
            If m_hst!KETHSLKERJA = "" Then
            Else
                If m_hst!JUMLAH = 0 Then
                Else
                    M_RPTCONN.Execute cmdsql
                End If
            End If
        End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
Exit Sub
add3:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub AmbilDtYgDiFU_PerAgent()
Dim m_hst As New ADODB.Recordset
Dim tglawal As String
Dim tglakhir As String
Dim LAgent As String
Dim cmdsql As String
Dim m_msgbox As Variant
DoEvents
On Error GoTo add3
tglawal = Format(TDBDate1(0).Value, "yyyymmdd") & Format(DTimeLastCall(0).Value, "hhnn")
tglakhir = Format(TDBDate1(1).Value, "yyyymmdd") & Format(DTimeLastCall(1).Value, "hhnn")
m_hst.CursorLocation = adUseClient
'cmdsql = "SELECT AGENT, KETHSLKERJA, COUNT(KETHSLKERJA) AS JUMLAH FROM (select custid, kethslkerja, recsource from cc_custtbl where custid in (Select distinct custid from where [datetime] between '" + tglawal + "' and '" + tglakhir + "')) A GROUP BY AGENT, KETHSLKERJA"
cmdsql = " SELECT AGENT, KETHSLKERJA, COUNT(KETHSLKERJA) AS JUMLAH FROM"
cmdsql = cmdsql + " (select custid, recsource,kethslkerja, agent from MGM"
cmdsql = cmdsql + " where custid in (Select distinct custid from MGM_HST "
'where [datetime] between '" + tglawal + "' and '" + tglakhir + "'))"
cmdsql = cmdsql + " where TGL Between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'))"
cmdsql = cmdsql + " A where LEFT(RECSOURCE,3) ='PRE' AND recsource between "
cmdsql = cmdsql + " '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' "
cmdsql = cmdsql + " GROUP BY AGENT, KETHSLKERJA"
m_hst.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent))
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + "[" + Trim(CStr(IIf(IsNull(m_hst!KETHSLKERJA), "", m_hst!KETHSLKERJA))) + "]"
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_hst!JUMLAH), 0, m_hst!JUMLAH)) + ""
        cmdsql = cmdsql + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!KETHSLKERJA) Then
        Else
            If m_hst!KETHSLKERJA = "" Then
            Else
                If m_hst!JUMLAH = 0 Then
                Else
                   
                    M_RPTCONN.Execute cmdsql
                End If
            End If
        End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
Exit Sub
add3:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub Combo1_Click(Index As Integer)
    Call Combo1_LostFocus(Index)
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim m_objrs As New ADODB.Recordset
On Error GoTo Combo1_LostFocusErr
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from datasourcetbl where kodeds ='" + Combo1(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not m_objrs.EOF Then
        Combo1(Index).Text = m_objrs!KODEDS
    Else
        Combo1(Index).Text = Empty
    End If
Exit Sub
Combo1_LostFocusErr:
    MsgBox Err.Description
End Sub

Private Sub Combo2_Click(Index As Integer)
    Call Combo2_LostFocus(Index)
End Sub

Private Sub Combo2_LostFocus(Index As Integer)
Dim m_objrs As New ADODB.Recordset
On Error GoTo Combo2_LostFocusErr
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from USERTBL where AKTIF = 0 AND USERID ='" + Combo2(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not m_objrs.EOF Then
        Combo2(Index).Text = m_objrs!USERID
    Else
        Combo2(Index).Text = Empty
    End If
Exit Sub
Combo2_LostFocusErr:
MsgBox Err.Description
End Sub

Private Sub Command1_Click(Index As Integer)
Dim m_msgbox As Variant
On Error GoTo Command1_ClickeR
If TDBDate1(0).ValueIsNull And TDBDate1(1).ValueIsNull Then
    TDBDate1(0).Value = "01/01/1990"
    TDBDate1(1).Value = "31/12/2020"
End If
If Combo1(0).Text = Empty And Combo1(1).Text = Empty Then
    Combo1(0).Text = "-----"
    Combo1(1).Text = "ZZZZZ"
End If
If Option1(0).Value = False And Option1(1).Value = False Then
    If Combo2(0).Text = Empty And Combo2(1).Text = Empty Then
        Combo2(0).Text = "-----"
        Combo2(1).Text = "ZZZZZ"
    End If
End If
ProgressBar1.Visible = True
Select Case Index
    Case 0
    Select Case ListView1.SelectedItem.Text
        Case 9
            Call Isi_Agent_MGM
            Call hitung_JmlData_PerAgent_MGM
            Call AmbilDtYgDiFU_PerAgent
            ' Call HitungJmlAgreePerAgent
            Call Hitung_JmlLeadsPerAgent
            Call hitung_BatchCallInitilized_PerAgent_MGM
            WaitSecs (1)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.ReportFileName = App.Path + "\Report\PreStepPerAgent.rpt"
            Call SHOW_PRN
        Case 10
            Call Isi_Agent_MGM
            Call hitung_JmlData_PerAgent_MGM
            Call AmbilDtYgDiFU_PerAgent
            ' Call HitungJmlAgreePerAgent
            Call Hitung_JmlLeadsPerAgent
            Call hitung_BatchCallInitilized_PerAgent_MGM
            WaitSecs (1)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.ReportFileName = App.Path + "\Report\PreStepPerTeam.rpt"
            Call SHOW_PRN
        Case 11
            Call ISI_DATASOURCE_MGM
            Call AmbilDtYgDiFU_PerBatch
            'Call HitungJmlAgreePerBatch
            Call hitung_JmlData_MGM
       '      Call Hitung_JmlLeadsPerDs
            Call hitung_BatchCallInitilized_MGM
            WaitSecs (1)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.ReportFileName = App.Path + "\Report\PreStepPerBatch.rpt"
            Call SHOW_PRN
        Case 12
            WaitSecs (1)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.Formulas(2) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & Format(DTimeLastCall(0).Value, "hh:nn")) + "')"
            RPT.Formulas(3) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & Format(DTimeLastCall(1).Value, "hh:nn")) + "')"
            RPT.Formulas(4) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(5) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(6) = "@RecsourceAwal = totext('" + CStr(Combo1(0).Text) + "')"
            RPT.Formulas(7) = "@RecsourceAkhir = totext('" + CStr(Combo1(1).Text) + "')"
            If Option1(0).Value = True Then
                RPT.Formulas(8) = "@AgentAwal = totext('" + CStr(Combo2(0).Text) + "')"
                RPT.Formulas(9) = "@AgentAkhir = totext('" + CStr(Combo2(1).Text) + "')"
                RPT.ReportFileName = App.Path + "\Report\pRERptSubmitedPerAgent.rpt"
            Else
                RPT.Formulas(8) = "@SpvAwal = totext('" + CStr(Combo3(0).Text) + "')"
                RPT.Formulas(9) = "@SpvAkhir = totext('" + CStr(Combo3(1).Text) + "')"
                RPT.ReportFileName = App.Path + "\Report\pRERptSubmitedPerSupervisor.rpt"
            End If
            Call SHOW_PRN
            
       Case 13
           Call Isi_Agent_MGM
           Call totincoming
           WaitSecs (2)
           RPT.Reset
           RPT.Formulas(1) = "@periode = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
           RPT.Formulas(2) = "@periode1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
           RPT.Formulas(3) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
           RPT.ReportFileName = App.Path + "\Report\RptAgreePerTSA.rpt"
           Call SHOW_PRN
            
    End Select
    ProgressBar1.Visible = False
    Case 1
        Unload Me
End Select
ProgressBar1.Visible = False
Exit Sub
Command1_ClickeR:
    MsgBox Err.Description
        m_msgbox = MsgBox("Retry..???", vbRetryCancel, "Telegrandi")
    If m_msgbox = vbRetry Then
        Resume
    End If
End Sub

Private Sub Hitung_JmlLeadsPerAgent()
Dim m_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim cmdsql As String
Dim LAgent As String
Dim m_msgbox As Variant
On Error GoTo Hitung_JmlLeadsPerAgent
    Me.MousePointer = vbHourglass
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    cmdsql = "Select Agent, count(custid) as jumlah from CC_Custtbl where tglsource between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and RecSourceRef >='" + Combo1(0).Text + "' and RecSourceRef <='" + Combo1(1).Text + "' and tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
    m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = m_objrs.RecordCount + 2
    While Not m_objrs.EOF
        ProgressBar1.Value = m_objrs.Bookmark
        LAgent = CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent))
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + " LEADS "
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + ""
        cmdsql = cmdsql + " Where AOC = '" + LAgent + "'"
        M_RPTCONN.Execute cmdsql
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_JmlLeadsPerAgent:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub Hitung_JmlLeadsPerDs()
Dim m_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim cmdsql As String
Dim LRecSourceRef As String
On Error GoTo Hitung_JmlLeadsPerAgent
    Me.MousePointer = vbHourglass
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select RecSourceRef, count(custid) as jumlah from CC_Custtbl where tglsource between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and RecSourceRef >='" + Combo1(0).Text + "' and RecSourceRef <='" + Combo1(1).Text + "'  group by RecSourceRef", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = m_objrs.RecordCount + 2
    While Not m_objrs.EOF
        ProgressBar1.Value = m_objrs.Bookmark
        LRecSourceRef = CStr(IIf(IsNull(m_objrs!RECSOURCEREF), "", m_objrs!RECSOURCEREF))
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + " LEADS "
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + ""
        cmdsql = cmdsql + " Where BATCH = '" + LRecSourceRef + "'"
        M_RPTCONN.Execute cmdsql
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_JmlLeadsPerAgent:
Me.MousePointer = vbNormal
MsgBox Err.Description
'Resume
End Sub


Private Sub ISI_DATASOURCE_Leads()
Dim m_objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
M_RPTCONN.Execute "Delete * from TrackingRptPerPrgBatch"
m_objrs.Open "Select * from datasourcetbl where kodeds >= '" + Combo1(0).Text + "' and kodeds <= '" + Combo1(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
m_objrs1.Open "Select * from TrackingRptPerPrgBatch", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1![batch] = CStr(m_objrs!KODEDS)
    m_objrs1.UPDATE
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Set m_objrs1 = Nothing
End Sub

Private Sub Hitung_TrackingReport_Leads()
Dim m_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim cmdsql As String
On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select recsourceRef, kethslkerja, count(custid) as jumlah from cc_custtbl   where tglstatus between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and RecSourceRef >='" + Combo1(0).Text + "' and RecSourceRef <='" + Combo1(1).Text + "'   group by RecSourceRef, kethslkerja", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = m_objrs.RecordCount + 2
    While Not m_objrs.EOF
        ProgressBar1.Value = m_objrs.Bookmark
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + Trim(CStr(IIf(IsNull(m_objrs!KETHSLKERJA), "", m_objrs!KETHSLKERJA)))
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + ""
        cmdsql = cmdsql + " where BATCH = '" + IIf(IsNull(m_objrs!RECSOURCEREF), "", m_objrs!RECSOURCEREF) + "'"
        If IsNull(m_objrs!KETHSLKERJA) Then
        Else
            If m_objrs!KETHSLKERJA = Empty Then
            Else
            M_RPTCONN.Execute cmdsql
            End If
        End If
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
Me.MousePointer = vbNormal
MsgBox Err.Description
'Resume
End Sub

Private Sub hitung_JmlData_Leads()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim cmdsql As String

On Error GoTo hitung_JmlDataer
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select RecSourceRef as RECSOURCE, count(custid) as jml from cc_custtbl where RecSourceRef >='" + Combo1(0).Text + "' and RecSourceRef <='" + Combo1(1).Text + "'   group by RecSourceRef", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    batch = CStr(IIf(IsNull(m_objrs!RECSOURCE), "", m_objrs!RECSOURCE))
    cmdsql = "Update TrackingRptPerPrgBatch set DATASIZE =" + JUMLAH + " where BATCH ='" + batch + "'"
    M_RPTCONN.Execute cmdsql
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_JmlDataer:
MsgBox Err.Description
End Sub

Private Sub hitung_BatchCallInitilized_Leads()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim cmdsql As String

On Error GoTo hitung_BatchCallInitilizeder
m_objrs.CursorLocation = adUseClient
cmdsql = "Select recsource,count(recsource) as jml from TblPhoneMonitorHst where tgl between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' group by recsource"
'cmdsql1 = "Select recsource, count(recsource) as jml from PHONENO_CALL group by recsource"
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    batch = CStr(IIf(IsNull(m_objrs!RECSOURCE), "", m_objrs!RECSOURCE))
    cmdsql = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + JUMLAH + " where BATCH ='" + batch + "'"
    M_RPTCONN.Execute cmdsql
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_BatchCallInitilizeder:
    MsgBox Err.Description
End Sub

Private Sub ISI_DATASOURCE_MGM()
Dim m_objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
M_RPTCONN.Execute "Delete * from TrackingRptPerPrgBatch"
m_objrs.Open "Select * from datasourcetbl where kodeds >= '" + Combo1(0).Text + "' and kodeds <= '" + Combo1(1).Text + "' and left(kodeds,3) ='PRE'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
m_objrs1.Open "Select * from TrackingRptPerPrgBatch", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1![batch] = CStr(m_objrs!KODEDS)
    m_objrs1.UPDATE
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Set m_objrs1 = Nothing
End Sub

Private Sub Hitung_TrackingReport_MGM()
Dim m_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim cmdsql As String
Dim m_msgbox As Variant
On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select recsource, kethslkerja, count(custid) as jumlah from MGM  where tglstatus between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "'  group by recsource, kethslkerja", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = m_objrs.RecordCount + 2
    While Not m_objrs.EOF
        ProgressBar1.Value = m_objrs.Bookmark
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + Trim(CStr(m_objrs!KETHSLKERJA))
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + ""
        cmdsql = cmdsql + " where BATCH = '" + IIf(IsNull(m_objrs!RECSOURCE), "", m_objrs!RECSOURCE) + "'"
        If IsNull(m_objrs!KETHSLKERJA) Then
        Else
            If m_objrs!KETHSLKERJA = Empty Then
            Else
            M_RPTCONN.Execute cmdsql
            End If
        End If
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub hitung_JmlData_MGM()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim cmdsql As String

On Error GoTo hitung_JmlDataer
m_objrs.CursorLocation = adUseClient
'cmdsql = "Select recsource, count(custid) as jml from MGM  where recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' and tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by recsource"
cmdsql = "Select recsource, count(custid) as jml from MGM  where recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' and tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by recsource"
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    batch = CStr(IIf(IsNull(m_objrs!RECSOURCE), "", m_objrs!RECSOURCE))
    cmdsql = "Update TrackingRptPerPrgBatch set DATASIZE =" + JUMLAH + " where BATCH ='" + batch + "'"
    M_RPTCONN.Execute cmdsql
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_JmlDataer:
MsgBox Err.Description
End Sub

Private Sub hitung_BatchCallInitilized_MGM()
Dim m_objrs As New ADODB.Recordset
Dim m_msgbox As Variant
Dim JUMLAH As Currency
Dim batch As String
Dim cmdsql As String
Dim LAgent As String
On Error GoTo hitung_BatchCallInitilizeder
m_objrs.CursorLocation = adUseClient
cmdsql = "Select recsource ,count(userid) as jml from TblPhoneMonitorHst where tgl between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' group by recsource "
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    'JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    JUMLAH = IIf(IsNull(m_objrs!jml), "", m_objrs!jml)
    LAgent = CStr(IIf(IsNull(m_objrs!RECSOURCE), "", m_objrs!RECSOURCE))
    cmdsql = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + CStr(JUMLAH) + " where BATCH ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.Execute cmdsql
    End If
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_BatchCallInitilizeder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
End Sub

' batas akhir
Private Sub Isi_Agent_MGM()
Dim m_objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
On Error GoTo Isi_AgentErr

M_RPTCONN.Execute "Delete * from TrackingRptPerPrgBatch"

m_objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
If Option1(1).Value = True Then
    m_objrs.Open "Select * from UserTbl where AKTIF = 0 AND SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "' AND USERTYPE =1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
        m_objrs.Open "Select * from UserTbl where AKTIF = 0 AND userid >='" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' AND USERTYPE =1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    Else
        m_objrs.Open "Select * from UserTbl where AKTIF = 0 AND USERTYPE =1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
End If
m_objrs1.Open "Select * from TrackingRptPerPrgBatch", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    m_objrs1.AddNew
    m_objrs1!TEAM = CStr(IIf(IsNull(m_objrs!TEAM), "", m_objrs!TEAM))
    m_objrs1!TSRNAME = CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent))
    m_objrs1!TEAM = CStr(IIf(IsNull(m_objrs!SPVCODE), "", m_objrs!SPVCODE))
    m_objrs1!AOC = CStr(IIf(IsNull(m_objrs!USERID), "", m_objrs!USERID))
    m_objrs1.UPDATE
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox Err.Description
    'Resume
End Sub

Private Sub hitung_JmlData_PerAgent_MGM()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim cmdsql As String
Dim LAgent As String
On Error GoTo hitung_JmlDataer
m_objrs.CursorLocation = adUseClient
'cmdsql = "Select Agent, count(custid) as jml from MGM  where LEFT(RECSOURCE,3) ='PRE' AND recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
cmdsql = "Select Agent, count(custid) as jml from MGM  where LEFT(RECSOURCE,3) ='PRE' AND recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    LAgent = CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent))
    cmdsql = "Update TrackingRptPerPrgBatch set DATASIZE =" + JUMLAH + " where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute cmdsql
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_JmlDataer:
MsgBox Err.Description

End Sub

Private Sub Hitung_TrackingReportPerAgent_MGM()
Dim m_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim cmdsql As String
Dim LAgent As String

On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    cmdsql = "Select AGENT, kethslkerja, count(AGENT) as jumlah from MGM where tglstatus between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "'group by AGENT, kethslkerja"
    m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = m_objrs.RecordCount + 1
    While Not m_objrs.EOF
        ProgressBar1.Value = m_objrs.Bookmark
'         WaitSecs (0.5)
        LAgent = CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent))
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + "[" + Trim(CStr(m_objrs!KETHSLKERJA)) + "]"
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + ""
        cmdsql = cmdsql + " Where AOC = '" + LAgent + "'"
        If IsNull(m_objrs!KETHSLKERJA) Then
        Else
            If m_objrs!KETHSLKERJA = "[]" Then
            Else
                If m_objrs!JUMLAH = 0 Then
                Else
                   
                    M_RPTCONN.Execute cmdsql
                End If
            End If
        End If
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
Me.MousePointer = vbNormal
MsgBox Err.Description
'Resume
End Sub

Private Sub hitung_BatchCallInitilized_PerAgent_MGM()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim cmdsql As String
Dim LAgent As String
On Error GoTo hitung_BatchCallInitilizeder
m_objrs.CursorLocation = adUseClient
cmdsql = "Select userid,count(userid) as jml from TblPhoneMonitorHst where tgl between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and LEFT(RECSOURCE,3) ='PRE' AND recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND RECSOURCE <>'MGM-REF' group by userid"
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    LAgent = CStr(IIf(IsNull(m_objrs!USERID), "", m_objrs!USERID))
    cmdsql = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + JUMLAH + " where AOC ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.Execute cmdsql
    End If
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_BatchCallInitilizeder:
    MsgBox Err.Description
End Sub

Private Sub SHOW_PRN()
    RPT.RetrieveDataFiles
    RPT.WindowLeft = 0
    RPT.WindowTop = 0
    RPT.WindowState = crptMaximized
    RPT.WindowShowPrintBtn = True
    RPT.WindowShowRefreshBtn = True
    RPT.WindowShowSearchBtn = True
    RPT.WindowShowPrintSetupBtn = True
    RPT.WindowControls = True
    RPT.PrintReport
    RPT.Reset
End Sub


Private Sub Isi_Agent_Leads()
Dim m_objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
On Error GoTo Isi_AgentErr

M_RPTCONN.Execute "Delete * from TrackingRptPerPrgBatch"

m_objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
If Option1(1).Value = True Then
    m_objrs.Open "Select * from UserTbl where AKTIF = 0 AND SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "' AND USERTYPE =1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    m_objrs.Open "Select * from UserTbl where AKTIF = 0 AND userid >='" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' AND USERTYPE =1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
End If
m_objrs1.Open "Select * from TrackingRptPerPrgBatch", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    m_objrs1.AddNew
    m_objrs1!TEAM = CStr(IIf(IsNull(m_objrs!TEAM), "", m_objrs!TEAM))
    m_objrs1!TSRNAME = CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent))
    m_objrs1!TEAM = CStr(IIf(IsNull(m_objrs!SPVCODE), "", m_objrs!SPVCODE))
    m_objrs1!AOC = CStr(IIf(IsNull(m_objrs!USERID), "", m_objrs!USERID))
    m_objrs1.UPDATE
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox Err.Description
End Sub

Private Sub hitung_JmlData_PerAgent_Leads()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim cmdsql As String
Dim LAgent As String

On Error GoTo hitung_JmlDataer
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select Agent, count(custid) as jml from CC_CUSTTBL  where recsourceref >='" + Combo1(0).Text + "' and recsourceref <='" + Combo1(1).Text + "' group by Agent", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    LAgent = CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent))
    cmdsql = "Update TrackingRptPerPrgBatch set DATASIZE =" + JUMLAH + " where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute cmdsql
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_JmlDataer:
MsgBox Err.Description
End Sub

Private Sub Hitung_TrackingReportPerAgent_Leads()
Dim m_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim cmdsql As String
Dim LAgent As String

On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    cmdsql = "Select AGENT, kethslkerja, count(AGENT) as jumlah from CC_CUSTTBL where tglstatus between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and RecSourceRef >='" + Combo1(0).Text + "' and RecSourceRef <='" + Combo1(1).Text + "' group by AGENT, kethslkerja"
    m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = m_objrs.RecordCount + 1
    While Not m_objrs.EOF
        ProgressBar1.Value = m_objrs.Bookmark
'         WaitSecs (0.5)
        LAgent = CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent))
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + "[" + Trim(CStr(IIf(IsNull(m_objrs!KETHSLKERJA), "", m_objrs!KETHSLKERJA))) + "]"
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + ""
        cmdsql = cmdsql + " Where AOC = '" + LAgent + "'"
        If IsNull(m_objrs!KETHSLKERJA) Then
        Else
            If m_objrs!KETHSLKERJA = "" Then
            Else
                If m_objrs!JUMLAH = 0 Then
                Else
                   
                    M_RPTCONN.Execute cmdsql
                End If
            End If
        End If
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
Me.MousePointer = vbNormal
MsgBox Err.Description
'Resume
End Sub

Private Sub hitung_BatchCallInitilized_PerAgent_Leads()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim cmdsql As String
Dim LAgent As String
On Error GoTo hitung_BatchCallInitilizeder

m_objrs.CursorLocation = adUseClient
cmdsql = "Select userid,count(userid) as jml from TblPhoneMonitorHst where tgl between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and recsource between 'MGM-REF' and 'MGM-REF'  group by userid"
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    LAgent = CStr(IIf(IsNull(m_objrs!USERID), "", m_objrs!USERID))
    cmdsql = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + JUMLAH + " where AOC ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.Execute cmdsql
    End If
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_BatchCallInitilizeder:
    MsgBox Err.Description
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "No", 4 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Report", 50 * TXT
End Sub

Private Sub Form_Load()
Dim LISTITEM As LISTITEM
Dim m_objrs As ADODB.Recordset
Set m_objrs = New ADODB.Recordset
DTimeLastCall(0).Text = "00:00"
DTimeLastCall(1).Text = "23:59"
m_objrs.CursorLocation = adUseClient
CmbCek.AddItem "Not Check"
CmbCek.AddItem "Accept"
CmbCek.AddItem "RETURN"
m_objrs.Open "SELECT * FROM USERTBL WHERE AKTIF = 0 ORDER BY USERID", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    Combo2(0).AddItem m_objrs!USERID
    Combo2(1).AddItem m_objrs!USERID
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "SELECT * FROM DATASOURCETBL WHERE LEFT(KODEDS,3) ='PRE' ORDER BY KODEDS", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    Combo1(0).AddItem m_objrs!KODEDS
    Combo1(1).AddItem m_objrs!KODEDS
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
'm_objrs.Open "SELECT * FROM spvtbl ORDER BY SPVCODE", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
m_objrs.Open "select distinct SPVTBL.SPVCODE from SPVTBL, USERTBL where SPVTBL.SPVCODE = USERTBL.SPVCODE AND USERTYPE = '6'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

While Not m_objrs.EOF
    Combo3(0).AddItem m_objrs!SPVCODE
    Combo3(1).AddItem m_objrs!SPVCODE
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Call header
'Set LISTITEM = ListView1.ListItems.ADD(, , "1")
'    LISTITEM.SubItems(1) = "Individual Call List Data Tracking Summary"
'Set LISTITEM = ListView1.ListItems.ADD(, , "2")
'    LISTITEM.SubItems(1) = "Individual Leads Data Tracking Summary"
'Set LISTITEM = ListView1.ListItems.ADD(, , "3")
'    LISTITEM.SubItems(1) = "Team Leads Data Tracking Summary"
'Set LISTITEM = ListView1.ListItems.ADD(, , "4")
'    LISTITEM.SubItems(1) = "Team Productivity Based On Data Source"
'Set LISTITEM = ListView1.ListItems.ADD(, , "5")
'    LISTITEM.SubItems(1) = "Team Leads Productivity Based On Data Source"
'Set LISTITEM = ListView1.ListItems.ADD(, , "6")
'    LISTITEM.SubItems(1) = "Team Call List Data Tracking Summary"
'Set LISTITEM = ListView1.ListItems.ADD(, , "7")
'    LISTITEM.SubItems(1) = "Inbound Data"
' report baru
Set LISTITEM = ListView1.ListItems.ADD(, , "9")
    LISTITEM.SubItems(1) = "Premboss Tracking Per Agent"
Set LISTITEM = ListView1.ListItems.ADD(, , "10")
    LISTITEM.SubItems(1) = "Preemboss Tracking Per Team"
Set LISTITEM = ListView1.ListItems.ADD(, , "11")
    LISTITEM.SubItems(1) = "Preemboss Tracking Per Batch Data"
Set LISTITEM = ListView1.ListItems.ADD(, , "12")
    LISTITEM.SubItems(1) = "Application Submitted Preembos"
Set LISTITEM = ListView1.ListItems.ADD(, , "13")
    LISTITEM.SubItems(1) = "Monthly Submitted Preembos"
End Sub

Private Sub ListView1_Click()
    Label2.Caption = ListView1.SelectedItem.SubItems(1)
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
    Case 0
        If Option1(Index).Value = False Then
            Option1(1).Value = False
        Else
            Combo2(0).Enabled = True
            Combo2(1).Enabled = True
            Combo3(0).Enabled = False
            Combo3(1).Enabled = False
        End If
    Case 1
        If Option1(Index).Value = False Then
            Option1(0).Value = False
        Else
            Combo2(0).Enabled = False
            Combo2(1).Enabled = False
            Combo3(0).Enabled = True
            Combo3(1).Enabled = True
        End If
End Select
End Sub

Private Sub Hitung_Tracking_Inc_2Step()
Dim m_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim cmdsql As String
On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select recsourceRef, kethslkerja, count(custid) as jumlah from cc_custtbl   where tglstatus between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and RecSourceRef >='" + Combo1(0).Text + "' and RecSourceRef <='" + Combo1(1).Text + "'  and kethslkerja ='I'  group by RecSourceRef, kethslkerja", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = m_objrs.RecordCount + 2
    While Not m_objrs.EOF
        ProgressBar1.Value = m_objrs.Bookmark
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + Trim(CStr(IIf(IsNull(m_objrs!KETHSLKERJA), "", m_objrs!KETHSLKERJA)))
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + ""
        cmdsql = cmdsql + " where BATCH = '" + IIf(IsNull(m_objrs!RECSOURCEREF), "", m_objrs!RECSOURCEREF) + "'"
        If IsNull(m_objrs!KETHSLKERJA) Then
        Else
            If m_objrs!KETHSLKERJA = Empty Then
            Else
            M_RPTCONN.Execute cmdsql
            End If
        End If
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
Me.MousePointer = vbNormal
MsgBox Err.Description
'Resume
End Sub


Private Sub CallInitilized_PerAgent_2Step()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim cmdsql As String
Dim LAgent As String
Dim m_msgbox As Variant
On Error GoTo hitung_BatchCallInitilizeder
DoEvents
m_objrs.CursorLocation = adUseClient
'cmdsql = "Select userid,count(userid) as jml from TblPhoneMonitorHst where tgl between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' group by userid"
cmdsql = "Select TblPhoneMonitorHst.userid,count(TblPhoneMonitorHst.userid) as jml from TblPhoneMonitorHst,cc_custtbl where TblPhoneMonitorHst.custid = cc_custtbl.custid and TblPhoneMonitorHst.tgl between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  group by userid"
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    LAgent = CStr(IIf(IsNull(m_objrs!USERID), "", m_objrs!USERID))
    cmdsql = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + JUMLAH + " where AOC ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.Execute cmdsql
    End If
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_BatchCallInitilizeder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub CallInitilized_PerBatchSource_2Step()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim cmdsql As String
Dim LAgent As String
Dim m_msgbox As Variant
On Error GoTo hitung_BatchCallInitilizeder
DoEvents
m_objrs.CursorLocation = adUseClient
'cmdsql = "Select userid,count(userid) as jml from TblPhoneMonitorHst where tgl between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' group by userid"
cmdsql = "Select cc_custtbl.RecSourceRef,count(cc_custtbl.RecSourceRef) as jml from TblPhoneMonitorHst,cc_custtbl where TblPhoneMonitorHst.custid = cc_custtbl.custid and TblPhoneMonitorHst.tgl between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  group by cc_custtbl.RecSourceRef"
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    LAgent = CStr(IIf(IsNull(m_objrs!RECSOURCEREF), "", m_objrs!RECSOURCEREF))
    cmdsql = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + JUMLAH + " where batch ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.Execute cmdsql
    End If
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_BatchCallInitilizeder:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub HitungJmlAgreePerAgent()
Dim m_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim cmdsql As String
Dim m_msgbox As Variant
'Dim cmdsql As String
On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    cmdsql = "Select Agent, kethslkerja, count(custid) as jumlah from MGM  where tglstatus between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and RecSource >='" + Combo1(0).Text + "' and RecSource <='" + Combo1(1).Text + "'  and kethslkerja ='A'  group by Agent, kethslkerja"
    m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = m_objrs.RecordCount + 2
    While Not m_objrs.EOF
        ProgressBar1.Value = m_objrs.Bookmark
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + Trim(CStr(IIf(IsNull(m_objrs!KETHSLKERJA), "", m_objrs!KETHSLKERJA)))
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + ""
        cmdsql = cmdsql + " where AOC = '" + IIf(IsNull(m_objrs!agent), "", m_objrs!agent) + "'"
        If IsNull(m_objrs!KETHSLKERJA) Then
        Else
            If m_objrs!KETHSLKERJA = Empty Then
            Else
            M_RPTCONN.Execute cmdsql
            End If
        End If
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
    If Err.Number = -2147217871 Then
        m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
End Sub


Private Sub HitungJmlAgreePerBatch()
Dim m_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim cmdsql As String
On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    cmdsql = "Select Recsource, kethslkerja, count(custid) as jumlah from MGM  where tglstatus between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and RecSource >='" + Combo1(0).Text + "' and RecSource <='" + Combo1(1).Text + "'  and kethslkerja ='A'  group by Recsource, Kethslkerja"
    m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = m_objrs.RecordCount + 2
    While Not m_objrs.EOF
        ProgressBar1.Value = m_objrs.Bookmark
        cmdsql = "Update TrackingRptPerPrgBatch Set "
        cmdsql = cmdsql + Trim(CStr(IIf(IsNull(m_objrs!KETHSLKERJA), "", m_objrs!KETHSLKERJA)))
        cmdsql = cmdsql + "=" + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + ""
        cmdsql = cmdsql + " where BATCH = '" + IIf(IsNull(m_objrs!RECSOURCE), "", m_objrs!RECSOURCE) + "'"
        If IsNull(m_objrs!KETHSLKERJA) Then
        Else
            If m_objrs!KETHSLKERJA = Empty Then
            Else
            M_RPTCONN.Execute cmdsql
            End If
        End If
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
Me.MousePointer = vbNormal
MsgBox Err.Description
Resume
End Sub


Public Sub totincoming()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim cmdsql As String
Dim LAgent As String
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select Agent, count(custid) as jml from MGM where left(recsource,3) = 'PRE' and kethslkerja='A' and tglstatus between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    LAgent = CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent))
    cmdsql = "Update TrackingRptPerPrgBatch set kethslkerja =" + JUMLAH + " where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute cmdsql
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_JmlDataer:
m_msgbox = MsgBox(Err.Description, vbRetryCancel, "Telegrandi")
If m_msgbox = vbRetry Then
    WaitSecs (3)
    Resume
End If
End Sub
