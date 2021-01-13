VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmWeeklyRpt 
   Caption         =   "Weekly Performance Indicator"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   Icon            =   "FrmWeeklyRpt.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport RPT 
      Left            =   525
      Top             =   3870
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Last Week"
      Height          =   1365
      Left            =   2115
      TabIndex        =   27
      Top             =   345
      Width           =   2100
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   990
         TabIndex        =   30
         Top             =   600
         Width           =   1005
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   990
         TabIndex        =   29
         Top             =   930
         Width           =   990
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   990
         TabIndex        =   28
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label Label9 
         Caption         =   "Bulan :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   33
         Top             =   645
         Width           =   600
      End
      Begin VB.Label Label8 
         Caption         =   "Tahun :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   420
         TabIndex        =   32
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label7 
         Caption         =   "Minggu Ke :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   31
         Top             =   315
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Week"
      Height          =   1365
      Left            =   30
      TabIndex        =   20
      Top             =   330
      Width           =   2040
      Begin VB.ComboBox CmbmingguKe 
         Height          =   315
         Left            =   990
         TabIndex        =   23
         Top             =   285
         Width           =   990
      End
      Begin VB.ComboBox CmbTahun 
         Height          =   315
         Left            =   1005
         TabIndex        =   22
         Top             =   930
         Width           =   960
      End
      Begin VB.ComboBox CmbBulan 
         Height          =   315
         Left            =   990
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Minggu Ke :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   26
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Tahun :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   420
         TabIndex        =   25
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Bulan :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   24
         Top             =   645
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1380
      Left            =   4230
      TabIndex        =   7
      Top             =   345
      Width           =   4815
      Begin VB.TextBox TxtWeekSubmission 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1530
         TabIndex        =   13
         Top             =   255
         Width           =   765
      End
      Begin VB.TextBox TxtWeekSubmission 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   4020
         TabIndex        =   12
         Top             =   255
         Width           =   690
      End
      Begin VB.TextBox TxtWeekSubmission 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   1530
         TabIndex        =   11
         Top             =   570
         Width           =   765
      End
      Begin VB.TextBox TxtWeekSubmission 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   4020
         TabIndex        =   10
         Top             =   570
         Width           =   690
      End
      Begin VB.TextBox TxtWeekSubmission 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   4020
         TabIndex        =   9
         Top             =   885
         Width           =   690
      End
      Begin VB.TextBox TxtWeekSubmission 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   1530
         TabIndex        =   8
         Top             =   885
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Tsa # :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   285
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Tsa Actual :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2460
         TabIndex        =   18
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Days # :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   17
         Top             =   600
         Width           =   1410
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Days Actual :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2460
         TabIndex        =   16
         Top             =   585
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Ábsent Actual :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   2340
         TabIndex        =   15
         Top             =   900
         Width           =   1650
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Inc # :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   165
         TabIndex        =   14
         Top             =   915
         Width           =   1290
      End
   End
   Begin RichTextLib.RichTextBox TxtObstacle 
      Height          =   780
      Left            =   915
      TabIndex        =   6
      Top             =   1725
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   1376
      _Version        =   393217
      TextRTF         =   $"FrmWeeklyRpt.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   360
      Index           =   1
      Left            =   8040
      TabIndex        =   4
      Top             =   3855
      Width           =   765
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   360
      Index           =   0
      Left            =   7230
      TabIndex        =   3
      Top             =   3855
      Width           =   765
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   0
      Left            =   3765
      TabIndex        =   1
      Top             =   30
      Width           =   2085
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pro&ses"
      Height          =   360
      Index           =   2
      Left            =   6405
      TabIndex        =   0
      Top             =   3855
      Width           =   765
   End
   Begin RichTextLib.RichTextBox TxtAction 
      Height          =   780
      Left            =   915
      TabIndex        =   34
      Top             =   2505
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   1376
      _Version        =   393217
      TextRTF         =   $"FrmWeeklyRpt.frx":0087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox TxtComment 
      Height          =   330
      Left            =   915
      TabIndex        =   36
      Top             =   3285
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   582
      _Version        =   393217
      TextRTF         =   $"FrmWeeklyRpt.frx":0102
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Comment :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   37
      Top             =   3315
      Width           =   915
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Action :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   35
      Top             =   2535
      Width           =   915
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Obstacle :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   5
      Top             =   1755
      Width           =   915
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Supervisor :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2820
      TabIndex        =   2
      Top             =   45
      Width           =   825
   End
End
Attribute VB_Name = "FrmWeeklyRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub Command1_Click(Index As Integer)
Dim m_objrs As ADODB.Recordset
Dim m_JmlIncAct As String
Dim lJmlTsaActLstW As String
Dim lJmlDaysActLstW As String
Dim lJmlAbsentActLstW As String
Dim lJmlIncActLstW As String
Dim lAvgPrdTeamActLstW As String
Dim lAvgPrdTSAActLstW As String
Dim cmdsql As String
Dim minggulst As String
Dim bulanlst As String
Dim tahunlst As String
Dim LHighPerfTsa As String
Dim LLowPerfTsa As String
Dim lIncWeek1 As String
Dim lIncWeek2 As String
Dim lIncWeek3 As String
Dim lIncWeek4 As String
Dim lIncWeek5 As String
Dim lperiode As String
Dim lbulan As String
Dim m_msgbox As Variant
    Select Case Index
        Case 0
        RPT.Formulas(1) = "@Minggu = totext('" + CStr(CmbmingguKe.Text) + "')"
        RPT.Formulas(2) = "@Bulan = totext('" + CmbBulan.Text + "')"
        RPT.Formulas(3) = "@Tahun = totext('" + CmbTahun.Text + "')"
        RPT.Formulas(4) = "@Supervisor = totext('" + CStr(Combo3(0).Text) + "')"
        
        RPT.Formulas(5) = "@NamaSpv = totext('" + CStr(MDIForm1.Text7.Text) + "')"
        Select Case CmbBulan
            Case "1"
                lbulan = "January "
            Case "2"
                lbulan = "February "
            Case "3"
                lbulan = "March "
            Case "4"
                lbulan = "April "
            Case "5"
                lbulan = "May "
            Case "6"
                lbulan = "June "
            Case "7"
                lbulan = "July "
            Case "8"
                lbulan = "August "
            Case "9"
                lbulan = "September "
            Case "10"
                lbulan = "October "
            Case "11"
                lbulan = "November "
            Case "12"
                lbulan = "December "
        End Select
        
        lperiode = "Week = " & CmbmingguKe.Text & " , " & lbulan & CmbTahun.Text
        RPT.Formulas(6) = "@Periode = totext('" + lperiode + "')"
        
        RPT.ReportFileName = App.Path + "\Report\WeeklyPerfRpt.rpt"
        Call SHOW_PRN
        Case 1
            Unload Me
        Case 2
            'insert nilai ke table rekapsubmissionweek
            Set m_objrs = New ADODB.Recordset
            m_objrs.CursorLocation = adUseClient
            m_objrs.Open "Select * from RekapSubmissionWeekly where Minggu = " + CmbmingguKe.Text + " and Bulan =" + CmbBulan.Text + " and Tahun =" + CmbTahun.Text + " and Spv ='" + Combo3(0).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs.RecordCount = 0 Then
                m_objrs.AddNew
                m_objrs!TEAM = Combo3(0).Text
                m_objrs!SPV = Combo3(0).Text
                m_objrs!Minggu = CmbmingguKe.Text
                m_objrs!Bulan = CmbBulan.Text
                m_objrs!tahun = CmbTahun.Text
                m_objrs!JmlTsaTarget = TxtWeekSubmission(0).Text
                m_objrs!JmlTsaAct = TxtWeekSubmission(1).Text
                m_objrs!JmlDaysTarget = TxtWeekSubmission(2).Text
                m_objrs!JmlDaysAct = TxtWeekSubmission(3).Text
                m_objrs!JmlAbsentAct = TxtWeekSubmission(4).Text
                m_objrs!JmlIncTarget = TxtWeekSubmission(5).Text
                m_objrs!ObStacle = TxtObstacle.Text
                m_objrs!Action1 = TxtAction.Text
                m_objrs!Comment1 = TxtComment.Text
                m_objrs.UPDATE
            Else
                m_msgbox = MsgBox("Data sudah pernah di proses!!.. Ingin Di Proses Ulang??", vbYesNo + vbQuestion, "Telegrandi")
                If m_msgbox = vbYes Then
                    M_OBJCONN.Execute "Delete from RekapSubmissionWeekly where Minggu = " + CmbmingguKe.Text + " and Bulan =" + CmbBulan.Text + " and Tahun =" + CmbTahun.Text + " and Spv ='" + Combo3(0).Text + "'"
                    m_objrs.Requery
                     m_objrs.AddNew
                    m_objrs!TEAM = Combo3(0).Text
                    m_objrs!SPV = Combo3(0).Text
                    m_objrs!Minggu = CmbmingguKe.Text
                    m_objrs!Bulan = CmbBulan.Text
                    m_objrs!tahun = CmbTahun.Text
                    m_objrs!JmlTsaTarget = TxtWeekSubmission(0).Text
                    m_objrs!JmlTsaAct = TxtWeekSubmission(1).Text
                    m_objrs!JmlDaysTarget = TxtWeekSubmission(2).Text
                    m_objrs!JmlDaysAct = TxtWeekSubmission(3).Text
                    m_objrs!JmlAbsentAct = TxtWeekSubmission(4).Text
                    m_objrs!JmlIncTarget = TxtWeekSubmission(5).Text
                    m_objrs!ObStacle = TxtObstacle.Text
                    m_objrs!Action1 = TxtAction.Text
                    m_objrs!Comment1 = TxtComment.Text
                    m_objrs.UPDATE
                Else
                    Exit Sub
                End If
            End If
            Set m_objrs = Nothing
            'cari aplikasi incoming
            Set m_objrs = New ADODB.Recordset
            m_objrs.CursorLocation = adUseClient
            m_objrs.Open "Select spvcode, sum(jumlah) as jumlah from rekapsubmission where minggu = " + CmbmingguKe.Text + " and bulan =" + CmbBulan.Text + " and tahun =" + CmbTahun.Text + "and spvcode ='" + Combo3(0).Text + "' group by spvcode", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs.RecordCount <> 0 Then
                m_JmlIncAct = CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
            Else
                m_JmlIncAct = "0"
               '' M_OBJCONN.Execute "Update RekapSubmissionWeekly set JmlIncAct = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + " where minggu = " + CmbmingguKe.Text + " and bulan =" + CmbBulan.Text + " and tahun =" + CmbTahun.Text + "and spv ='" + Combo3(0).Text + "'"
            End If
            Set m_objrs = Nothing
            'cari nilai last week
            Set m_objrs = New ADODB.Recordset
            m_objrs.CursorLocation = adUseClient
            m_objrs.Open "Select * from RekapSubmissionWeekly where minggu = " + Combo1.Text + " and bulan =" + Combo4.Text + " and tahun =" + Combo2.Text + "and spv ='" + Combo3(0).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs.RecordCount <> 0 Then
                lJmlTsaActLstW = IIf(IsNull(m_objrs!JmlTsaAct), 0, m_objrs!JmlTsaAct)
                lJmlDaysActLstW = IIf(IsNull(m_objrs!JmlDaysAct), 0, m_objrs!JmlDaysAct)
                lJmlAbsentActLstW = IIf(IsNull(m_objrs!JmlAbsentAct), 0, m_objrs!JmlAbsentAct)
                lJmlIncActLstW = IIf(IsNull(m_objrs!JmlIncAct), 0, m_objrs!JmlIncAct)
                lAvgPrdTeamActLstW = IIf(IsNull(m_objrs!AvgPrdTeamAct), 0, m_objrs!AvgPrdTeamAct)
                lAvgPrdTSAActLstW = IIf(IsNull(m_objrs!AvgPrdTSAAct), 0, m_objrs!AvgPrdTSAAct)
            Else
                lJmlTsaActLstW = 0
                lJmlDaysActLstW = 0
                lJmlAbsentActLstW = 0
                lJmlIncActLstW = 0
                lAvgPrdTeamActLstW = 0
                lAvgPrdTSAActLstW = 0
            End If
            Set m_objrs = Nothing
            
            'find the hightest inc
            Set m_objrs = New ADODB.Recordset
            m_objrs.CursorLocation = adUseClient
            cmdsql = " SELECT VwCariMaxMinInc.*, UserTbl.Agent FROM VwCariMaxMinInc,UserTbl WHERE VwCariMaxMinInc.Userid = UserTbl.UserId and JUMLAH IN (select  max (jumlah) from VwCariMaxMinInc"
            cmdsql = cmdsql + " Where SPVCODE ='" + Combo3(0).Text + "' AND MINGGU =" + CmbmingguKe.Text + " AND BULAN ='" + CmbBulan.Text + "' AND TAHUN ='" + CmbTahun.Text + "')"
            cmdsql = cmdsql + " AND VwCariMaxMinInc.SPVCODE ='" + Combo3(0).Text + "' AND MINGGU =" + CmbmingguKe.Text + " AND BULAN ='" + CmbBulan.Text + "' AND TAHUN ='" + CmbTahun.Text + "'"
            m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            While Not m_objrs.EOF
                LHighPerfTsa = LHighPerfTsa + " , " + m_objrs!agent + " " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + " Inc"
                m_objrs.MoveNext
            Wend
            Set m_objrs = Nothing
            
            'find the low inc
            Set m_objrs = New ADODB.Recordset
            m_objrs.CursorLocation = adUseClient
            cmdsql = " SELECT VwCariMaxMinInc.*, UserTbl.Agent FROM VwCariMaxMinInc,UserTbl WHERE VwCariMaxMinInc.Userid = UserTbl.UserId and JUMLAH IN (select  min (jumlah) from VwCariMaxMinInc"
            cmdsql = cmdsql + " Where SPVCODE ='" + Combo3(0).Text + "' AND MINGGU =" + CmbmingguKe.Text + " AND BULAN ='" + CmbBulan.Text + "' AND TAHUN ='" + CmbTahun.Text + "')"
            cmdsql = cmdsql + " AND VwCariMaxMinInc.SPVCODE ='" + Combo3(0).Text + "' AND MINGGU =" + CmbmingguKe.Text + " AND BULAN ='" + CmbBulan.Text + "' AND TAHUN ='" + CmbTahun.Text + "'"
            m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            While Not m_objrs.EOF
                LLowPerfTsa = LLowPerfTsa + " , " + m_objrs!agent + " " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + " Inc"
                m_objrs.MoveNext
            Wend
            Set m_objrs = Nothing
            cmdsql = "UPdate RekapSubmissionWeekly set"
            cmdsql = cmdsql + " JmlTsaActLstW= " + lJmlTsaActLstW + ", "
            cmdsql = cmdsql + " JmlDaysActLstW= " + lJmlDaysActLstW + ", "
            cmdsql = cmdsql + " JmlAbsentActLstW= " + lJmlAbsentActLstW + ", "
            cmdsql = cmdsql + " JmlIncActLstW= " + lJmlIncActLstW + ", "
            cmdsql = cmdsql + " AvgPrdTeamActLstW= " + lAvgPrdTeamActLstW + ", "
            cmdsql = cmdsql + " JmlIncAct= " + m_JmlIncAct + ", "
            cmdsql = cmdsql + " HighPerfTsa = '" + LHighPerfTsa + "', "
            cmdsql = cmdsql + " LowPerfTsa = '" + LLowPerfTsa + "' "
            cmdsql = cmdsql + " where SPV ='" + Combo3(0).Text + "' AND MINGGU =" + CmbmingguKe.Text + " AND BULAN =" + CmbBulan.Text + " AND TAHUN =" + CmbTahun.Text + ""
            M_OBJCONN.Execute cmdsql
            'find last week incoming
            lIncWeek1 = "0"
            lIncWeek2 = "0"
            lIncWeek3 = "0"
            lIncWeek4 = "0"
            lIncWeek5 = "0"
            Set m_objrs = New ADODB.Recordset
            m_objrs.CursorLocation = adUseClient
            cmdsql = "select Minggu, JmlIncAct from rekapsubmissionweekly where bulan = " + CmbBulan.Text + " and tahun =" + CmbTahun.Text + " and spv ='" + Combo3(0).Text + "' order by minggu"
            m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            While Not m_objrs.EOF
                Select Case m_objrs!Minggu
                    Case 1
                        lIncWeek1 = CStr(IIf(IsNull(m_objrs!JmlIncAct), 0, m_objrs!JmlIncAct))
                    Case 2
                        lIncWeek2 = CStr(IIf(IsNull(m_objrs!JmlIncAct), 0, m_objrs!JmlIncAct))
                    Case 3
                        lIncWeek3 = CStr(IIf(IsNull(m_objrs!JmlIncAct), 0, m_objrs!JmlIncAct))
                    Case 4
                        lIncWeek4 = CStr(IIf(IsNull(m_objrs!JmlIncAct), 0, m_objrs!JmlIncAct))
                    Case 5
                        lIncWeek5 = CStr(IIf(IsNull(m_objrs!JmlIncAct), 0, m_objrs!JmlIncAct))
                End Select
                m_objrs.MoveNext
            Wend
            Set m_objrs = Nothing
            cmdsql = "UPdate RekapSubmissionWeekly set"
            cmdsql = cmdsql + " IncWeek1 = " + lIncWeek1 + ", "
            cmdsql = cmdsql + " IncWeek2 = " + lIncWeek2 + ", "
            cmdsql = cmdsql + " IncWeek3 = " + lIncWeek3 + ", "
            cmdsql = cmdsql + " IncWeek4 = " + lIncWeek4 + ", "
            cmdsql = cmdsql + " IncWeek5 = " + lIncWeek5 + " "
            cmdsql = cmdsql + " where SPV ='" + Combo3(0).Text + "' AND MINGGU =" + CmbmingguKe.Text + " AND BULAN =" + CmbBulan.Text + " AND TAHUN =" + CmbTahun.Text + ""
            M_OBJCONN.Execute cmdsql
            MsgBox "Done"
    End Select
End Sub

Private Sub Form_Load()
Dim m_spv As ADODB.Recordset
Dim i As Integer
Set m_spv = New ADODB.Recordset
m_spv.CursorLocation = adUseClient
'm_spv.CursorLocation = adUseClient
'm_spv.Open "SELECT * FROM spvtbl ORDER BY SPVCODE", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

m_spv.Open "select SPVTBL.SPVCODE from SPVTBL, USERTBL where SPVTBL.SPVCODE = USERTBL.SPVCODE AND USERTYPE = '6'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

While Not m_spv.EOF
    Combo3(0).AddItem m_spv!SPVCODE
    m_spv.MoveNext
Wend
Set m_spv = Nothing
For i = 1 To 5
    CmbmingguKe.AddItem i
    Combo1.AddItem i
Next i

CmbTahun.AddItem 2005
CmbTahun.AddItem 2006
CmbTahun.AddItem 2007
CmbTahun.AddItem 2008
CmbTahun.AddItem 2009
CmbTahun.AddItem 2010

Combo2.AddItem 2005
Combo2.AddItem 2006
Combo2.AddItem 2007
Combo2.AddItem 2008
Combo2.AddItem 2009
Combo2.AddItem 2010

For i = 1 To 12
    CmbBulan.AddItem i
    Combo4.AddItem i
Next i
TxtWeekSubmission(0).Text = 0
TxtWeekSubmission(1).Text = 0
TxtWeekSubmission(2).Text = 0
TxtWeekSubmission(3).Text = 0
TxtWeekSubmission(4).Text = 0
TxtWeekSubmission(5).Text = 0
End Sub
