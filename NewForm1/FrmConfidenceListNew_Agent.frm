VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmConfidenceListNew_Agent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confidence Analisys Agent"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   17985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerConfidenceAnalisys 
      Interval        =   1500
      Left            =   4740
      Top             =   60
   End
   Begin VB.TextBox TxtPaymentAnalisis 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   16980
      TabIndex        =   25
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox TxtListPaymentDetail 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   16980
      TabIndex        =   24
      Text            =   "0"
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton CmdSavePtpValid 
      Caption         =   "&Save PTP Valid"
      Enabled         =   0   'False
      Height          =   495
      Left            =   15840
      TabIndex        =   23
      Top             =   7680
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   15840
      TabIndex        =   22
      Top             =   8220
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CommandButton CmdCekAutoPtpValid 
      Caption         =   "&Cek PTP Valid By Program"
      Enabled         =   0   'False
      Height          =   495
      Left            =   15840
      TabIndex        =   21
      Top             =   7140
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CommandButton CmdLoadData 
      Caption         =   "&Load Data"
      Height          =   495
      Left            =   15840
      TabIndex        =   20
      Top             =   6600
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.TextBox TxtAllPayment 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   16980
      TabIndex        =   18
      Text            =   "0"
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Summary"
      Height          =   2055
      Left            =   0
      TabIndex        =   6
      Top             =   6600
      Width           =   15675
      Begin VB.CheckBox Check2 
         BackColor       =   &H0000FF00&
         Caption         =   "Ya, saya sudah siap"
         Height          =   315
         Left            =   7680
         TabIndex        =   41
         Top             =   1380
         Width           =   7515
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000FF00&
         Caption         =   "Ya, saya sudah baca dan paham"
         Height          =   315
         Left            =   7680
         TabIndex        =   39
         Top             =   660
         Width           =   7515
      End
      Begin VB.TextBox TxtCountPtpValid 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1860
         TabIndex        =   13
         Text            =   "0"
         Top             =   1500
         Width           =   1695
      End
      Begin VB.TextBox TxtTotalPtpValid 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1860
         TabIndex        =   12
         Text            =   "0"
         Top             =   1140
         Width           =   1695
      End
      Begin VB.TextBox TxtTotalPayment 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1860
         TabIndex        =   11
         Text            =   "0"
         Top             =   780
         Width           =   1695
      End
      Begin VB.TextBox TxtTotalPtp 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1860
         TabIndex        =   10
         Text            =   "0"
         Top             =   420
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H000EC1CB&
         Caption         =   "Your Confidence Analisis"
         Height          =   1575
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   3015
         Begin VB.Label LblCA 
            Alignment       =   2  'Center
            BackColor       =   &H00004040&
            Caption         =   "0"
            ForeColor       =   &H0000FFFF&
            Height          =   435
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   2715
         End
         Begin VB.Label Label13 
            Caption         =   "(Total Payment+Total PTP Valid)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   780
            Width           =   2715
         End
      End
      Begin VB.Label Label15 
         Caption         =   "Sudahkah siapkah  anda untuk hari ini dan esok?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   7080
         TabIndex        =   40
         Top             =   1020
         Width           =   8355
      End
      Begin VB.Label Label14 
         Caption         =   "Sudahkah anda baca dan paham confidence analisys anda hari ini?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   7140
         TabIndex        =   38
         Top             =   300
         Width           =   8355
      End
      Begin VB.Label Label10 
         BackColor       =   &H000080FF&
         Caption         =   "Count PTP Valid:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   17
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H000080FF&
         Caption         =   "Total PTP Valid:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   16
         Top             =   1140
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H000080FF&
         Caption         =   "Total Payment:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   15
         Top             =   780
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H000080FF&
         Caption         =   "Total PTP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   14
         Top             =   420
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report"
      Enabled         =   0   'False
      Height          =   1995
      Left            =   6960
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   3135
      Begin VB.OptionButton OptSemua 
         Caption         =   "Semua data di list"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton OptValid 
         Caption         =   "Hanya PTP Valid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2775
      End
      Begin VB.OptionButton OptNotValid 
         Caption         =   "Not PTP Valid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   900
         Width           =   2775
      End
      Begin VB.CommandButton CmdPreview 
         Caption         =   "&Preview"
         Height          =   435
         Left            =   300
         TabIndex        =   2
         Top             =   1200
         Width           =   2595
      End
      Begin MSComctlLib.ProgressBar Pb2 
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   1680
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   435
      Left            =   0
      TabIndex        =   19
      Top             =   8700
      Width           =   17955
      _ExtentX        =   31671
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ListView LvPTPPayment 
      Height          =   5820
      Left            =   0
      TabIndex        =   26
      Top             =   420
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   10266
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   1
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
   Begin MSComctlLib.ListView LvPayment 
      Height          =   1260
      Left            =   11640
      TabIndex        =   27
      Top             =   420
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   2223
      View            =   3
      LabelEdit       =   1
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
   Begin MSComctlLib.ListView LvPaymentDetail 
      Height          =   1740
      Left            =   11640
      TabIndex        =   28
      Top             =   2040
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   3069
      View            =   3
      LabelEdit       =   1
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
   Begin MSComctlLib.ListView LvPaymentAll 
      Height          =   2460
      Left            =   11640
      TabIndex        =   29
      Top             =   4140
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   4339
      View            =   3
      LabelEdit       =   1
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
   Begin VB.Label Label1 
      Caption         =   "List PTP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   37
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Payment account PTP terpilih bulan ini"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11700
      TabIndex        =   36
      Top             =   120
      Width           =   4395
   End
   Begin VB.Label Label4 
      Caption         =   "*)PTP yang dicentang merupakan PTP yang Valid hasil analisa TL anda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   35
      Top             =   6240
      Width           =   11535
   End
   Begin VB.Label Label6 
      Caption         =   "Payment detail account PTP terpilih"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11700
      TabIndex        =   34
      Top             =   1740
      Width           =   3315
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   16200
      TabIndex        =   33
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   16140
      TabIndex        =   32
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Semua payment dari agent terpilih"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   31
      Top             =   3840
      Width           =   3315
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000FF00&
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   16140
      TabIndex        =   30
      Top             =   3840
      Width           =   855
   End
End
Attribute VB_Name = "FrmConfidenceListNew_Agent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub IsiHeader()
    LvPTPPayment.ColumnHeaders.ADD 1, , "No.", 800
    LvPTPPayment.ColumnHeaders.ADD 2, , "Agent", 700
    LvPTPPayment.ColumnHeaders.ADD 3, , "Nama CH", 2500
    LvPTPPayment.ColumnHeaders.ADD 4, , "Custid", 2000
    LvPTPPayment.ColumnHeaders.ADD 5, , "Tgl.PTP", 1300
    LvPTPPayment.ColumnHeaders.ADD 6, , "Amount PTP", 1500
    LvPTPPayment.ColumnHeaders.ADD 7, , "LPD", 1500
    LvPTPPayment.ColumnHeaders.ADD 8, , "LPA", 1500
    LvPTPPayment.ColumnHeaders.ADD 9, , "Id", 700
    LvPTPPayment.ColumnHeaders.ADD 10, , "Valid", 700
    LvPTPPayment.ColumnHeaders.ADD 11, , "CekAuto", 700
    
End Sub
Private Sub HeaderPayment()
    LvPayment.ColumnHeaders.ADD 1, , "Agent", 1000
    LvPayment.ColumnHeaders.ADD 2, , "Custid", 2000
    LvPayment.ColumnHeaders.ADD 3, , "Tgl.Payment", 2000
    LvPayment.ColumnHeaders.ADD 4, , "Payment", 20000
End Sub

Private Sub HeaderPaymentDetail()
    LvPaymentDetail.ColumnHeaders.ADD 1, , "Agent", 1000
    LvPaymentDetail.ColumnHeaders.ADD 2, , "Custid", 2000
    LvPaymentDetail.ColumnHeaders.ADD 3, , "Tgl.Payment", 2000
    LvPaymentDetail.ColumnHeaders.ADD 4, , "Payment", 20000
End Sub

Private Sub HeaderPaymentAll()
    LvPaymentAll.ColumnHeaders.ADD 1, , "Agent", 1000
    LvPaymentAll.ColumnHeaders.ADD 2, , "Custid", 2000
    LvPaymentAll.ColumnHeaders.ADD 3, , "Tgl.Payment", 2000
    LvPaymentAll.ColumnHeaders.ADD 4, , "Payment", 20000
End Sub



Private Sub Check1_Click()
    If Check1.Value = vbChecked And Check2.Value = vbChecked Then
        MsgBox " Terima Kasih! Mulakan dengan doa dan niat yang baik,semangat, jadikan hari ini lebih baik dari kemarin, dan sukses :)", vbOKOnly, "Welcome"
        Unload Me
    End If
End Sub

Private Sub Check2_Click()
    If Check1.Value = vbChecked And Check2.Value = vbChecked Then
        MsgBox " Terima Kasih! Mulakan dengan doa dan niat yang baik,semangat, jadikan hari ini lebih baik dari kemarin, dan sukses :)", vbOKOnly, "Welcome"
        Unload Me
    End If
End Sub

Private Sub CmdCekAutoPtpValid_Click()
    Dim a As String
    Dim W As Integer
    
    a = MsgBox("Anda yakin akan menandakan, semua account PTP bertanda merah sebagai PTP Valid?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbYes Then
        For W = 1 To LvPTPPayment.ListItems.Count
            If LvPTPPayment.ListItems(W).SubItems(10) = "1" Then
                LvPTPPayment.ListItems(W).Checked = True
            End If
        Next W
    End If
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdLoadData_Click()
    Call IsiPtp
    Call TotalPayment
    Call TotalConfidenceAnalisis
    Call LpdLpa
End Sub

Private Sub CmdPreview_Click()
    
    If LvPTPPayment.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Jika memilih data semua
    If OptSemua.Value Then
        Call IsiPreviewAll
    End If
    'Jika Hanya memilih data yang Valid
    If OptValid.Value Then
        Call IsiPreviewValid
    End If
    'Jika Hanya memilih data yang tidak valid
    If OptNotValid.Value Then
        Call IsiPreviewTidakValid
    End If
    
    Call UpdatePaymentPerAgent
    
    'WaitSecs (2)
   ' RPT.Reset
'            RPT.Formulas(1) = "@User = totext('" + CStr(mdiform1.txtusername.text) + "')"
'            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
'            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
   ' RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptConfidenceAnalisys.rpt"
'    Call SHOW_PRN
    
End Sub

Private Sub UpdatePaymentPerAgent()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    CMDSQL = "select agent,sum(payment) from tbllunas where "
    CMDSQL = CMDSQL + "date_part('month',tbllunas.paydate)=date_part('month',now()) and "
    CMDSQL = CMDSQL + "date_part('year',tbllunas.paydate)=date_part('year',now()) and agent in ("
    'Jika yang login Agent
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        CMDSQL = CMDSQL + "'" + MDIForm1.TxtUsername.Text + "')"
    End If
    'Jika yang login TL
    If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        CMDSQL = CMDSQL + "select userid from usertbl where team='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "')"
    End If
    'Jika yang login SPV/Admin
    If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.Text) = "ADMIN" Or UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Then
        CMDSQL = CMDSQL + "select userid from usertbl)"
    End If
    CMDSQL = CMDSQL + " group by agent order by agent asc "
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    PB2.Max = M_objrs.RecordCount
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            PB2.Value = M_objrs.Bookmark
            CMDSQL = "update tblconfidenceanalisys set total_payment='"
            CMDSQL = CMDSQL + CStr(M_objrs(1)) + "' where agent='"
            CMDSQL = CMDSQL + M_objrs("agent") + "'"
            M_RPTCONN.Execute CMDSQL
            M_objrs.MoveNext
        Wend
    End If
    
    Set M_objrs = Nothing
End Sub


Private Sub IsiPreviewAll()
    Dim CMDSQL As String
    Dim W As Integer
    Dim StatusValid As String
    
    PB2.Max = LvPTPPayment.ListItems.Count
    CMDSQL = "delete from tblconfidenceanalisys"
    M_RPTCONN.Execute CMDSQL
    For W = 1 To LvPTPPayment.ListItems.Count
        PB2.Value = W
        If LvPTPPayment.ListItems(W).SubItems(9) = "1" Then
            StatusValid = "Valid"
        Else
            StatusValid = "Not Valid"
        End If
        CMDSQL = "insert into tblconfidenceanalisys (agent,nama_ch,custid,"
        CMDSQL = CMDSQL + "tgl_ptp,amount_ptp,lpd,lpa,status_valid) values ('"
        CMDSQL = CMDSQL + LvPTPPayment.ListItems(W).SubItems(1) + "','"
        CMDSQL = CMDSQL + LvPTPPayment.ListItems(W).SubItems(2) + "','"
        CMDSQL = CMDSQL + LvPTPPayment.ListItems(W).SubItems(3) + "','"
        CMDSQL = CMDSQL + CStr(Format(LvPTPPayment.ListItems(W).SubItems(4), "yyyy-mm-dd")) + "','"
        CMDSQL = CMDSQL + CStr(IIf(LvPTPPayment.ListItems(W).SubItems(5) = "", "0", Format(LvPTPPayment.ListItems(W).SubItems(5), "############"))) + "',"
        CMDSQL = CMDSQL + IIf(LvPTPPayment.ListItems(W).SubItems(6) = "", "null", "'" + CStr(Format(LvPTPPayment.ListItems(W).SubItems(6), "yyyy-mm-dd")) + "'") + ",'"
        CMDSQL = CMDSQL + CStr(IIf(LvPTPPayment.ListItems(W).SubItems(7) = "", "0", Format(LvPTPPayment.ListItems(W).SubItems(7), "#############"))) + "','"
        CMDSQL = CMDSQL + StatusValid + "')"
        M_RPTCONN.Execute CMDSQL
    Next W
End Sub

Private Sub IsiPreviewValid()
    Dim CMDSQL As String
    Dim W As Integer
    Dim StatusValid As String
    
    PB2.Max = LvPTPPayment.ListItems.Count
    CMDSQL = "delete from tblconfidenceanalisys"
    M_RPTCONN.Execute CMDSQL
    For W = 1 To LvPTPPayment.ListItems.Count
        PB2.Value = W
        If LvPTPPayment.ListItems(W).SubItems(9) = "1" Then
            If LvPTPPayment.ListItems(W).SubItems(9) = "1" Then
                StatusValid = "Valid"
            Else
                StatusValid = "Not Valid"
            End If
            CMDSQL = "insert into tblconfidenceanalisys (agent,nama_ch,custid,"
            CMDSQL = CMDSQL + "tgl_ptp,amount_ptp,lpd,lpa,status_valid) values ('"
            CMDSQL = CMDSQL + LvPTPPayment.ListItems(W).SubItems(1) + "','"
            CMDSQL = CMDSQL + LvPTPPayment.ListItems(W).SubItems(2) + "','"
            CMDSQL = CMDSQL + LvPTPPayment.ListItems(W).SubItems(3) + "','"
            CMDSQL = CMDSQL + CStr(Format(LvPTPPayment.ListItems(W).SubItems(4), "yyyy-mm-dd")) + "','"
            CMDSQL = CMDSQL + CStr(Format(LvPTPPayment.ListItems(W).SubItems(5), "##########")) + "',"
            CMDSQL = CMDSQL + IIf(LvPTPPayment.ListItems(W).SubItems(6) = "", "null", "'" + CStr(Format(LvPTPPayment.ListItems(W).SubItems(6), "yyyy-mm-dd")) + "'") + ",'"
            CMDSQL = CMDSQL + CStr(IIf(LvPTPPayment.ListItems(W).SubItems(7) = "", "0", Format(LvPTPPayment.ListItems(W).SubItems(7), "###########"))) + "','"
            CMDSQL = CMDSQL + StatusValid + "')"
            M_RPTCONN.Execute CMDSQL
        End If
    Next W
End Sub

Private Sub IsiPreviewTidakValid()
    Dim CMDSQL As String
    Dim W As Integer
    Dim StatusValid As String
    
    PB2.Max = LvPTPPayment.ListItems.Count
    CMDSQL = "delete from tblconfidenceanalisys"
    M_RPTCONN.Execute CMDSQL
    For W = 1 To LvPTPPayment.ListItems.Count
        PB2.Value = W
        If LvPTPPayment.ListItems(W).SubItems(9) <> "1" Then
            If LvPTPPayment.ListItems(W).SubItems(9) = "1" Then
                StatusValid = "Valid"
            Else
                StatusValid = "Not Valid"
            End If
            CMDSQL = "insert into tblconfidenceanalisys (agent,nama_ch,custid,"
            CMDSQL = CMDSQL + "tgl_ptp,amount_ptp,lpd,lpa,status_valid) values ('"
            CMDSQL = CMDSQL + LvPTPPayment.ListItems(W).SubItems(1) + "','"
            CMDSQL = CMDSQL + LvPTPPayment.ListItems(W).SubItems(2) + "','"
            CMDSQL = CMDSQL + LvPTPPayment.ListItems(W).SubItems(3) + "','"
            CMDSQL = CMDSQL + CStr(Format(LvPTPPayment.ListItems(W).SubItems(4), "yyyy-mm-dd")) + "','"
            CMDSQL = CMDSQL + CStr(IIf(LvPTPPayment.ListItems(W).SubItems(5) = "", "0", Format(LvPTPPayment.ListItems(W).SubItems(5), "############"))) + "',"
            CMDSQL = CMDSQL + IIf(LvPTPPayment.ListItems(W).SubItems(6) = "", "null", "'" + CStr(Format(LvPTPPayment.ListItems(W).SubItems(6), "yyyy-mm-dd")) + "'") + ",'"
            CMDSQL = CMDSQL + CStr(IIf(LvPTPPayment.ListItems(W).SubItems(7) = "", "0", Format(LvPTPPayment.ListItems(W).SubItems(7), "#############"))) + "','"
            CMDSQL = CMDSQL + StatusValid + "')"
            M_RPTCONN.Execute CMDSQL
        End If
    Next W
End Sub



Private Sub CmdSavePtpValid_Click()
    Dim W As Integer
    Dim CMDSQL As String
    Dim TotalPtpValid As Long
    Dim CountPtpValid As Integer
    Dim a As String
    
    If LvPTPPayment.ListItems.Count = 0 Then
        MsgBox "Data PTP tidak ada!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Yakin data PTP Valid akan disimpan?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        Exit Sub
    End If
    
    TotalPtpValid = 0
    CountPtpValid = 0
    For W = 1 To LvPTPPayment.ListItems.Count
        If LvPTPPayment.ListItems(W).Checked = True Then
            CMDSQL = "update tblnegoptp set f_valid='1' where id='"
            CMDSQL = CMDSQL + LvPTPPayment.ListItems(W).SubItems(8) + "'"
            M_OBJCONN.Execute CMDSQL
            TotalPtpValid = TotalPtpValid + Val(Format(LvPTPPayment.ListItems(W).SubItems(5), "#############"))
            CountPtpValid = CountPtpValid + 1
        Else
            CMDSQL = "update tblnegoptp set f_valid=null where id='"
            CMDSQL = CMDSQL + LvPTPPayment.ListItems(W).SubItems(8) + "'"
            M_OBJCONN.Execute CMDSQL
        End If
    Next W
    TxtTotalPtpValid.Text = Format(TotalPtpValid, "##,###")
    TxtCountPtpValid.Text = Format(CountPtpValid, "##,###")
    Call TotalConfidenceAnalisis
    MsgBox "Data PTP Valid berhasil disimpan!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub Form_Load()
    Call IsiHeader
    Call HeaderPaymentDetail
    Call HeaderPayment
    Call HeaderPaymentAll
End Sub

Private Sub IsiPtp()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    Dim m_objrs_payment As ADODB.Recordset
    Dim cmdsql_ptp As ADODB.Recordset
    Dim TotalPtp As Double
    
    
    CMDSQL = " select m.agent,m.name, m.custid, ptp.promisedate, ptp.promisepay,ptp.id,ptp.f_valid "
    CMDSQL = CMDSQL + " from tblnegoptp as ptp, mgm as m "
    CMDSQL = CMDSQL + " where ptp.custid = m.custid and "
    CMDSQL = CMDSQL + " m.agent in ("
    'Jika yang login Agent
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        CMDSQL = CMDSQL + "'" + MDIForm1.TxtUsername.Text + "')"
    End If
    'Jika yang login TL
    If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        CMDSQL = CMDSQL + "select userid from usertbl where team='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "')"
    End If
    'Jika yang login SPV/Admin
    If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.Text) = "ADMIN" Or UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Then
        CMDSQL = CMDSQL + "select userid from usertbl)"
    End If
    CMDSQL = CMDSQL + " and date_part('month',ptp.promisedate)=date_part('month',now()) and "
    CMDSQL = CMDSQL + " date_part('year',ptp.promisedate)=date_part('year',now()) "
    CMDSQL = CMDSQL + " order by m.agent,ptp.promisedate,m.custid"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvPTPPayment.ListItems.CLEAR
    
    If M_objrs.RecordCount > 0 Then
        no = 0
        Dim TotalPtpValid As Long
        PB1.Max = M_objrs.RecordCount
        While Not M_objrs.EOF
            PB1.Value = M_objrs.Bookmark
            no = no + 1
            Set ListItem = LvPTPPayment.ListItems.ADD(, , no)
                ListItem.SubItems(1) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("name")), "", M_objrs("name"))
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
                ListItem.SubItems(4) = IIf(IsNull(M_objrs("promisedate")), "", Format(M_objrs("promisedate"), "yyyy-mm-dd"))
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("promisepay")), "", Format(M_objrs("promisepay"), "##,###"))
                ListItem.SubItems(8) = IIf(IsNull(M_objrs("id")), "", M_objrs("id"))
                ListItem.SubItems(9) = IIf(IsNull(M_objrs("f_valid")), "", M_objrs("f_valid"))
            
            'Hitung Total PTP Valid
            If M_objrs("f_valid") = "1" Then
                TotalPtpValid = TotalPtpValid + Val(M_objrs("promisepay"))
            End If
            
            CMDSQL = "select * from tbllunas where custid='"
            CMDSQL = CMDSQL + M_objrs("custid") + "' and agent='"
            CMDSQL = CMDSQL + M_objrs("agent") + "' and date_part('month',paydate)=date_part('month',date('"
            CMDSQL = CMDSQL + Format(M_objrs("promisedate"), "yyyy-mm-dd") + "')) and date_part('year',paydate)=date_part('year',date('"
            CMDSQL = CMDSQL + Format(M_objrs("promisedate"), "yyyy-mm-dd") + "')) "
            CMDSQL = CMDSQL + " order by paydate asc"
            
            Set m_objrs_payment = New ADODB.Recordset
            m_objrs_payment.CursorLocation = adUseClient
            m_objrs_payment.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If m_objrs_payment.RecordCount > 0 Then
                ListItem.ForeColor = vbRed
                ListItem.ListSubItems(1).ForeColor = vbRed
                ListItem.ListSubItems(2).ForeColor = vbRed
                ListItem.ListSubItems(3).ForeColor = vbRed
                ListItem.ListSubItems(4).ForeColor = vbRed
                ListItem.ListSubItems(5).ForeColor = vbRed
                ListItem.ListSubItems(6).ForeColor = vbRed
                ListItem.ListSubItems(7).ForeColor = vbRed
                ListItem.ListSubItems(8).ForeColor = vbRed
                ListItem.SubItems(10) = "1"
            End If
            Set m_objrs_payment = Nothing
            
            'Buat centang yang sudah jadi ptpvalid
            If ListItem.SubItems(9) = "1" Then
                ListItem.Checked = True
                CountPtpValid = CountPtpValid + 1
            End If
            
            
            TotalPtp = TotalPtp + Val(IIf(IsNull(M_objrs("promisepay")), 0, M_objrs("promisepay")))
            
            M_objrs.MoveNext
        Wend
    End If
    
    TxtTotalPtp.Text = Format(TotalPtp, "##,###")
    TxtTotalPtpValid.Text = Format(TotalPtpValid, "##,###")
    TxtCountPtpValid.Text = Format(CountPtpValid, "##,###")
   
    Set M_objrs = Nothing
End Sub


Private Sub LvPayment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvPayment.SortKey = ColumnHeader.Index - 1
    LvPayment.Sorted = True
End Sub



Private Sub LvPaymentDetail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvPaymentDetail.SortKey = ColumnHeader.Index - 1
    LvPaymentDetail.Sorted = True
End Sub

Private Sub LvPTPPayment_Click()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    If LvPTPPayment.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    CMDSQL = "select * from tbllunas where custid='"
    CMDSQL = CMDSQL + LvPTPPayment.SelectedItem.SubItems(3) + "' and agent='"
    CMDSQL = CMDSQL + LvPTPPayment.SelectedItem.SubItems(1) + "' and date_part('month',paydate)=date_part('month',date('"
    CMDSQL = CMDSQL + Format(LvPTPPayment.SelectedItem.SubItems(4), "yyyy-mm-dd") + "')) and date_part('year',paydate)=date_part('year',date('"
    CMDSQL = CMDSQL + Format(LvPTPPayment.SelectedItem.SubItems(4), "yyyy-mm-dd") + "')) "
    CMDSQL = CMDSQL + " order by paydate asc"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvPayment.ListItems.CLEAR
    Dim TotalAnalisisPayment As Double
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvPayment.ListItems.ADD(, , IIf(IsNull(M_objrs("agent")), "", M_objrs("agent")))
                ListItem.SubItems(1) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("paydate")), "", Format(M_objrs("paydate"), "yyyy-mm-dd"))
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("payment")), "", Format(M_objrs("payment"), "##,###"))
            TotalAnalisisPayment = TotalAnalisisPayment + Val(IIf(IsNull(M_objrs("payment")), "0", M_objrs("payment")))
            
            M_objrs.MoveNext
        Wend
    End If
    TxtPaymentAnalisis.Text = Format(TotalAnalisisPayment, "##,###")
    Set M_objrs = Nothing
    
    
    Dim TotalDetailPayment As Double
    'Menghitung payment detail
    CMDSQL = "select * from tbllunas where custid='"
    CMDSQL = CMDSQL + LvPTPPayment.SelectedItem.SubItems(3) + "' and agent='"
    CMDSQL = CMDSQL + LvPTPPayment.SelectedItem.SubItems(1) + "' "
    CMDSQL = CMDSQL + " order by paydate desc"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    LvPaymentDetail.ListItems.CLEAR
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvPaymentDetail.ListItems.ADD(, , IIf(IsNull(M_objrs("agent")), "", M_objrs("agent")))
                ListItem.SubItems(1) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("paydate")), "", Format(M_objrs("paydate"), "yyyy-mm-dd"))
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("payment")), "", Format(M_objrs("payment"), "##,###"))
                TotalDetailPayment = TotalDetailPayment + Val(IIf(IsNull(M_objrs("payment")), "0", M_objrs("payment")))
            M_objrs.MoveNext
        Wend
    End If
    TxtListPaymentDetail.Text = Format(TotalDetailPayment, "##,###")
    Set M_objrs = Nothing
    
    'Menghitung All Payment
    Dim TotalAllPayment As Double
    CMDSQL = "select * from tbllunas where agent='"
    CMDSQL = CMDSQL + LvPTPPayment.SelectedItem.SubItems(1) + "' "
    CMDSQL = CMDSQL + " order by paydate desc"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    LvPaymentAll.ListItems.CLEAR
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvPaymentAll.ListItems.ADD(, , IIf(IsNull(M_objrs("agent")), "", M_objrs("agent")))
                ListItem.SubItems(1) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("paydate")), "", Format(M_objrs("paydate"), "yyyy-mm-dd"))
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("payment")), "", Format(M_objrs("payment"), "##,###"))
                TotalAllPayment = TotalAllPayment + Val(IIf(IsNull(M_objrs("payment")), "0", M_objrs("payment")))
            M_objrs.MoveNext
        Wend
    End If
    TxtAllPayment.Text = Format(TotalAllPayment, "##,###")
    Set M_objrs = Nothing
End Sub

Private Sub TotalPayment()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
'    CMDSQL = "select sum(tbllunas.payment) from tbllunas,mgm "
'    CMDSQL = CMDSQL + "where  tbllunas.agent=mgm.agent "
'    CMDSQL = CMDSQL + "and tbllunas.agent in ("
    CMDSQL = "select sum(payment) from tbllunas where agent in ("
    'Jika yang login Agent
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        CMDSQL = CMDSQL + "'" + MDIForm1.TxtUsername.Text + "')"
    End If
    'Jika yang login TL
    If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        CMDSQL = CMDSQL + "select userid from usertbl where team='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "')"
    End If
    'Jika yang login SPV/Admin
    If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.Text) = "ADMIN" Or UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Then
        CMDSQL = CMDSQL + "select userid from usertbl)"
    End If
    
    CMDSQL = CMDSQL + " and date_part('month',tbllunas.paydate)=date_part('month',now()) and "
    CMDSQL = CMDSQL + " date_part('year',tbllunas.paydate)=date_part('year',now())"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    TxtTotalPayment.Text = Format(M_objrs(0), "##,###")
    Set M_objrs = Nothing
End Sub

Private Sub LvPTPPayment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvPTPPayment.SortKey = ColumnHeader.Index - 1
    LvPTPPayment.Sorted = True
End Sub

Private Sub LvPTPPayment_DblClick()
    If LvPTPPayment.ListItems.Count > 0 Then
        VIEW_MGMDATA.txtnocard.Text = LvPTPPayment.SelectedItem.SubItems(3)
        FrmConfidenceListNew_Agent.Hide
        VIEW_MGMDATA.Show
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If
End Sub

Private Sub LvPTPPayment_KeyUp(KeyCode As Integer, Shift As Integer)
    LvPTPPayment_Click
End Sub

Private Sub TotalConfidenceAnalisis()
    Dim CA As Double
    
    CA = Val(Format(TxtTotalPayment.Text, "############")) + Val(Format(TxtTotalPtpValid.Text, "############"))
    LblCA.Caption = Format(CA, "##,###")
End Sub

Private Sub LpdLpa()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim W As Integer
        
    If LvPTPPayment.ListItems.Count = 0 Then
        Exit Sub
    End If

    PB1.Max = LvPTPPayment.ListItems.Count
    For W = 1 To LvPTPPayment.ListItems.Count
        PB1.Value = W
        CMDSQL = "select * from tbllunas where custid='"
        CMDSQL = CMDSQL + Trim(LvPTPPayment.ListItems(W).SubItems(3)) + "' "
        CMDSQL = CMDSQL + " order by paydate desc limit 1"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_objrs.RecordCount > 0 Then
            LvPTPPayment.ListItems(W).SubItems(6) = Format(M_objrs("paydate"), "yyyy-mm-dd")
            LvPTPPayment.ListItems(W).SubItems(7) = Format(M_objrs("payment"), "##,###")
        End If
        Set M_objrs = Nothing
    Next W
    
End Sub

Private Sub SHOW_PRN()
'    RPT.RetrieveDataFiles
'    RPT.WindowLeft = 0
'    RPT.WindowTop = 0
'    RPT.WindowState = crptMaximized
'    RPT.WindowShowPrintBtn = True
'    RPT.WindowShowRefreshBtn = True
'    RPT.WindowShowSearchBtn = True
'    RPT.WindowShowPrintSetupBtn = True
'    RPT.WindowControls = True
'    RPT.PrintReport
'    'RPT.Action = 1
'    'RPT.Reset
End Sub

Private Sub TimerConfidenceAnalisys_Timer()
    CmdLoadData_Click
    TimerConfidenceAnalisys.Enabled = False
End Sub
