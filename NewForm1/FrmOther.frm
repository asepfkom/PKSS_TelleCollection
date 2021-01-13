VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmOther 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu Other...."
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport Rpt 
      Left            =   5640
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3795
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6694
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Surat Pemeberitahuan"
      TabPicture(0)   =   "FrmOther.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TxtNamaDihubungi"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TxtTempat"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TxtTelpDihubungi"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CmdPreviewSuratPemberitahuan"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtNoSurat"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtAgent"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtTL"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtAlamat"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.TextBox TxtAlamat 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   1380
         Width           =   4815
      End
      Begin VB.TextBox TxtTL 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   13
         Top             =   1080
         Width           =   2115
      End
      Begin VB.TextBox TxtAgent 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   12
         Top             =   720
         Width           =   2115
      End
      Begin VB.TextBox TxtNoSurat 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Top             =   420
         Width           =   2115
      End
      Begin VB.CommandButton CmdPreviewSuratPemberitahuan 
         Caption         =   "&Tampilkan surat"
         Height          =   435
         Left            =   3000
         TabIndex        =   7
         Top             =   3120
         Width           =   1395
      End
      Begin VB.TextBox TxtTelpDihubungi 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Text            =   "(021) 2927 5876"
         Top             =   2700
         Width           =   2115
      End
      Begin VB.TextBox TxtTempat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Text            =   "Plasa Sentral"
         Top             =   2340
         Width           =   2115
      End
      Begin VB.TextBox TxtNamaDihubungi 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   1980
         Width           =   2115
      End
      Begin VB.Label Label7 
         Caption         =   "Alamat:"
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "TL:"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Agent:"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   780
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Nomor Surat:"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Telp.dihubungi:"
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   2700
         Width           =   1875
      End
      Begin VB.Label Label2 
         Caption         =   "Tempat:"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   2340
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Nama yang dihubungi:"
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   2040
         Width           =   1875
      End
   End
End
Attribute VB_Name = "FrmOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdPreviewSuratPemberitahuan_Click()
    Dim CMDSQL As String
    Dim JenisKartu As String
    
    
    
    With FrmCC_Colection
        
        If Trim(UCase(.lbltype.Caption)) = "PIL" Or Trim(UCase(.lbltype.Caption)) = "GRF" Then
            JenisKartu = "Personal Instalment Loan"
        Else
            JenisKartu = "Kartu Kredit"
        End If
        
        'Catet dalam Log
        CMDSQL = " insert into log_surat_pemberitahuan (tglsurat,no_surat,custid,"
        CMDSQL = CMDSQL + "nama_ch,alamat_ch,curbal,agent,tl,cetak_by) values ('"
        CMDSQL = CMDSQL + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "','"
        CMDSQL = CMDSQL + TxtNoSurat.text + "','"
        CMDSQL = CMDSQL + .lblCustId.text + "','"
        CMDSQL = CMDSQL + .lblNama.text + "','"
        CMDSQL = CMDSQL + Trim(TxtAlamat.text) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.TDB_cur_bal.Value), "0", .TDB_cur_bal.Value)) + "','"
        CMDSQL = CMDSQL + TxtAgent.text + "','"
        CMDSQL = CMDSQL + TxtTL.text + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "')"
        M_OBJCONN.Execute CMDSQL
    
        WaitSecs (2)
        RPT.Reset
        RPT.Formulas(1) = "@nama = totext('" + CStr(Trim(.lblNama.text)) + "')"
        RPT.Formulas(2) = "@alamat = totext('" + CStr(Trim(TxtAlamat.text)) + "')"
        RPT.Formulas(3) = "@custid = totext('" + CStr(Trim(.lblCustId.text)) + "')"
        RPT.Formulas(4) = "@curbal = totext('" + CStr("Rp. " & IIf(IsNull(.TDB_cur_bal.text), "0", .TDB_cur_bal.text)) + "')"
        RPT.Formulas(5) = "@hubungi = totext('" + CStr(UCase(Trim(TxtNamaDihubungi.text))) + "')"
        RPT.Formulas(6) = "@tempat = totext('" + CStr(Trim(TxtTempat.text)) + "')"
        RPT.Formulas(7) = "@notelpon = totext('" + CStr(Trim(TxtTelpDihubungi.text)) + "')"
        RPT.Formulas(8) = "@no_surat = totext('" + CStr(Trim(TxtNoSurat.text)) + "')"
        RPT.Formulas(9) = "@jenis = totext('" + CStr(Trim(JenisKartu)) + "')"
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptSuratPenawaran.rpt"
        Call SHOW_PRN
    End With
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
    'RPT.Action = 1
    'RPT.Reset
End Sub


Private Sub NoSuratPemberitahuan()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim NoUrut As String
    Dim TglSurat As String
    Dim DigitCurbal As String
    
    TglSurat = Format(MDIForm1.TDBDate1.Value, "ddmmyy")
    DigitCurbal = CStr(Right(IIf(IsNull(FrmCC_Colection.TDB_cur_bal.Value), "0", FrmCC_Colection.TDB_cur_bal.Value), 4))
    TxtNamaDihubungi.text = MDIForm1.TxtUsername.text
    TxtAgent.text = FrmCC_Colection.lblaoc.Caption
    
    'Cari nama TL
    CMDSQL = "select * from usertbl where userid='"
    CMDSQL = CMDSQL + Trim(FrmCC_Colection.lblaoc.Caption) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        TxtTL.text = IIf(IsNull(M_objrs("team")), "", M_objrs("team"))
    End If
    
    Set M_objrs = Nothing
    
    'Cari no_urut surat
    CMDSQL = "select * from log_surat_pemberitahuan where date_part('month',now())='"
    CMDSQL = CMDSQL + Format(MDIForm1.TDBDate1.Value, "m") + "' and date_part('year',now())='"
    CMDSQL = CMDSQL + Format(MDIForm1.TDBDate1.Value, "yyyy") + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        NoUrut = "1"
    Else
        NoUrut = CStr(M_objrs.RecordCount + 1)
    End If
    
    Set M_objrs = Nothing
    
    
    TxtNoSurat.text = TglSurat & " " & TxtAgent.text & "/" & TxtTL.text & "/" & DigitCurbal & "-" & NoUrut
    
End Sub
    
    


Private Sub Form_Load()
    Call NoSuratPemberitahuan
    TxtAlamat.text = FrmCC_Colection.lblAddr.text
End Sub
