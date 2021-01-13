VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmAccLunas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Account Lunas"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12675
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      Caption         =   "Check All Batal"
      Height          =   195
      Left            =   4560
      TabIndex        =   38
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check All Belum Lunas"
      Height          =   195
      Left            =   2160
      TabIndex        =   37
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check All Lunas"
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   7560
      Width           =   1935
   End
   Begin Crystal.CrystalReport RPT 
      Left            =   300
      Top             =   5940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   435
      Left            =   10920
      TabIndex        =   21
      Top             =   8340
      Width           =   1575
   End
   Begin VB.Frame FrameFilter 
      Caption         =   "Filter Data Lunas"
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   12435
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   10680
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Remove"
         Height          =   375
         Left            =   8760
         TabIndex        =   45
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search cust"
         Height          =   375
         Left            =   7500
         TabIndex        =   44
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   7500
         TabIndex        =   43
         Top             =   720
         Width           =   2535
      End
      Begin VB.ComboBox CmbStatus 
         Height          =   315
         ItemData        =   "FrmAccLunas.frx":0000
         Left            =   1800
         List            =   "FrmAccLunas.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   840
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   29
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   319750147
         CurrentDate     =   41444
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10980
         TabIndex        =   13
         Top             =   780
         Width           =   1335
      End
      Begin VB.CommandButton CmdFilter 
         Caption         =   "&Filter"
         Height          =   375
         Left            =   10980
         TabIndex        =   12
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox TxtCustid 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7500
         TabIndex        =   11
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox CmbJenisPTP 
         Height          =   315
         Left            =   4380
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox CmbTipeKartu 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   1
         Left            =   4380
         TabIndex        =   32
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   110952451
         CurrentDate     =   41444
      End
      Begin VB.Label Label4 
         Caption         =   "Status "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Sampai"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3600
         TabIndex        =   31
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Tanggal Pelunasan"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Custid:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6240
         TabIndex        =   10
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Jenis PTP:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3540
         TabIndex        =   8
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Tipe Kartu:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   7980
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame FrameList 
      Caption         =   "List Data Lunas"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   12435
      Begin VB.CommandButton Command2 
         Caption         =   "Export Coints"
         Height          =   555
         Left            =   10920
         TabIndex        =   42
         Top             =   3960
         Width           =   1395
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   3720
         TabIndex        =   34
         Top             =   1320
         Visible         =   0   'False
         Width           =   3975
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "CALCULATION . . ."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   720
            TabIndex        =   35
            Top             =   600
            Width           =   2415
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Export To Excel"
         Height          =   555
         Left            =   10920
         TabIndex        =   33
         Top             =   3360
         Width           =   1395
      End
      Begin VB.CommandButton CmdViewCPAElektornik 
         Caption         =   "View CPA Elektronik..."
         Height          =   555
         Left            =   10920
         TabIndex        =   28
         Top             =   2760
         Width           =   1395
      End
      Begin VB.CommandButton CmdViewReport 
         Caption         =   "View Report..."
         Height          =   555
         Left            =   10920
         TabIndex        =   27
         Top             =   2160
         Width           =   1395
      End
      Begin VB.CommandButton CmdPindahLunas 
         Caption         =   "Pindah ke coding lunas...."
         Height          =   555
         Left            =   10920
         TabIndex        =   26
         Top             =   1560
         Width           =   1395
      End
      Begin VB.CommandButton CmdUnCekAll 
         Caption         =   "UnCek All"
         Height          =   375
         Left            =   10920
         TabIndex        =   25
         Top             =   1140
         Width           =   1395
      End
      Begin VB.CommandButton CmdCekAll 
         Caption         =   "Cek All"
         Height          =   375
         Left            =   10920
         TabIndex        =   24
         Top             =   720
         Width           =   1395
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   10920
         TabIndex        =   14
         Top             =   300
         Width           =   1395
      End
      Begin VB.TextBox TxtJmlData 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1260
         TabIndex        =   3
         Text            =   "0"
         Top             =   4680
         Width           =   1035
      End
      Begin MSComctlLib.ListView LvAccLunas 
         Height          =   4335
         Left            =   60
         TabIndex        =   1
         Top             =   300
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin TDBNumber6Ctl.TDBNumber TxtBalanceMMU 
         Height          =   255
         Left            =   4200
         TabIndex        =   16
         Top             =   4680
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   450
         Calculator      =   "FrmAccLunas.frx":0004
         Caption         =   "FrmAccLunas.frx":0024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmAccLunas.frx":0090
         Keys            =   "FrmAccLunas.frx":00AE
         Spin            =   "FrmAccLunas.frx":00F8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   8421631
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
      Begin TDBNumber6Ctl.TDBNumber TxtJumlahCurbal 
         Height          =   255
         Left            =   7380
         TabIndex        =   18
         Top             =   4680
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   450
         Calculator      =   "FrmAccLunas.frx":0120
         Caption         =   "FrmAccLunas.frx":0140
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmAccLunas.frx":01AC
         Keys            =   "FrmAccLunas.frx":01CA
         Spin            =   "FrmAccLunas.frx":0214
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   8421631
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
      Begin TDBNumber6Ctl.TDBNumber TxtJumlahPayment 
         Height          =   255
         Left            =   10620
         TabIndex        =   20
         Top             =   4680
         Width           =   1710
         _Version        =   65536
         _ExtentX        =   3016
         _ExtentY        =   450
         Calculator      =   "FrmAccLunas.frx":023C
         Caption         =   "FrmAccLunas.frx":025C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmAccLunas.frx":02C8
         Keys            =   "FrmAccLunas.frx":02E6
         Spin            =   "FrmAccLunas.frx":0330
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   8454016
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
      Begin VB.Label Label8 
         Caption         =   "Total Payment:"
         Height          =   195
         Left            =   9540
         TabIndex        =   19
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Total CurBal:"
         Height          =   195
         Left            =   6420
         TabIndex        =   17
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Total Balance MMU:"
         Height          =   195
         Left            =   2640
         TabIndex        =   15
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah Data:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   4680
         Width           =   1035
      End
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   4680
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   "Account Batal"
      Height          =   195
      Left            =   5160
      TabIndex        =   39
      Top             =   8460
      Width           =   1995
   End
   Begin VB.Label Label10 
      Caption         =   "Account Belum Lunas"
      Height          =   195
      Left            =   2640
      TabIndex        =   23
      Top             =   8460
      Width           =   1995
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00008080&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   2160
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "Account Lunas"
      Height          =   195
      Left            =   600
      TabIndex        =   22
      Top             =   8460
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   180
      Top             =   8400
      Width           =   375
   End
End
Attribute VB_Name = "FrmAccLunas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CekPayment As String
Dim b_exportExcel As Boolean
Private Const warna_lunas = vbBlue
Private Const warna_blm_lunas = &H8080&
Private Const warna_batal = vbRed
Private b_exportSID As Boolean
Private sLokasiExcel As String
Private bCari_Bylist_cust As Boolean

Private warna_x As String

Private Sub HeaderLunas()
    LvAccLunas.ColumnHeaders.ADD 1, , "Custid", 2000
    LvAccLunas.ColumnHeaders.ADD 2, , "Nama Customer", 3000
    LvAccLunas.ColumnHeaders.ADD 3, , "Tipe Kartu", 1500
    LvAccLunas.ColumnHeaders.ADD 4, , "Tipe PTP", 1500
    LvAccLunas.ColumnHeaders.ADD 5, , "Balance dari MMU", 1500
    LvAccLunas.ColumnHeaders.ADD 6, , "Angka Deal (Khusus PTP Discount)", 1500
    LvAccLunas.ColumnHeaders.ADD 7, , "Total Payment Saat ini", 1500
    LvAccLunas.ColumnHeaders.ADD 8, , "Sisa Payment", 1500
    LvAccLunas.ColumnHeaders.ADD 9, , "Status Account", 1500
    LvAccLunas.ColumnHeaders.ADD 10, , "Agent Saat ini", 1500
    LvAccLunas.ColumnHeaders.ADD 11, , "Proposal Date", 1500
    LvAccLunas.ColumnHeaders.ADD 12, , "Tenor", 1500
    LvAccLunas.ColumnHeaders.ADD 13, , "Batas Akhir", 1500
    LvAccLunas.ColumnHeaders.ADD 14, , "statuslastcall", 1500
End Sub

Private Sub IsiLunas()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim m_objrs_waktu As ADODB.Recordset
    Dim ListItem As ListItem
    Dim AngkaDeal As Double
    Dim m_objrs_payment As ADODB.Recordset
    Dim SisaPayment As Double
    Dim JumlahPayment As Double
    Dim BalanceMMU As Double
    Dim K As Integer
    Dim M_WHERE As String
    
    Dim GrandTotalBalanceMMU As Double
    Dim GrandTotalCurBal As Double
    Dim TotalPayment As Double
    
    Dim dProposal As Date
    Dim Tenor As Integer
    Dim batas_akhir_bayar As Date
    
    Dim vstatus As String
    Dim sts_acc_lunas As String
    Dim z As Integer
    Dim xx As Integer
    
    Dim cust_exist As Boolean
    Dim list_cust_sel As String
        
    'On Error GoTo SALAH
    
    Set m_objrs_payment = New ADODB.Recordset
    m_objrs_payment.CursorLocation = adUseClient
    m_objrs_payment.ActiveConnection = M_OBJCONN
    m_objrs_payment.LockType = adLockOptimistic
    m_objrs_payment.CursorType = adOpenDynamic
    
    M_WHERE = ""
    
    
    
    Set m_objrs_waktu = New ADODB.Recordset
    m_objrs_waktu.CursorLocation = adUseClient
    m_objrs_waktu.Open "SELECT now() as waktu_server ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    waktu_sekarang = Format(m_objrs_waktu!waktu_server, "yyyy-mm-dd")
    Set m_objrs_waktu = Nothing
    
    Frame1.Visible = True
    If CekPayment = "0" Then
        DoEvents
        
        'Kasih pesan dulu deh
        MsgBox "sistem akan melakukan kalkulasi payment, setelah anda menekan tombol OK dari pesan ini :) Mohon tunggu sebentar! ", vbOKOnly + vbInformation, "Informasi"

        'Update dulu paymentnya di mgm
        CMDSQL = "update mgm set total_payment=total_payment_new"
        CMDSQL = CMDSQL & " From "
        CMDSQL = CMDSQL & " (select b.custid_new,b.total_payment_new from mgm, "
        CMDSQL = CMDSQL & " (select custid as custid_new,sum(payment) as total_payment_new "
        CMDSQL = CMDSQL & " from tbllunas group by custid) as b "
        CMDSQL = CMDSQL & " where mgm.custid=b.custid_new) as c "
        CMDSQL = CMDSQL & " Where "
        CMDSQL = CMDSQL & " mgm.CustId = c.custid_new "

        ' 01 APRIL 2014 ==========================================
        'cmdsql = "UPDATE mgm SET total_payment=total_payment_new "
        'cmdsql = cmdsql & " FROM (SELECT tbllunas.custid as custid_new,sum(payment) as total_payment_new FROM mgm,tbllunas WHERE mgm.custid=tbllunas.custid AND date(tbllunas.paydate)+1 > mgm.tglsource GROUP by tbllunas.custid) x "
        'cmdsql = cmdsql & " WHERE mgm.CustId = x.custid_new "
        ' ========================================================
        
        M_OBJCONN.Execute CMDSQL
        
        CekPayment = "1"
    End If
    
    If Trim(LCase(CmbStatus.Text) = "lunas") Then
        ' Tambahan Filter Tanggal Pencarian Acc Lunas 19 Juni 2013
        M_WHERE = " AND (date(mgm.tgl_paid_off) BETWEEN '" & Format(DTPicker1(0).Value, "yyyy-mm-dd") & "' AND '" & Format(DTPicker1(1).Value, "yyyy-mm-dd") & "' )"
    End If
    
    'Buat Filter Data
    If CmbTipeKartu.Text <> "ALL" Then
        If M_WHERE = "" Then
            M_WHERE = " and mgm.acc_type='" & CmbTipeKartu.Text & "' "
        Else
            M_WHERE = M_WHERE & " and mgm.acc_type='" & CmbTipeKartu.Text & "' "
        End If
    End If
    
    If CmbJenisPTP.Text <> "ALL" Then
        If M_WHERE = "" Then
            M_WHERE = " and tblsendptp_log_approve.jenis_ptp='" & CmbJenisPTP.Text & "' "
        Else
            M_WHERE = M_WHERE & " and tblsendptp_log_approve.jenis_ptp='" & CmbJenisPTP.Text & "' "
        End If
    End If
    
    If CmbStatus.Text <> "ALL" Then
        If M_WHERE = "" Then
            M_WHERE = " and mgm.status_lunas='" & CmbStatus.Text & "' "
        Else
            M_WHERE = M_WHERE & " and mgm.status_lunas='" & CmbStatus.Text & "' "
        End If
    End If
    
    ' By List
    If bCari_Bylist_cust = True Then
        For xx = 0 To List1.ListCount - 1
            list_cust_sel = list_cust_sel & "'" & List1.list(xx) & "',"
        Next xx
        
        list_cust_sel = Mid(list_cust_sel, 1, Len(list_cust_sel) - 1)
        
        If M_WHERE = "" Then
            M_WHERE = " and mgm.custid in (" & list_cust_sel & ")"
        Else
            M_WHERE = M_WHERE & " and mgm.custid in in (" & list_cust_sel & ")"
        End If
    Else
        If txtcustid.Text <> "" Then
            If M_WHERE = "" Then
                M_WHERE = " and mgm.custid like '%" & txtcustid.Text & "%' "
            Else
                M_WHERE = M_WHERE & " and mgm.custid like '%" & txtcustid.Text & "%' "
            End If
        End If
    End If
    
    CMDSQL = "SELECT * FROM mgm,tblsendptp_log_approve,tblcpa WHERE "
    CMDSQL = CMDSQL & " tblcpa.vcustid = mgm.custid AND "
    CMDSQL = CMDSQL & " tblsendptp_log_approve.custid=tblcpa.vcustid AND "
    CMDSQL = CMDSQL & " tblsendptp_log_approve.custid=mgm.custid AND "
    CMDSQL = CMDSQL & " date(tblsendptp_log_approve.tgl_proposal)=date(tblcpa.dpropsal) AND "
    CMDSQL = CMDSQL & " tblcpa.nid in (select max(nid) from tblcpa group by vcustid)  "
    'Cmdsql = Cmdsql & " mgm.custid in (select distinct custid from tbllunas) "
    CMDSQL = CMDSQL & " AND mgm.agent<>'LUNAS' "
    CMDSQL = CMDSQL & M_WHERE
    CMDSQL = CMDSQL & " ORDER BY tblsendptp_log_approve.jenis_ptp,mgm.acc_type,mgm.name asc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Frame1.Visible = False
    
    LvAccLunas.ListItems.CLEAR
    TxtJmlData.Text = M_objrs.RecordCount
    
    If M_objrs.RecordCount > 0 Then
        cust_exist = False
'        If Trim(TxtCustid.Text) <> "" Then
'            For xx = 0 To List1.ListCount - 1
'                If TxtCustid.Text = List1.list(xx) Then
'                    cust_exist = True
'                End If
'            Next xx
'            ' Add list customer
'            If cust_exist = False Then
'                List1.AddItem TxtCustid.Text
'            End If
'        End If
        PB1.Max = M_objrs.RecordCount
        TotalPayment = 0
        GrandTotalBalanceMMU = 0
        GrandTotalCurBal = 0
        
        M_OBJCONN.Execute "DELETE FROM temp_proses_lunas ;"
        
        If m_objrs_payment.State = 1 Then m_objrs_payment.Close
        m_objrs_payment.Open "SELECT * FROM temp_proses_lunas WHERE custid='XXXX'"
        
        While Not M_objrs.EOF
            DoEvents
            PB1.Value = M_objrs.Bookmark
            
            TotalPayment = TotalPayment + Val(IIf(IsNull(M_objrs("total_payment")), 0, M_objrs("total_payment")))
            GrandTotalBalanceMMU = GrandTotalBalanceMMU + Val(IIf(IsNull(M_objrs("amountwo")), 0, M_objrs("amountwo")))
            GrandTotalCurBal = GrandTotalCurBal + Val(IIf(IsNull(M_objrs("curbal")), 0, M_objrs("curbal")))
            
            'Ambil nilai Deal khusus PTP yang discount
            If UCase(Trim(M_objrs("jenis_ptp"))) = "PTP DISCOUNT" Then
                AngkaDeal = IIf(IsNull(M_objrs("nttlpayment")), "0", M_objrs("nttlpayment"))
            Else
                AngkaDeal = 0
            End If
            
            BalanceMMU = M_objrs("amountwo")
            vstatus = IIf(IsNull(M_objrs("f_cek_new")), "", M_objrs("f_cek_new"))
                
            dProposal = IIf(IsNull(M_objrs("dpropsal")), "", M_objrs("dpropsal"))
            JumlahPayment = IIf(IsNull(M_objrs("total_payment")), 0, M_objrs("total_payment"))
                
            SisaPayment = 0
            'Hitung Sisa Payment
            If UCase(Trim(M_objrs("jenis_ptp"))) = "PTP DISCOUNT" Then
                'Jika CH PTP discount sisa payment yang harus dibayar Angka Deal-Jumlah Payment:
                SisaPayment = AngkaDeal - JumlahPayment
            ElseIf UCase(Trim(M_objrs("jenis_ptp"))) = "PTP NO DISCOUNT" Then
                'Jika CH ptp no discount , payment yang harus dibayar:
                SisaPayment = BalanceMMU - JumlahPayment
            End If
                
            Tenor = IIf(IsNull(M_objrs("nperiod")), 0, M_objrs("nperiod"))
            'If M_Objrs!CustId = "4544931104468120" Then MsgBox "4544931104468120 - batal"
            
            'If Right(M_Objrs!CustId, 6) = "147123" Then MsgBox "ok"
            
            If CDate(waktu_sekarang) <= DateAdd("m", Tenor, dProposal) Then
'                If Format(dProposal, "dd") <= 15 Then
'                batas_akhir_bayar = "20" + "-" + Format(waktu_sekarang, "mm-yyyy")
'                ElseIf Format(dProposal, "dd") > 15 Then
'                batas_akhir_bayar = "05" + "-" + DateAdd("m", 1, Format(waktu_sekarang, "mm")) + Format(waktu_sekarang, "yyyy")
'
'                End If
                If Format(dProposal, "dd") <= 15 Then
                    batas_akhir_bayar = Format(waktu_sekarang, "yyyy-mm-20")
                ElseIf Format(dProposal, "dd") > 15 Then
                    batas_akhir_bayar = Format(waktu_sekarang, "yyyy-") & Format(DateAdd("m", 1, waktu_sekarang), "mm-") & "05"
                End If
            Else
               batas_akhir_bayar = DateAdd("m", Tenor, dProposal)
            End If
            'batas_akhir_bayar = DateAdd("m", tenor, dProposal)
                
'            Set ListItem = LvAccLunas.ListItems.ADD(, , M_Objrs("custid"))
'            ListItem.SubItems(1) = M_Objrs("name")
'            ListItem.SubItems(2) = IIf(IsNull(M_Objrs("acc_type")), "", M_Objrs("acc_type"))
'            ListItem.SubItems(3) = IIf(IsNull(M_Objrs("jenis_ptp")), "", M_Objrs("jenis_ptp"))
'            ListItem.SubItems(4) = BalanceMMU
'            ListItem.SubItems(5) = AngkaDeal
'            ListItem.SubItems(6) = JumlahPayment
'            ListItem.SubItems(7) = SisaPayment
'            ListItem.SubItems(9) = M_Objrs("agent")
'            ' TAMBAHAN UNTUK TENOR DISC ----------------------------------------
'            ListItem.SubItems(10) = Format(dProposal, "yyyy-mm-dd")
'            ListItem.SubItems(11) = tenor
'            ListItem.SubItems(12) = Format(batas_akhir_bayar, "yyyy-mm-dd")
'            ' ------------------------------------------------------------------
'            ListItem.SubItems(13) = vstatus
            
'            sts_acc_lunas = update_status_acc(M_Objrs("custid"), SisaPayment, IIf(IsNull(M_Objrs("jenis_ptp")), "", M_Objrs("jenis_ptp")), Format(batas_akhir_bayar, "yyyy-mm-dd"))
'            ListItem.SubItems(8) = sts_acc_lunas
            
            
            m_objrs_payment.AddNew
            m_objrs_payment!CustId = M_objrs("custid")
            m_objrs_payment!custname = M_objrs("name")
            m_objrs_payment!acc_type = IIf(IsNull(M_objrs("acc_type")), "", M_objrs("acc_type"))
            m_objrs_payment!jenis_ptp = IIf(IsNull(M_objrs("jenis_ptp")), "", M_objrs("jenis_ptp"))
            m_objrs_payment!balance_MMU = BalanceMMU
            m_objrs_payment!angka_deal = AngkaDeal
            If AngkaDeal <= 0 Then
                m_objrs_payment!jml_payment = JumlahPayment
            Else
                m_objrs_payment!jml_payment = 0
            End If
            m_objrs_payment!sisa_payment = SisaPayment
            m_objrs_payment!status_lunas = ""
            m_objrs_payment!AGENT = M_objrs("agent")
            m_objrs_payment!tgl_proposal = Format(dProposal, "yyyy-mm-dd")
            m_objrs_payment!Tenor = Tenor
            m_objrs_payment!batas_akhir = Format(batas_akhir_bayar, "yyyy-mm-dd")
            m_objrs_payment!status_acc = vstatus
            'm_objrs_payment!lpd = M_Objrs("lpd_from_payment")
            m_objrs_payment.update

'            If SisaPayment <= 0 Then
'
'                ListItem.SubItems(8) = "LUNAS"
'                ListItem.ForeColor = vbBlue
'
'                'Update statusnya ke lunas
'                cmdsql = "update mgm set status_lunas='LUNAS' where custid='"
'                cmdsql = cmdsql & CStr(M_Objrs("custid")) & "'"
'                M_OBJCONN.Execute cmdsql
'                For K = 1 To 12
'                    ListItem.ListSubItems(K).ForeColor = vbBlue
'                Next K
'
'            Else
'
'                If UCase(Trim(M_Objrs("jenis_ptp"))) = "PTP DISCOUNT" Then
'                    If waktu_sekarang > Format(batas_akhir_bayar, "yyyy-mm-dd") Then
'                        ListItem.SubItems(8) = "BATAL"
'                        ListItem.ForeColor = vbRed
'                        For K = 1 To 12
'                            ListItem.ListSubItems(K).ForeColor = vbRed
'                        Next K
'                        'Update statusnya ke belum lunas
'                        cmdsql = "update mgm set status_lunas='BATAL' where custid='"
'                        cmdsql = cmdsql & CStr(M_Objrs("custid")) & "'"
'                        M_OBJCONN.Execute cmdsql
'                        '-------------------------------
'                    Else
'                        ListItem.SubItems(8) = "BELUM LUNAS"
'                        ListItem.ForeColor = &H8080&
'                        For K = 1 To 12
'                            ListItem.ListSubItems(K).ForeColor = &H8080&
'                        Next K
'                        'Update statusnya ke belum lunas
'                        cmdsql = "update mgm set status_lunas='BELUM LUNAS' where custid='"
'                        cmdsql = cmdsql & CStr(M_Objrs("custid")) & "'"
'                        M_OBJCONN.Execute cmdsql
'                        '-------------------------------
'                    End If
'
'                Else
'
'                    ListItem.SubItems(8) = "BELUM LUNAS"
'                    ListItem.ForeColor = &H8080&
'
'                    'Update statusnya ke belum lunas
'                    cmdsql = "update mgm set status_lunas='BELUM LUNAS' where custid='"
'                    cmdsql = cmdsql & CStr(M_Objrs("custid")) & "'"
'                    M_OBJCONN.Execute cmdsql
'                    For K = 1 To 12
'                        ListItem.ListSubItems(K).ForeColor = &H8080&
'                    Next K
'
'                End If
'            End If
            M_objrs.MoveNext
        Wend
        
        ' UPDATE PEMBAYARAN UNTUK YG PTP ================
'        M_OBJCONN.Execute "UPDATE temp_proses_lunas SET jml_payment=z.total_bayar FROM " & _
'                            "(SELECT x.custid,sum(y.payment) as total_bayar FROM temp_proses_lunas x," & _
'                            "(SELECT custid,paydate,payment FROM tbllunas) y WHERE x.custid=y.custid AND " & _
'                            "date(y.paydate) >= date(x.tgl_proposal) GROUP BY x.custid) z " & _
'                            "WHERE temp_proses_lunas.custid=z.custid AND temp_proses_lunas.angka_deal > 0;"
        M_OBJCONN.Execute "UPDATE temp_proses_lunas SET jml_payment=z.total_bayar,lpd=z.Tgl_lpd FROM " & _
                            "(SELECT x.custid,sum(y.payment) as total_bayar,max(paydate) as Tgl_lpd FROM temp_proses_lunas x," & _
                            "(SELECT custid,paydate,payment FROM tbllunas) y,(select custid,tglsource FROM mgm) a WHERE a.custid=x.custid AND x.custid=y.custid AND " & _
                            "date(y.paydate) >= date(x.tgl_proposal) AND date(y.paydate)>date(a.tglsource) GROUP BY x.custid) z " & _
                            "WHERE temp_proses_lunas.custid=z.custid AND temp_proses_lunas.angka_deal > 0;"
        ' ===============================================
        
        If M_objrs.State = 1 Then M_objrs.Close
        M_objrs.Open "SELECT * FROM temp_proses_lunas"
        
        If M_objrs.RecordCount > 0 Then
            PB1.Max = M_objrs.RecordCount
            While Not M_objrs.EOF
                DoEvents
                PB1.Value = M_objrs.Bookmark
    '            If Left(LvAccLunas.ListItems(z).SubItems(13), 3) = "PTP" Then
    '                 '-------- cek lunas ---------
    '                cmdsql = "SELECT sum(payment) as total_pembayaran_x FROM tbllunas WHERE custid='" & LvAccLunas.ListItems(z).Text & "' AND date(paydate)>='" & Format(LvAccLunas.ListItems(z).SubItems(10), "yyyy-mm-dd") & "'"
    '
    '                If m_objrs_payment.state = 1 Then m_objrs_payment.Close
    '                m_objrs_payment.Open cmdsql
    '
    '                BalanceMMU = ListItem.SubItems(4)
    '                AngkaDeal = ListItem.SubItems(5)
    '
    '                If m_objrs_payment.RecordCount > 0 Then
    '                    JumlahPayment = IIf(IsNull(m_objrs_payment("total_pembayaran_x")), "0", m_objrs_payment("total_pembayaran_x"))
    '                    LvAccLunas.ListItems(z).SubItems(6) = JumlahPayment
    '                Else
    '                    JumlahPayment = 0
    '                    LvAccLunas.ListItems(z).SubItems(6) = JumlahPayment
    '                End If
    '
    '                SisaPayment = 0
    '                If UCase(LvAccLunas.ListItems(z).SubItems(3)) = "PTP DISCOUNT" Then
    '                    'Jika CH PTP discount sisa payment yang harus dibayar Angka Deal-Jumlah Payment:
    '                    SisaPayment = AngkaDeal - JumlahPayment
    '                ElseIf UCase(LvAccLunas.ListItems(z).SubItems(3)) = "PTP NO DISCOUNT" Then
    '                    'Jika CH ptp no discount , payment yang harus dibayar:
    '                    SisaPayment = BalanceMMU - JumlahPayment
    '                End If
    '
    '                LvAccLunas.ListItems(z).SubItems(7) = SisaPayment
    '
    '                sts_acc_lunas = update_status_acc(LvAccLunas.ListItems(z).Text, SisaPayment, LvAccLunas.ListItems(z).SubItems(3), Format(LvAccLunas.ListItems(z).SubItems(12), "yyyy-mm-dd"))
    '                LvAccLunas.ListItems(z).SubItems(8) = sts_acc_lunas
    '
    '                Set m_objrs_payment = Nothing
    '            End If
                'If M_Objrs!CustId = "045444197840" Then MsgBox "ok"
                'If Right(M_Objrs!CustId, 6) = "147123" Then MsgBox "ok"
                
                BalanceMMU = IIf(IsNull(M_objrs("balance_MMU")), 0, M_objrs("balance_MMU"))
                AngkaDeal = IIf(IsNull(M_objrs("angka_deal")), 0, M_objrs("angka_deal"))
                JumlahPayment = IIf(IsNull(M_objrs("jml_payment")), 0, M_objrs("jml_payment"))
                
                If UCase(Trim(M_objrs("jenis_ptp"))) = "PTP DISCOUNT" Then
                    'Jika CH PTP discount sisa payment yang harus dibayar Angka Deal-Jumlah Payment:
                    SisaPayment = AngkaDeal - JumlahPayment
                ElseIf UCase(Trim(M_objrs("jenis_ptp"))) = "PTP NO DISCOUNT" Then
                    'Jika CH ptp no discount , payment yang harus dibayar:
                    SisaPayment = BalanceMMU - JumlahPayment
                End If
    
                Set ListItem = LvAccLunas.ListItems.ADD(, , M_objrs("custid"))
                ListItem.SubItems(1) = M_objrs("custname")
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("acc_type")), "", M_objrs("acc_type"))
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("jenis_ptp")), "", M_objrs("jenis_ptp"))
                ListItem.SubItems(4) = BalanceMMU
                ListItem.SubItems(5) = AngkaDeal
                ListItem.SubItems(6) = JumlahPayment
                ListItem.SubItems(7) = SisaPayment 'SisaPayment
                
                sts_acc_lunas = update_status_acc(M_objrs("custid"), SisaPayment, IIf(IsNull(M_objrs("jenis_ptp")), "", M_objrs("jenis_ptp")), IIf(IsNull(M_objrs("batas_akhir")), "", Format(M_objrs("batas_akhir"), "yyyy-mm-dd")), IIf(IsNull(M_objrs("lpd")), "1970-01-01", Format(M_objrs("lpd"), "yyyy-mm-dd")))
                ListItem.SubItems(8) = sts_acc_lunas
                
                ListItem.SubItems(9) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent")) 'M_Objrs("agent")
                ' TAMBAHAN UNTUK TENOR DISC ----------------------------------------
                ListItem.SubItems(10) = IIf(IsNull(M_objrs("tgl_proposal")), "", M_objrs("tgl_proposal")) 'Format(dProposal, "yyyy-mm-dd")
                ListItem.SubItems(11) = IIf(IsNull(M_objrs("tenor")), "", M_objrs("tenor")) 'tenor
                ListItem.SubItems(12) = IIf(IsNull(M_objrs("batas_akhir")), "", M_objrs("batas_akhir")) 'Format(batas_akhir_bayar, "yyyy-mm-dd")
                ' ------------------------------------------------------------------
                ListItem.SubItems(13) = IIf(IsNull(M_objrs("status_acc")), "", M_objrs("status_acc")) 'vstatus
                
                Select Case sts_acc_lunas
                Case "LUNAS"
                    warna_x = warna_lunas
                Case "BELUM LUNAS"
                    warna_x = warna_blm_lunas
                Case "BATAL"
                    warna_x = warna_batal
                End Select
                
                For K = 1 To 13
                    ListItem.ForeColor = warna_x
                    ListItem.ListSubItems(K).ForeColor = warna_x
                Next K
                
            M_objrs.MoveNext
            Wend
            
            MsgBox "Progress Done !!!", vbOKOnly, "INFO"
        End If
    Else
        MsgBox "Data Tidak ditemukan!!"
    End If
    
    Set M_objrs = Nothing
    
    TxtBalanceMMU.Value = GrandTotalBalanceMMU
    TxtJumlahCurbal.Value = GrandTotalCurBal
    TxtJumlahPayment.Value = TotalPayment
    Exit Sub
Salah:
    MsgBox "Mohon maaf! Ada error: " & err.Description, vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub IsiTipeKartu()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    On Error GoTo Salah
    
    CmbTipeKartu.CLEAR
    CmbTipeKartu.AddItem "ALL"
    
    CMDSQL = "select distinct acc_type from mgm where acc_type is not null "
    CMDSQL = CMDSQL & " or acc_type='' order by acc_type asc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            CmbTipeKartu.AddItem M_objrs("acc_type")
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
    CmbTipeKartu.Text = "ALL"
    Exit Sub
Salah:
    MsgBox "Maaf, ada error: " & err.Description, vbOKOnly + vbExclamation, "Peringatan"
    
End Sub


Private Sub IsiJenisPTP()
    CmbJenisPTP.CLEAR
    CmbJenisPTP.AddItem "ALL"
    CmbJenisPTP.AddItem "PTP Discount"
    CmbJenisPTP.AddItem "PTP No Discount"
    CmbJenisPTP.Text = "ALL"
End Sub

Private Sub IsiStatusLunas()
    CmbStatus.CLEAR
    CmbStatus.AddItem "ALL"
    CmbStatus.AddItem "LUNAS"
    CmbStatus.AddItem "BATAL"
    CmbStatus.AddItem "BELUM LUNAS"
    CmbStatus.Text = "ALL"
End Sub

Private Sub Check1_Click()
    Dim W As Integer
    
    If LvAccLunas.ListItems.Count = 0 Then
        MsgBox "Maaf data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If Check1.Value = 1 Then
        For W = 1 To LvAccLunas.ListItems.Count
            If LCase(LvAccLunas.ListItems(W).SubItems(8)) = "lunas" Then
                LvAccLunas.ListItems(W).Checked = True
            End If
        Next W
    Else
        For W = 1 To LvAccLunas.ListItems.Count
            If LCase(LvAccLunas.ListItems(W).SubItems(8)) = "lunas" Then
                LvAccLunas.ListItems(W).Checked = False
            End If
        Next W
    End If
End Sub

Private Sub Check2_Click()
    Dim W As Integer
    
    If LvAccLunas.ListItems.Count = 0 Then
        MsgBox "Maaf data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If Check2.Value = 1 Then
        For W = 1 To LvAccLunas.ListItems.Count
            If LCase(LvAccLunas.ListItems(W).SubItems(8)) = "belum lunas" Then
                LvAccLunas.ListItems(W).Checked = True
            End If
        Next W
    Else
        For W = 1 To LvAccLunas.ListItems.Count
            If LCase(LvAccLunas.ListItems(W).SubItems(8)) = "belum lunas" Then
                LvAccLunas.ListItems(W).Checked = False
            End If
        Next W
    End If
End Sub

Private Sub Check3_Click()
    Dim W As Integer
    
    If LvAccLunas.ListItems.Count = 0 Then
        MsgBox "Maaf data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If Check3.Value = 1 Then
        For W = 1 To LvAccLunas.ListItems.Count
            If LCase(LvAccLunas.ListItems(W).SubItems(8)) = "batal" Then
                LvAccLunas.ListItems(W).Checked = True
            End If
        Next W
    Else
        For W = 1 To LvAccLunas.ListItems.Count
            If LCase(LvAccLunas.ListItems(W).SubItems(8)) = "batal" Then
                LvAccLunas.ListItems(W).Checked = False
            End If
        Next W
    End If
End Sub

Private Sub CmbStatus_Click()
    If UCase(CmbStatus.Text) = "LUNAS" Then
        DTPicker1(0).Enabled = True
        DTPicker1(1).Enabled = True
        CmdPindahLunas.Enabled = False
    Else
        DTPicker1(0).Enabled = False
        DTPicker1(1).Enabled = False
        CmdPindahLunas.Enabled = True
    End If
End Sub

Private Sub CmdCekAll_Click()
    Dim W As Integer
    
    If LvAccLunas.ListItems.Count = 0 Then
        MsgBox "Maaf data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvAccLunas.ListItems.Count
        LvAccLunas.ListItems(W).Checked = True
    Next W
End Sub

Private Sub CmdClear_Click()
    CmbTipeKartu.Text = "ALL"
    CmbJenisPTP.Text = "ALL"
    CmbStatus.Text = "ALL"
    txtcustid.Text = ""
End Sub

Private Sub CmdFilter_Click()
    Call IsiLunas
End Sub

Private Sub CmdPindahLunas_Click()
    Dim a As String
    Dim W, K, S As Integer
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    On Error GoTo Salah
    
    If LvAccLunas.ListItems.Count = 0 Then
        MsgBox "Maaf, data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin akan memindahkan account yang dicentang ke coding lunas? Account yang dapat dpindah ke coding lunas hanya data yang sudah lunas!", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbNo Then
        MsgBox "Proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    S = 0
    
    For W = 1 To LvAccLunas.ListItems.Count
        If LvAccLunas.ListItems(W).Checked = True Then
            S = S + 1
            Exit For
        End If
    Next W
    
    If S = 0 Then
        MsgBox "Maaf, anda belum memilih data yang akan dipindahkan ke coding lunas!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    PB1.Max = LvAccLunas.ListItems.Count
    
    DoEvents
    MousePointer = vbHourglass
    FrameFilter.Enabled = False
    FrameList.Enabled = False
    For K = 1 To LvAccLunas.ListItems.Count
        PB1.Value = K
        If LvAccLunas.ListItems(K).Checked = True Then
            'Cek dulu apakah statusnya Lunas atau tidak ?
            If UCase(Trim(LvAccLunas.ListItems(K).SubItems(8))) = "LUNAS" Then
                CMDSQL = "update mgm set log_agent_lunas=agent, agent='LUNAS' where custid='"
                CMDSQL = CMDSQL & CStr(LvAccLunas.ListItems(K).Text) & "'"
                M_OBJCONN.Execute CMDSQL
            End If
        End If
    Next K
    
    MsgBox "Data berhasil dipindahkan ke coding lunas!", vbOKOnly + vbInformation, "Informasi"
    
    CmdRefresh_Click
    MousePointer = vbNormal
    FrameFilter.Enabled = True
    FrameList.Enabled = True
    Exit Sub
Salah:
    MsgBox "Maaf, ada kesalahan! " & err.Description, vbOKOnly + vbInformation, "Informasi"
    
End Sub

Private Sub CmdRefresh_Click()
    CmbTipeKartu.Text = "ALL"
    CmbJenisPTP.Text = "ALL"
    CmbStatus.Text = "ALL"
    txtcustid.Text = ""
    FrameFilter.Enabled = False
    FrameList.Enabled = False
    Call IsiLunas
    FrameFilter.Enabled = True
    FrameList.Enabled = True
End Sub

Private Sub CmdTutup_Click()
    Unload Me
End Sub

Private Sub CmdUnCekAll_Click()
    Dim W As Integer
    
    If LvAccLunas.ListItems.Count = 0 Then
        MsgBox "Maaf data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvAccLunas.ListItems.Count
        LvAccLunas.ListItems(W).Checked = False
    Next W
End Sub

Private Sub CmdViewCPAElektornik_Click()
    Dim CMDSQL As String
    Dim rsTemporary As ADODB.Recordset
    Dim W, K, S As Integer
    Dim CustId As String
    Dim rsTemp1 As ADODB.Recordset
    Dim M_OBJRS_LPD_LPA As ADODB.Recordset
    
    On Error GoTo Salah
    CustId = ""
    
    If LvAccLunas.ListItems.Count = 0 Then
        MsgBox "Maaf, data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    S = 0
    For K = 1 To LvAccLunas.ListItems.Count
        If LvAccLunas.ListItems(K).Checked = True Then
            S = S + 1
            Exit For
        End If
    Next K
    
    If S = 0 Then
        MsgBox "Maaf anda belum memilih data yang akan di lihat!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    FrameFilter.Enabled = False
    FrameList.Enabled = False
    'Ambil nilai custid
    For W = 1 To LvAccLunas.ListItems.Count
        If LvAccLunas.ListItems(W).Checked = True Then
            If CustId = "" Then
                CustId = "'" & CStr(LvAccLunas.ListItems(W).Text) & "'"
            Else
                CustId = CustId & ",'" & CStr(LvAccLunas.ListItems(W).Text) & "'"
            End If
        End If
    Next W
    
'     Cmdsql = "select * from "
'        '--------------------------- Joint Antara Tabel MGM Sama CPA ------------------------
'        Cmdsql = Cmdsql + "(SELECT  * FROM ( "
'        Cmdsql = Cmdsql + " SELECT * FROM TBLCPA) AS A"
'        Cmdsql = Cmdsql + " Inner Join (SELECT * FROM   MGM) AS B  ON A.VCUSTID=B.CUSTID ) as cpa_mgm, "
'        '--------------------------- Joint Antara Tabel MGM Sama CPA ------------------------
'
'    Cmdsql = Cmdsql + "  tblsendptp_log_approve  as send_ptp_app "
'    Cmdsql = Cmdsql + " where cpa_mgm.vcustid=send_ptp_app.custid and "
'    Cmdsql = Cmdsql + " date(cpa_mgm.dpropsal)=date(send_ptp_app.tgl_proposal) and "
'    Cmdsql = Cmdsql + " cpa_mgm.custid in (" & CustId & ") "
    
    ' UPDATE 18042013
'    cmdsql = "SELECT a.*, b.maxdrpopsal FROM (select * from "
'    '--------------------------- Joint Antara Tabel MGM Sama CPA ------------------------
'    cmdsql = cmdsql + "(SELECT  * FROM ( "
'    cmdsql = cmdsql + " SELECT * FROM TBLCPA) AS A"
'    cmdsql = cmdsql + " Inner Join (SELECT * FROM MGM) AS B  ON A.VCUSTID=B.CUSTID ) as cpa_mgm, "
'    '--------------------------- Joint Antara Tabel MGM Sama CPA ------------------------
'
'    cmdsql = cmdsql + " (select custid, tgl_proposal, max(tgldata) as tgldata from tblsendptp_log_approve group by custid, tgl_proposal)  as send_ptp_app "
'    cmdsql = cmdsql + " where cpa_mgm.vcustid=send_ptp_app.custid and "
'    cmdsql = cmdsql + " date(cpa_mgm.dpropsal)=date(send_ptp_app.tgl_proposal) and "
'    cmdsql = cmdsql + " cpa_mgm.custid in (" & CustId & ")) a"
'
'    cmdsql = cmdsql + " ,(select vcustid, max(dpropsal)as maxdrpopsal from tblcpa where vcustid in (" & CustId & ") group by vcustid) b " & _
'                        " where a.vcustid=b.vcustid and date(b.maxdrpopsal)=date(a.dpropsal)"

    CMDSQL = "SELECT * FROM (select * from "
    '--------------------------- Joint Antara Tabel MGM Sama CPA ------------------------
    CMDSQL = CMDSQL + "(SELECT  *,B.dob as d_dob FROM ( "
    CMDSQL = CMDSQL + " SELECT x.* FROM TBLCPA x,(select vcustid, max(dpropsal)as maxdrpopsal from tblcpa group by vcustid) b  WHERE x.vcustid=b.vcustid AND x.dpropsal=b.maxdrpopsal) AS A"
    CMDSQL = CMDSQL + " Inner Join (SELECT * FROM MGM) AS B  ON A.VCUSTID=B.CUSTID ) as cpa_mgm, "
    '--------------------------- Joint Antara Tabel MGM Sama CPA ------------------------
                  
    CMDSQL = CMDSQL + " (select custid, tgl_proposal, max(tgldata) as tgldata from tblsendptp_log_approve group by custid, tgl_proposal)  as send_ptp_app "
    CMDSQL = CMDSQL + " where cpa_mgm.vcustid=send_ptp_app.custid and "
    CMDSQL = CMDSQL + " date(cpa_mgm.dpropsal)=date(send_ptp_app.tgl_proposal) and "
    CMDSQL = CMDSQL + " cpa_mgm.custid in (" & CustId & ")) a"
    
    Set rsTemporary = New ADODB.Recordset
    rsTemporary.CursorLocation = adUseClient
    
    If b_exportSID = True Then
'        M_OBJCONN.Execute "DELETE FROM tblcpa_sid;"
'        M_OBJCONN.Execute "INSERT INTO tblcpa_sid(ref_num,prd_type,name_,dob,id_no,requestor_name,mother_maiden_name) " & _
'                            " SELECT vcustid,vproduct,name,d_dob,ktpno,approve_by||'/ID/HBAP/HSBC',mother FROM ( " & cmdsql & " )"
'
'        rsTemporary.Open "SELECT * FROM tblcpa_sid;", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'        Call ConvertToExcel(rsTemporary, sLokasiExcel)
'
'        'Set rsTemp1 = Nothing
'        Set rsTemporary = Nothing
    Else
        rsTemporary.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        M_RPTCONN.Execute "delete from tblreportcpa"
        
        If rsTemporary.RecordCount > 0 Then
            PB1.Max = rsTemporary.RecordCount
    
            'Buka koneksi access
            CMDSQL = "select * from tblreportcpa "
            Set rsTemp1 = New ADODB.Recordset
            rsTemp1.CursorLocation = adUseClient
            rsTemp1.Open CMDSQL, M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
            While Not rsTemporary.EOF
                DoEvents
                PB1.Value = rsTemporary.Bookmark
                rsTemp1.AddNew
    
                rsTemp1("jenis") = "APPROVED"
                rsTemp1("status_ptp") = IIf(IsNull(rsTemporary("status_ptp")), "", rsTemporary("status_ptp"))
                rsTemp1("vregion") = IIf(IsNull(rsTemporary("region")), "", rsTemporary("region"))
                rsTemp1("dproposal") = IIf(IsNull(rsTemporary("dpropsal")), Null, rsTemporary("dpropsal"))
                rsTemp1("vreffno") = IIf(IsNull(rsTemporary("vreffno")), "", rsTemporary("vreffno"))
                rsTemp1("product") = IIf(IsNull(rsTemporary("vproduct")), "", rsTemporary("vproduct"))
                rsTemp1("arrangement") = IIf(IsNull(rsTemporary("varragement")), "", rsTemporary("varragement"))
                rsTemp1("cardno") = IIf(IsNull(rsTemporary("nocard")), "", rsTemporary("nocard"))
                rsTemp1("custname") = IIf(IsNull(rsTemporary("name")), "", rsTemporary("name"))
                rsTemp1("cardopen") = IIf(IsNull(rsTemporary("opendate")), Null, rsTemporary("opendate"))
                rsTemp1("agent") = IIf(IsNull(rsTemporary("agent")), "", rsTemporary("agent"))
                rsTemp1("outbalance") = IIf(IsNull(rsTemporary("nbalance")), 0, rsTemporary("nbalance"))
                rsTemp1("ttlpayment") = IIf(IsNull(rsTemporary("nttlpayment")), 0, rsTemporary("nttlpayment"))
                rsTemp1("downpayment") = IIf(IsNull(rsTemporary("ndownpay")), 0, rsTemporary("ndownpay"))
                'rsTemp1("futurepayment") = IIf(IsNull(rsTemporary("nfuturepay")), 0, rsTemporary("nfuturepay"))
                ' # Update 17 April 2013 By Izuddin
                rsTemp1("futurepayment") = IIf(IsNull(rsTemporary("nttlpayment")), 0, rsTemporary("nttlpayment")) - IIf(IsNull(rsTemporary("ndownpay")), 0, rsTemporary("ndownpay"))
                rsTemp1("nprincipal") = IIf(IsNull(rsTemporary("nprincipal")), 0, rsTemporary("nprincipal"))
                rsTemp1("ncharge") = IIf(IsNull(rsTemporary("ncharge")), 0, rsTemporary("ncharge"))
                rsTemp1("ndiskon") = IIf(IsNull(rsTemporary("ndiscountamt")), 0, rsTemporary("ndiscountamt"))
                rsTemp1("osfrombalance") = IIf(IsNull(rsTemporary("vosbalance")), "", rsTemporary("vosbalance"))
                rsTemp1("osfromprincipal") = IIf(IsNull(rsTemporary("vosprincipal")), "", rsTemporary("vosprincipal"))
                rsTemp1("custid") = IIf(IsNull(rsTemporary("vcustid")), "", rsTemporary("vcustid"))
                rsTemp1("nperiod") = IIf(IsNull(rsTemporary("nperiod")), 0, rsTemporary("nperiod"))
                'rsTemp1("approve") = IIf(IsNull(rsTemporary("vnameapprovel")), "", rsTemporary("vnameapprovel"))
                rsTemp1("approve") = IIf(IsNull(rsTemporary("approve_by")), "", rsTemporary("approve_by"))
                rsTemp1("sts_approve") = IIf(IsNull(rsTemporary("sts_approve")), "", rsTemporary("sts_approve"))
                rsTemp1("payment_after_tenor") = CStr(IIf(IsNull(rsTemporary("payment_after_tenor")), "0", rsTemporary("payment_after_tenor")))
    
                rsTemp1("vjust") = CStr(IIf(IsNull(rsTemporary("vjust")), "", Mid(rsTemporary("vjust"), 1, 250)))
    
                rsTemp1("voccupation") = CStr(IIf(IsNull(rsTemporary("voccupation")), "", rsTemporary("voccupation")))
                rsTemp1("vreason") = CStr(IIf(IsNull(rsTemporary("vreason")), "", rsTemporary("vreason")))
    
                rsTemp1("chkfaxed") = CStr(IIf(IsNull(rsTemporary("chkfaxed")), "0", rsTemporary("chkfaxed")))
                rsTemp1("chkwentalking") = CStr(IIf(IsNull(rsTemporary("chkwentalking")), "0", rsTemporary("chkwentalking")))
                rsTemp1("chkktp") = CStr(IIf(IsNull(rsTemporary("chkktp")), "0", rsTemporary("chkktp")))
                rsTemp1("chksup") = CStr(IIf(IsNull(rsTemporary("chksup")), "0", rsTemporary("chksup")))
                rsTemp1("chkbillings") = CStr(IIf(IsNull(rsTemporary("chkbillings")), "0", rsTemporary("chkbillings")))
                rsTemp1("chkothers") = CStr(IIf(IsNull(rsTemporary("chkothers")), "0", rsTemporary("chkothers")))
    
                rsTemp1("pay_off_date") = IIf(IsNull(rsTemporary("tgl_paid_off")), Null, rsTemporary("tgl_paid_off"))
                rsTemp1("txt_map") = IIf(IsNull(rsTemporary("map")), 0, rsTemporary("map"))
    
                '@@25072012, Catet f_cek_new yang paid off
                rsTemp1("f_cek_new") = CStr(IIf(Trim(rsTemporary("f_cek_new")) = "PO-", "PO-", ""))
                '@@26Juli2012, Simpan Wo Date nya
                rsTemp1("wo_date") = IIf(IsNull(rsTemporary("b_d")), Null, Format(rsTemporary("b_d"), "yyyy-mm-dd"))
    
                '-------------------------------- Cari LPD dan LPA ----------------------------------------
                CMDSQL = "select paydate,payment from tbllunas where custid='"
                CMDSQL = CMDSQL + CStr(Trim(rsTemporary("custid"))) + "' order by paydate desc limit 1 "
    
                Set M_OBJRS_LPD_LPA = New ADODB.Recordset
                M_OBJRS_LPD_LPA.CursorLocation = adUseClient
                M_OBJRS_LPD_LPA.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                If M_OBJRS_LPD_LPA.RecordCount > 0 Then
                    rsTemp1("lpd_from_payment") = IIf(IsNull(M_OBJRS_LPD_LPA("paydate")), Null, Format(M_OBJRS_LPD_LPA("paydate"), "yyyy-mm-dd"))
                    rsTemp1("lpa_from_payment") = IIf(IsNull(M_OBJRS_LPD_LPA("payment")), Null, M_OBJRS_LPD_LPA("payment"))
                Else
    
                End If
                Set M_OBJRS_LPD_LPA = Nothing
    
                '-------------------------------- Cari LPD dan LPA ----------------------------------------
                    rsTemp1.update
                    rsTemporary.MoveNext
            Wend
            Set rsTemp1 = Nothing
        End If
        
        'Set rsTemp1 = Nothing
        Set rsTemporary = Nothing
        
        FrameFilter.Enabled = True
        FrameList.Enabled = True
        
        ' Ini codingan CR OLD 16042013
        WaitSecs (2)

        RPT.Reset
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\Rptcpaelektronik.rpt"
        SHOW_PRN
    End If
    
    Exit Sub
Salah:
    MsgBox "Maaf ada error! " & err.Description, vbOKOnly + vbInformation, "Informasi"
    
    
End Sub

Private Sub CmdViewReport_Click()
    Dim a As String
    Dim W, K, S As Integer
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    On Error GoTo Salah
    
    If LvAccLunas.ListItems.Count = 0 Then
        MsgBox "Maaf, data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    
    S = 0
    
    For W = 1 To LvAccLunas.ListItems.Count
        If LvAccLunas.ListItems(W).Checked = True Then
            S = S + 1
            Exit For
        End If
    Next W
    
    If S = 0 Then
        MsgBox "Maaf, anda belum memilih data yang akan ditampilkan dalam report!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    M_RPTCONN.Execute "delete from rpt_lunas "
    
    PB1.Max = LvAccLunas.ListItems.Count
    FrameFilter.Enabled = False
    FrameList.Enabled = False
    
    For K = 1 To LvAccLunas.ListItems.Count
        PB1.Value = K
        If LvAccLunas.ListItems(K).Checked = True Then
            CMDSQL = "insert into rpt_lunas(custid,nama,card_tipe,jenis_ptp,balance_mmu,"
            CMDSQL = CMDSQL & "balance_deal,total_payment,sisa_payment,status_account) values ('"
            CMDSQL = CMDSQL & Trim(IIf(IsNull(LvAccLunas.ListItems(K).Text), "", LvAccLunas.ListItems(K).Text)) & "','"
            CMDSQL = CMDSQL & Trim(IIf(IsNull(LvAccLunas.ListItems(K).SubItems(1)), "", LvAccLunas.ListItems(K).SubItems(1))) & "','"
            CMDSQL = CMDSQL & IIf(IsNull(LvAccLunas.ListItems(K).SubItems(2)), "", LvAccLunas.ListItems(K).SubItems(2)) & "','"
            CMDSQL = CMDSQL & IIf(IsNull(LvAccLunas.ListItems(K).SubItems(3)), "", LvAccLunas.ListItems(K).SubItems(3)) & "',"
            CMDSQL = CMDSQL & IIf(IsNull(LvAccLunas.ListItems(K).SubItems(4)), "0", LvAccLunas.ListItems(K).SubItems(4)) & ","
            CMDSQL = CMDSQL & IIf(IsNull(LvAccLunas.ListItems(K).SubItems(5)), "0", LvAccLunas.ListItems(K).SubItems(5)) & ","
            CMDSQL = CMDSQL & IIf(IsNull(LvAccLunas.ListItems(K).SubItems(6)), "0", LvAccLunas.ListItems(K).SubItems(6)) & ","
            CMDSQL = CMDSQL & IIf(IsNull(LvAccLunas.ListItems(K).SubItems(7)), "0", LvAccLunas.ListItems(K).SubItems(7)) & ",'"
            CMDSQL = CMDSQL & IIf(IsNull(LvAccLunas.ListItems(K).SubItems(8)), "", LvAccLunas.ListItems(K).SubItems(8)) & "')"
            M_RPTCONN.Execute CMDSQL
        End If
    Next K
    
    FrameFilter.Enabled = True
    FrameList.Enabled = True
    WaitSecs (2)
    RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\Rptlunas.rpt"
    SHOW_PRN
    
    Exit Sub
Salah:
    MsgBox "Maaf, ada kesalahan! " & err.Description, vbOKOnly + vbExclamation, "Informasi"
End Sub

Private Sub Command1_Click()
    b_exportExcel = True
    CmdViewCPAElektornik_Click
    b_exportExcel = False
End Sub

'Private Sub Command2_Click()
'    On Error GoTo hell
'    b_exportSID = True
'
'    CD_save.Filter = "Excel Files |*.xls"
'    CD_save.ShowSave
'
'    sLokasiExcel = CD_save.FileName
'    Call CmdViewCPAElektornik_Click
'hell:
'    b_exportSID = False
'End Sub

Private Sub Command2_Click()
    Call ExportCoint
End Sub

Private Sub ExportCoint()
    Dim sQuery As String
    Dim RsCoint As ADODB.Recordset
    Dim W, K, S As Integer
    Dim CustId As String
    Dim ExlObj As Excel.Application
    
    CustId = ""
    
    If LvAccLunas.ListItems.Count = 0 Then
        MsgBox "Maaf, data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    S = 0
    For K = 1 To LvAccLunas.ListItems.Count
        If LvAccLunas.ListItems(K).Checked = True Then
            S = S + 1
            Exit For
        End If
    Next K
    
    If S = 0 Then
        MsgBox "Maaf anda belum memilih data yang akan di lihat!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
        
    FrameFilter.Enabled = False
    FrameList.Enabled = False
    'Ambil nilai custid
    For W = 1 To LvAccLunas.ListItems.Count
        If LvAccLunas.ListItems(W).Checked = True Then
            If CustId = "" Then
                CustId = "'" & CStr(LvAccLunas.ListItems(W).Text) & "'"
            Else
                CustId = CustId & ",'" & CStr(LvAccLunas.ListItems(W).Text) & "'"
            End If
        End If
    Next W
    
    sQuery = "SELECT region, now() as proposal_date, Ref_Number, ndiscountamt, produk,  customer_id,"
    sQuery = sQuery + " name, opendate, 'RIT' as area, nbalance, jml_payment, nttlpayment,"
    sQuery = sQuery + " nprincipal, vosprincipal, vosbalance , 'FINANCIAL PROBLEM' AS justification,"
    sQuery = sQuery + " dpropsal,aoc, kode_deskcol, nama_deskcol, nperiod FROM ("
    sQuery = sQuery + " SELECT region, now() as proposal_date, Ref_Number, ndiscountamt, produk,  customer_id, name, opendate, "
    sQuery = sQuery + " 'RIT' as area, nbalance, jml_payment, nttlpayment,  nprincipal, vosprincipal, vosbalance ,"
    sQuery = sQuery + " 'FINANCIAL PROBLEM' AS justification, dpropsal, descol, nperiod FROM ("
    sQuery = sQuery + " SELECT region, now() as proposal_date, Ref_Number, ndiscountamt, produk, "
    sQuery = sQuery + " customer_id, name, opendate, 'RIT' as area, nbalance, jml_payment, nttlpayment, "
    sQuery = sQuery + " nprincipal, vosprincipal, vosbalance , 'FINANCIAL PROBLEM' AS justification , dpropsal, descol, nperiod"
    sQuery = sQuery + " FROM ("
    sQuery = sQuery + " SELECT region, now() as proposal_date, mgm.acc_type as produk, "
    sQuery = sQuery + " mgm.custid as customer_id, name, opendate, 'RIT' as area, tblcpa.nbalance, "
    sQuery = sQuery + " tblcpa.nttlpayment, tblcpa.nprincipal, tblcpa.ndiscountamt,"
    sQuery = sQuery + " CASE WHEN  tblcpa.ndiscountamt = 0 THEN 'X' "
    sQuery = sQuery + " WHEN  tblcpa.ndiscountamt > 0 THEN 'D'"
    sQuery = sQuery + " END AS Ref_Number, vosprincipal, vosbalance, dpropsal, mgm.agent as descol, nperiod  "
    sQuery = sQuery + " FROM mgm inner join tblcpa on tblcpa.vcustid = mgm.custid ) AS a1"
    sQuery = sQuery + " INNER JOIN ( "
    sQuery = sQuery + " SELECT * FROM temp_proses_lunas ) AS a2"
    sQuery = sQuery + " ON a1.customer_id=a2.custid WHERE customer_id in(" & CustId & ") ) as query1"
    sQuery = sQuery + " INNER JOIN ( "
    sQuery = sQuery + " SELECT vcustid, max(dpropsal)AS maxdrpopsal FROM tblcpa WHERE vcustid in(" & CustId & ")"
    sQuery = sQuery + " group by vcustid ) as query2 ON query1.dpropsal = query2.maxdrpopsal) as queryjoin1"
    sQuery = sQuery + " LEFT JOIN"
    sQuery = sQuery + " (SELECT kode_deskcol, nama_deskcol, aoc FROM tbl_data_karyawan) as queryjoin2"
    sQuery = sQuery + " ON queryjoin1.descol = queryjoin2.kode_deskcol"
    Set RsCoint = New ADODB.Recordset
    RsCoint.CursorLocation = adUseClient
    RsCoint.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Set ExlObj = CreateObject("excel.application")
    ExlObj.Workbooks.ADD
    ExlObj.Visible = True
    
    With ExlObj.ActiveSheet
        .Cells(1, 1).Value = "NO"
        .Cells(1, 2).Value = "AOC"
        .Cells(1, 3).Value = "NAMA DESK COLLECTION"
        .Cells(1, 4).Value = "REGION"
        .Cells(1, 5).Value = "PROPOSAL DATE"
        .Cells(1, 6).Value = "REF. No."
        .Cells(1, 7).Value = "PRODUCT"
        .Cells(1, 8).Value = "CARD / Acc. No."
        .Cells(1, 9).Value = "CUSTOMER NAME"
        .Cells(1, 10).Value = "CARD OPEN DATE"
        .Cells(1, 11).Value = "AREA"
        .Cells(1, 12).Value = "OUTSTANDING BALANCE"
        .Cells(1, 13).Value = "TOTAL PAYMENT"
        .Cells(1, 14).Value = "PRINCIPAL"
        .Cells(1, 15).Value = "DISCOUNT AMOUNT"
        .Cells(1, 16).Value = "TENOR"
        .Cells(1, 17).Value = "+/- FROM O/S %"
        .Cells(1, 18).Value = "+/- FROM PRINCIPAL %"
        .Cells(1, 19).Value = "JUSTIFICATION"
        .Cells(1, 20).Value = "PAY DATE"
        .Cells(1, 21).Value = "APPROVE BY"
        
        
        iRow = 2
        If RsCoint.RecordCount > 0 Then
            i = 0
            Do Until RsCoint.EOF
                i = i + 1
                iRow = iRow + 1
                
                .Cells(iRow, 1).Value = i
                .Cells(iRow, 2).Value = IIf(IsNull(RsCoint!aoc), "", RsCoint!aoc)
                .Cells(iRow, 3).Value = IIf(IsNull(RsCoint!nama_deskcol), "", RsCoint!nama_deskcol)
                .Cells(iRow, 4).Value = IIf(IsNull(RsCoint!region), "", RsCoint!region)
                .Cells(iRow, 5).Value = Format(IIf(IsNull(RsCoint!proposal_date), "", RsCoint!proposal_date), "DD-MMM-YYYY")
                .Cells(iRow, 6).Value = IIf(IsNull(RsCoint!Ref_Number), "", RsCoint!Ref_Number)
                .Cells(iRow, 7).Value = IIf(IsNull(RsCoint!produk), "", RsCoint!produk)
                .Cells(iRow, 8).Value = IIf(IsNull(RsCoint!customer_id), "", RsCoint!customer_id)
                .Cells(iRow, 9).Value = IIf(IsNull(RsCoint!Name), "", RsCoint!Name)
                .Cells(iRow, 10).Value = Format(IIf(IsNull(RsCoint!opendate), "", RsCoint!opendate), "DD-MMM-YYYY")
                .Cells(iRow, 11).Value = IIf(IsNull(RsCoint!area), "", RsCoint!area)
                .Cells(iRow, 12).Value = IIf(IsNull(RsCoint!nbalance), "", RsCoint!nbalance)
                .Cells(iRow, 13).Value = IIf(IsNull(RsCoint!jml_payment), "", RsCoint!jml_payment)
                .Cells(iRow, 14).Value = IIf(IsNull(RsCoint!nprincipal), "", RsCoint!nprincipal)
                .Cells(iRow, 15).Value = IIf(IsNull(RsCoint!ndiscountamt), "", RsCoint!ndiscountamt)
                .Cells(iRow, 16).Value = IIf(IsNull(RsCoint!nperiod), "1", RsCoint!nperiod)
                .Cells(iRow, 17).Value = IIf(IsNull(RsCoint!vosbalance), "", RsCoint!vosbalance)
                .Cells(iRow, 18).Value = IIf(IsNull(RsCoint!vosprincipal), "", RsCoint!vosprincipal)
                .Cells(iRow, 19).Value = IIf(IsNull(RsCoint!justification), "", RsCoint!justification)
       
                RsCoint.MoveNext
            Loop
        End If
    
        'OTOMATISASI CELL
        For iColom = 1 To 16
            ExlObj.Cells(2, iColom).EntireColumn.AutoFit
        Next
        
        MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"

        FrameList.Enabled = True
        Set ExlObj = Nothing
        Set RsCoint = Nothing
    End With
End Sub

Private Sub Command3_Click()
    If List1.ListCount > 0 Then
        bCari_Bylist_cust = True
        Call IsiLunas
    End If
    bCari_Bylist_cust = False
End Sub

Private Sub Command4_Click()
    If List1.ListCount > 0 Then
        List1.RemoveItem List1.ListIndex
    End If
End Sub

Private Sub Form_Load()
    Call HeaderLunas
    Call IsiTipeKartu
    Call IsiJenisPTP
    Call IsiStatusLunas
    
    CekPayment = "0"
    'MsgBox "Mohon tunggu sebentar! sistem akan melakukan kalkulasi payment, setelah anda menekan tombol OK dari pesan ini :)", vbOKOnly + vbInformation, "Informasi"
    
    'MousePointer = vbHourglass
    'CmdRefresh_Click
    'MousePointer = vbNormal
    
    DTPicker1(0).Value = Now - 1
    DTPicker1(1).Value = Now - 1
    b_exportExcel = False
    b_exportSID = False
    
    LvAccLunas.ListItems.CLEAR
    'CmdRefresh_Click
End Sub

Private Sub SHOW_PRN()
    If b_exportExcel = False Then
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
    Else
        b_exportExcel = False
        RPT.RetrieveDataFiles
        RPT.Destination = crptToFile
        RPT.PrintFileType = crptExcel50
        RPT.Action = 1
    End If
End Sub

Private Sub LvAccLunas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvAccLunas.SortKey = ColumnHeader.Index - 1
    LvAccLunas.Sorted = True
End Sub

Private Sub export_data(M_objrs As ADODB.Recordset)
    On Error GoTo Salah
'    Dim M_objrs         As ADODB.Recordset
    Dim CMDSQL          As String
    Dim ListItem        As ListItem
    Dim cmdsql_update   As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i               As Integer
    Dim m_msgbox        As String
    
    i = 1
    
'    Set M_objrs = New ADODB.Recordset
'    M_objrs.CursorLocation = adUseClient
'    M_objrs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
       
Form_Save:
    Cd_save.ShowSave
    TxtPath.Text = Cd_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.Text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Download dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo Form_Save        '-> maka goto form_save
        End If
    End If
    
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
    On Error GoTo Salah
    'Proses pengsisian nama field ke excel
    Dim X, Y    As Integer
    If M_objrs.State = 1 Then
        X = 0
        Y = M_objrs.fields().Count - 1
        Do Until X > Y
            DoEvents
            objSheet.Cells(1, i).Value = CStr(M_objrs.fields(X).Name)
            i = i + 1
            X = X + 1
        Loop
    End If
        
    ' COLUMN HEADER
    objSheet.Cells(1, 1).Value = "Request Info" 'MERGE
    '-------------------------------------------------
    objSheet.Cells(2, 1).Value = "No"
    objSheet.Cells(2, 2).Value = "Region"
    objSheet.Cells(2, 3).Value = "Proposal Date"
    objSheet.Cells(2, 4).Value = "Product Card"
    objSheet.Cells(2, 5).Value = "Arrangement"
    
    objSheet.Cells(1, 6).Value = "Account Overview" 'MERGE
    '-----------------------------------------------------
    objSheet.Cells(2, 6).Value = "Card No"
    objSheet.Cells(2, 7).Value = "Customer Name"
    objSheet.Cells(2, 8).Value = "Card Open Date"
    objSheet.Cells(2, 9).Value = "Bucket"
    objSheet.Cells(2, 10).Value = "Account Status"
    objSheet.Cells(2, 11).Value = "Straight Flow"
    objSheet.Cells(2, 12).Value = "Responsinble Collector"
    objSheet.Cells(2, 13).Value = "Area"
    objSheet.Cells(2, 14).Value = "Agency Name"
    
    objSheet.Cells(1, 15).Value = "Payment Arrangement" 'MERGE
    '--------------------------------------------------------
    objSheet.Cells(2, 15).Value = "Outs. Balance"
    objSheet.Cells(2, 16).Value = "Total Payment"
    objSheet.Cells(2, 17).Value = "Down Payment"
    objSheet.Cells(2, 18).Value = "Future Payment"
    objSheet.Cells(2, 19).Value = "Payment Period"
    objSheet.Cells(2, 20).Value = "Principal"
    
    objSheet.Cells(1, 21).Value = "Calculation" 'MERGE
    '------------------------------------------------
    objSheet.Cells(2, 21).Value = "Charges"
    objSheet.Cells(2, 22).Value = "Discount Amount"
    objSheet.Cells(2, 23).Value = "+/- From O/S Balance (%)"
    objSheet.Cells(2, 24).Value = "+/- From O/S principal (%)"
    
    objSheet.Cells(1, 25).Value = "Background" 'MERGE
    '-----------------------------------------------
    objSheet.Cells(2, 25).Value = "Occupation"
    objSheet.Cells(2, 26).Value = "Reason"
    objSheet.Cells(2, 27).Value = "No. Of other delinguent debt"
    objSheet.Cells(2, 28).Value = "Payment Handle By"
    
    objSheet.Cells(1, 29).Value = "Others" 'MERGE
    '-------------------------------------------
    objSheet.Cells(2, 29).Value = "Mapping Accounts"
    objSheet.Cells(2, 30).Value = "Justifications"
    
    objSheet.Cells(1, 1).Value = "Approval" 'MERGE
    '---------------------------------------------
    objSheet.Cells(2, 25).Value = "Exception Level"
    objSheet.Cells(2, 26).Value = "Authority"
    objSheet.Cells(2, 27).Value = "Approver Name"
    
    objSheet.Cells(1, 28).Value = "Collection Support" 'MERGE
    '-------------------------------------------------------
    objSheet.Cells(2, 28).Value = "Balance to be written of"
    
    objSheet.Cells(1, 1).Value = "Surat Lunas" 'MERGE
    '------------------------------------------------
    objSheet.Cells(2, 25).Value = "Tanggal Pelunasan"
    objSheet.Cells(2, 26).Value = "Tempat Pengambilan Surat"
    
    objSheet.Cells(1, 1).Value = "DPA" 'MERGE
    '----------------------------------------
    
    objSheet.Range("A3").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A3
    objBook.SaveAs TxtPath.Text, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
        
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
     
Salah:
    Exit Sub
End Sub

Private Function update_status_acc(custid_x As String, Sisa_payment_x As Double, jenis_ptp As String, Optional batas_akhir_bayar_x As Date, Optional lpd As Date) As String
    Dim set_status_acc As String
    Dim waktu_sekarang As Date

    Set m_objrs_waktu = New ADODB.Recordset
    m_objrs_waktu.CursorLocation = adUseClient
    m_objrs_waktu.Open "SELECT now() as waktu_server ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    waktu_sekarang = Format(m_objrs_waktu!waktu_server, "yyyy-mm-dd")
    If custid_x = "4544931104468120" Then MsgBox "4544931104468120 - cek batal"""
    'If custid_x = "5184940104172675" Then MsgBox "ok"
    
    If Sisa_payment_x <= 0 Then
        'set_status_acc = "LUNAS"
        If UCase(jenis_ptp) = "PTP DISCOUNT" Then
            ' CEK TANGGAL PAYMENT DAN TENOR
            'If m_objrs_waktu.state = 1 Then m_objrs_waktu.Close
            'm_objrs_waktu.Open "SELECT max(paydate) as Tgl_akhirbayar FROM tbllunas a inner join mgm b on a.custid=b.custid WHERE a.custid='" & custid_x & "' AND date(a.Paydate)+1  > b.tglsource "
            'If Right(custid_x, 6) = "147123" Then MsgBox "OK"
            If Format(lpd, "yyyy-mm-dd") > Format(batas_akhir_bayar_x, "yyyy-mm-dd") Then
                set_status_acc = "BATAL"
            Else
                set_status_acc = "LUNAS"
            End If
        Else
            set_status_acc = "LUNAS"
        End If
        
    Else
        If UCase(jenis_ptp) = "PTP DISCOUNT" Then
            If waktu_sekarang > Format(batas_akhir_bayar_x, "yyyy-mm-dd") Then
                set_status_acc = "BATAL"
            Else
                set_status_acc = "BELUM LUNAS"
            End If
        Else
            set_status_acc = "BELUM LUNAS"
        End If
    End If
    
    CMDSQL = "UPDATE mgm SET status_lunas='" & set_status_acc & "' WHERE custid='"
    CMDSQL = CMDSQL & CStr(custid_x) & "'"
    M_OBJCONN.Execute CMDSQL
    
    update_status_acc = set_status_acc
End Function

