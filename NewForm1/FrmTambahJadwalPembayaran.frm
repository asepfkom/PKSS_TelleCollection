VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form FrmTambahJadwalPembayaran 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tambah Jadwal Pembayaran"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2100
      TabIndex        =   15
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   2760
      Width           =   1155
   End
   Begin VB.CommandButton CmdTambah 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   2760
      Width           =   1155
   End
   Begin TDBNumber6Ctl.TDBNumber TxtTunggakan 
      Height          =   315
      Left            =   2100
      TabIndex        =   3
      Top             =   780
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   556
      Calculator      =   "FrmTambahJadwalPembayaran.frx":0000
      Caption         =   "FrmTambahJadwalPembayaran.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmTambahJadwalPembayaran.frx":008C
      Keys            =   "FrmTambahJadwalPembayaran.frx":00AA
      Spin            =   "FrmTambahJadwalPembayaran.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
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
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
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
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBDate6Ctl.TDBDate TxtTglBayar 
      Height          =   285
      Left            =   2100
      TabIndex        =   1
      Top             =   420
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   494
      Calendar        =   "FrmTambahJadwalPembayaran.frx":011C
      Caption         =   "FrmTambahJadwalPembayaran.frx":0234
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmTambahJadwalPembayaran.frx":02A0
      Keys            =   "FrmTambahJadwalPembayaran.frx":02BE
      Spin            =   "FrmTambahJadwalPembayaran.frx":031C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
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
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   6815745
      Value           =   39876
      CenturyMode     =   0
   End
   Begin TDBNumber6Ctl.TDBNumber TxtPembayaran 
      Height          =   315
      Left            =   2100
      TabIndex        =   5
      Top             =   1140
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   556
      Calculator      =   "FrmTambahJadwalPembayaran.frx":0344
      Caption         =   "FrmTambahJadwalPembayaran.frx":0364
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmTambahJadwalPembayaran.frx":03D0
      Keys            =   "FrmTambahJadwalPembayaran.frx":03EE
      Spin            =   "FrmTambahJadwalPembayaran.frx":0438
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
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
      MaxValue        =   999999999
      MinValue        =   -9999999999
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
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber TxtHutangPokok 
      Height          =   315
      Left            =   2100
      TabIndex        =   7
      Top             =   1500
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   556
      Calculator      =   "FrmTambahJadwalPembayaran.frx":0460
      Caption         =   "FrmTambahJadwalPembayaran.frx":0480
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmTambahJadwalPembayaran.frx":04EC
      Keys            =   "FrmTambahJadwalPembayaran.frx":050A
      Spin            =   "FrmTambahJadwalPembayaran.frx":0554
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
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
      MaxValue        =   999999999
      MinValue        =   -999999999
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
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber TxtInterest 
      Height          =   315
      Left            =   2100
      TabIndex        =   9
      Top             =   1860
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   556
      Calculator      =   "FrmTambahJadwalPembayaran.frx":057C
      Caption         =   "FrmTambahJadwalPembayaran.frx":059C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmTambahJadwalPembayaran.frx":0608
      Keys            =   "FrmTambahJadwalPembayaran.frx":0626
      Spin            =   "FrmTambahJadwalPembayaran.frx":0670
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
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
      MaxValue        =   99999999
      MinValue        =   -999999999
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
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtBalanceTerakhir 
      Height          =   315
      Left            =   2100
      TabIndex        =   11
      Top             =   2220
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   556
      Calculator      =   "FrmTambahJadwalPembayaran.frx":0698
      Caption         =   "FrmTambahJadwalPembayaran.frx":06B8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmTambahJadwalPembayaran.frx":0724
      Keys            =   "FrmTambahJadwalPembayaran.frx":0742
      Spin            =   "FrmTambahJadwalPembayaran.frx":078C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
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
      MaxValue        =   999999999
      MinValue        =   -9999999999
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
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label7 
      Caption         =   "Id Data: "
      Height          =   255
      Left            =   180
      TabIndex        =   14
      Top             =   60
      Width           =   1875
   End
   Begin VB.Label Label6 
      Caption         =   "Balance Terakhir:"
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   2220
      Width           =   1875
   End
   Begin VB.Label Label5 
      Caption         =   "Interest:"
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   1860
      Width           =   1875
   End
   Begin VB.Label Label4 
      Caption         =   "Hutang Pokok:"
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   1500
      Width           =   1875
   End
   Begin VB.Label Label3 
      Caption         =   "Pembayaran:"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1140
      Width           =   1875
   End
   Begin VB.Label Label2 
      Caption         =   "Tunggakan:"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   780
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal Pembayaran:"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   420
      Width           =   1875
   End
End
Attribute VB_Name = "FrmTambahJadwalPembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdTambah_Click()
    Dim CMDSQL As String
    Dim Tahun As String
    Dim Bulan As String
    
    'Cek Data
    
    If txtpembayaran.Value = 0 Then
        MsgBox "Pembayaran tidak boleh 0!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtTglBayar.Value = Empty Or TxtTglBayar.ValueIsNull = True Then
        MsgBox "Tanggal bayar tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
  
    
    If Me.Caption = "EDIT" Then
        Tahun = Format(TxtTglBayar.Value, "yyyy")
        Bulan = Format(TxtTglBayar.Value, "dd-mmm")
        
        CMDSQL = "update tbljadwalpembayaran set tahun='"
        CMDSQL = CMDSQL + Tahun + "',bulan='"
        CMDSQL = CMDSQL + Bulan + "',tunggakan='"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(TxtTunggakan.Value), "0", TxtTunggakan.Value)) + "',pembayaran='"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(txtpembayaran.Value), "0", txtpembayaran.Value)) + "',hutang_pokok='"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(TxtHutangPokok.Value), "0", TxtHutangPokok.Value)) + "',interest='"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(TxtInterest.Value), "0", TxtInterest.Value)) + "',balance_terakhir='"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(txtBalanceTerakhir.Value), "0", txtBalanceTerakhir.Value)) + "',tglbayar='"
        CMDSQL = CMDSQL + CStr(Format(TxtTglBayar.Value, "yyyy-mm-dd")) + "' where id='"
        CMDSQL = CMDSQL + Trim(txtId.Text) + "'"
        M_OBJCONN.Execute CMDSQL
        MsgBox "Data berhasil diupdate!", vbOKOnly + vbInformation, "Informasi"
    Else
        If TxtTunggakan.Value = 0 Then
            MsgBox "CH sudah tidak ada tunggakan!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        
        If txtBalanceTerakhir.Value < 0 Then
            MsgBox "Balance terakhir tidak boleh minus!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        
        Tahun = Format(TxtTglBayar.Value, "yyyy")
        Bulan = Format(TxtTglBayar.Value, "dd-mmm")
    
        CMDSQL = "insert into tbljadwalpembayaran (nid,tahun,bulan,tunggakan,"
        CMDSQL = CMDSQL + "pembayaran,hutang_pokok,interest,balance_terakhir,tglbayar) values ('"
        CMDSQL = CMDSQL + FrmJadwalPembayaranCpa.TxtIdCpa.Text + "','"
        CMDSQL = CMDSQL + Tahun + "','"
        CMDSQL = CMDSQL + Bulan + "','"
        CMDSQL = CMDSQL + CStr(TxtTunggakan.Value) + "','"
        CMDSQL = CMDSQL + CStr(txtpembayaran.Value) + "','"
        CMDSQL = CMDSQL + CStr(TxtHutangPokok.Value) + "','"
        CMDSQL = CMDSQL + CStr(TxtInterest.Value) + "','"
        CMDSQL = CMDSQL + CStr(txtBalanceTerakhir.Value) + "','"
        CMDSQL = CMDSQL + CStr(Format(TxtTglBayar.Value, "yyyy-mm-dd")) + "')"
        M_OBJCONN.Execute CMDSQL
        
        MsgBox "Data berhasil disimpan!", vbOKOnly + vbInformation, "Informasi"
    End If
    
  
    Unload Me
    
End Sub



Private Sub BalanceTerakhir()
    txtBalanceTerakhir.Value = TxtTunggakan.Value - txtpembayaran.Value
End Sub



Private Sub TxtPembayaran_Change()
    BalanceTerakhir
End Sub
