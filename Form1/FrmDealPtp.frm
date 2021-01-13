VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form FrmDealPtp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfirmasi Pembayaran Awal"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3960
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2460
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin TDBNumber6Ctl.TDBNumber txtPembayaranAwal 
      Height          =   255
      Left            =   1620
      TabIndex        =   1
      Top             =   2040
      Width           =   2190
      _Version        =   65536
      _ExtentX        =   3863
      _ExtentY        =   450
      Calculator      =   "FrmDealPtp.frx":0000
      Caption         =   "FrmDealPtp.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDealPtp.frx":008C
      Keys            =   "FrmDealPtp.frx":00AA
      Spin            =   "FrmDealPtp.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648384
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
   Begin VB.Label LblInformasi 
      Caption         =   "LblInformasi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   3675
   End
   Begin VB.Label Label1 
      Caption         =   "Pembayaran Awal:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   2040
      Width           =   1395
   End
End
Attribute VB_Name = "FrmDealPtp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PaymentTenor As Double

Private Sub cmdOK_Click()
    Dim CMDSQL As String
    Dim i As Integer
    
    bcekptp = True
    With FrmCC_Colection
        'Cek dulu, payment awal tidak boleh 0
        If txtPembayaranAwal.Value = 0 Then
            MsgBox "Pembayaran awal tidak boleh = 0!", vbOKOnly + vbExclamation, "Peringatan"
            Exit Sub
        End If
        'Cek dulu, payment awal tidak boleh lebih kecil dari 0
        If txtPembayaranAwal.Value < 0 Then
            MsgBox "Pembayaran awal tidak boleh < 0!", vbOKOnly + vbExclamation, "Peringatan"
            Exit Sub
        End If
        
        'Cek dulu, payment awal tidak boleh lebih besar dari total payment
        If txtPembayaranAwal.Value >= .txtPayment.Value Then
            MsgBox "Pembayaran awal tidak boleh lebih besar atau sama dengan dari total payment!", vbOKOnly + vbExclamation, "Peringatan"
            Exit Sub
        End If
        
    
        Dim M_Objrs_Cek_Tgl As ADODB.Recordset
        If .Chktenor.Value = 0 Then
                  
                '@@14-04-2012 Cek Data
                CMDSQL = "select * from mandiri.tblnegoptp where custid='"
                CMDSQL = CMDSQL + .lblcustid.Caption + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(.TDBDate3.Value, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from mandiri.tblnegoptp where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                  
            jatuhtempo = Format(.TDBDate3.Value, "yyyy-mm-dd")
            CMDSQL = "INSERT INTO mandiri.TblNegoPTP "
            CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + .lblcustid.Caption + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(txtPembayaranAwal.Value) + " , "
            'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'IPO')"
            M_OBJCONN.Execute CMDSQL
            
            
                '@@14-04-2012 Cek Data
                CMDSQL = "select * from mandiri.tblnegoptp_log where custid='"
                CMDSQL = CMDSQL + .lblcustid.Caption + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(.TDBDate3.Value, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from mandiri.tblnegoptp_log where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
            
            
            ' isi ke tbl log_ptp
            CMDSQL = "INSERT INTO mandiri.tblnegoptp_log "
            CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + .lblcustid.Caption + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(txtPembayaranAwal.Value) + " , "
            'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'" + .lblaoc.Caption + "','P')"
            M_OBJCONN.Execute CMDSQL
            
            Set ListItem = .LstPayment.ListItems.ADD(, , "")
                ListItem.SubItems(1) = ""
                ListItem.SubItems(2) = Format(.TDBDate3.Value, "dd/mm/yyyy")
                ListItem.SubItems(3) = CStr(txtPembayaranAwal.Value)
                ListItem.SubItems(4) = "IPO"
                ListItem.SubItems(5) = MDIForm1.TDBDate1.Value
                
        Else
                        
            jatuhtempo = Format(.TDBDate3.Value, "yyyy-mm-dd")
            
                '@@14-04-2012 Cek Data
                CMDSQL = "select * from mandiri.tblnegoptp where custid='"
                CMDSQL = CMDSQL + .lblcustid.Caption + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from mandiri.tblnegoptp where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
            
            CMDSQL = "INSERT INTO mandiri.TblNegoPTP "
            CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + .lblcustid.Caption + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(txtPembayaranAwal.Value) + " , "
            'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'IPO')"
            M_OBJCONN.Execute CMDSQL
            
                '@@14-04-2012 Cek Data
                CMDSQL = "select * from mandiri.tblnegoptp_log where custid='"
                CMDSQL = CMDSQL + .lblcustid.Caption + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(.TDBDate3.Value, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from mandiri.tblnegoptp_log where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
            
            
            'isi ke tbl log_ptp
            CMDSQL = "INSERT INTO mandiri.tblnegoptp_log "
            CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + .lblcustid.Caption + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(txtPembayaranAwal.Value) + " , "
            'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'" + .lblaoc.Caption + "','P')"
            M_OBJCONN.Execute CMDSQL
                
            Set ListItem = .LstPayment.ListItems.ADD(, , "")
                ListItem.SubItems(1) = ""
                ListItem.SubItems(2) = Format(.TDBDate3.Value, "dd/mm/yyyy")
                ListItem.SubItems(3) = CStr(txtPembayaranAwal.Value)
                ListItem.SubItems(4) = "IPO"
                ListItem.SubItems(5) = MDIForm1.TDBDate1.Value
                
        
    
            n = 0
            
            HitungInstallmentPtp
            
            For i = 1 To Val(.txttenor - 1)
                    n = n + 1
                    'JMLPAY = ((.TxtPayment - txtPembayaranAwal.Value) - PaymentTenor) / (.txttenor.Value - 1)
                    JmlPay = PaymentTenor
                    Vrdate = DateAdd("m", n, Format(.TDBDate3.Value, "yyyy-mm-dd"))
                    
                '@@14-04-2012 Cek Data
                CMDSQL = "select * from mandiri.tblreserve where custid='"
                CMDSQL = CMDSQL + .lblcustid.Caption + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(Vrdate, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from mandiri.tblreserve where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                    
                    CMDSQL = "INSERT INTO mandiri.tblreserve "
                    CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
                    CMDSQL = CMDSQL + "VALUES "
                    CMDSQL = CMDSQL + "('" + .lblcustid.Caption + "', "
                    CMDSQL = CMDSQL + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
                    CMDSQL = CMDSQL + "" + CStr(JmlPay) + " , "
                    'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
                    CMDSQL = CMDSQL + "now(), "
                    CMDSQL = CMDSQL + "'IPO')"
                    M_OBJCONN.Execute CMDSQL
                    
                    '@@14-04-2012 Cek Data
                CMDSQL = "select * from mandiri.tblnegoptp_log where custid='"
                CMDSQL = CMDSQL + .lblcustid.Caption + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(Vrdate, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from mandiri.tblnegoptp_log where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                    
                    
                    CMDSQL = "INSERT INTO mandiri.TblNegoptp_log "
                    CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglinput,agent,stsacc) "
                    CMDSQL = CMDSQL + "VALUES "
                    CMDSQL = CMDSQL + "('" + .lblcustid.Caption + "', "
                    CMDSQL = CMDSQL + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
                    CMDSQL = CMDSQL + "" + CStr(JmlPay) + " , "
                    'CMDSQL = CMDSQL + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
                    CMDSQL = CMDSQL + "now(), "
                    CMDSQL = CMDSQL + "'" + .lblaoc.Caption + "','R')"
                    M_OBJCONN.Execute CMDSQL
        
                Set ListItem = .LstReserve.ListItems.ADD(, , "")
                    ListItem.SubItems(1) = ""
                    ListItem.SubItems(2) = Format(Vrdate, "dd/mm/yyyy")
                    ListItem.SubItems(3) = JmlPay
                    ListItem.SubItems(4) = "IPO"
                    ListItem.SubItems(5) = MDIForm1.TDBDate1.Value
            Next i
       End If
    End With
    
    PaymentTenor = 0
    
    MsgBox "PTP berhasil ditambahkan!", vbOKOnly + vbInformation, "Informasi"
    Unload Me
End Sub

Private Sub Form_Load()
    Dim TeksInfo As String
    
    With FrmCC_Colection
        TeksInfo = "Anda melakukan PTP sebesar : "
        TeksInfo = TeksInfo + .txtPayment.Text + vbCrLf
        TeksInfo = TeksInfo + "Tenor: " + .txttenor.Text + " kali " + vbCrLf
        TeksInfo = TeksInfo + "Anda dapat mengeset nilai pembayaran pertama pada kolom di bawah, " + vbCrLf
        TeksInfo = TeksInfo + "Selanjutnya program akan menghitung PTP di bulan selanjutnya (reserved PTP) " + vbCrLf
        TeksInfo = TeksInfo + "secara otomatis sesuai dengan besarnya amount PTP dan Tenor."
    
    
    LblInformasi.Caption = TeksInfo
    On Error GoTo Salah
    txtPembayaranAwal.Value = .txtPayment.Value / .txttenor.Value   '.Tdabamoint.Value
    End With
    Exit Sub
Salah:
    MsgBox "Ada error: " & err.Description
End Sub


'@@22-09-2011 Hitung InstallmentPtp
Private Sub HitungInstallmentPtp()
    Dim installment As Double
    
    With FrmCC_Colection
        If .txttenor.Value = 0 Then
            installment = .txtPayment.Value / 1
        Else
            installment = (.txtPayment.Value - txtPembayaranAwal.Value) / (.txttenor.Value - 1)
        End If
        PaymentTenor = installment
    End With
End Sub
