VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBersihkanNegoPTP 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Bersihkan PTP"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   5145
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdHapusReservedGanda 
      Caption         =   "&Hapus reserved ganda"
      Height          =   495
      Left            =   180
      TabIndex        =   5
      Top             =   2640
      Width           =   4875
   End
   Begin VB.CommandButton CmdHapusNegoPTPLog 
      Caption         =   "Hapus Negoptp_log yang ganda"
      Height          =   495
      Left            =   180
      TabIndex        =   4
      Top             =   600
      Width           =   4875
   End
   Begin VB.CommandButton CmdInputNegoptpFromNegoptplog 
      Caption         =   "&Input negoptp dari tblnegoptp_log"
      Height          =   495
      Left            =   180
      TabIndex        =   3
      Top             =   1320
      Width           =   4875
   End
   Begin VB.CommandButton CmdUpdateDatePTPMgm 
      Caption         =   "&Update dateptp sesuai data di tblnegoptp"
      Height          =   495
      Left            =   180
      TabIndex        =   2
      Top             =   1860
      Width           =   4875
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   3720
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdHapusNegoPTP 
      Caption         =   "&Hapus data ganda di tabel negoptp"
      Height          =   495
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   4875
   End
End
Attribute VB_Name = "FrmBersihkanNegoPTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdHapusNegoPTP_Click()
'    Dim Cmdsql As String
'    Dim M_Objrs_Custid As ADODB.Recordset
'    Dim W As Integer
'    Dim K As Integer
'    Dim M_Objrs_Hapus As ADODB.Recordset
'
'    '--1. Cari Custid yang datanya lebih dari 1 dan memiliki promisedate yang sama --
'    Cmdsql = "select custid,promisedate,count(custid) as jumlah from tblnegoptp "
'    Cmdsql = Cmdsql + " group by custid,promisedate having count(custid) > 1"
'    Set M_Objrs_Custid = New ADODB.Recordset
'    M_Objrs_Custid.CursorLocation = adUseClient
'    M_Objrs_Custid.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If M_Objrs_Custid.RecordCount = 0 Then
'        MsgBox "Tidak ada data di tabel tblnegoptp!", vbOKOnly + vbInformation, "Informasi"
'        Set M_Objrs_Custid = Nothing
'        Exit Sub
'    End If
'
'    PB1.Max = M_Objrs_Custid.RecordCount
'
'    While Not M_Objrs_Custid.EOF
'        PB1.Value = M_Objrs_Custid.Bookmark
'        W = M_Objrs_Custid("jumlah") - 1
'        '--2. Proses penghapusan dicari berdasarkan tanggal input yang terkecil
'        '---- Data yang disisakan adalah 1 data dengan tanggal input terakhir
'        Cmdsql = "select * from tblnegoptp where custid='"
'        Cmdsql = Cmdsql + M_Objrs_Custid("custid") + "' and date(promisedate)='"
'        Cmdsql = Cmdsql + Format(M_Objrs_Custid("promisedate"), "yyyy-mm-dd") + "' "
'        Cmdsql = Cmdsql + "order by inputdate asc "
'        Set M_Objrs_Hapus = New ADODB.Recordset
'        M_Objrs_Hapus.CursorLocation = adUseClient
'        M_Objrs_Hapus.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'        If M_Objrs_Hapus.RecordCount > 1 Then
'            For K = 1 To W
'                Cmdsql = "delete from tblnegoptp where id='"
'                Cmdsql = Cmdsql + CStr(M_Objrs_Hapus("id")) + "'"
'                M_OBJCONN.Execute Cmdsql
'            Next K
'            Set M_Objrs_Hapus = Nothing
'        Else
'            Set M_Objrs_Hapus = Nothing
'        End If
'
'        M_Objrs_Custid.MoveNext
'    Wend
    
    Call BersihkanNegoPTPGanda
End Sub

Private Sub BersihkanNegoPTPGanda()
    Dim Cmdsql_Custid As String
    Dim M_Objrs_Custid As ADODB.Recordset
    Dim W As Integer
    Dim K As Integer
    Dim M_Objrs_Tgl As ADODB.Recordset
    Dim CMDSQL As String
    Dim M_Objrs_Data As ADODB.Recordset
    Dim M_Objrs_Cek As ADODB.Recordset
    
    'Hapus yang memeliki promisepay =<0
    CMDSQL = "delete from mandiri.tblnegoptp where promisepay<=0"
    M_OBJCONN.Execute CMDSQL
    
    'Hapus Tabel tblnegoptp_temp
    CMDSQL = "delete from mandiri.tblnegoptp_temp "
    M_OBJCONN.Execute CMDSQL
    
    'Inputkan Data Negoptp yang memilik f_valid='1' ke tabel tblnegoptp_temp
    CMDSQL = "insert into mandiri.tblnegoptp_temp "
    CMDSQL = CMDSQL + "select * from mandiri.tblnegoptp where f_valid='1'"
    M_OBJCONN.Execute CMDSQL
    
cek_lagi:
    
    '--1. Cari Distinct Custid
    Cmdsql_Custid = "select distinct custid from mandiri.tblnegoptp group by custid having count(custid) > 1"
    Set M_Objrs_Custid = New ADODB.Recordset
    M_Objrs_Custid.CursorLocation = adUseClient
    M_Objrs_Custid.Open Cmdsql_Custid, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs_Custid.RecordCount = 0 Then
        MsgBox "Data Nego PTP tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Set M_Objrs_Custid = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_Objrs_Custid.RecordCount
    
    While Not M_Objrs_Custid.EOF
        
        PB1.Value = M_Objrs_Custid.Bookmark
        '--2. Cari Yang memiliki promisedate yang ganda
        CMDSQL = "select promisedate,count(promisedate) from mandiri.tblnegoptp "
        CMDSQL = CMDSQL + "where custid='"
        CMDSQL = CMDSQL + M_Objrs_Custid("custid") + "' group by promisedate "
        CMDSQL = CMDSQL + " having count(promisedate) > 1"
        Set M_Objrs_Tgl = New ADODB.Recordset
        M_Objrs_Tgl.CursorLocation = adUseClient
        M_Objrs_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs_Tgl.RecordCount > 0 Then
            While Not M_Objrs_Tgl.EOF
                '--3. Hapus Salah satu data yang promisedate nya ganda
                CMDSQL = "select * from mandiri.tblnegoptp where custid='"
                CMDSQL = CMDSQL + M_Objrs_Custid("custid") + "' and date(promisedate)='"
                CMDSQL = CMDSQL + Format(M_Objrs_Tgl("promisedate"), "yyyy-mm-dd") + "' "
                'Cmdsql = Cmdsql + "and f_valid is null "
                CMDSQL = CMDSQL + "order by inputdate asc"
                Set M_Objrs_Data = New ADODB.Recordset
                M_Objrs_Data.CursorLocation = adUseClient
                M_Objrs_Data.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                                
                
'                If M_Objrs_Data.RecordCount = 1 Then
'                     Cmdsql = "delete from tblnegoptp where id='"
'                     Cmdsql = Cmdsql + CStr(M_Objrs_Data("id")) + "'"
'                     M_OBJCONN.Execute Cmdsql
'                End If
                If M_Objrs_Data.RecordCount > 1 Then
                    For W = 1 To (M_Objrs_Data.RecordCount - 1)
                        CMDSQL = "delete from mandiri.tblnegoptp where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Data("id")) + "' "
                        M_OBJCONN.Execute CMDSQL
                    Next W
                End If
                
                Set M_Objrs_Data = Nothing
                M_Objrs_Tgl.MoveNext
            Wend
        End If
        Set M_Objrs_Tgl = Nothing
        
        
        M_Objrs_Custid.MoveNext
    Wend
    
    Set M_Objrs_Custid = Nothing
    
    'Cek data dulu
    CMDSQL = "select custid,promisedate,count(custid) as jumlah from mandiri.tblnegoptp  "
    CMDSQL = CMDSQL + "group by custid,promisedate having count(custid) > 1"
    Set M_Objrs_Cek = New ADODB.Recordset
    M_Objrs_Cek.CursorLocation = adUseClient
    M_Objrs_Cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs_Cek.RecordCount > 0 Then
        Set M_Objrs_Cek = Nothing
        GoTo cek_lagi
    End If
    
    'Update yang memliki data valid
    CMDSQL = "update mandiri.tblnegoptp set f_valid='1' "
    CMDSQL = CMDSQL + "from ( select tblnegoptp_temp.custid as custid_new,tblnegoptp_temp.custid as custid_new1,"
    CMDSQL = CMDSQL + " tblnegoptp_temp.promisedate as promisedate_new,"
    CMDSQL = CMDSQL + " tblnegoptp_temp.promisepay as promisepay_new,"
    CMDSQL = CMDSQL + " tblnegoptp_temp.inputdate as inputdate_new "
    CMDSQL = CMDSQL + " from mandiri.tblnegoptp,mandiri.tblnegoptp_temp "
    CMDSQL = CMDSQL + "  where  tblnegoptp.custid=tblnegoptp_temp.custid and "
    CMDSQL = CMDSQL + " tblnegoptp.promisedate=tblnegoptp_temp.promisedate and "
    CMDSQL = CMDSQL + "  tblnegoptp.promisepay=tblnegoptp_temp.promisepay and "
    CMDSQL = CMDSQL + " tblnegoptp.inputdate=tblnegoptp_temp.inputdate) as a"
    CMDSQL = CMDSQL + " where tblnegoptp.custid=a.custid_new1 and "
    CMDSQL = CMDSQL + " tblnegoptp.promisedate=a.promisedate_new and "
    CMDSQL = CMDSQL + " tblnegoptp.promisepay=a.promisepay_new and "
    CMDSQL = CMDSQL + " tblnegoptp.inputdate=a.inputdate_new  "
    M_OBJCONN.Execute CMDSQL
    
    'Bersihkan lagi data tblnegoptp_temp
    CMDSQL = "delete from mandiri.tblnegoptp_temp"
    M_OBJCONN.Execute CMDSQL
    MsgBox "Proses penghapusan data ganda negoptp selesai!", vbOKOnly + vbInformation, "Informasi"
    
End Sub


Private Sub UpdateTanggalPTPdenganNegoPTP()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    CMDSQL = "update mandiri.mgm set dateptp = tgl_bayar "
    CMDSQL = CMDSQL + " from (select b.custid_nego,mgm.custid,b.tgl_bayar "
    CMDSQL = CMDSQL + " from mandiri.mgm,(select custid as custid_nego,max(promisedate) as tgl_bayar from "
    CMDSQL = CMDSQL + " mandiri.tblnegoptp group by custid) as b "
    CMDSQL = CMDSQL + " where mgm.custid=b.custid_nego and mgm.f_cek_new like 'PTP-%' ) as c "
    CMDSQL = CMDSQL + " where mgm.custid=c.custid_nego and mgm.f_cek_new like 'PTP-%' "
    M_OBJCONN.Execute CMDSQL
    
    MsgBox "Proses update date ptp selesai!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub CmdHapusNegoPTPLog_Click()
    HapusNegoPTPLogGanda_NEW
End Sub

Private Sub CmdHapusReservedGanda_Click()
    HapusReservedGanda
End Sub

Private Sub CmdInputNegoptpFromNegoptplog_Click()
    InputNegoPTPFromNegoptpLog
End Sub

Private Sub CmdUpdateDatePTPMgm_Click()
    UpdateTanggalPTPdenganNegoPTP
End Sub

Private Sub InputNegoPTPFromNegoptpLog()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim M_Objrs_Data As ADODB.Recordset
    
    '1.-- Ambil data custid dan promisedate dari tabel negoptp_log berdasarkan custid
    'di tabel mgm yang status acc=PTP tetapi data PTPnya tidak ada di tabel tblbnegoptp
    CMDSQL = "select distinct custid,max(promisedate) as tgl  from mandiri.tblnegoptp_log where custid in ("
    CMDSQL = CMDSQL + " select custid from mandiri.mgm where custid not in "
    CMDSQL = CMDSQL + " (select custid_nego from  "
    CMDSQL = CMDSQL + " (select custid as custid_nego,max(promisedate) tanggal_bayar "
    CMDSQL = CMDSQL + " from mandiri.tblnegoptp group by custid) as b) and f_cek_new like 'PTP-%') "
    CMDSQL = CMDSQL + " and stsacc='P' "
    CMDSQL = CMDSQL + " group by custid "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    
    '2.-- Cari data terakhir berdasarkan inputdate
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        CMDSQL = "select * from mandiri.tblnegoptp_log where custid='"
        CMDSQL = CMDSQL + Trim(M_objrs("custid")) + "' and date(promisedate)='"
        CMDSQL = CMDSQL + CStr(Format(M_objrs("tgl"), "yyyy-mm-dd")) + "' "
        CMDSQL = CMDSQL + " and stsacc='P' and promisepay>0 order by tglinput desc limit 1"
        Set M_Objrs_Data = New ADODB.Recordset
        M_Objrs_Data.CursorLocation = adUseClient
        M_Objrs_Data.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Data.RecordCount > 0 Then
            'Inputkan ke data tblnegoptp
            CMDSQL = "insert into mandiri.tblnegoptp (custid,promisedate,promisepay,inputdate) values ('"
            CMDSQL = CMDSQL + M_Objrs_Data("custid") + "','"
            CMDSQL = CMDSQL + CStr(Format(M_Objrs_Data("promisedate"), "yyyy-mm-dd")) + "','"
            CMDSQL = CMDSQL + CStr(M_Objrs_Data("promisepay")) + "','"
            CMDSQL = CMDSQL + CStr(Format(M_Objrs_Data("tglinput"), "yyyy-mm-dd")) + "')"
            M_OBJCONN.Execute CMDSQL
        End If
        Set M_Objrs_Data = Nothing
        
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub HapusNegoPTPLogGanda_NEW()
    Dim CMDSQL As String
    Dim M_Objrs_Custid As ADODB.Recordset
    Dim W As Integer
    Dim K As Integer
    Dim M_Objrs_Hapus As ADODB.Recordset
    
     'Hapus yang memeliki promisepay =<0
    CMDSQL = "delete from mandiri.tblnegoptp_log where promisepay<=0"
    M_OBJCONN.Execute CMDSQL
    
cek_lagi:

    '--1. Cari Custid yang datanya lebih dari 1 dan memiliki promisedate yang sama --
    CMDSQL = "select custid,promisedate,count(custid) as jumlah from mandiri.tblnegoptp_log "
    CMDSQL = CMDSQL + " group by custid,promisedate having count(custid) > 1"
    Set M_Objrs_Custid = New ADODB.Recordset
    M_Objrs_Custid.CursorLocation = adUseClient
    M_Objrs_Custid.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    If M_Objrs_Custid.RecordCount = 0 Then
        MsgBox "Tidak ada data di tabel tblnegoptpLog!", vbOKOnly + vbInformation, "Informasi"
        Set M_Objrs_Custid = Nothing
        Exit Sub
    End If

    PB1.Max = M_Objrs_Custid.RecordCount

    While Not M_Objrs_Custid.EOF
        PB1.Value = M_Objrs_Custid.Bookmark
        W = M_Objrs_Custid("jumlah") - 1
        '--2. Proses penghapusan dicari berdasarkan tanggal input yang terkecil
        '---- Data yang disisakan adalah 1 data dengan tanggal input terakhir
        CMDSQL = "select * from mandiri.tblnegoptp_log where custid='"
        CMDSQL = CMDSQL + M_Objrs_Custid("custid") + "' and date(promisedate)='"
        CMDSQL = CMDSQL + Format(M_Objrs_Custid("promisedate"), "yyyy-mm-dd") + "' "
        CMDSQL = CMDSQL + "order by tglinput asc "
        Set M_Objrs_Hapus = New ADODB.Recordset
        M_Objrs_Hapus.CursorLocation = adUseClient
        M_Objrs_Hapus.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

        If M_Objrs_Hapus.RecordCount > 1 Then
            For K = 1 To W
                CMDSQL = "delete from mandiri.tblnegoptp_log where id='"
                CMDSQL = CMDSQL + CStr(M_Objrs_Hapus("id")) + "'"
                M_OBJCONN.Execute CMDSQL
            Next K
            Set M_Objrs_Hapus = Nothing
        Else
            Set M_Objrs_Hapus = Nothing
        End If

        M_Objrs_Custid.MoveNext
    Wend
    
    Set M_Objrs_Custid = Nothing
    
     'Cek data dulu
    CMDSQL = "select custid,promisedate,count(custid) as jumlah from mandiri.tblnegoptp_log  "
    CMDSQL = CMDSQL + "group by custid,promisedate having count(custid) > 1"
    Set M_Objrs_Cek = New ADODB.Recordset
    M_Objrs_Cek.CursorLocation = adUseClient
    M_Objrs_Cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs_Cek.RecordCount > 0 Then
        Set M_Objrs_Cek = Nothing
        GoTo cek_lagi
    End If
    
    MsgBox "Proses Hapus NegoPTPLOG Berhasil!", vbOKOnly + vbInformation, "Informasi"
    
End Sub


Private Sub HapusReservedGanda()
     Dim CMDSQL As String
    Dim M_Objrs_Custid As ADODB.Recordset
    Dim W As Integer
    Dim K As Integer
    Dim M_Objrs_Hapus As ADODB.Recordset
    
     'Hapus yang memiliki promisepay =<0
    CMDSQL = "delete from mandiri.tblreserve where stsmove='1'"
    M_OBJCONN.Execute CMDSQL
    
    
     'Hapus yang memiliki promisepay =<0
    CMDSQL = "delete from mandiri.tblreserve where promisepay<=0"
    M_OBJCONN.Execute CMDSQL
    
cek_lagi:

    '--1. Cari Custid yang datanya lebih dari 1 dan memiliki promisedate yang sama --
    CMDSQL = "select custid,promisedate,count(custid) as jumlah from mandiri.tblreserve where stsmove='0' "
    CMDSQL = CMDSQL + " group by custid,promisedate having count(custid) > 1"
    Set M_Objrs_Custid = New ADODB.Recordset
    M_Objrs_Custid.CursorLocation = adUseClient
    M_Objrs_Custid.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    If M_Objrs_Custid.RecordCount = 0 Then
        MsgBox "Tidak ada data di tabel tblreserve!", vbOKOnly + vbInformation, "Informasi"
        Set M_Objrs_Custid = Nothing
        Exit Sub
    End If

    PB1.Max = M_Objrs_Custid.RecordCount

    While Not M_Objrs_Custid.EOF
        PB1.Value = M_Objrs_Custid.Bookmark
        W = M_Objrs_Custid("jumlah") - 1
        '--2. Proses penghapusan dicari berdasarkan tanggal input yang terkecil
        '---- Data yang disisakan adalah 1 data dengan tanggal input terakhir
        CMDSQL = "select * from mandiri.tblreserve where custid='"
        CMDSQL = CMDSQL + M_Objrs_Custid("custid") + "' and date(promisedate)='"
        CMDSQL = CMDSQL + Format(M_Objrs_Custid("promisedate"), "yyyy-mm-dd") + "' "
        CMDSQL = CMDSQL + "order by inputdate asc "
        Set M_Objrs_Hapus = New ADODB.Recordset
        M_Objrs_Hapus.CursorLocation = adUseClient
        M_Objrs_Hapus.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

        If M_Objrs_Hapus.RecordCount > 1 Then
            For K = 1 To W
                CMDSQL = "delete from mandiri.tblreserve where id='"
                CMDSQL = CMDSQL + CStr(M_Objrs_Hapus("id")) + "'"
                M_OBJCONN.Execute CMDSQL
            Next K
            Set M_Objrs_Hapus = Nothing
        Else
            Set M_Objrs_Hapus = Nothing
        End If

        M_Objrs_Custid.MoveNext
    Wend
    
    Set M_Objrs_Custid = Nothing
    
     'Cek data dulu
    CMDSQL = "select custid,promisedate,count(custid) as jumlah from mandiri.tblreserve  "
    CMDSQL = CMDSQL + "group by custid,promisedate having count(custid) > 1"
    Set M_Objrs_Cek = New ADODB.Recordset
    M_Objrs_Cek.CursorLocation = adUseClient
    M_Objrs_Cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs_Cek.RecordCount > 0 Then
        Set M_Objrs_Cek = Nothing
        GoTo cek_lagi
    End If
    
    MsgBox "Proses Hapus TblReserve Berhasil!", vbOKOnly + vbInformation, "Informasi"
End Sub

