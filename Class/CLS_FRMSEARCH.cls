VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_FRMSEARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim CMDSQLOPEN As String
Dim M_OBJCONN As New ADODB.Connection

Public Function QUERY_SEARCH_UPLOAD(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, DATASOURCE As String, TGLLAHIR As String, OFFPHONE As String, OFFPHONE2 As String, HOMEPHONE As String, HOMEPHONE2 As String, MOBILEPHONE As String, MOBILEPHONE2 As String, FAXPHONE As String, FAXPHONE2 As String, TIPE_PRODUK As String) As Object
Dim CMDSQL As String
Dim SYARAT As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM TEMPCC_CUSTTBL"

 If NamaAgent <> Empty Then
    SYARAT = SYARAT + " WHERE " + NamaAgent
 End If

 If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If
 If DATASOURCE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + DATASOURCE
    Else
        SYARAT = SYARAT + " WHERE " + DATASOURCE
    End If
 End If
 If TGLLAHIR <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + TGLLAHIR
    Else
        SYARAT = SYARAT + " WHERE " + TGLLAHIR
    End If
 End If
 If OFFPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + OFFPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + OFFPHONE
    End If
 End If
 If OFFPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + OFFPHONE2 + ")"
 End If
 If HOMEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + HOMEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
    End If
 End If
 If HOMEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ")"
 End If
 If MOBILEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + MOBILEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
    End If
 End If
 If MOBILEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ")"
 End If
 If FAXPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + FAXPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + FAXPHONE
    End If
 End If
 If FAXPHONE2 <> Empty Then
     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
 End If
CMDSQL = CMDSQL + " " + SYARAT + " ORDER BY NAME"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_UPLOAD = M_objrs
Set M_objrs = Nothing
End Function
Public Function QUERY_SEARCH(M_OBJCONN As ADODB.Connection, M_WHERE As String, TIPE_PRODUK As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

Select Case UCase(TIPE_PRODUK)
Case "KTA"
    CMDSQL = "SELECT * FROM KTA_CUSTTBL"
Case "CREDIT CARD"
    CMDSQL = "SELECT * FROM CC_CUSTTBL"
Case "KTA - CROSS SELL"
    CMDSQL = "SELECT * FROM CS_CUSTTBL"
Case "CC - CROSS SELL"
    CMDSQL = "SELECT * FROM CCCS_CUSTTBL"
Case Else
    CMDSQL = "SELECT * FROM CC_CUSTTBL"
End Select
If Len(M_WHERE) <> 0 Then
   CMDSQL = CMDSQL + " WHERE " + M_WHERE
End If
CMDSQL = CMDSQL + " ORDER BY NAME"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_SEARCH_CONDITION(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, DATASOURCE As String, TGLLAHIR As String, OFFPHONE As String, OFFPHONE2 As String, HOMEPHONE As String, HOMEPHONE2 As String, MOBILEPHONE As String, MOBILEPHONE2 As String, FAXPHONE As String, FAXPHONE2 As String, TIPE_PRODUK As String, Lcustid1 As String, KETHSLKERJA As String, lLastCallDate1 As String, lStatusCek1 As String) As Object
Dim CMDSQL As String
Dim SYARAT As String
Dim M_objrs As ADODB.Recordset


    CMDSQL = "SELECT * FROM CC_CUSTTBL"
    
 If Lcustid1 <> Empty Then
    SYARAT = SYARAT + " WHERE " + Lcustid1
 End If
 If NamaAgent <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " and " + NamaAgent
    Else
        SYARAT = SYARAT + " WHERE " + NamaAgent
    End If
 End If
 If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If
 If DATASOURCE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + DATASOURCE
    Else
        SYARAT = SYARAT + " WHERE " + DATASOURCE
    End If
 End If
 If TGLLAHIR <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + TGLLAHIR
    Else
        SYARAT = SYARAT + " WHERE " + TGLLAHIR
    End If
 End If
    
     If lLastCallDate1 <> Empty Then
        If Len(SYARAT) > 0 Then
            SYARAT = SYARAT + " AND " + lLastCallDate1
        Else
            SYARAT = SYARAT + " WHERE " + lLastCallDate1
        End If
     End If
 If lStatusCek1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + lStatusCek1
    Else
        SYARAT = SYARAT + " WHERE " + lStatusCek1
    End If
 End If
 
 If KETHSLKERJA <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + KETHSLKERJA
    Else
        SYARAT = SYARAT + " WHERE " + KETHSLKERJA
    End If
  End If

 If OFFPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + OFFPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + OFFPHONE
    End If
 End If
 If OFFPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + OFFPHONE2 + ")"
 End If
 If HOMEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + HOMEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
    End If
 End If
 If HOMEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ")"
 End If
 If MOBILEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + MOBILEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
    End If
 End If
 If MOBILEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ")"
 End If
 If FAXPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + FAXPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + FAXPHONE
    End If
 End If
 
 If FAXPHONE2 <> Empty Then
     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
 End If
 
 
'tambahan tian AND retur is null
If MDIForm1.txtlevel.text = "Admin" Then
    CMDSQL = CMDSQL + " " + SYARAT
Else
    CMDSQL = CMDSQL + " " + SYARAT + " AND retur is null ORDER BY NAME"
End If
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_CONDITION = M_objrs
Set M_objrs = Nothing

End Function



Public Function QUERY_SEARCH_nonDist(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, DATASOURCE As String, TGLLAHIR As String, OFFPHONE As String, OFFPHONE2 As String, HOMEPHONE As String, HOMEPHONE2 As String, MOBILEPHONE As String, MOBILEPHONE2 As String, FAXPHONE As String, FAXPHONE2 As String, TIPE_PRODUK As String) As Object
Dim CMDSQL As String
Dim SYARAT As String
Dim M_objrs As ADODB.Recordset


 CMDSQL = "SELECT * FROM tempcc_custtbl"

 If NamaAgent <> Empty Then
    SYARAT = SYARAT + " WHERE " + NamaAgent
 End If

 If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If
 If DATASOURCE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + DATASOURCE
    Else
        SYARAT = SYARAT + " WHERE " + DATASOURCE
    End If
 End If
 If TGLLAHIR <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + TGLLAHIR
    Else
        SYARAT = SYARAT + " WHERE " + TGLLAHIR
    End If
 End If
 If OFFPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + OFFPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + OFFPHONE
    End If
 End If
 If OFFPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + OFFPHONE2 + ")"
 End If
 If HOMEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + HOMEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
    End If
 End If
 If HOMEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ")"
 End If
 If MOBILEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + MOBILEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
    End If
 End If
 If MOBILEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ")"
 End If
 If FAXPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + FAXPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + FAXPHONE
    End If
 End If
 If FAXPHONE2 <> Empty Then
     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
 End If
CMDSQL = CMDSQL + " " + SYARAT + " ORDER BY NAME"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_nonDist = M_objrs
Set M_objrs = Nothing
End Function

Private Sub bukakoneksi()
    'CMDSQLOPEN = "Driver={PostgreSQL ANSI}; Server=localhost; PORT=5432; Database=indium; UID=postgres; PWD=123"
    'M_OBJCONN.Open CMDSQLOPEN
End Sub

Public Function QUERY_SEARCH_CONDITION_mgm(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, DATASOURCE As String, TGLLAHIR As String, OFFPHONE As String, OFFPHONE2 As String, HOMEPHONE As String, HOMEPHONE2 As String, MOBILEPHONE As String, MOBILEPHONE2 As String, FAXPHONE As String, FAXPHONE2 As String, TIPE_PRODUK As String, _
                                        AHOMENO As String, AHOMENO2 As String, AOFFICENO As String, AOFFICENO2 As String, extoffice As String, extoffice2 As String, homenoadd1 As String, ahomenoadd1 As String, homenoadd2 As String, ahomenoadd2 As String, officenoadd1 As String, aofficenoadd2 As String, officenoadd2 As String, mobilenoadd1 As String, _
                                        mobilenoadd2 As String, ec_telp As String, alamatrumah As String, alamatkantor As String, alamatec As String, _
                                        Lcustid1 As String, F_CEK As String, lLastCallDate1 As String, lStatusCek1 As String, f_Pending As String, FCEKSTATUS As String, WO_DATE As String, strverify As String, strapprovel As String, strapprovelyet As String, strReject As String, strSukes As String, strBlocked As String, strBlokentry As String, strblokedmarkup As String, nmagentprev As String, LOCKTL As String, BlokedPayment, BlokedPTPNoPayment, Optional page As Long, Optional slimitpage As Double, Optional statuscall As String, Optional Lcustno As String, Optional client As String) As Object
Dim CMDSQL, qs, CMDSQL1 As String
Dim SYARAT As String
Dim iOffset As Integer
Dim iLimit As String
Dim M_objrs As ADODB.Recordset
Dim spvnye As String
Dim tandanye As Integer

Call bukakoneksi

'    Cmdsql = "SELECT CUSTID,PRIOR,NAME, NEXTACTDATE, NEXTACT, REMARKS, F_cek,F_cek_NEW,kethslkerja_new, AGENT,  Principal,DOB,stscpa ,intapprovel,intverify,resultcpa,tglstscpa ,flaglead,"
'    Cmdsql = Cmdsql + "AmountWo, OpenDate, Pay_dt,B_D,TtlPTP, RECSOURCE,TGLSOURCE,TGLSTATUS, Kethslkerja, TGLINCOMING, AmountWo,f_Pending,statusprior,StatusCall,TGLCALL,F_LOCK,tglinsertfrmcpa,curbal,ktpno,dob,lastpay,pay_dt FROM mgm"
    ' UPDATE 12 SEPT 2013 UNTUK tgl cek claim acc
    CMDSQL = "SELECT CUSTID,STATUSCALL,region,PRIOR,NAME, NEXTACTDATE, NEXTACT, REMARKS, F_cek,F_cek_NEW,kethslkerja_new, AGENT,  Principal,DOB,stscpa ,intapprovel,intverify,resultcpa,tglstscpa ,flaglead,"
    CMDSQL = CMDSQL + "ahomeno, ahomeno2, aofficeno, aofficeno2, extoffice, extoffice2, homenoadd1, ahomenoadd1,homenoadd2 , ahomenoadd2, officenoadd1, aofficenoadd2, officenoadd2, mobilenoadd1, mobilenoadd2,ec_telp,addrnow,ecaddr,addrpt,"
    CMDSQL = CMDSQL + "AmountWo, OpenDate, Pay_dt,B_D,TtlPTP, RECSOURCE,TGLSOURCE,TGLSTATUS, Kethslkerja, TGLINCOMING, AmountWo,f_Pending,statusprior,StatusCall,TGLCALL,F_LOCK,tglinsertfrmcpa,curbal,ktpno,dob,lastpay,pay_dt,app_claim,agent_asli, tgl_exp_claim,statuscall FROM mgm"
'==asep==='
SYARAT = "WHERE 1=1 and f_parkir='0'"
'========='
'SYARAT = "WHERE 1=1 "
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        CMDSQL1 = "select spvcode from usertbl where userid = '" & MDIForm1.TxtUsername.text & "'"
        M_objrs.Open CMDSQL1, M_OBJCONN, adOpenDynamic, adLockOptimistic
        spvnye = LCase(M_objrs!SPVCODE)
        Set M_objrs = Nothing
        
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        CMDSQL1 = "select * from tbl_lock_log_" & spvnye & " where agent = '" & MDIForm1.TxtUsername.text & "' and jam_awal < now() and jam_akhir > now()"
        M_objrs.Open CMDSQL1, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        If M_objrs.RecordCount > 0 Then
            tandanye = M_objrs!tanda
            'SYARAT = SYARAT + " and custid in (select custid from tbl_lock_custid_" & spvnye & " where tanda = " & tandanye & ")"
        Else
            tandanye = 0
        End If
        
        Set M_objrs = Nothing
            
    End If
    
    
If MDIForm1.txtlevel.text = "Supervisor" Then
    SYARAT = SYARAT + " and (agent in (select userid from usertbl where spvcode='" + MDIForm1.TxtUsername.text + "')  OR AGENT ='" + MDIForm1.TxtUsername.text + "')"
End If

If Lcustid1 <> Empty Then
   If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + Lcustid1
    Else
       SYARAT = SYARAT + " WHERE " + Lcustid1
    End If
End If

If NamaAgent <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + NamaAgent
    Else
       SYARAT = SYARAT + " WHERE " + NamaAgent
    End If
End If
   
   
If Lcustno <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + Lcustno
    Else
       SYARAT = SYARAT + " WHERE " + Lcustno
    End If
End If
If nmagentprev <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + nmagentprev
    Else
       SYARAT = SYARAT + " WHERE " + nmagentprev
    End If
End If

If LOCKTL <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + LOCKTL
    Else
       SYARAT = SYARAT + " WHERE " + LOCKTL
    End If
        
End If

If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If
 
 If statuscall <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + statuscall
    Else
        SYARAT = SYARAT + " WHERE " + statuscall
    End If
 End If

If FCEKSTATUS <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + FCEKSTATUS
    Else
        SYARAT = SYARAT + " WHERE " + FCEKSTATUS
    End If
 End If
 If DATASOURCE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + DATASOURCE
    Else
        SYARAT = SYARAT + " WHERE " + DATASOURCE
    End If
 End If
 
 If client <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + client
    Else
        SYARAT = SYARAT + " WHERE " + client
    End If
 End If
 
 
 If TGLLAHIR <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + TGLLAHIR
    Else
        SYARAT = SYARAT + " WHERE " + TGLLAHIR
    End If
 End If
 
 If tandanye <> 0 Then
    MsgBox "sedang ada lock data"
    SYARAT = SYARAT + " and custid in (select custid from tbl_lock_custid_" & spvnye & " where tanda = " & tandanye & ")"
 End If
  
  If lLastCallDate1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + lLastCallDate1
    Else
        SYARAT = SYARAT + " WHERE " + lLastCallDate1
    End If
 End If
 If lStatusCek1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + lStatusCek1
    Else
        SYARAT = SYARAT + " WHERE " + lStatusCek1
    End If
 End If
 
  
'  If WO_DATE <> Empty Then
'    If Len(SYARAT) <> 0 Then
'        SYARAT = SYARAT + " AND " + WO_DATE
'    Else
'        SYARAT = SYARAT + " WHERE " + WO_DATE
'    End If
'
'  End If
  
 If OFFPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + OFFPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + OFFPHONE
    End If
 End If
' If OFFPHONE2 <> Empty Then
'    SYARAT = SYARAT + " OR " + OFFPHONE2 + ""
' End If
' If HOMEPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + HOMEPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
'    End If
' End If
' If HOMEPHONE2 <> Empty Then
'    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ""
' End If
 If MOBILEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + MOBILEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
    End If
 End If
' If AHOMENO <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + AHOMENO
'    Else
'        SYARAT = SYARAT + " WHERE (" + AHOMENO
'    End If
' End If
' If AHOMENO2 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + AHOMENO2
'    Else
'        SYARAT = SYARAT + " WHERE (" + AHOMENO2
'    End If
' End If
' If AOFFICENO <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + AOFFICENO
'    Else
'        SYARAT = SYARAT + " WHERE (" + AOFFICENO
'    End If
' End If
' If AOFFICENO2 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + AOFFICENO2
'    Else
'        SYARAT = SYARAT + " WHERE (" + AOFFICENO2
'    End If
' End If
' If extoffice <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + extoffice
'    Else
'        SYARAT = SYARAT + " WHERE (" + extoffice
'    End If
' End If
' If extoffice2 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + extoffice2
'    Else
'        SYARAT = SYARAT + " WHERE (" + extoffice2
'    End If
' End If
' If homenoadd1 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + homenoadd1
'    Else
'        SYARAT = SYARAT + " WHERE (" + homenoadd1
'    End If
' End If
' If ahomenoadd1 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + ahomenoadd1
'    Else
'        SYARAT = SYARAT + " WHERE (" + ahomenoadd1
'    End If
' End If
' If homenoadd2 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + homenoadd2
'    Else
'        SYARAT = SYARAT + " WHERE (" + homenoadd2
'    End If
' End If
' If ahomenoadd2 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + ahomenoadd2
'    Else
'        SYARAT = SYARAT + " WHERE (" + ahomenoadd2
'    End If
' End If
' If officenoadd1 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + officenoadd1
'    Else
'        SYARAT = SYARAT + " WHERE (" + officenoadd1
'    End If
' End If
'  If aofficenoadd2 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + aofficenoadd2
'    Else
'        SYARAT = SYARAT + " WHERE (" + aofficenoadd2
'    End If
' End If
'   If aofficenoadd2 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + aofficenoadd2
'    Else
'        SYARAT = SYARAT + " WHERE (" + aofficenoadd2
'    End If
' End If
' If officenoadd2 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + officenoadd2
'    Else
'        SYARAT = SYARAT + " WHERE (" + officenoadd2
'    End If
' End If
'  If mobilenoadd1 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + mobilenoadd1
'    Else
'        SYARAT = SYARAT + " WHERE (" + mobilenoadd1
'    End If
' End If
'   If mobilenoadd2 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + mobilenoadd2
'    Else
'        SYARAT = SYARAT + " WHERE (" + mobilenoadd2
'    End If
' End If
    If ec_telp <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR " + ec_telp
    Else
        SYARAT = SYARAT + " WHERE (" + ec_telp
    End If
 End If
         
 
 If strverify <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strverify
    Else
        SYARAT = SYARAT + " WHERE " + strverify
    End If
 End If
 
 If strapprovel <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strapprovel
    Else
        SYARAT = SYARAT + " WHERE " + strapprovel
    End If
 End If
 
 If strapprovelyet <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strapprovelyet
    Else
        SYARAT = SYARAT + " WHERE " + strapprovelyet
    End If
 End If
 
 If strReject <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strReject
    Else
        SYARAT = SYARAT + " WHERE " + strReject
    End If
 End If
 
' If strSukes <> Empty Then
'    If Len(SYARAT) > 0 Then
'              SYARAT = SYARAT + " AND " + strSukes
'    Else
'        SYARAT = SYARAT + " WHERE " + strSukes
'    End If
' End If
'
' If f_Pending <> Empty Then
'    If Len(SYARAT) <> 0 Then
'        SYARAT = SYARAT + " AND " + f_Pending
'    Else
'        SYARAT = SYARAT + " WHERE " + f_Pending
'    End If
'  End If
  

 If MOBILEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ""
 End If
' If FAXPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + FAXPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + FAXPHONE
'    End If
' End If
' If FAXPHONE2 <> Empty Then
'     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
' End If
 
If F_LOCK = True Then
    SYARAT = SYARAT + " AND F_LOCK='Y'"
 End If
 
 '@@201010 Ini buat gabungin blok data
 If strBlokentry <> Empty Or F_CEK <> Empty Or strblokedmarkup <> Empty Or BlokedPayment <> Empty Then
    SYARAT = SYARAT + " AND ( "
    Dim SyaratBloked As String

    SyaratBloked = ""

    If strBlokentry <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or " + strBlokentry
        Else
            SyaratBloked = SyaratBloked + strBlokentry
        End If
    End If

    If F_CEK <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or " + F_CEK
        Else
            SyaratBloked = SyaratBloked + F_CEK
        End If
    End If
    
    '@@15 Agustus 2011 Blok Payment Request Gaby
    If BlokedPayment <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or (" + BlokedPayment + ") "
        Else
            SyaratBloked = SyaratBloked + BlokedPayment
        End If
    End If
    
    'BlokedPTPNoPayment
    If BlokedPTPNoPayment <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or (" + BlokedPTPNoPayment + ") "
        Else
            SyaratBloked = SyaratBloked + BlokedPTPNoPayment
        End If
    End If
    
    If strblokedmarkup <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or exclude in (" + Replace(strblokedmarkup, "@", "'") + ")"
        Else
            SyaratBloked = SyaratBloked + " exclude in (" + Replace(strblokedmarkup, "@", "'") + ")"
        End If
    End If
    
    SYARAT = SYARAT + SyaratBloked + ")"
 End If
  

 'jejaktian searchbyalamat
' If alamatrumah <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + alamatrumah
'    Else
'        SYARAT = SYARAT + " WHERE " + alamatrumah
'    End If
' End If
' If alamatkantor <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + alamatkantor
'    Else
'        SYARAT = SYARAT + " WHERE (" + alamatkantor
'    End If
' End If
' If alamatec <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR " + alamatec
'    Else
'        SYARAT = SYARAT + " WHERE (" + alamatec
'    End If
' End If
'===================================
 
    CMDSQL = CMDSQL + " " + SYARAT + ""
  
  'tambahan tian AND retur is null
  If MDIForm1.txtlevel.text = "Admin" Then
    CMDSQL = CMDSQL + " ORDER BY ID "
  Else
    CMDSQL = CMDSQL + " AND retur is null ORDER BY ID "
  End If
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient

M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
Set QUERY_SEARCH_CONDITION_mgm = M_objrs

Set M_objrs = Nothing
End Function
Public Function QUERY_SEARCH_mgm(M_OBJCONN As ADODB.Connection, M_WHERE As String, TIPE_PRODUK As String, F_CEK As String, f_Pending As String, strBlok As String, strBlokentry As String, strblokedmarkup As String, WO_DATE As String, STSLOCKTL As String) As Object
Dim CMDSQL As String
Dim SYARAT As String
Dim M_objrs As ADODB.Recordset

  
    CMDSQL = "SELECT * from mgm"
    'CMDSQL = "SELECT CUSTID,PRIOR,NAME, NEXTACTDATE, NEXTACT, REMARKS, F_cek, AGENT,  Principal, "
    'CMDSQL = CMDSQL + "AmountWo, OpenDate, TtlPTP, RECSOURCE, TGLSTATUS, Kethslkerja, TGLINCOMING,  AmountWo FROM mgm"
If Len(M_WHERE) <> 0 Then
   CMDSQL = CMDSQL + " WHERE " + M_WHERE
End If

 If F_CEK <> Empty Then
    
        CMDSQL = CMDSQL + " AND " + F_CEK
'    Else
'        SYARAT = SYARAT + " WHERE " + F_CEK
    End If
 
 If f_Pending <> Empty Then
        CMDSQL = CMDSQL + " AND " + f_Pending
 End If
  
 '@@Start 19 Juli 2010
  If strBlok <> Empty Then
    If Len(CMDSQL) > 0 Then
        CMDSQL = CMDSQL + " AND " + strBlok
    Else
        CMDSQL = CMDSQL + " WHERE " + strBlok
    End If
 End If
 If strBlokentry <> Empty Then
    If Len(CMDSQL) > 0 Then
        CMDSQL = CMDSQL + " AND " + strBlokentry
    Else
        CMDSQL = CMDSQL + " WHERE " + strBlokentry
    End If
 End If
 
 
 If STSLOCKTL <> Empty Then
 If Len(CMDSQL) > 0 Then
        CMDSQL = CMDSQL + " AND " + STSLOCKTL
    Else
        CMDSQL = CMDSQL + " WHERE " + STSLOCKTL
    End If
 
 End If
 
  If strblokedmarkup <> Empty Then
    If Len(SYARAT) > 0 Then
        CMDSQL = CMDSQL + " AND  exclude in (" + Replace(strblokedmarkup, "@", "'") + ") "
    Else
        CMDSQL = CMDSQL + " and exclude in (" + Replace(strblokedmarkup, "@", "'") + ") "
    End If
 End If

If WO_DATE <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + WO_DATE
    Else
        SYARAT = SYARAT + " WHERE " + WO_DATE
    End If
  
  End If
 
 '--------------- @@End 19 Juli 2010
  
CMDSQL = CMDSQL
'ORDER BY NAME"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_mgm = M_objrs
Set M_objrs = Nothing

End Function
Public Function QUERY_SEARCH_jmlrow(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, DATASOURCE As String, TGLLAHIR As String, OFFPHONE As String, OFFPHONE2 As String, HOMEPHONE As String, HOMEPHONE2 As String, MOBILEPHONE As String, MOBILEPHONE2 As String, FAXPHONE As String, FAXPHONE2 As String, TIPE_PRODUK As String, Lcustid1 As String, F_CEK As String, lLastCallDate1 As String, lStatusCek1 As String, f_Pending As String, FCEKSTATUS As String, WO_DATE As String, strverify As String, strapprovel As String, strapprovelyet As String, strReject As String, strSukes As String, strBlocked As String, strBlokentry As String, strblokedmarkup As String, nmagentprev As String, LOCKTL As String, BlokedPayment, BlokedPTPNoPayment, Optional page As Long, Optional slimitpage As Double, Optional statuscall As String, Optional Lcustno As String) As Object
Dim CMDSQL As String
Dim SYARAT As String
Dim M_objrs As ADODB.Recordset

  
    CMDSQL = "SELECT custid from mgm"
SYARAT = " WHERE 1=1"
If MDIForm1.txtlevel.text = "Supervisor" Then
    SYARAT = SYARAT + " AND (agent in (select userid from usertbl where spvcode='" + MDIForm1.TxtUsername.text + "')  OR AGENT ='" + MDIForm1.TxtUsername.text + "')"
End If
If Lcustid1 <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + Lcustid1
    Else
       SYARAT = SYARAT + " WHERE " + Lcustid1
    End If
End If

If Lcustno <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + Lcustno
    Else
       SYARAT = SYARAT + " WHERE " + Lcustno
    End If
End If
If NamaAgent <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + NamaAgent
    Else
       SYARAT = SYARAT + " WHERE " + NamaAgent
    End If
End If
   
If statuscall <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + statuscall
    Else
       SYARAT = SYARAT + " WHERE " + statuscall
    End If
End If
   
If nmagentprev <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + nmagentprev
    Else
       SYARAT = SYARAT + " WHERE " + nmagentprev
    End If
End If

If LOCKTL <> Empty Then
    If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + LOCKTL
    Else
       SYARAT = SYARAT + " WHERE " + LOCKTL
    End If
        
End If

If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If

If FCEKSTATUS <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + FCEKSTATUS
    Else
        SYARAT = SYARAT + " WHERE " + FCEKSTATUS
    End If
 End If
 If DATASOURCE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + DATASOURCE
    Else
        SYARAT = SYARAT + " WHERE " + DATASOURCE
    End If
 End If
 
' If TGLLAHIR <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND " + TGLLAHIR
'    Else
'        SYARAT = SYARAT + " WHERE " + TGLLAHIR
'    End If
' End If
'
'  If lLastCallDate1 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND " + lLastCallDate1
'    Else
'        SYARAT = SYARAT + " WHERE " + lLastCallDate1
'    End If
' End If
' If lStatusCek1 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND " + lStatusCek1
'    Else
'        SYARAT = SYARAT + " WHERE " + lStatusCek1
'    End If
' End If
'
'
'  If WO_DATE <> Empty Then
'    If Len(SYARAT) <> 0 Then
'        SYARAT = SYARAT + " AND " + WO_DATE
'    Else
'        SYARAT = SYARAT + " WHERE " + WO_DATE
'    End If
'
'  End If
  
' If OFFPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND (" + OFFPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + OFFPHONE
'    End If
' End If
' If OFFPHONE2 <> Empty Then
'    SYARAT = SYARAT + " OR " + OFFPHONE2 + ")"
' End If
' If HOMEPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR (" + HOMEPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
'    End If
' End If
' If HOMEPHONE2 <> Empty Then
'    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ")"
' End If
' If MOBILEPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND (" + MOBILEPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
'    End If
' End If
 
 If strverify <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strverify
    Else
        SYARAT = SYARAT + " WHERE " + strverify
    End If
 End If
 
 If strapprovel <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strapprovel
    Else
        SYARAT = SYARAT + " WHERE " + strapprovel
    End If
 End If
 
 If strapprovelyet <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strapprovelyet
    Else
        SYARAT = SYARAT + " WHERE " + strapprovelyet
    End If
 End If
 
 If strReject <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strReject
    Else
        SYARAT = SYARAT + " WHERE " + strReject
    End If
 End If
 
 If strSukes <> Empty Then
    If Len(SYARAT) > 0 Then
              SYARAT = SYARAT + " AND " + strSukes
    Else
        SYARAT = SYARAT + " WHERE " + strSukes
    End If
 End If
 
' If f_Pending <> Empty Then
'    If Len(SYARAT) <> 0 Then
'        SYARAT = SYARAT + " AND " + f_Pending
'    Else
'        SYARAT = SYARAT + " WHERE " + f_Pending
'    End If
'  End If
'
'
' If MOBILEPHONE2 <> Empty Then
'    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ")"
' End If
' If FAXPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR (" + FAXPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + FAXPHONE
'    End If
' End If
' If FAXPHONE2 <> Empty Then
'     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
' End If
 
If F_LOCK = True Then
    SYARAT = SYARAT + " AND F_LOCK='Y'"
 End If
 
 '@@201010 Ini buat gabungin blok data
 If strBlokentry <> Empty Or F_CEK <> Empty Or strblokedmarkup <> Empty Or BlokedPayment <> Empty Then
    SYARAT = SYARAT + " AND ( "
    Dim SyaratBloked As String

    SyaratBloked = ""

    If strBlokentry <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or " + strBlokentry
        Else
            SyaratBloked = SyaratBloked + strBlokentry
        End If
    End If

    If F_CEK <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or " + F_CEK
        Else
            SyaratBloked = SyaratBloked + F_CEK
        End If
    End If
    
    '@@15 Agustus 2011 Blok Payment Request Gaby
    If BlokedPayment <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or (" + BlokedPayment + ") "
        Else
            SyaratBloked = SyaratBloked + BlokedPayment
        End If
    End If
    
    'BlokedPTPNoPayment
    If BlokedPTPNoPayment <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or (" + BlokedPTPNoPayment + ") "
        Else
            SyaratBloked = SyaratBloked + BlokedPTPNoPayment
        End If
    End If
    
    If strblokedmarkup <> Empty Then
        If Len(SyaratBloked) > 0 Then
            SyaratBloked = SyaratBloked + " or exclude in (" + Replace(strblokedmarkup, "@", "'") + ")"
        Else
            SyaratBloked = SyaratBloked + " exclude in (" + Replace(strblokedmarkup, "@", "'") + ")"
        End If
    End If
    
    SYARAT = SYARAT + SyaratBloked + ")"
 End If
 
 '--------------- @@End 19 Juli 2010
'tambahan tian " AND retur is null"
  If MDIForm1.txtlevel.text = "Admin" Then
    CMDSQL = CMDSQL + SYARAT
  Else
    CMDSQL = CMDSQL + SYARAT + " AND retur is null"
  End If

'ORDER BY NAME"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
DoEvents
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
Set QUERY_SEARCH_jmlrow = M_objrs




Set M_objrs = Nothing

End Function
Public Function QUERY_SEARCH_CALLINITIATED(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, DATASOURCE As String, TGLLAHIR As String, OFFPHONE As String, OFFPHONE2 As String, HOMEPHONE As String, HOMEPHONE2 As String, MOBILEPHONE As String, MOBILEPHONE2 As String, FAXPHONE As String, FAXPHONE2 As String, TIPE_PRODUK As String, Lcustid1 As String, F_CEK As String, lLastCallDate1 As String, lStatusCek1 As String) As Object
Dim SYARAT As String
Dim CMDSQL2 As String
Dim m_objrs2 As ADODB.Recordset

    
    CMDSQL2 = "SELECT CUSTID, COUNT(CUSTID) AS CALL FROM mgm_hst WHERE CUSTID IN "
    CMDSQL2 = CMDSQL2 + " (SELECT CUSTID FROM mgm"
'CMDSQL = "SELECT * FROM SEARCH_mgm"

If Lcustid1 <> Empty Then
   SYARAT = SYARAT + " WHERE " + Lcustid1
End If
If NamaAgent <> Empty Then
   If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + NamaAgent
   Else
       SYARAT = SYARAT + " WHERE " + NamaAgent
   End If
End If
If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If
 If DATASOURCE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + DATASOURCE
    Else
        SYARAT = SYARAT + " WHERE " + DATASOURCE
    End If
 End If
 If TGLLAHIR <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + TGLLAHIR
    Else
        SYARAT = SYARAT + " WHERE " + TGLLAHIR
    End If
 End If
  
  If lLastCallDate1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + lLastCallDate1
    Else
        SYARAT = SYARAT + " WHERE " + lLastCallDate1
    End If
 End If
 If lStatusCek1 <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND " + lStatusCek1
    Else
        SYARAT = SYARAT + " WHERE " + lStatusCek1
    End If
 End If
 
 If F_CEK <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + F_CEK
    Else
        SYARAT = SYARAT + " WHERE " + F_CEK
    End If
  End If

 If OFFPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + OFFPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + OFFPHONE
    End If
 End If
 If OFFPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + OFFPHONE2 + ")"
 End If
 If HOMEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + HOMEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
    End If
 End If
 If HOMEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ")"
 End If
 If MOBILEPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " AND (" + MOBILEPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
    End If
 End If
 If MOBILEPHONE2 <> Empty Then
    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ")"
 End If
 If FAXPHONE <> Empty Then
    If Len(SYARAT) > 0 Then
        SYARAT = SYARAT + " OR (" + FAXPHONE
    Else
        SYARAT = SYARAT + " WHERE (" + FAXPHONE
    End If
 End If
 If FAXPHONE2 <> Empty Then
     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
 End If
CMDSQL2 = CMDSQL2 + " " + SYARAT + ") GROUP BY CUSTID "
CMDSQL2 = "SELECT * FROM ( " + CMDSQL2 + ")A ORDER BY CUSTID "
Set m_objrs2 = New ADODB.Recordset
m_objrs2.CursorLocation = adUseClient
m_objrs2.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_CALLINITIATED = m_objrs2
Set m_objrs2 = Nothing

End Function
Public Function QUERY_SEARCH_VISIT(M_OBJCONN As ADODB.Connection, NAMACUST As String, NamaAgent As String, Lcustid1 As String, STS As String) As Object
Dim SYARAT As String
Dim CMDSQL2 As String
Dim m_objrs2 As ADODB.Recordset


    CMDSQL2 = "SELECT TblVisit.*, mgm.NAME, mgm.AGENT AS AGENT"
    CMDSQL2 = CMDSQL2 + " FROM  TblVisit INNER JOIN "
    CMDSQL2 = CMDSQL2 + " mgm ON TblVisit.CUSTID = mgm.CUSTID  "

If Lcustid1 <> Empty Then
   SYARAT = SYARAT + " WHERE " + Lcustid1
End If
If NamaAgent <> Empty Then
   If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + NamaAgent
   Else
       SYARAT = SYARAT + " WHERE " + NamaAgent
   End If
End If
If NAMACUST <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + NAMACUST
    Else
        SYARAT = SYARAT + " WHERE " + NAMACUST
    End If
 End If
 
 If STS <> Empty Then
 If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + STS
    Else
        SYARAT = SYARAT + " WHERE " + STS
    End If
 End If
 
 CMDSQL2 = CMDSQL2 + " " + SYARAT + "  ORDER BY tblvisit.CUSTID"
Set m_objrs2 = New ADODB.Recordset
m_objrs2.CursorLocation = adUseClient
m_objrs2.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_VISIT = m_objrs2
Set m_objrs2 = Nothing
End Function


Public Function QUERY_AGENT_JADWAL(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim mwhere As String
Dim M_objrs As ADODB.Recordset
    
    CMDSQL = "SELECT USERID,agent from usertbl "
    
    If MDIForm1.txtlevel = "Supervisor" Then
        mwhere = " where spvcode='" + MDIForm1.TxtUsername.text + "' and aktif='1'"
    ElseIf MDIForm1.txtlevel = "Admin" Or MDIForm1.txtlevel = "Administrator" Then
        mwhere = " where aktif='1' and userid <> 'ADM01' "
    End If
    CMDSQL = CMDSQL + mwhere + " ORDER BY USERID"
        
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient

M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_AGENT_JADWAL = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_AGENT(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset


    CMDSQL = "SELECT usertbl.*, SPVTBL.SPVNAME"
    CMDSQL = CMDSQL + " FROM usertbl INNER JOIN"
    CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
    CMDSQL = CMDSQL + " WHERE USERTYPE ='1'"
    If Len(M_WHERE) <> 0 Then
        CMDSQL = CMDSQL + " AND " + M_WHERE
    End If
    
    CMDSQL = CMDSQL + " ORDER BY USERID"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_AGENT = M_objrs
Set M_objrs = Nothing

End Function
Public Function QUERY_FIELD(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset


    CMDSQL = "SELECT usertbl.*, SPVTBL.SPVNAME"
    CMDSQL = CMDSQL + " FROM usertbl INNER JOIN"
    CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
    CMDSQL = CMDSQL + " WHERE USERTYPE ='2'"
    If Len(M_WHERE) <> 0 Then
        CMDSQL = CMDSQL + " AND " + M_WHERE
    End If
    
    CMDSQL = CMDSQL + " ORDER BY USERID"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_FIELD = M_objrs
Set M_objrs = Nothing

End Function

Public Function QUERY_DATASOURCE(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset


'CMDSQL = "SELECT * FROM datasourcetbl"
' If Len(M_WHERE) <> 0 Then
'    CMDSQL = CMDSQL + " WHERE " + M_WHERE
' End If
'
'CMDSQL = CMDSQL + " ORDER BY KODEDS"

'===asep==='
CMDSQL = "SELECT distinct recsource FROM mgm"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY recsource"
'=========='
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient

M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_DATASOURCE = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_AGENT_xl(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset


    CMDSQL = "SELECT usertbl.*, SPVTBL.SPVNAME"
    CMDSQL = CMDSQL + " FROM usertbl INNER JOIN"
    CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
'    CMDSQL = CMDSQL + " WHERE USERTYPE ='1'"
    If Len(M_WHERE) <> 0 Then
        CMDSQL = CMDSQL + " where " + M_WHERE
    End If
    
    CMDSQL = CMDSQL + " ORDER BY USERID"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_AGENT_xl = M_objrs
Set M_objrs = Nothing

End Function

Public Function QUERY_SEARCH_CONDITION_mgm2(M_OBJCONN As ADODB.Connection, Reff_Num As String, No_Laporan As String, tgl_laporan As String, Idi_User As String, Debitur_Name As String, Born_Place As String, Born_Place2 As String, Born_Date As String, Din As String, NPWP As String, NPWP2 As String, KTP_NUM As String, KTP_NUM2 As String, Passport_Num As String, Passport_Num2 As String, HTML_SOURCE As String) As Object
Dim CMDSQL As String
Dim SYARAT As String
Dim M_objrs As ADODB.Recordset
'    CMDSQL = "SELECT CUSTID,PRIOR,NAME, NEXTACTDATE, NEXTACT, REMARKS, F_cek, AGENT,  Principal,DOB "
'    CMDSQL = CMDSQL + "AmountWo, OpenDate, TtlPTP, RECSOURCE, TGLSTATUS, Kethslkerja, TGLINCOMING, AmountWo,f_Pending FROM mgm"
    CMDSQL = "SELECT * From MAIN_DEBITUR_INFO "
    
'If Lcustid1 <> Empty Then
'   SYARAT = SYARAT + " WHERE " + Lcustid1
'End If
If Reff_Num <> Empty Then
   If Len(SYARAT) <> 0 Then
       SYARAT = SYARAT + " AND " + Reff_Num
   Else
       SYARAT = SYARAT + " WHERE " + Reff_Num
   End If
End If
If Debitur_Name <> Empty Then
    If Len(SYARAT) <> 0 Then
        SYARAT = SYARAT + " AND " + Debitur_Name
    Else
        SYARAT = SYARAT + " WHERE " + Debitur_Name
    End If
 End If
' If DATASOURCE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND " + DATASOURCE
'    Else
'        SYARAT = SYARAT + " WHERE " + DATASOURCE
'    End If
' End If
' If TGLLAHIR <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND " + TGLLAHIR
'    Else
'        SYARAT = SYARAT + " WHERE " + TGLLAHIR
'    End If
' End If
'
'  If lLastCallDate1 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND " + lLastCallDate1
'    Else
'        SYARAT = SYARAT + " WHERE " + lLastCallDate1
'    End If
' End If
' If lStatusCek1 <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND " + lStatusCek1
'    Else
'        SYARAT = SYARAT + " WHERE " + lStatusCek1
'    End If
' End If
'
' If F_CEK <> Empty Then
'    If Len(SYARAT) <> 0 Then
'        SYARAT = SYARAT + " AND " + F_CEK
'    Else
'        SYARAT = SYARAT + " WHERE " + F_CEK
'    End If
'  End If
'
' If OFFPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND (" + OFFPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + OFFPHONE
'    End If
' End If
' If OFFPHONE2 <> Empty Then
'    SYARAT = SYARAT + " OR " + OFFPHONE2 + ")"
' End If
' If HOMEPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR (" + HOMEPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + HOMEPHONE
'    End If
' End If
' If HOMEPHONE2 <> Empty Then
'    SYARAT = SYARAT + " OR " + HOMEPHONE2 + ")"
' End If
' If MOBILEPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " AND (" + MOBILEPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + MOBILEPHONE
'    End If
' End If
' If f_Pending <> Empty Then
'    If Len(SYARAT) <> 0 Then
'        SYARAT = SYARAT + " AND " + f_Pending
'    Else
'        SYARAT = SYARAT + " WHERE " + f_Pending
'    End If
'  End If
'
' If MOBILEPHONE2 <> Empty Then
'    SYARAT = SYARAT + " OR " + MOBILEPHONE2 + ")"
' End If
' If FAXPHONE <> Empty Then
'    If Len(SYARAT) > 0 Then
'        SYARAT = SYARAT + " OR (" + FAXPHONE
'    Else
'        SYARAT = SYARAT + " WHERE (" + FAXPHONE
'    End If
' End If
' If FAXPHONE2 <> Empty Then
'     SYARAT = SYARAT + " OR " + FAXPHONE2 + ")"
' End If
CMDSQL = CMDSQL + " " + SYARAT + ""

Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient


M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SEARCH_CONDITION_mgm2 = M_objrs
Set M_objrs = Nothing


End Function


