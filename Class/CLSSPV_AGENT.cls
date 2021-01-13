VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLSSPV_AGENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ADD_OK As Boolean

Public Function QUERY_KOMISI(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM KOMISI"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY KODE"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_KOMISI = M_objrs
Set M_objrs = Nothing
End Function


Public Function UPDATE_KOMISI(M_OBJCONN As ADODB.Connection, KODE As String, BTSAWAL As String, BTSAKHR As String, KOMISI As String)
Dim CMDSQL As String
On Error GoTo add_error
CMDSQL = " UPDATE KOMISI"
CMDSQL = CMDSQL + " SET"
CMDSQL = CMDSQL + " BTSAWAL = " + CStr(CCur(BTSAWAL)) + ","
CMDSQL = CMDSQL + " BTSAKHR = " + CStr(CCur(BTSAKHR)) + ","
CMDSQL = CMDSQL + " KOMISI = " + CStr(CCur(KOMISI)) + ""
CMDSQL = CMDSQL + " WHERE"
CMDSQL = CMDSQL + " KODE = '" + UBAH_QUOTE(KODE) + "'"
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    If err.number = 13 Then
        MsgBox "Angka yang dimasukan tidak valid"
    End If
    ADD_OK = False
End Function


Public Function query(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM SPVTBL"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY SPVCODE"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set query = M_objrs
Set M_objrs = Nothing

End Function

Public Function DELETE(M_OBJCONN As ADODB.Connection, SPVCODE As String)
Dim CMDSQL As String

CMDSQL = "DELETE from SPVTBL"
CMDSQL = CMDSQL + " WHERE SPVCODE = '" + UBAH_QUOTE(SPVCODE) + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
  
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function ADD(M_OBJCONN As ADODB.Connection, SPVCODE As String, SPVNAME As String, TEAM As String, UNIT As String, TARGET As String, JABATAN As String)
Dim CMDSQL As String

CMDSQL = " INSERT INTO SPVTBL"
CMDSQL = CMDSQL + " (SPVCODE,"
CMDSQL = CMDSQL + " TEAM,"
CMDSQL = CMDSQL + " UNIT,"
'CMDSQL = CMDSQL + " TARGET,"
CMDSQL = CMDSQL + " SPVNAME)"
CMDSQL = CMDSQL + " VALUES"
CMDSQL = CMDSQL + " ('" + UBAH_QUOTE(SPVCODE) + "',"
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(TEAM) + "',"
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(UNIT) + "',"
'CMDSQL = CMDSQL + " " + CStr(CCur(TARGET)) + ","
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(SPVNAME) + "')"
    On Error GoTo add_error
     M_OBJCONN.Execute CMDSQL
    'Set M_OBJCONN = Nothing
    Call ADD_SPV_DITBLAGENT(M_OBJCONN, SPVCODE, SPVNAME, UNIT, TEAM, JABATAN)
    'Call ADD_CLIENTTBL(M_OBJCONN, SPVCODE)
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Private Function ADD_SPV_DITBLAGENT(M_OBJCONN As ADODB.Connection, USERID As String, AGENT As String, UNIT As String, TEAM As String, JABATAN As String)
Dim CMDSQL As String
    CMDSQL = " INSERT INTO usertbl"
    CMDSQL = CMDSQL + " (USERID,"
    CMDSQL = CMDSQL + " AGENT,"
    CMDSQL = CMDSQL + " UNIT,"
    CMDSQL = CMDSQL + " SPVCODE,"
    CMDSQL = CMDSQL + " PWD,"
    CMDSQL = CMDSQL + " ACCREC,"
    CMDSQL = CMDSQL + " TEAM,"
    CMDSQL = CMDSQL + " USERTYPE)"
    CMDSQL = CMDSQL + " VALUES"
    CMDSQL = CMDSQL + " ('" + UBAH_QUOTE(USERID) + "',"
    CMDSQL = CMDSQL + " '" + UBAH_QUOTE(AGENT) + "',"
    CMDSQL = CMDSQL + " '" + UBAH_QUOTE(UNIT) + "',"
    CMDSQL = CMDSQL + " '" + UBAH_QUOTE(USERID) + "',"
    CMDSQL = CMDSQL + " '" + Encrypt(Len(USERID), "PASS12345") + "',"
    CMDSQL = CMDSQL + " '" + Encrypt(Len(USERID), "PASS12345") + "',"
    CMDSQL = CMDSQL + " '" + UBAH_QUOTE(TEAM) + "',"
    If JABATAN = "AM" Then
        CMDSQL = CMDSQL + " 25)"
    Else
        CMDSQL = CMDSQL + " 20)"
    End If
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function update(M_OBJCONN As ADODB.Connection, SPVCODE As String, SPVNAME As String, TEAM As String, UNIT As String, TARGET As String)
Dim CMDSQL As String
CMDSQL = " UPDATE SPVTBL"
CMDSQL = CMDSQL + " SET"
CMDSQL = CMDSQL + " SPVNAME = '" + UBAH_QUOTE(SPVNAME) + "',"
CMDSQL = CMDSQL + " UNIT = '" + UBAH_QUOTE(UNIT) + "',"
'CMDSQL = CMDSQL + " TARGET = " + CStr(CCur(TARGET)) + ","
CMDSQL = CMDSQL + " TEAM = '" + UBAH_QUOTE(TEAM) + "'"
CMDSQL = CMDSQL + " WHERE"
CMDSQL = CMDSQL + " SPVCODE ='" + UBAH_QUOTE(SPVCODE) + "'"
    On Error GoTo add_error
  
    M_OBJCONN.Execute CMDSQL
  
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function QUERY_AGENT(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset


'    cmdsql = "SELECT usertbl.*, SPVTBL.SPVNAME"
'    cmdsql = cmdsql + " FROM usertbl INNER JOIN"
'    cmdsql = cmdsql + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"

    CMDSQL = "select *,teamleader from "
    CMDSQL = CMDSQL + " (SELECT usertbl.* FROM usertbl ORDER BY USERID) chris "
    CMDSQL = CMDSQL + " Left Join "
    CMDSQL = CMDSQL + " (select agent as teamleader,spvcode from usertbl where userid like '%TL%') tian "
    CMDSQL = CMDSQL + " on chris.spvcode = tian.spvcode "

    If Len(M_WHERE) <> 0 Then
        CMDSQL = CMDSQL + " WHERE " + M_WHERE
    End If
    
    CMDSQL = CMDSQL + " ORDER BY USERID"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_AGENT = M_objrs
Set M_objrs = Nothing

End Function

Public Function DELETE_AGENT(M_OBJCONN As ADODB.Connection, USERID As String)
Dim CMDSQL As String
CMDSQL = "DELETE from usertbl"
CMDSQL = CMDSQL + " WHERE USERID = '" + UBAH_QUOTE(USERID) + "'"
    On Error GoTo add_error
      M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function DELETE_ordering(M_OBJCONN As ADODB.Connection, USERID As String)
Dim CMDSQL As String
CMDSQL = "DELETE from tbloffering"
CMDSQL = CMDSQL + " WHERE id_offering  = " + CStr(USERID) + ""
    On Error GoTo add_error
      M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function ADD_AGENT(M_OBJCONN As ADODB.Connection, USERID As String, AGENT As String, SPVCODE As String, SALARYBSC As String, TEAM As String, aktif As String, LVL As String, UNIT As String, AM As String, sId As Integer)
Dim CMDSQL As String
    If Len(SALARYBSC) = 0 Then
        SALARYBSC = "0"
    End If
    CMDSQL = " INSERT INTO usertbl"
    CMDSQL = CMDSQL + " (id,USERID,"
    CMDSQL = CMDSQL + " AGENT,"
    CMDSQL = CMDSQL + " SALARYBSC,"
    CMDSQL = CMDSQL + " TEAM,"
    CMDSQL = CMDSQL + " AKTIF,"
    CMDSQL = CMDSQL + " LVL,"
    CMDSQL = CMDSQL + " USERTYPE,"
    CMDSQL = CMDSQL + " ACCREC,"
    CMDSQL = CMDSQL + " UNIT,"
    CMDSQL = CMDSQL + " AM,"
    CMDSQL = CMDSQL + " SPVCODE)"
    CMDSQL = CMDSQL + " VALUES"
    CMDSQL = CMDSQL + " (" & sId & ",'" + UBAH_QUOTE(USERID) + "',"
    CMDSQL = CMDSQL + " '" + UBAH_QUOTE(AGENT) + "',"
    CMDSQL = CMDSQL + " " + CStr(CCur(SALARYBSC)) + ","
    CMDSQL = CMDSQL + " '" + TEAM + "',"
    CMDSQL = CMDSQL + " " + CStr(CCur(aktif)) + ","
    CMDSQL = CMDSQL + " '" + LVL + "',"
    CMDSQL = CMDSQL + " 1,"
    CMDSQL = CMDSQL + " '" & Encrypt(Len(USERID), "PASS12345") & "',"
    CMDSQL = CMDSQL + " '" + UNIT + "',"
    CMDSQL = CMDSQL + " '" + AM + "',"
    CMDSQL = CMDSQL + " '" + UBAH_QUOTE(SPVCODE) + "')"
    On Error GoTo add_error
    
    M_OBJCONN.Execute CMDSQL
    
    Call ADD_CLIENTTBL(M_OBJCONN, USERID)
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function ADD_offering(M_OBJCONN As ADODB.Connection, keterangan As String, fld As String, discount As String, Remarks As String, operator As String, keyrms As String, existpersentase As String)
Dim CMDSQL As String
  
'    CMDSQL = " INSERT INTO tbloffering"
'    CMDSQL = CMDSQL + " (keterangan,"
'    CMDSQL = CMDSQL + " fldrms ,"
'    CMDSQL = CMDSQL + " persentase,"
'    CMDSQL = CMDSQL + " remarks,operand,idkey,exispersentase)"
'    CMDSQL = CMDSQL + " VALUES"
'    CMDSQL = CMDSQL + " ('" + keterangan + "',"
'    CMDSQL = CMDSQL + " '" + fld + "',"
'    CMDSQL = CMDSQL + " " + CStr(discount) + ","
'    CMDSQL = CMDSQL + " '" + frminputoreder.txtremarks.Text + "','" + operator + "','" + keyrms + "','" + existpersentase + "')"
'    On Error GoTo add_error
'
'    M_OBJCONN.Execute CMDSQL
'
'    ADD_OK = True
'    Exit Function
'add_error:
'    ADD_OK = False
End Function

Public Function UPDATE_AGENT(M_OBJCONN As ADODB.Connection, USERID As String, AGENT As String, SPVCODE As String, SALARYBSC As String, TEAM As String, aktif As String, LVL As String, UNIT As String, AM As String)
Dim CMDSQL As String

    If Len(SALARYBSC) = 0 Then
    SALARYBSC = 0
    End If

    CMDSQL = " UPDATE usertbl"
    CMDSQL = CMDSQL + " SET"
    CMDSQL = CMDSQL + " AGENT ='" + UBAH_QUOTE(AGENT) + "' ,"
    CMDSQL = CMDSQL + " SALARYBSC = " + CStr(CCur(SALARYBSC)) + ","
    CMDSQL = CMDSQL + " TEAM = '" + TEAM + "',"
    CMDSQL = CMDSQL + " LVL = '" + LVL + "',"
    CMDSQL = CMDSQL + " UNIT = '" + UNIT + "',"
    CMDSQL = CMDSQL + " AM = '" + AM + "',"
    CMDSQL = CMDSQL + " AKTIF = " + CStr(CCur(aktif)) + ","
    CMDSQL = CMDSQL + " SPVCODE = '" + UBAH_QUOTE(SPVCODE) + "'"
    CMDSQL = CMDSQL + " WHERE"
    CMDSQL = CMDSQL + " USERID = '" + UBAH_QUOTE(USERID) + "'"
    
    On Error GoTo add_error
    
    M_OBJCONN.Execute CMDSQL
    
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function
'Public Function UPDATE_ordering(M_OBJCONN As ADODB.Connection, id_offer As String, keterangan As String, fld As String, discount As String, Remarks As String, operator As String, keyrms As String, exispersentase As String)
'Dim CMDSQL As String
'    CMDSQL = " UPDATE tbloffering "
'    CMDSQL = CMDSQL + " SET keterangan ='" + keterangan + "' ,"
'    CMDSQL = CMDSQL + " idkey ='" + keyrms + "' ,"
'    CMDSQL = CMDSQL + " fldrms ='" + fld + "' ,"
'     CMDSQL = CMDSQL + "  exispersentase ='" + exispersentase + "' ,"
'    CMDSQL = CMDSQL + " persentase = " + discount + ","
'     CMDSQL = CMDSQL + " operand  = '" + operator + "',"
'    CMDSQL = CMDSQL + " remarks = '" + frmheaderoffeer.TxtRemarks.Text + "'"
'    CMDSQL = CMDSQL + " WHERE "
'    CMDSQL = CMDSQL + "  id_offering = " + CStr(id_offer) + ""
'
'    On Error GoTo add_error
'
'    M_OBJCONN.Execute CMDSQL
'
'    ADD_OK = True
'    Exit Function
'add_error:
'    ADD_OK = False
'End Function

Public Function COMBO_SPV(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim M_objrs As ADODB.Recordset
Dim CMDSQL As String
    
        CMDSQL = " SELECT * FROM usertbl where userid like 'TL%' and spvcode like 'SPV%'"
    If Len(M_WHERE) <> 0 Then
        CMDSQL = CMDSQL + " AND " + M_WHERE
    End If
    CMDSQL = CMDSQL + " ORDER BY SPVCODE"
   
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set COMBO_SPV = M_objrs
Set M_objrs = Nothing

End Function

Public Function QUERY_PRODUCT(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM PRODUCTTBL"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY CODE"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_PRODUCT = M_objrs
Set M_objrs = Nothing
End Function

Public Function ADD_PRODUCT(M_OBJCONN As ADODB.Connection, CODE As String, STATUS1 As String, PRODUCT As String, UNIT As String)
Dim CMDSQL As String
CMDSQL = " INSERT INTO PRODUCTTBL"
CMDSQL = CMDSQL + " (CODE,"
CMDSQL = CMDSQL + " STATUS,"
CMDSQL = CMDSQL + " UNIT,"
CMDSQL = CMDSQL + " PRODUCT)"
CMDSQL = CMDSQL + " VALUES"
CMDSQL = CMDSQL + " ('" + UBAH_QUOTE(CODE) + "',"
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(STATUS1) + "',"
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(UNIT) + "',"
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(PRODUCT) + "')"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function DELETE_PRODUCT(M_OBJCONN As ADODB.Connection, CODE As String)
Dim CMDSQL As String
CMDSQL = "DELETE From PRODUCTTBL"
CMDSQL = CMDSQL + " WHERE CODE = '" + CODE + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function UPDATE_PRODUCT(M_OBJCONN As ADODB.Connection, CODE As String, PRODUCT As String, STATUS1 As String, UNIT As String)
Dim CMDSQL As String
CMDSQL = " UPDATE PRODUCTTBL"
CMDSQL = CMDSQL + " SET"
CMDSQL = CMDSQL + " PRODUCT = '" + UBAH_QUOTE(PRODUCT) + "',"
CMDSQL = CMDSQL + " UNIT = '" + UBAH_QUOTE(UNIT) + "',"
CMDSQL = CMDSQL + " STATUS = '" + UBAH_QUOTE(STATUS1) + "'"
CMDSQL = CMDSQL + " WHERE"
CMDSQL = CMDSQL + " CODE ='" + UBAH_QUOTE(CODE) + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function


Public Function UPDATE_Password(M_OBJCONN As ADODB.Connection, USERID As String, ACCREC As String, AUTH As String)
Dim CMDSQL As String
CMDSQL = " UPDATE usertbl"
CMDSQL = CMDSQL + " SET"
CMDSQL = CMDSQL + " ACCREC = '" + UBAH_QUOTE(ACCREC) + "',"
CMDSQL = CMDSQL + " AUTH = '" + UBAH_QUOTE(AUTH) + "'"
CMDSQL = CMDSQL + " WHERE"
CMDSQL = CMDSQL + " USERID ='" + UBAH_QUOTE(USERID) + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function QUERY_SET_PWDAGENT(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

    CMDSQL = "SELECT * FROM usertbl"
    CMDSQL = CMDSQL + " WHERE USERID <> 'ADMINISTRATOR'"
    
    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
         CMDSQL = CMDSQL + " and spvcode = '" & UCase(MDIForm1.TxtUsername.text) & "'"
    End If
    
    If Len(M_WHERE) <> 0 Then
        CMDSQL = CMDSQL + " AND " + M_WHERE
    End If
    
    CMDSQL = CMDSQL + " ORDER BY USERID"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SET_PWDAGENT = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_CLOSSING(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM CLSREASONTBL"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
   
CMDSQL = CMDSQL + " ORDER BY KDCLS"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_CLOSSING = M_objrs
Set M_objrs = Nothing
End Function

Public Function ADD_CLOSSING(M_OBJCONN As ADODB.Connection, KDCLS As String, KETCLS As String, jenis As String)
Dim CMDSQL As String

CMDSQL = " INSERT INTO CLSREASONTBL"
CMDSQL = CMDSQL + " (KDCLS,"
CMDSQL = CMDSQL + " jenis,"
CMDSQL = CMDSQL + " KETCLS)"
CMDSQL = CMDSQL + " VALUES"
CMDSQL = CMDSQL + " ('" + UBAH_QUOTE(KDCLS) + "',"
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(jenis) + "',"
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(KETCLS) + "')"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function UPDATE_CLOSSING(M_OBJCONN As ADODB.Connection, KDCLS As String, KETCLS As String, jenis As String)
Dim CMDSQL As String
CMDSQL = " UPDATE CLSREASONTBL"
CMDSQL = CMDSQL + " SET"
CMDSQL = CMDSQL + " KETCLS = '" + UBAH_QUOTE(KETCLS) + "',"
CMDSQL = CMDSQL + " jenis = '" + UBAH_QUOTE(jenis) + "'"
CMDSQL = CMDSQL + " WHERE"
CMDSQL = CMDSQL + " KDCLS ='" + UBAH_QUOTE(KDCLS) + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function DELETE_CLOSSING(M_OBJCONN As ADODB.Connection, KDCLS As String)
Dim CMDSQL As String
CMDSQL = "DELETE from CLSREASONTBL"
CMDSQL = CMDSQL + " WHERE KDCLS = '" + UBAH_QUOTE(KDCLS) + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function ADD_DATASOURCE(M_OBJCONN As ADODB.Connection, KODEDS As String, STATUS As String, keterangan As String, KDPROGRAM As String)
Dim CMDSQL As String

CMDSQL = " INSERT INTO DATASOURCETBL"
CMDSQL = CMDSQL + " (KODEDS,"
CMDSQL = CMDSQL + " STATUS,"
CMDSQL = CMDSQL + " KDPROGRAM,"
CMDSQL = CMDSQL + " KETERANGAN)"
CMDSQL = CMDSQL + " VALUES"
CMDSQL = CMDSQL + " ('" + UBAH_QUOTE(KODEDS) + "',"
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(STATUS) + "',"
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(KDPROGRAM) + "',"
CMDSQL = CMDSQL + " '" + UBAH_QUOTE(keterangan) + "')"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function DELETE_DATASOURCE(M_OBJCONN As ADODB.Connection, KODEDS As String)
Dim CMDSQL As String
CMDSQL = "DELETE from DATASOURCETBL"
CMDSQL = CMDSQL + " WHERE KODEDS = '" + UBAH_QUOTE(KODEDS) + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function


Public Function QUERY_DATASOURCE(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

CMDSQL = "SELECT * FROM DATASOURCETBL"
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY KODEDS"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_DATASOURCE = M_objrs
Set M_objrs = Nothing
End Function

Public Function UPDATE_DATASOURCE(M_OBJCONN As ADODB.Connection, KODEDS As String, STATUS As String, keterangan As String, KDPROGRAM As String)
Dim CMDSQL As String
CMDSQL = " UPDATE DATASOURCETBL"
CMDSQL = CMDSQL + " SET"
CMDSQL = CMDSQL + " KETERANGAN = '" + UBAH_QUOTE(keterangan) + "', "
CMDSQL = CMDSQL + " KDPROGRAM = '" + UBAH_QUOTE(KDPROGRAM) + "', "
CMDSQL = CMDSQL + " STATUS =  '" + UBAH_QUOTE(STATUS) + "' "
CMDSQL = CMDSQL + " WHERE"
CMDSQL = CMDSQL + " KODEDS = '" + KODEDS + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    'MsgBox Err.Description
End Function

Public Function ADD_CLIENTTBL(M_OBJCONN As ADODB.Connection, AGENT As String)
Dim CMDSQL As String
CMDSQL = " INSERT INTO CLIENTTBL"
CMDSQL = CMDSQL + " (AGENT)"
CMDSQL = CMDSQL + " VALUES"
CMDSQL = CMDSQL + " ('" + UBAH_QUOTE(AGENT) + "')"
    On Error GoTo add_error
     M_OBJCONN.Execute CMDSQL
        ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function QUERY_AOC(M_OBJCONN As ADODB.Connection) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset


    CMDSQL = "SELECT * FROM usertbl WHERE USERID != AGENT ORDER BY USERID"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_AOC = M_objrs
Set M_objrs = Nothing

End Function

'Public Function QUERY_AOC_CHECK(M_OBJCONN As ADODB.Connection) As Boolean
'    Dim CMDSQL As String
'    Dim M_objrs As ADODB.Recordset
'    Dim cuserid As String
'    Dim cagent As String
'    Dim caoc As String
'
'    cuserid = FormAOC.txtuserid.Text
'    cagent = FormAOC.lblnama.Caption
'    caoc = FormAOC.txtaoc.Text
'    CMDSQL = "SELECT USERID FROM tblaoc WHERE USERID != AGENT AND USERID = '" + cagent + "'"
'
'    Set M_objrs = New ADODB.Recordset
'    M_objrs.CursorLocation = adUseClient
'    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If M_objrs.RecordCount > 0 Then
'        QUERY_AOC_CHECK = True
'    Else
'        QUERY_AOC_CHECK = False
'    End If
'    Set M_objrs = Nothing
'End Function

Public Function SHOW_AOC(M_OBJCONN As ADODB.Connection)
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    CMDSQL = "SELECT * FROM tblaoc"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    Set SHOW_AOC = M_objrs
    Set M_objrs = Nothing
End Function

'Public Function UPDATE_AOC(M_OBJCONN As ADODB.Connection)
'    Dim CMDSQL As String
'    Dim cuserid As String
'    Dim cagent As String
'    Dim caoc As String
'
'    cuserid = FormAOC.txtuserid.Text
'    cagent = FormAOC.lblnama.Caption
'    caoc = FormAOC.txtaoc.Text
'        CMDSQL = "UPDATE TBLAOC SET agent = '" + cuserid + "' , aoc = '" + caoc + "' where userid = '" + cagent + "'"
'
'    M_OBJCONN.Execute CMDSQL
'        ADD_OK = True
'        Exit Function
'add_error:
'        ADD_OK = False
'        FormAOC.txtuserid.SetFocus
'End Function

'Public Function ADD_AOC(M_OBJCONN As ADODB.Connection)
'    Dim CMDSQL As String
'    Dim cuserid As String
'    Dim cagent As String
'    Dim caoc As String
'
'    cuserid = FormAOC.txtuserid.Text
'    cagent = FormAOC.lblnama.Caption
'    caoc = FormAOC.txtaoc.Text
'        CMDSQL = "INSERT INTO tblaoc values ('" + cagent + "' , '" + cuserid + "' , '" + caoc + "')"
'    M_OBJCONN.Execute CMDSQL
'        ADD_OK = True
'        Exit Function
'add_error:
'        ADD_OK = False
'        FormAOC.txtuserid.SetFocus
'End Function

'Public Function DELETE_AOC(M_OBJCONN As ADODB.Connection)
'Dim CMDSQL As String
'Dim cuserid As String
'cuserid = FormAOC.lblnama.Caption
'CMDSQL = "DELETE from tblaoc where userid = '" + cuserid + "'"
'      M_OBJCONN.Execute CMDSQL
'End Function
