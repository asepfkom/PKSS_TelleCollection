VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsNegoPTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ADD_OK As Boolean
Public Function ADD_NegoPTP(M_OBJCONN As ADODB.Connection, CustId As String, PromiseDate As String, PromisePay As String, inputdate As String, Jns As String)
Dim CMDSQL As String
PromiseDate = Format(PromiseDate, "yyyy-mm-dd")
inputdate = Format(inputdate, "yyyy-mm-dd")
CMDSQL = "INSERT INTO TblNegoPTP "
CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
CMDSQL = CMDSQL + "VALUES "
CMDSQL = CMDSQL + "('" + CustId + "', "
CMDSQL = CMDSQL + "'" + PromiseDate + "', "
CMDSQL = CMDSQL + "" + PromisePay + " , "
CMDSQL = CMDSQL + "'" + inputdate + "', "
CMDSQL = CMDSQL + "'" + Jns + "')"
On Error GoTo add_error
M_OBJCONN.Execute CMDSQL
ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox err.Description

End Function

Public Function UPDATE_NegoPTP(M_OBJCONN As ADODB.Connection, CustId As String, PromiseDate As String, PromisePay As String, ID As String)
Dim CMDSQL As String
PromiseDate = Format(PromiseDate, "yyyy-mm-dd")
inputdate = Format(inputdate, "yyyy-mm-dd")
CMDSQL = "UPDATE tblnegoPTP SET "
CMDSQL = CMDSQL + "Promisedate = '" + PromiseDate + "', "
CMDSQL = CMDSQL + "Promisepay = " + PromisePay + " "
CMDSQL = CMDSQL + "WHERE id = '" + ID + "'"
On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    
    '@@ 29 Des 2011 Update Status PTP ke tabel mgm
    CMDSQL = "update mgm set dateptp='"
    CMDSQL = CMDSQL + PromiseDate + "' where custid='"
    CMDSQL = CMDSQL + CustId + "'"
    M_OBJCONN.Execute CMDSQL
    
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function


Public Function DELETE_Nego_PTP(M_OBJCONN As ADODB.Connection, ID As String)
    Dim CMDSQL As String
    Dim M_Objrs_Log As ADODB.Recordset
    
    '@@ 11042012 Catet yang menghapus data PTP
    CMDSQL = "select * from tblnegoptp where id='"
    CMDSQL = CMDSQL + CStr(ID) + "'"
    Set M_Objrs_Log = New ADODB.Recordset
    M_Objrs_Log.CursorLocation = adUseClient
    M_Objrs_Log.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs_Log.RecordCount > 0 Then
        CMDSQL = "insert into tblnegoptp_logdel (custid,promisedate,promisepay,inputdate,"
        CMDSQL = CMDSQL + "user_del,tgl_del) values ('"
        CMDSQL = CMDSQL + M_Objrs_Log("custid") + "','"
        CMDSQL = CMDSQL + CStr(Format(M_Objrs_Log("promisedate"), "yyyy-mm-dd")) + "','"
        CMDSQL = CMDSQL + CStr(M_Objrs_Log("promisepay")) + "','"
        CMDSQL = CMDSQL + CStr(Format(M_Objrs_Log("inputdate"), "yyyy-mm-dd")) + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "',"
        CMDSQL = CMDSQL + "now())"
        M_OBJCONN.Execute CMDSQL
    End If
    Set M_Objrs_Log = Nothing
    
    
    CMDSQL = "delete from tblnegoptp WHERE "
    CMDSQL = CMDSQL + "ID = '" + ID + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    
  
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function DELETE_Nego_Reserve(M_OBJCONN As ADODB.Connection, ID As String)
Dim CMDSQL As String
CMDSQL = "DELETE from tblreserve WHERE "
CMDSQL = CMDSQL + "ID = '" + ID + "'"
On Error GoTo add_error
M_OBJCONN.Execute CMDSQL
ADD_OK = True
Exit Function
add_error:
ADD_OK = False
End Function

