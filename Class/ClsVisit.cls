VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ADD_OK As Boolean
Public Function ADD_RequestVisit(M_OBJCONN As ADODB.Connection, CustId As String, F_CEK As String, VisitNo As String, RequestDate As String, DetailsR As String, VisitKe As String, AddressToVisit As String, AGENT As String)
Dim CMDSQL As String
RequestDate = Format(RequestDate, "yyyy-mm-dd")
CMDSQL = "INSERT INTO TblVisit "

CMDSQL = CMDSQL + "(custid,f_cek,VisitNo,RequestDate,DetailsR,VisitKe, AddresstoVisit,AGENT) "
CMDSQL = CMDSQL + "VALUES "
CMDSQL = CMDSQL + "('" + CustId + "', "
CMDSQL = CMDSQL + "'" + F_CEK + "', "
CMDSQL = CMDSQL + "'" + VisitNo + "', "
CMDSQL = CMDSQL + "'" + RequestDate + "', "
CMDSQL = CMDSQL + "'" + DetailsR + "', "
CMDSQL = CMDSQL + "'" + VisitKe + "', "
CMDSQL = CMDSQL + "'" + AddressToVisit + "', "
CMDSQL = CMDSQL + "'" + AGENT + "' )"
On Error GoTo add_error
M_OBJCONN.Execute CMDSQL
ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox err.Description

End Function
Public Function SELECT_RequestVisit(M_OBJCONN As ADODB.Connection, Lcustid As String)
Dim CMDSQL As String
CMDSQL = "SELECT * FROM tblVisit where CUSTID = "
CMDSQL = CMDSQL + "'" + Lcustid + "'"

On Error GoTo add_error
M_OBJCONN.Execute CMDSQL
ADD_OK = True
Exit Function
add_error:
    ADD_OK = False
   MsgBox err.Description

End Function
Public Function SELECT_VisitKe(M_OBJCONN As ADODB.Connection, Lcustid As String)
Dim CMDSQL As String
CMDSQL = "SELECT count(custid) as jml FROM tblVisit where CUSTID = "
CMDSQL = CMDSQL + "'" + Lcustid + "'"

On Error GoTo add_error
M_OBJCONN.Execute CMDSQL
ADD_OK = True
Exit Function
add_error:
    ADD_OK = False

End Function

Public Function UPDATE_RequestVisit(M_OBJCONN As ADODB.Connection, RequestDate As String, VisitDate As String, DetailsV As String, StatusVisit As String, ID As String)
Dim CMDSQL As String
RequestDate = Format(PromiseDate, "yyyy-mm-dd")
VisitDate = Format(PromiseDate, "yyyy-mm-dd")

CMDSQL = "UPDATE tblvisit SET "
CMDSQL = CMDSQL + "RequestDate = '" + RequestDate + "', "
CMDSQL = CMDSQL + "VisitDate = '" + VisitDate + "' "
CMDSQL = CMDSQL + "DetailsV = '" + DetailsV + "' "
CMDSQL = CMDSQL + "StatusVisit = '" + StatusVisit + "' "
CMDSQL = CMDSQL + "Sts = '1' "
CMDSQL = CMDSQL + "WHERE id = '" + ID + "'"
On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function


Public Function DELETE_RequestVisit(M_OBJCONN As ADODB.Connection, ID As String)
Dim CMDSQL As String

CMDSQL = "DELETE tblNegoPTP WHERE "
CMDSQL = CMDSQL + "ID = '" + ID + "'"
On Error GoTo add_error
M_OBJCONN.Execute CMDSQL
ADD_OK = True
Exit Function
add_error:
ADD_OK = False
End Function


