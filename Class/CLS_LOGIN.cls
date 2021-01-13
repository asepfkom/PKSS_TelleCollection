VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ADD_OK As Boolean

Public Function INSERT_LOGIN(M_OBJCONN As ADODB.Connection, Tanggal As String, WSTATION As String, ACTIVITY As String, AGENT As String)
Dim CMDSQL As String
CMDSQL = " INSERT INTO USERLOG"
CMDSQL = CMDSQL + " (DATETIME,"
CMDSQL = CMDSQL + " WSTATION,"
CMDSQL = CMDSQL + " ACTIVITY,"
CMDSQL = CMDSQL + " AGENT)"
CMDSQL = CMDSQL + " VALUES"
CMDSQL = CMDSQL + " ('" + Tanggal + "',"
CMDSQL = CMDSQL + " '" + WSTATION + "',"
CMDSQL = CMDSQL + " '" + ACTIVITY + "',"
CMDSQL = CMDSQL + " '" + AGENT + "')"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox err.Description
End Function

Public Function INSERT_CLIENT(M_OBJCONN As ADODB.Connection, WSSTATIONIP As String, WSSTATIONPORT As String, AGENT As String)
Dim CMDSQL As String
CMDSQL = " INSERT INTO CLIENTTBL"
CMDSQL = CMDSQL + " (AGENT,"
CMDSQL = CMDSQL + " WSSTATIONIP,"
CMDSQL = CMDSQL + " WSSTATIONPORT)"
CMDSQL = CMDSQL + " VALUES"
CMDSQL = CMDSQL + " ('" + AGENT + "',"
CMDSQL = CMDSQL + " '" + WSSTATIONIP + "',"
CMDSQL = CMDSQL + " '" + WSSTATIONPORT + "')"

    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
End Function

Public Function UPDATE_CLIENT(M_OBJCONN As ADODB.Connection, WSSTATIONIP As String, AGENT As String, WSSTATION As String)
Dim CMDSQL As String
CMDSQL = " UPDATE CLIENTTBL"
CMDSQL = CMDSQL + " SET WSSTATIONIP ='" + WSSTATIONIP + "',"
CMDSQL = CMDSQL + " WSSTATION ='" + WSSTATION + "'"
CMDSQL = CMDSQL + " WHERE"
CMDSQL = CMDSQL + " AGENT ='" + AGENT + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox err.Description
End Function

Public Function UPDATE_LOGIN(M_OBJCONN As ADODB.Connection, Tanggal As String, WSTATION As String, ACTIVITY As String, AGENT As String)
Dim CMDSQL As String
CMDSQL = " UPDATE USERLOG"
CMDSQL = CMDSQL + " SET WSTATION ='" + WSTATION + "',"
CMDSQL = CMDSQL + " ACTIVITY='" + ACTIVITY + "' "
CMDSQL = CMDSQL + " WHERE"
CMDSQL = CMDSQL + " AGENT ='" + AGENT + "' "
CMDSQL = CMDSQL + " AND DATETIME = '" + Tanggal + "'"
    On Error GoTo add_error
    M_OBJCONN.Execute CMDSQL
    ADD_OK = True
    Exit Function
add_error:
    ADD_OK = False
    MsgBox err.Description
End Function

