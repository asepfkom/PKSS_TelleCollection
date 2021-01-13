VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function QUERY_DATASOURCE(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

    CMDSQL = "SELECT * FROM DATASOURCETBL"
    
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY KETERANGAN"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_DATASOURCE = M_objrs
Set M_objrs = Nothing
End Function


Public Function QUERY_SPV(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

    CMDSQL = "SELECT * FROM SPVTBL"
    
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " WHERE " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY SPVNAME"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_SPV = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_AGENT(M_OBJCONN As ADODB.Connection, M_WHERE As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

    CMDSQL = "SELECT * FROM usertbl"
    CMDSQL = CMDSQL + " WHERE USERTYPE = 1"
    
 If Len(M_WHERE) <> 0 Then
    CMDSQL = CMDSQL + " AND " + M_WHERE
 End If
    
CMDSQL = CMDSQL + " ORDER BY AGENT"
    
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_AGENT = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_VIEW_ALL_AVAILABLE(M_OBJCONN As ADODB.Connection, M_WHERE As String, M_ORDER As String, TIPE_PRODUK As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

Select Case UCase(TIPE_PRODUK)
Case "KTA"
           CMDSQL = "SELECT KTA_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
           CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
           CMDSQL = CMDSQL + " FROM KTA_CUSTTBL INNER JOIN"
           CMDSQL = CMDSQL + " usertbl ON"
           CMDSQL = CMDSQL + " KTA_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
           CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
           CMDSQL = CMDSQL + " WHERE LEFT(RECSTATUS,2) ='1A'"
            If Len(M_WHERE) <> 0 Then
               CMDSQL = CMDSQL + " AND " + M_WHERE
            End If
Case "CREDIT CARD"
           CMDSQL = "SELECT CC_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
           CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
           CMDSQL = CMDSQL + " FROM CC_CUSTTBL INNER JOIN"
           CMDSQL = CMDSQL + " usertbl ON"
           CMDSQL = CMDSQL + " CC_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
           CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
           CMDSQL = CMDSQL + " WHERE LEFT(RECSTATUS,2) ='1A'"
            If Len(M_WHERE) <> 0 Then
               CMDSQL = CMDSQL + " AND " + M_WHERE
            End If
Case "KTA - CROSS SELL"
           CMDSQL = "SELECT CS_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
           CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
           CMDSQL = CMDSQL + " FROM CS_CUSTTBL INNER JOIN"
           CMDSQL = CMDSQL + " usertbl ON"
           CMDSQL = CMDSQL + " CS_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
           CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
           CMDSQL = CMDSQL + " WHERE LEFT(RECSTATUS,2) ='1A'"
            If Len(M_WHERE) <> 0 Then
               CMDSQL = CMDSQL + " AND " + M_WHERE
            End If
Case "CC - CROSS SELL"
           CMDSQL = "SELECT CCCS_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
           CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
           CMDSQL = CMDSQL + " FROM CCCS_CUSTTBL INNER JOIN"
           CMDSQL = CMDSQL + " usertbl ON"
           CMDSQL = CMDSQL + " CCCS_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
           CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
           CMDSQL = CMDSQL + " WHERE LEFT(RECSTATUS,2) ='1A'"
            If Len(M_WHERE) <> 0 Then
               CMDSQL = CMDSQL + " AND " + M_WHERE
            End If
Case Else
    Exit Function
End Select
    
 If Len(M_ORDER) <> 0 Then
    CMDSQL = CMDSQL + " ORDER BY " + M_ORDER
 End If
   
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_VIEW_ALL_AVAILABLE = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_VIEW_ALL_NEW(M_OBJCONN As ADODB.Connection, m_data As String, M_AGENT As String, M_SPV As String, M_ORDER As String, TIPE_PRODUK As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset
    
Select Case UCase(TIPE_PRODUK)
Case "KTA"
       CMDSQL = "SELECT KTA_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
       CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
       CMDSQL = CMDSQL + " FROM KTA_CUSTTBL INNER JOIN"
       CMDSQL = CMDSQL + " usertbl ON"
       CMDSQL = CMDSQL + " KTA_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
       CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
       CMDSQL = CMDSQL + " WHERE LEFT(RECSTATUS,2) ='1A'"
    If m_data <> Empty Then
       CMDSQL = CMDSQL + " AND " + m_data
    End If
       If M_AGENT <> Empty Then
       CMDSQL = CMDSQL + " AND " + M_AGENT
    End If
    If M_SPV <> Empty Then
       CMDSQL = CMDSQL + " AND " + M_SPV
    End If
Case "CREDIT CARD"
       CMDSQL = "SELECT CC_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
       CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
       CMDSQL = CMDSQL + " FROM CC_CUSTTBL INNER JOIN"
       CMDSQL = CMDSQL + " usertbl ON"
       CMDSQL = CMDSQL + " CC_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
       CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
       CMDSQL = CMDSQL + " WHERE LEFT(RECSTATUS,2) ='1A'"
    If m_data <> Empty Then
       CMDSQL = CMDSQL + " AND " + m_data
    End If
       If M_AGENT <> Empty Then
       CMDSQL = CMDSQL + " AND " + M_AGENT
    End If
    If M_SPV <> Empty Then
       CMDSQL = CMDSQL + " AND " + M_SPV
    End If
Case "KTA - CROSS SELL"
       CMDSQL = "SELECT CS_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
       CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
       CMDSQL = CMDSQL + " FROM CS_CUSTTBL INNER JOIN"
       CMDSQL = CMDSQL + " usertbl ON"
       CMDSQL = CMDSQL + " CS_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
       CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
       CMDSQL = CMDSQL + " WHERE LEFT(RECSTATUS,2) ='1A'"
    If m_data <> Empty Then
       CMDSQL = CMDSQL + " AND " + m_data
    End If
       If M_AGENT <> Empty Then
       CMDSQL = CMDSQL + " AND " + M_AGENT
    End If
    If M_SPV <> Empty Then
       CMDSQL = CMDSQL + " AND " + M_SPV
    End If
Case "CC - CROSS SELL"
       CMDSQL = "SELECT CCCS_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
       CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
       CMDSQL = CMDSQL + " FROM CCCS_CUSTTBL INNER JOIN"
       CMDSQL = CMDSQL + " usertbl ON"
       CMDSQL = CMDSQL + " CCCS_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
       CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
       CMDSQL = CMDSQL + " WHERE LEFT(RECSTATUS,2) ='1A'"
    If m_data <> Empty Then
       CMDSQL = CMDSQL + " AND " + m_data
    End If
       If M_AGENT <> Empty Then
       CMDSQL = CMDSQL + " AND " + M_AGENT
    End If
    If M_SPV <> Empty Then
       CMDSQL = CMDSQL + " AND " + M_SPV
    End If
Case Else
    Exit Function
End Select
 
 If Len(M_ORDER) <> 0 Then
    CMDSQL = CMDSQL + " ORDER BY " + M_ORDER
 End If
   
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_VIEW_ALL_NEW = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_INCOMING(M_OBJCONN As ADODB.Connection, M_WHERE As String, M_ORDER As String, TIPE_PRODUK As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset

Select Case UCase(TIPE_PRODUK)
Case "KTA"
       CMDSQL = "SELECT KTA_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
       CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
       CMDSQL = CMDSQL + " FROM KTA_CUSTTBL INNER JOIN"
       CMDSQL = CMDSQL + " usertbl ON"
       CMDSQL = CMDSQL + " KTA_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
       CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
       CMDSQL = CMDSQL + " WHERE (RECSTATUS ='2C' OR RECSTATUS ='3A' OR RECSTATUS ='XX')"
    If Len(M_WHERE) <> 0 Then
       CMDSQL = CMDSQL + " AND " + M_WHERE
    End If
Case "CREDIT CARD"
       CMDSQL = "SELECT CC_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
       CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
       CMDSQL = CMDSQL + " FROM CC_CUSTTBL INNER JOIN"
       CMDSQL = CMDSQL + " usertbl ON"
       CMDSQL = CMDSQL + " CC_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
       CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
       CMDSQL = CMDSQL + " WHERE (RECSTATUS ='2C' OR RECSTATUS ='3A' OR RECSTATUS ='XX')"
    If Len(M_WHERE) <> 0 Then
       CMDSQL = CMDSQL + " AND " + M_WHERE
    End If
Case "KTA - CROSS SELL"
       CMDSQL = "SELECT CS_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
       CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
       CMDSQL = CMDSQL + " FROM CS_CUSTTBL INNER JOIN"
       CMDSQL = CMDSQL + " usertbl ON"
       CMDSQL = CMDSQL + " CS_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
       CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
       CMDSQL = CMDSQL + " WHERE (RECSTATUS ='2C' OR RECSTATUS ='3A' OR RECSTATUS ='XX')"
    If Len(M_WHERE) <> 0 Then
       CMDSQL = CMDSQL + " AND " + M_WHERE
    End If
Case "CC - CROSS SELL"
       CMDSQL = "SELECT CCCS_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
       CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
       CMDSQL = CMDSQL + " FROM CCCS_CUSTTBL INNER JOIN"
       CMDSQL = CMDSQL + " usertbl ON"
       CMDSQL = CMDSQL + " CCCS_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
       CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
       CMDSQL = CMDSQL + " WHERE (RECSTATUS ='2C' OR RECSTATUS ='3A' OR RECSTATUS ='XX')"
    If Len(M_WHERE) <> 0 Then
       CMDSQL = CMDSQL + " AND " + M_WHERE
    End If
Case Else
    Exit Function
End Select
If Len(M_ORDER) <> 0 Then
    CMDSQL = CMDSQL + " ORDER BY " + M_ORDER
 End If
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_INCOMING = M_objrs
Set M_objrs = Nothing
End Function

Public Function QUERY_INCOMING_NEW(M_OBJCONN As ADODB.Connection, m_data As String, M_AGENT As String, M_SPV As String, M_ORDER As String, TIPE_PRODUK As String) As Object
Dim CMDSQL As String
Dim M_objrs As ADODB.Recordset
Select Case UCase(TIPE_PRODUK)
Case "KTA"
       CMDSQL = "SELECT KTA_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
       CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
       CMDSQL = CMDSQL + " FROM KTA_CUSTTBL INNER JOIN"
       CMDSQL = CMDSQL + " usertbl ON"
       CMDSQL = CMDSQL + " KTA_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
       CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
       CMDSQL = CMDSQL + " WHERE RECSTATUS ='2C'"
       'OR RECSTATUS1 ='4'"
        If m_data <> Empty Then
           CMDSQL = CMDSQL + " AND " + m_data
        End If
           If M_AGENT <> Empty Then
           CMDSQL = CMDSQL + " AND " + M_AGENT
        End If
        If M_SPV <> Empty Then
           CMDSQL = CMDSQL + " AND " + M_SPV
        End If
Case "CREDIT CARD"
       CMDSQL = "SELECT CC_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
       CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
       CMDSQL = CMDSQL + " FROM CC_CUSTTBL INNER JOIN"
       CMDSQL = CMDSQL + " usertbl ON"
       CMDSQL = CMDSQL + " CC_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
       CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
       CMDSQL = CMDSQL + " WHERE RECSTATUS ='2C'"
       'OR RECSTATUS1 ='4'"
        If m_data <> Empty Then
           CMDSQL = CMDSQL + " AND " + m_data
        End If
           If M_AGENT <> Empty Then
           CMDSQL = CMDSQL + " AND " + M_AGENT
        End If
        If M_SPV <> Empty Then
           CMDSQL = CMDSQL + " AND " + M_SPV
        End If
Case "KTA - CROSS SELL"
       CMDSQL = "SELECT CS_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
       CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
       CMDSQL = CMDSQL + " FROM CS_CUSTTBL INNER JOIN"
       CMDSQL = CMDSQL + " usertbl ON"
       CMDSQL = CMDSQL + " CS_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
       CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
       CMDSQL = CMDSQL + " WHERE RECSTATUS ='2C'"
       'OR RECSTATUS1 ='4'"
        If m_data <> Empty Then
           CMDSQL = CMDSQL + " AND " + m_data
        End If
           If M_AGENT <> Empty Then
           CMDSQL = CMDSQL + " AND " + M_AGENT
        End If
        If M_SPV <> Empty Then
           CMDSQL = CMDSQL + " AND " + M_SPV
        End If
Case "CC - CROSS SELL"
       CMDSQL = "SELECT CCCS_CUSTTBL.*, usertbl.AGENT AS NAMAAGENT,"
       CMDSQL = CMDSQL + " SPVTBL.SPVNAME AS SPVNAME"
       CMDSQL = CMDSQL + " FROM CCCS_CUSTTBL INNER JOIN"
       CMDSQL = CMDSQL + " usertbl ON"
       CMDSQL = CMDSQL + " CCCS_CUSTTBL.AGENT = usertbl.USERID INNER JOIN"
       CMDSQL = CMDSQL + " SPVTBL ON usertbl.SPVCODE = SPVTBL.SPVCODE"
       CMDSQL = CMDSQL + " WHERE RECSTATUS ='2C'"
       'OR RECSTATUS1 ='4'"
        If m_data <> Empty Then
           CMDSQL = CMDSQL + " AND " + m_data
        End If
           If M_AGENT <> Empty Then
           CMDSQL = CMDSQL + " AND " + M_AGENT
        End If
        If M_SPV <> Empty Then
           CMDSQL = CMDSQL + " AND " + M_SPV
        End If
Case Else
    Exit Function
End Select
If Len(M_ORDER) <> 0 Then
   CMDSQL = CMDSQL + " ORDER BY " + M_ORDER
End If
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Set QUERY_INCOMING_NEW = M_objrs
Set M_objrs = Nothing
End Function

