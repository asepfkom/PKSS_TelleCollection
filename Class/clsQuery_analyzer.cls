VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQuery_analyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function getAllTable() As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select distinct table_name from information_schema.columns where table_catalog='crm' and table_schema='public'"
Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenKeyset, adLockReadOnly
   Set getAllTable = RSTEMP
   Set RSTEMP = Nothing
End Function


Public Function getFieldTable(sTable As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " SELECT column_name as nama_kolom From information_schema.Columns WHERE table_name='" + sTable + "' ORDER BY ordinal_position "
Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenKeyset, adLockReadOnly
   Set getFieldTable = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function getQuerys(sCmdText As String) As ADODB.Recordset
    On Error GoTo ke
    Dim RSTEMP As New ADODB.Recordset
    Dim sName, sNoKartu As String
    strsql = sCmdText
    
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    
    'sName = Decrypt_PGP("m.nama") + " as nama "
    'sNoKartu = Decrypt_PGP("m.no_kartu") + " as no_kartu "
    
    'strsql = Replace(strsql, "m.nama", sName)
    'strsql = Replace(strsql, "m.no_kartu", sNoKartu)
    
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenKeyset, adLockReadOnly
    
    Set getQuerys = RSTEMP
    Set RSTEMP = Nothing
    Exit Function
ke:
    MsgBox err.Description, vbInformation + vbOKOnly
End Function
Public Function getLoghst() As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " SELECT * From tbllogexport "
Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenKeyset, adLockReadOnly
   Set getLoghst = RSTEMP
   Set RSTEMP = Nothing
End Function
