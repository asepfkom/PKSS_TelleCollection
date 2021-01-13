Attribute VB_Name = "FuncCekSegment"
Public Function FuncCekSegmen(notlp As String) As Double
    Dim sQuery As String
    Dim Rs_Segmen As ADODB.Recordset
    
    FuncCekSegmen = 0
    
    sQuery = "SELECT * FROM tbl_temp_segment_call "
    sQuery = sQuery + " WHERE no_telpon = '" & notlp & "' AND date(tgl_call) = '" & Format(waktu_server_sekarang, "yyyy-mm-dd") & "' "
    Set Rs_Segmen = New ADODB.Recordset
    Rs_Segmen.CursorLocation = adUseClient
    Rs_Segmen.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Rs_Segmen.RecordCount > 0 Then
        FuncCekSegmen = Rs_Segmen!jumlah_call
    Else
        FuncCekSegmen = 0
    End If
End Function


Public Function FuncCekReview(notlp As String) As Double
    Dim sQuery As String
    Dim Rs_Jumlah_Call As ADODB.Recordset

    CustId = Trim(FrmCC_Colection.lblCustId.text)
    
    FuncCekReview = 0

    sQuery = "SELECT * FROM tbl_temp_telfon_review "
    sQuery = sQuery + " WHERE no_telfon = '" & notlp & "' AND date(tanggal_telfon) = '" & Format(waktu_server_sekarang, "yyyy-mm-dd") & "' AND custId = '" & CustId & "'" 'UPDATETIAN23FEBRUARI2016
    Set Rs_Jumlah_Call = New ADODB.Recordset
    Rs_Jumlah_Call.CursorLocation = adUseClient
    Rs_Jumlah_Call.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    If Rs_Jumlah_Call.RecordCount > 0 Then
        FuncCekReview = Rs_Jumlah_Call!jumlah_call
    Else
        FuncCekReview = 0
    End If
End Function

Public Sub ConnectRS(ByRef rsS As ADODB.Recordset)
    Set rsS = New ADODB.Recordset
        rsS.CursorLocation = adUseClient
        rsS.CursorType = adOpenDynamic
        rsS.LockType = adLockOptimistic
        rsS.ActiveConnection = M_OBJCONN
End Sub
Public Sub UnConnectRs(ByRef rsS As ADODB.Recordset)
    On Error Resume Next
    If rsS.State = 1 Then rsS.Close
End Sub

Public Function getQuerys(sCmdText As String) As ADODB.Recordset
    On Error GoTo ke
    Dim RSTEMP As New ADODB.Recordset
    strsql = sCmdText
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere, M_OBJCONN, adOpenKeyset, adLockReadOnly
    query_public = strsql + mwhere
    Set getQuerys = RSTEMP
    Set RSTEMP = Nothing
    Exit Function
ke:
    MsgBox "Query Wrong", vbInformation + vbOKOnly
End Function
