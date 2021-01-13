Attribute VB_Name = "ValidasiCall"
Private Sub ValidasiCall(sId As String)
    Dim countAll, countDay As Integer
    Dim rsValidasi_1 As New ADODB.Recordset
    
    
    countAll = 0
    countDay = 0
    sQuerySelect = "SELECT tglcall,jml_call FROM tbl_hst_call_summary WHERE id_cust=" & sId
    rsValidasi_1.Open sQuerySelect, M_OBJCONN, adOpenDynamic, adLockOptimistic
    While rsValidasi_1.EOF = False
        countAll = countAll + cnull(rsValidasi_1!jml_call)
        If cnull(rsValidasi_1!tglcall) = cnull(Now()) Then
            countDay = cnull(rsValidasi_1!jml_call)
        End If
        rsValidasi_1.MoveNext
    Wend
    txt_jumlah_call.Text = cnull(countAll)
    'If last_statuscall = "Followup" Or last_statuscall = "Call Again" Then Exit Sub
    If last_statuscall = "Agree" Then
        If f_agree = "0" Then
            CmdCall.Enabled = True
            Exit Sub
        Else
            CmdCall.Enabled = False
            Exit Sub
        End If
    End If
    If countAll >= call_per_month Then
        CmdCall.Enabled = False
    Else
        If countDay >= call_per_day Then CmdCall.Enabled = False
    End If
    Set rsValidasi_1 = Nothing
End Sub

Private Sub SaveHstCall(sId As String)
    Dim rs_ As New ADODB.Recordset
    Dim iId, iIdHstCall, iJmlCall As Integer
    iId = sId
    sQuerySelect = "SELECT id::integer,jml_call::integer FROM tbl_hst_call_summary WHERE id_cust=" & iId & " AND tglcall=date(now())"
    rs_.Open sQuerySelect, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If rs_.EOF Then
        sQueryInsert = " INSERT INTO tbl_hst_call_summary (id_cust,tglcall,jml_call) VALUES (" & _
                       " " & iId & ",date(now()),1" & _
                       " ) "
        M_OBJCONN.Execute sQueryInsert
    Else
        iIdHstCall = cnull(rs_!ID)
        iJmlCall = cnull(rs_!jml_call)
        sQueryUpdate = "UPDATE tbl_hst_call_summary set jml_call=jml_call+1 WHERE id=" & iIdHstCall
        M_OBJCONN.Execute sQueryUpdate
    End If
End Sub


