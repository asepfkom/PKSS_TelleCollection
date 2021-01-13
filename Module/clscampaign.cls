VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clscampaign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function FindCampaign(Optional strcodeCampaign As String, Optional strNama As String) As ADODB.Recordset
Dim RSTEMP As New ADODB.Recordset
strsql = " select * from datasourcetbl where kodeds <>'' "

If strcodeCampaign <> Empty Then
            mwhere = " and kodeds ='" + strcodeCampaign + "'"
End If

If strNama <> Empty Then
            mwhere = mwhere + " and keterangan  like '%" + strNama + "%'"
End If

Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.Open strsql + mwhere + "", M_OBJCONN, adOpenDynamic, adLockOptimistic
   Set FindCampaign = RSTEMP
   Set RSTEMP = Nothing
End Function
Public Function insertCampaign(M_OBJCONN As ADODB.Connection, sCodeCampaign As String, sKetproduct As String, sAktif, tdbTgl As Object) As Boolean
On Error GoTo out
Dim strsql As String
    insertCampaign = True
    strsql = "insert into  datasourcetbl (kodeds,keterangan,status,tglinsert,last_tglupdate,tglexpire) "
    strsql = strsql + " values ('" + sCodeCampaign + "','" + sKetproduct + "','" + sAktif + "',date(now()),date(now()),"
    
    If Not (tdbTgl.ValueIsNull) Then
      strsql = strsql + "'" + CStr(Format(tdbTgl.Value, "yyyy-mm-dd")) + "'"
    Else
      strsql = strsql + "NULL"
    End If
    strsql = strsql + ")"
    M_OBJCONN.Execute (strsql)
    
'    STRSQL = "insert Into tbl_datasourcetbl_hst ( kodeds,status ,keterangan, tglexpire,tgl_lead ,user_input,nama_input, activity) "
'    STRSQL = STRSQL + " select kodeds,status,keterangan,tglexpire,tgl_lead,   "
'    STRSQL = STRSQL + "'" + mdiform1.txtusername.text + "' as userinput,'" + MDIForm1.txtnama.Text + "' as nminput,'ADD'"
'    STRSQL = STRSQL + " from  datasourcetbl where  kodeds='" + sCodeCampaign + "'"
'    M_OBJCONN.Execute (STRSQL)
    
    Exit Function
    
    
    
out:
    insertCampaign = False
End Function

'Public Function editCampaign(M_OBJCONN As ADODB.Connection, sCodeCampaign As String, sKetproduct As String, sAktif, tdbTgl As Object) As Boolean
'On Error GoTo out
'Dim STRSQL As String
'editCampaign = True
'
'   STRSQL = "insert Into tbl_datasourcetbl_hst ( kodeds,status ,keterangan, tglexpire,tgl_lead ,user_input,nama_input, activity) "
'    STRSQL = STRSQL + " select kodeds,status,keterangan,tglexpire,tgl_lead,   "
'    STRSQL = STRSQL + "'" + MDIForm1.txtUserName.Text + "' as userinput,'" + MDIForm1.txtnama.Text + "' as nminput,'EDIT'"
'    STRSQL = STRSQL + " from  datasourcetbl where  kodeds='" + sCodeCampaign + "'"
'    M_OBJCONN.Execute (STRSQL)
'
'    STRSQL = "update datasourcetbl set  keterangan ='" + sKetproduct + "',status='" + sAktif + "',"
'
'    If Not (tdbTgl.ValueIsNull) Then
'      STRSQL = STRSQL + " TGLEXPIRE='" + CStr(Format(tdbTgl.Value, "yyyy-mm-dd")) + "',"
'    Else
'      STRSQL = STRSQL + " TGLEXPIRE =NULL,"
'    End If
'
'    STRSQL = STRSQL + "last_tglupdate =date(now()),LAST_USER='" + MDIForm1.txtUserName.Text + "' where kodeds='" + sCodeCampaign + "'"
'    M_OBJCONN.Execute (STRSQL)
'    Exit Function
'out:
'    editCampaign = False
'End Function
'Public Function deleteCampaign(M_OBJCONN As ADODB.Connection, sCodeCampaign As String) As Boolean
'On Error GoTo out
'Dim STRSQL As String
'    deleteCampaign = True
'
'
'    STRSQL = "insert Into tbl_datasourcetbl_hst ( kodeds,status ,keterangan, tglexpire,tgl_lead ,user_input,nama_input, activity) "
'    STRSQL = STRSQL + " select kodeds,status,keterangan,tglexpire,tgl_lead,   "
'    STRSQL = STRSQL + "'" + MDIForm1.txtUserName.Text + "' as userinput,'" + MDIForm1.txtnama.Text + "' as nminput,'DELETE'"
'    STRSQL = STRSQL + " from  datasourcetbl where  kodeds='" + sCodeCampaign + "'"
'    M_OBJCONN.Execute (STRSQL)
'
'    STRSQL = "delete from datasourcetbl where kodeds='" + sCodeCampaign + "'"
'    M_OBJCONN.Execute (STRSQL)
'    Exit Function
'out:
'    deleteCampaign = False
'End Function
'
'
'
'

