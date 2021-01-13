Attribute VB_Name = "Module1"
Public m_objconn As New ADODB.Connection

Sub Main()
Dim cmdsql As String
On Error GoTo SqlConnErr
    cmdsql = "Provider=MSDASQL.1;Password=;Persist Security Info=False;User ID=TELE;Data Source=TELEGRANDISQL"
    m_objconn.Open cmdsql
    FRMTERIMAMSG.Show
Exit Sub
SqlConnErr:
    MsgBox Err.Description
End Sub

Public Function BUKA_FILE_KONEKSI(M_FILE As String) As String
Dim f As String
Dim t As TextStream
On Error GoTo HELL
'Set t = fso.OpenTextFile(App.Path & "\" & M_FILE, ForReading)
Set t = fso.OpenTextFile(M_FILE, ForReading)
BUKA_FILE_KONEKSI = t.ReadAll
t.Close
Exit Function
HELL:
    BUKA_FILE_KONEKSI = ""
'    MsgBox Err.Description
End Function

