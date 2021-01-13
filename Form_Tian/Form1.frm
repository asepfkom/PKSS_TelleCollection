VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd1 
      Caption         =   "Make"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = " select tblstatuscall_keterangan as stts from tblstatuscall where tblstatuscall_kdstatus = '1' order by 1"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    a = ""
    
    While Not M_objrs.EOF
        a = a + " ,case when kodeds = '" & "'" & M_objrs!stts & "'" & "' then 1 else 0 end as """ & "" & M_objrs!stts & """"
        M_objrs.MoveNext
    Wend
    
    q = "select agent " & " & a & "
    q = q & "from mgm_hst"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
End Sub
