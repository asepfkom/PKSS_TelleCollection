VERSION 5.00
Begin VB.Form Frmlock 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6735
   ClientLeft      =   4500
   ClientTop       =   1875
   ClientWidth     =   5385
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   6105
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Batal"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proses"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Agent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   4095
   End
End
Attribute VB_Name = "Frmlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim ds As ADODB.Recordset
CMDSQL = "update mandiri.usertbl set F_LOCK='Y' where USERID='" & Mid(Label1.Caption, 8, 10) & "'"
M_OBJCONN.Execute CMDSQL

For X = 0 To List1.ListCount - 1
    CMDSQL = "UPDATE mandiri.mgm set F_LOCK='Y' where CUSTID='" & List1.list(X) & "'"
    M_OBJCONN.Execute CMDSQL
Next X

MsgBox "Finish ... "
End Sub

Private Sub Form_Load()
'LSTACCESS.ColumnHeaders.ADD 2, , "Cust Id", 10 * TXT
'Call HEADER_VIEW_mgm
'Call ISI_VIEW_mgm
Label1.Caption = "Data : " & VIEW_MGMDATA.cmb_kdagent.Text
End Sub


Private Sub HEADER_VIEW_mgm()
'LstVwSearchmgm.ColumnHeaders.ADD 1, , "No", 3 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 1, , "Cust Id", 5 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 2, , "Priority", 5 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 3, , "Nama Customer", 10 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 4, , "Batch Expire", 10 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 5, , "Tgl Schedule", 5 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 6, , "Next Action", 10 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 7, , "Remarks", 17 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 8, , "Sts Account", 17 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 9, , "Sts LastCall", 17 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 10, , "Call Initial", 5 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 11, , "SalesCode", 5 * TXT
'LstVwSearchmgm.ColumnHeaders.ADD 10, , "Agent", 1 * 1
    
LstVwSearchMgm.ColumnHeaders.ADD 12, , "Principle", 10 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 13, , "Total Amount", 10 * TXT
    
LstVwSearchMgm.ColumnHeaders.ADD 14, , "Join Date", 5 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 15, , "PTP Amount", 5 * TXT
    
LstVwSearchMgm.ColumnHeaders.ADD 16, , "DataBase", 5 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 17, , "LastStatus Date", 5 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 18, , "LastCall Date", 5 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 19, , "Sts Account", 5 * TXT
    
LstVwSearchMgm.ColumnHeaders.ADD 20, , "PTP Date", 5 * TXT
LstVwSearchMgm.ColumnHeaders.ADD 21, , "Complaint Note", 5 * TXT
    'LstVwSearchmgm.ColumnHeaders.ADD 16, , "ID", 10 * TXT
End Sub


Private Sub ISI_VIEW_mgm()
Dim ds As ADODB.Recordset
Dim CMDSQL As String
Set ds = New ADODB.Recordset
ds.CursorLocation = adUseClient
If VIEW_MGMDATA.Combo1(2).Text = "" Then
    CMDSQL = "select * from mandiri.mgm where agent='" & VIEW_MGMDATA.cmb_kdagent.Text & "'"
ElseIf VIEW_MGMDATA.Combo1(2).Text = "" Then
    CMDSQL = "select * from mandiri.mgm where agent='" & VIEW_MGMDATA.cmb_kdagent.Text & "' and Recsource='" & VIEW_MGMDATA.Combo1(2).Text & "'"
End If
ds.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

Set ds = Nothing
End Sub

