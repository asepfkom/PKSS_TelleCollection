VERSION 5.00
Begin VB.Form fmunlock 
   Caption         =   "Un lock Data"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form2"
   ScaleHeight     =   2505
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Unlock data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Agent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "fMUNLOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Dim ds As New ADODB.Recordset
ds.CursorLocation = adUseClient
ds.Open "select AGENT FROM usertbl WHERE USERID='" & Combo1.Text & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If ds.BOF And ds.EOF Then
    MsgBox "User tidak ada !"
Else
    Text1.Text = ds!agent
End If
End Sub


Private Sub Command1_Click()
Cmdsql = "update usertbl set F_LOCK='' WHErE USERID='" & Combo1.Text & "'"
M_OBJCONN.Execute Cmdsql

Cmdsql = "UPDATE mgm set F_LOCK='' where agent='" & Combo1.Text & "'"
M_OBJCONN.Execute Cmdsql

MsgBox "Finish ... "
End Sub

Private Sub Form_Load()
Dim ds As New ADODB.Recordset
ds.CursorLocation = adUseClient
'ds.Open "select USERID,AGENT FROM usertbl WHERE TEAM='" & MDIForm1.Text1.Text & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ds.Open "select USERID,AGENT FROM usertbl ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If ds.BOF And ds.EOF Then
Else
    ds.MoveFirst
    Do While Not ds.EOF
        Combo1.AddItem ds!Userid
        ds.MoveNext
    Loop
End If
End Sub
