VERSION 5.00
Begin VB.Form Frm_Script_SMS 
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5025
   LinkTopic       =   "Form2"
   ScaleHeight     =   3705
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3660
      TabIndex        =   9
      Top             =   3210
      Width           =   1125
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2490
      TabIndex        =   8
      Top             =   3210
      Width           =   1125
   End
   Begin VB.TextBox Txtpanjang 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2550
      Width           =   945
   End
   Begin VB.TextBox TxtSms 
      Height          =   1455
      Left            =   1050
      MaxLength       =   320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1020
      Width           =   3675
   End
   Begin VB.TextBox TxtSubOption 
      Height          =   315
      Left            =   1050
      TabIndex        =   3
      Top             =   630
      Width           =   3675
   End
   Begin VB.ComboBox CmbOption 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   210
      Width           =   2865
   End
   Begin VB.Label Label4 
      Caption         =   "(Max 160 Char)"
      Height          =   285
      Left            =   3570
      TabIndex        =   6
      Top             =   2550
      Width           =   1245
   End
   Begin VB.Label Label3 
      Caption         =   "Script sms:"
      Height          =   285
      Left            =   60
      TabIndex        =   4
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Sub option:"
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Option:"
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   645
   End
End
Attribute VB_Name = "Frm_Script_SMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ok           As Boolean

Private Sub IsiOption()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    CMDSQL = "select distinct option from tblscriptsms"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    CmbOption.CLEAR
    While Not M_objrs.EOF
        CmbOption.AddItem M_objrs("option")
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End Sub

Private Sub CmdBatal_Click()
    ok = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim VSAVE As Boolean
    
    VSAVE = True
    VSAVE = VSAVE And CmbOption.Text <> ""
    VSAVE = VSAVE And TxtSubOption.Text <> ""
    VSAVE = VSAVE And TxtSms.Text <> ""
    
    If VSAVE Then
        ok = True
        Me.Hide
    Else
        MsgBox "Textbox ada yang masih kosong!", vbOKOnly + vbInformation, "Informasi"
        ok = False
        Exit Sub
    End If
    
End Sub

Private Sub Form_Load()
    ok = False
    IsiOption
End Sub

Private Sub TxtSms_Change()
    Txtpanjang.Text = Len(Trim(TxtSms.Text))
End Sub

