VERSION 5.00
Begin VB.Form FrmUserTarget 
   Caption         =   "Target"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   Icon            =   "FrmUserTarget.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtTarget 
      Height          =   315
      Index           =   4
      Left            =   7515
      TabIndex        =   13
      Top             =   990
      Width           =   570
   End
   Begin VB.TextBox TxtAbsen 
      Height          =   315
      Index           =   4
      Left            =   7515
      TabIndex        =   12
      Top             =   675
      Width           =   570
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   255
      TabIndex        =   28
      Top             =   1365
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox TxtAbsen 
      Height          =   315
      Index           =   3
      Left            =   6000
      TabIndex        =   10
      Top             =   660
      Width           =   570
   End
   Begin VB.TextBox TxtTarget 
      Height          =   315
      Index           =   3
      Left            =   6000
      TabIndex        =   11
      Top             =   975
      Width           =   570
   End
   Begin VB.TextBox TxtAbsen 
      Height          =   315
      Index           =   2
      Left            =   4305
      TabIndex        =   8
      Top             =   675
      Width           =   600
   End
   Begin VB.TextBox TxtTarget 
      Height          =   315
      Index           =   2
      Left            =   4305
      TabIndex        =   9
      Top             =   990
      Width           =   600
   End
   Begin VB.TextBox TxtAbsen 
      Height          =   315
      Index           =   1
      Left            =   2670
      TabIndex        =   6
      Top             =   675
      Width           =   660
   End
   Begin VB.TextBox TxtTarget 
      Height          =   315
      Index           =   1
      Left            =   2670
      TabIndex        =   7
      Top             =   990
      Width           =   660
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   7065
      TabIndex        =   15
      Top             =   1365
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   315
      Left            =   6000
      TabIndex        =   14
      Top             =   1365
      Width           =   945
   End
   Begin VB.TextBox TxtTarget 
      Height          =   315
      Index           =   0
      Left            =   1230
      TabIndex        =   5
      Top             =   990
      Width           =   585
   End
   Begin VB.TextBox TxtAbsen 
      Height          =   315
      Index           =   0
      Left            =   1230
      TabIndex        =   4
      Top             =   675
      Width           =   585
   End
   Begin VB.ComboBox CmbBulan 
      Height          =   315
      Left            =   1050
      TabIndex        =   0
      Top             =   0
      Width           =   810
   End
   Begin VB.ComboBox CmbTahun 
      Height          =   315
      Left            =   1050
      TabIndex        =   1
      Top             =   315
      Width           =   945
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   3375
      TabIndex        =   3
      Top             =   330
      Width           =   3705
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   3375
      TabIndex        =   2
      Top             =   0
      Width           =   2130
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Target W5 :"
      Height          =   270
      Index           =   10
      Left            =   6690
      TabIndex        =   30
      Top             =   1005
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Hadir W5 :"
      Height          =   270
      Index           =   9
      Left            =   6705
      TabIndex        =   29
      Top             =   690
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Hadir W4 :"
      Height          =   270
      Index           =   8
      Left            =   5190
      TabIndex        =   27
      Top             =   675
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Target W4 :"
      Height          =   270
      Index           =   7
      Left            =   5175
      TabIndex        =   26
      Top             =   990
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Hadir W3 :"
      Height          =   270
      Index           =   6
      Left            =   3495
      TabIndex        =   25
      Top             =   690
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Target W3 :"
      Height          =   270
      Index           =   5
      Left            =   3480
      TabIndex        =   24
      Top             =   1005
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Hadir W2 :"
      Height          =   270
      Index           =   4
      Left            =   1890
      TabIndex        =   23
      Top             =   690
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Target W2 :"
      Height          =   270
      Index           =   3
      Left            =   1815
      TabIndex        =   22
      Top             =   1005
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Target W1 :"
      Height          =   270
      Index           =   2
      Left            =   360
      TabIndex        =   21
      Top             =   1005
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Hadir W1 :"
      Height          =   270
      Index           =   1
      Left            =   450
      TabIndex        =   20
      Top             =   690
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Bulan :"
      Height          =   270
      Left            =   270
      TabIndex        =   19
      Top             =   30
      Width           =   750
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Tahun :"
      Height          =   270
      Index           =   0
      Left            =   270
      TabIndex        =   18
      Top             =   345
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Agent Name :"
      Height          =   285
      Index           =   1
      Left            =   2130
      TabIndex        =   17
      Top             =   375
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Agent Code :"
      Height          =   285
      Index           =   0
      Left            =   2130
      TabIndex        =   16
      Top             =   45
      Width           =   1245
   End
End
Attribute VB_Name = "FrmUserTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ok As Boolean

Private Sub Combo1_Click(Index As Integer)
    Call Combo1_LostFocus(Index)
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim m_objrs As ADODB.Recordset
On Error GoTo Combo2_LostFocusErr
Select Case Index
    Case 0
        Set m_objrs = New ADODB.Recordset
        m_objrs.CursorLocation = adUseClient
        m_objrs.Open "Select * from USERTBL where USERID ='" + Combo1(0).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If Not m_objrs.EOF Then
            Combo1(0).Text = m_objrs!USERID
            Combo1(1).Text = m_objrs!agent
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 1
        Set m_objrs = New ADODB.Recordset
        m_objrs.CursorLocation = adUseClient
        m_objrs.Open "Select * from USERTBL where AGENT ='" + Combo1(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If Not m_objrs.EOF Then
            Combo1(0).Text = m_objrs!USERID
            Combo1(1).Text = m_objrs!agent
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
End Select
Set m_objrs = Nothing
Exit Sub
Combo2_LostFocusErr:
    MsgBox Err.Description
End Sub

Private Sub Command1_Click()
    ok = True
    Me.Hide
    
End Sub

Private Sub Command2_Click()
    ok = False
    Unload Me
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "SELECT * FROM USERTBL ORDER BY USERID", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    Combo1(0).AddItem m_objrs!USERID
    Combo1(1).AddItem m_objrs!USERID
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing

TxtAbsen(0).Text = 0
TxtAbsen(1).Text = 0
TxtAbsen(2).Text = 0
TxtAbsen(3).Text = 0
TxtAbsen(4).Text = 0
TxtTarget(0).Text = 0
TxtTarget(1).Text = 0
TxtTarget(2).Text = 0
TxtTarget(3).Text = 0
TxtTarget(4).Text = 0

CmbTahun.AddItem 2005
CmbTahun.AddItem 2006
CmbTahun.AddItem 2007
CmbTahun.AddItem 2008
CmbTahun.AddItem 2009
CmbTahun.AddItem 2010
For i = 1 To 12
    CmbBulan.AddItem i
Next i

End Sub


Private Sub TxtAbsen_LostFocus(Index As Integer)
If Trim(TxtAbsen(Index).Text) = "" Then
    TxtAbsen(Index).Text = 0
End If
End Sub

Private Sub TxtTarget_LostFocus(Index As Integer)
If Trim(TxtAbsen(Index).Text) = "" Then
    TxtAbsen(Index).Text = 0
End If
End Sub

