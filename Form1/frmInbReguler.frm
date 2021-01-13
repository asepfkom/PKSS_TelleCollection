VERSION 5.00
Begin VB.Form frmInbReguler 
   Caption         =   "Inbound Reguler"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4650
   Icon            =   "frmInbReguler.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   2940
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   345
      Index           =   1
      Left            =   3495
      TabIndex        =   12
      Top             =   2370
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   345
      Index           =   0
      Left            =   2520
      TabIndex        =   11
      Top             =   2370
      Width           =   945
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Top             =   1815
      Width           =   1755
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   3030
      TabIndex        =   8
      Top             =   1500
      Width           =   1005
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   1500
      Width           =   1545
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   1185
      Width           =   1545
   End
   Begin VB.TextBox Text2 
      Height          =   765
      Left            =   1440
      TabIndex        =   2
      Top             =   405
      Width           =   2985
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   75
      Width           =   2505
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Hp :"
      Height          =   315
      Left            =   30
      TabIndex        =   10
      Top             =   1830
      Width           =   1395
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Telp Kantor :"
      Height          =   315
      Left            =   30
      TabIndex        =   7
      Top             =   1500
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Telp Rumah :"
      Height          =   315
      Left            =   30
      TabIndex        =   6
      Top             =   1185
      Width           =   1395
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Alamat :"
      Height          =   315
      Left            =   30
      TabIndex        =   5
      Top             =   405
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nama Customer :"
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "frmInbReguler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Dim m_objrs As ADODB.Recordset
    Select Case Index
        Case 0
            Set m_objrs = New ADODB.Recordset
            m_objrs.Open "Select * from InbReguler where CustId =0", M_OBJCONN, adOpenDynamic, adLockOptimistic, admcdtext
            m_objrs.AddNew
            m_objrs!NAMA = Text1.Text
            m_objrs!Alamat = Text2.Text
            m_objrs!TelpRumah = Text3.Text
            m_objrs!TelpKantor = Text4.Text
            m_objrs!ExtKantor = Text5.Text
            m_objrs!HP = Text6.Text
            m_objrs!AOC = MDIForm1.Text1.Text
            m_objrs!AgentName = MDIForm1.Text7.Text
            m_objrs!TGL = Format(MDIForm1.TDBDate1.Text, "MM/DD/YYYY")
            m_objrs.UPDATE
            Unload Me
        Case 1
            Unload Me
    End Select
End Sub

