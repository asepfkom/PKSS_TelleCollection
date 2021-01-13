VERSION 5.00
Begin VB.Form FrmConfidenceAnalysis 
   Caption         =   "Confidence Analysis"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3270
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   435
      Left            =   5760
      TabIndex        =   5
      Top             =   2700
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "[ Klik disini untuk detail  ... ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   4740
      TabIndex        =   4
      Top             =   2100
      Width           =   2535
   End
   Begin VB.Label LblTotalPTP 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "PTP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   3
      Top             =   1320
      Width           =   7215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "Total PTP Anda Sampai Hari ini:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   2
      Top             =   900
      Width           =   7215
   End
   Begin VB.Label LblTotalPayment 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Total Payment Anda Sampai Hari ini:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "FrmConfidenceAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim m_objrs_ptp As ADODB.Recordset
    Dim cmdsql_ptp As String
    Dim cmdsql_payment As String
    
    cmdsql_payment = "select sum(payment) as payment from mandiri.tbllunas  where agent='"
    cmdsql_payment = cmdsql_payment + Trim(MDIForm1.TxtUsername.Text) + "'"
    cmdsql_payment = cmdsql_payment + " group by agent"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql_payment, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LblTotalPayment.Caption = Format(M_objrs(0), "##,###")
    
    Set M_objrs = Nothing
    
    cmdsql_ptp = "select sum(promisepay) as promisepay from mandiri.tblnegoptp "
    cmdsql_ptp = cmdsql_ptp + "where custid in (select custid from mandiri.mgm where agent='"
    cmdsql_ptp = cmdsql_ptp + Trim(MDIForm1.TxtUsername.Text) + "')"
    
    Set m_objrs_ptp = New ADODB.Recordset
    m_objrs_ptp.CursorLocation = adUseClient
    m_objrs_ptp.Open cmdsql_ptp, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    LblTotalPTP.Caption = Format(m_objrs_ptp(0), "##,###")
    Set m_objrs_ptp = Nothing
End Sub

Private Sub Label4_Click()
    FrmConfidenceAnalisisDetail.Show vbModal
End Sub
