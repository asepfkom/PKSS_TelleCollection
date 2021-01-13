VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form FrmConfidenceAnalisisDetail 
   Caption         =   "Confidence Analisis Detail"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Payment"
      TabPicture(0)   =   "FrmConfidenceAnalisisDetail.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblPayment"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LstPayment"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "PTP"
      TabPicture(1)   =   "FrmConfidenceAnalisisDetail.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LstPTP"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "LblTotalPTP"
      Tab(1).ControlCount=   3
      Begin MSComctlLib.ListView LstPayment 
         Height          =   4620
         Left            =   180
         TabIndex        =   1
         Top             =   420
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   8149
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LstPTP 
         Height          =   4620
         Left            =   -74880
         TabIndex        =   2
         Top             =   540
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   8149
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Total PTP:"
         Height          =   315
         Left            =   -68100
         TabIndex        =   6
         Top             =   5220
         Width           =   1155
      End
      Begin VB.Label LblTotalPTP 
         Caption         =   "0"
         Height          =   315
         Left            =   -66780
         TabIndex        =   5
         Top             =   5220
         Width           =   1755
      End
      Begin VB.Label LblPayment 
         Caption         =   "0"
         Height          =   315
         Left            =   8280
         TabIndex        =   4
         Top             =   5280
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "Total Payment:"
         Height          =   315
         Left            =   6960
         TabIndex        =   3
         Top             =   5280
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmConfidenceAnalisisDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call HeaderPayment
    Call HeaderPTP
    Call IsiPayment
    Call IsiPtp
End Sub

Private Sub HeaderPayment()
    LstPayment.ColumnHeaders.ADD , , "AOC", 2000
    LstPayment.ColumnHeaders.ADD , , "Paydate", 2000
    LstPayment.ColumnHeaders.ADD , , "Amount", 3000
End Sub

Private Sub HeaderPTP()
    LstPTP.ColumnHeaders.ADD , , "AOC", 2000
    LstPTP.ColumnHeaders.ADD , , "Tgl.PTP", 2000
    LstPTP.ColumnHeaders.ADD , , "Amount", 2000
End Sub

Private Sub IsiPayment()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim Payment As Long
    
    CMDSQL = "select * from mandiri.tbllunas where agent='"
    CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.Text) + "' and custid is not null"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Payment = 0
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LstPayment.ListItems.ADD(, , M_objrs("custid"))
            ListItem.SubItems(1) = Format(M_objrs("paydate"), "dd-mm-yyyy")
            ListItem.SubItems(2) = Format(M_objrs("payment"), "##,###")
            Payment = Payment + Val(M_objrs("payment"))
            M_objrs.MoveNext
        Wend
    End If
    lblpayment.Caption = Format(Payment, "##,###")
    Set M_objrs = Nothing
End Sub


Private Sub IsiPtp()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim ptp As Long
    
    CMDSQL = "select * from mandiri.tblnegoptp where custid in (select custid from mandiri.mgm where agent='"
    CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.Text) + "')"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    ptp = 0
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LstPTP.ListItems.ADD(, , M_objrs("custid"))
            ListItem.SubItems(1) = Format(M_objrs("promisedate"), "dd-mm-yyyy")
            ListItem.SubItems(2) = Format(M_objrs("promisepay"), "##,###")
            ptp = ptp + Val(M_objrs("promisepay"))
            M_objrs.MoveNext
        Wend
    End If
    LblTotalPTP.Caption = Format(ptp, "##,###")
    Set M_objrs = Nothing
End Sub

