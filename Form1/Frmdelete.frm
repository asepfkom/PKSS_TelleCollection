VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Frmdelete 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5700
   ClientLeft      =   8910
   ClientTop       =   2370
   ClientWidth     =   4200
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   4200
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3255
      TabIndex        =   3
      Top             =   915
      Width           =   855
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3270
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin MSComctlLib.ListView LstPayment 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   661
      _Version        =   196610
      Font3D          =   4
      ForeColor       =   12582912
      Caption         =   "Delete PTP"
      BevelWidth      =   2
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
End
Attribute VB_Name = "Frmdelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim exdel As Boolean
Dim m_data As New ClsNegoPTP
Dim ListItem As ListItem
Dim showlist As New ADODB.Recordset
Dim TOTPTP As Currency
Private Sub cmddel_Click()
    Dim Cmdsql_Cek As String
    Dim M_Cek_Status As ADODB.Recordset
    
    If LstPayment.ListItems.Count = 0 Then
        Exit Sub
    End If
    m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
    If m_msgbox = 1 Then
        
        '@@ 11-04-2012 Cek status account terlebih dahulu, data bisa diedit jika status account PTP
        Cmdsql_Cek = "select f_cek_new from mandiri.mgm where custid='"
        Cmdsql_Cek = Cmdsql_Cek + FrmCC_Colection.lblcustid.Caption + "'"
        Set M_Cek_Status = New ADODB.Recordset
        M_Cek_Status.CursorLocation = adUseClient
        M_Cek_Status.Open Cmdsql_Cek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If Mid(M_Cek_Status("f_cek_new"), 1, 3) = "PTP" Then
            MsgBox "Data nego ptp tidak dapat dihapus jika status account=PTP!", vbOKOnly + vbExclamation, "Peringatan!"
            Set M_Cek_Status = Nothing
            Exit Sub
        End If
        
        If Mid(M_Cek_Status("f_cek_new"), 1, 2) = "BP" Then
            MsgBox "Data nego ptp tidak dapat dihapus jika status account=BP!", vbOKOnly + vbExclamation, "Peringatan!"
            Set M_Cek_Status = Nothing
            Exit Sub
        End If
        
        If Mid(M_Cek_Status("f_cek_new"), 1, 3) = "POP" Then
            MsgBox "Data nego ptp tidak dapat dihapus jika status account=POP!", vbOKOnly + vbExclamation, "Peringatan!"
            Set M_Cek_Status = Nothing
            Exit Sub
        End If
        
        
        m_data.DELETE_Nego_PTP M_OBJCONN, LstPayment.SelectedItem.SubItems(1)
        If m_data.ADD_OK Then
            LstPayment.ListItems.Remove LstPayment.SelectedItem.Index
        End If
    End If
    Exit Sub
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
LstPayment.ColumnHeaders.ADD 1, , "", 0 * TXT
LstPayment.ColumnHeaders.ADD 2, , "ID", 2 * TXT
LstPayment.ColumnHeaders.ADD 3, , "PROMISE DATE", 15 * TXT
LstPayment.ColumnHeaders.ADD 4, , "PAYMENT", 30 * TXT
LstPayment.ColumnHeaders.ADD 5, , "TYPE", 30 * TXT
LstPayment.ColumnHeaders.ADD 6, , "INPUT DATE", 30 * TXT

Call Show_NEGOPTP
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Public Sub Show_NEGOPTP()
Dim CMDSQL As String
Dim ssql As String
ssql = "SELECT CUSTID,sum(PAYMENT) as Jum FROM mandiri.tbllunas WHERE custid = '" + FrmCC_Colection.lblcustid.Caption + "' GROUP BY CUSTID"
showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If showlist.BOF And showlist.EOF Then
    TOTPTP = 0
Else
    TOTPTP = IIf(IsNull(showlist!jum), 0, showlist!jum)
End If


'@@ 09-04-2012, Filter Tanggal Dimatikan
CMDSQL = "SELECT * FROM mandiri.tblnegoPTP where custid = '" + Trim(FrmCC_Colection.lblcustid.Caption) + "' "
'Cmdsql = Cmdsql + " and date_part('month',promisedate)>=date_part('month',now()) and "
'Cmdsql = Cmdsql + " date_part('year',promisedate)=date_part('year',now()) "
CMDSQL = CMDSQL + "order by promisedate desc "

Set showlist = New ADODB.Recordset
showlist.CursorLocation = adUseClient
showlist.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

Dim n As Currency
If exdel Then
    With FrmCC_Colection
        .LstPayment.ListItems.CLEAR
        While Not showlist.EOF
            Set ListItem = .LstPayment.ListItems.ADD(, , "")
            Call showlst
        Wend
    End With
Else
    LstPayment.ListItems.CLEAR
    While Not showlist.EOF
        Set ListItem = LstPayment.ListItems.ADD(, , "")
        Call showlst
    Wend
End If
Set showlist = Nothing
exdel = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
exdel = True
Call Show_NEGOPTP
End Sub

Sub showlst()
'listitem.SubItems(1) = ""
ListItem.SubItems(1) = CStr(IIf(IsNull(showlist!ID), "", (showlist!ID)))
ListItem.SubItems(2) = CStr(IIf(IsNull(showlist!PromiseDate), "", Format(showlist!PromiseDate, "dd/mm/yyyy")))
ListItem.SubItems(3) = CStr(IIf(IsNull(showlist!PromisePay), "", Round((showlist!PromisePay), 0)))
n = n + Val(ListItem.SubItems(3))
If n <= TOTPTP Then
    ListItem.ListSubItems(1).ForeColor = vbRed
    ListItem.ListSubItems(2).ForeColor = vbRed
    ListItem.ListSubItems(3).ForeColor = vbRed
End If
ListItem.SubItems(4) = IIf(IsNull(showlist!Type), "", showlist!Type)
ListItem.SubItems(5) = CStr(IIf(IsNull(showlist!inputdate), "", Format(showlist!inputdate, "dd/mm/yyyy")))
showlist.MoveNext
End Sub
