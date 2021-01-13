VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmdeletereserve 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Reserve"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4500
      TabIndex        =   0
      Top             =   1620
      Width           =   855
   End
   Begin VB.CommandButton CmdUncekAll 
      Caption         =   "&UnCek All"
      Height          =   375
      Left            =   4500
      TabIndex        =   5
      Top             =   780
      Width           =   855
   End
   Begin VB.CommandButton cmdcekall 
      Caption         =   "&Cek All"
      Height          =   375
      Left            =   4500
      TabIndex        =   4
      Top             =   420
      Width           =   855
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4500
      TabIndex        =   1
      Top             =   1260
      Width           =   855
   End
   Begin MSComctlLib.ListView LstPayment 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
      TabIndex        =   3
      Top             =   0
      Width           =   5400
      _ExtentX        =   9525
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
Attribute VB_Name = "frmdeletereserve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim exdel As Boolean
Dim m_data As New ClsNegoPTP
Dim ListItem As ListItem
Dim showlist As New ADODB.Recordset
Dim TOTPTP As Currency

Private Sub CmdCekAll_Click()
    Dim K As Integer
    
    If LstPayment.ListItems.Count = 0 Then
        MsgBox "Tidak ada data reserve ptp!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For K = 1 To LstPayment.ListItems.Count
        LstPayment.ListItems(K).Checked = True
    Next K
End Sub

Private Sub cmddel_Click()
    If LstPayment.ListItems.Count = 0 Then
        Exit Sub
    End If
    
'        m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
'        If m_msgbox = 1 Then
'            M_DATA.DELETE_Nego_Reserve M_OBJCONN, LstPayment.SelectedItem.SubItems(1)
'            If M_DATA.ADD_OK Then
'                LstPayment.ListItems.Remove LstPayment.SelectedItem.Index
'            End If
'        End If
'        Exit Sub
    
    m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
    If m_msgbox = 1 Then
        For i = 1 To LstPayment.ListItems.Count
          If LstPayment.ListItems(i).Checked = True Then
            CMDSQL = "delete from tblreserve where id='"
            CMDSQL = CMDSQL + LstPayment.ListItems(i).SubItems(1) + "'"
            M_OBJCONN.Execute CMDSQL
          End If
        Next i
    End If
    
    MsgBox "Data berhasil dihapus!", vbOKOnly + vbInformation, "Informasi"
    Unload Me
End Sub

Private Sub CmdUnCekAll_Click()
    Dim K As Integer
    
    If LstPayment.ListItems.Count = 0 Then
        MsgBox "Tidak ada data reserve ptp!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For K = 1 To LstPayment.ListItems.Count
        LstPayment.ListItems(K).Checked = False
    Next K
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
LstPayment.ColumnHeaders.ADD 1, , "", 3 * TXT
LstPayment.ColumnHeaders.ADD 2, , "ID", 5 * TXT
LstPayment.ColumnHeaders.ADD 3, , "PROMISE DATE", 15 * TXT
LstPayment.ColumnHeaders.ADD 4, , "PAYMENT", 30 * TXT
LstPayment.ColumnHeaders.ADD 5, , "TYPE", 30 * TXT
LstPayment.ColumnHeaders.ADD 6, , "INPUT DATE", 30 * TXT

Call Show_Reserve
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Public Sub Show_Reserve()
Dim CMDSQL As String
Dim ssql As String
ssql = "SELECT CUSTID,sum(PAYMENT) as Jum FROM mandiri.tbllunas WHERE custid = '" + FrmCC_Colection.lblcustid.Caption + "' GROUP BY CUSTID"
showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If showlist.BOF And showlist.EOF Then
    TOTPTP = 0
Else
    TOTPTP = IIf(IsNull(showlist!jum), 0, showlist!jum)
End If


CMDSQL = "SELECT * FROM mandiri.tblreserve where custid = '" + FrmCC_Colection.lblcustid.Caption + "' order by promisedate"

Set showlist = New ADODB.Recordset
showlist.CursorLocation = adUseClient
showlist.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

Dim n As Currency
If exdel Then
    With FrmCC_Colection
        .LstReserve.ListItems.CLEAR
        While Not showlist.EOF
            Set ListItem = .LstReserve.ListItems.ADD(, , "")
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
Call Show_Reserve
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

