VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_reminder 
   BackColor       =   &H008389EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reminder PTP"
   ClientHeight    =   5265
   ClientLeft      =   3255
   ClientTop       =   2100
   ClientWidth     =   8475
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H008389EC&
      Caption         =   "REMINDER PTP"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4755
      Left            =   255
      TabIndex        =   0
      Top             =   165
      Width           =   7815
      Begin MSComctlLib.ListView lvreminder 
         Height          =   4260
         Left            =   90
         TabIndex        =   1
         Top             =   390
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   7514
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "Form_reminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Call header
    Call tampil_reminder
End Sub


Private Sub tampil_reminder()
 Dim CMDSQL As String
 Dim M_objrs As ADODB.Recordset
 Dim list As ListItem

If MDIForm1.txtlevel.text = "Agent" Then
    CMDSQL = "select * from tblnegoptp_log where date(promisedate)=date(now())+1 and agent='" + MDIForm1.TxtUsername.text + "'"
Else
    CMDSQL = "select * from tblnegoptp_log where date(promisedate)=date(now())+1 "
End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    no = 0
lvreminder.ListItems.CLEAR
While Not M_objrs.EOF
     no = no + 1
     Set list = lvreminder.ListItems.ADD(, , no)
     list.SubItems(1) = IIf(IsNull(M_objrs!CustId), "", M_objrs!CustId)
     list.SubItems(2) = IIf(IsNull(M_objrs!PromiseDate), "", M_objrs!PromiseDate)
     list.SubItems(3) = IIf(IsNull(M_objrs!PromisePay), "", M_objrs!PromisePay)
     list.SubItems(4) = IIf(IsNull(M_objrs!AGENT), "", M_objrs!AGENT)
     M_objrs.MoveNext
Wend
Set M_objrs = Nothing
End Sub


Private Sub header()
    lvreminder.ColumnHeaders.ADD 1, , "No", 5 * TXT
    lvreminder.ColumnHeaders.ADD 2, , "Customer ID", 15 * TXT
    lvreminder.ColumnHeaders.ADD 3, , "PTP Date", 13 * TXT
    lvreminder.ColumnHeaders.ADD 4, , "Payment", 15 * TXT
    lvreminder.ColumnHeaders.ADD 5, , "Agent", 15 * TXT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MDIForm1.Label6.Caption = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If sReminder_CUST_ID = "" Then
        MDIForm1.Timer100.Enabled = True
    End If
End Sub

Private Sub lvreminder_DblClick()
    sReminder_CUST_ID = lvreminder.SelectedItem.SubItems(1)
    Unload Me
    If bAktif_form_customer Then
        Unload FrmCC_Colection
    End If
    bReminder_agent = True
    FrmCC_Colection.Show vbModal
End Sub

