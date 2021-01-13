VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Frm_ProdKnowledg 
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   Icon            =   "Frm_ProdKnowledg.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5715
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Bulletin Board"
      Height          =   2550
      Left            =   -15
      TabIndex        =   4
      Top             =   2475
      Width           =   6570
      Begin VB.CommandButton Command4 
         Caption         =   "&Show"
         Height          =   360
         Left            =   5760
         TabIndex        =   5
         Top             =   240
         Width           =   645
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2190
         Left            =   30
         TabIndex        =   6
         Top             =   240
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   3863
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   330
      Left            =   6000
      TabIndex        =   3
      Top             =   5280
      Width           =   720
   End
   Begin VB.Frame Frame1 
      Caption         =   "Product Knowledge"
      Height          =   2430
      Left            =   -15
      TabIndex        =   0
      Top             =   15
      Width           =   6555
      Begin VB.CommandButton Command2 
         Caption         =   "&Show"
         Height          =   360
         Left            =   5760
         TabIndex        =   2
         Top             =   240
         Width           =   645
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2190
         Left            =   30
         TabIndex        =   1
         Top             =   210
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   3863
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "Frm_ProdKnowledg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
If ListView1.ListItems.Count < 1 Then
    Exit Sub
End If
'    If StartMeUp(ListView1.SelectedItem.SubItems(1)) <= 32 Then
'        MsgBox "File Tidak Ada / File Tidak Bisa Dibuka"
'    End If
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Keterangan", 50 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Lokasi File", 1 * 1
End Sub

Private Sub Command4_Click()
If ListView2.ListItems.Count < 1 Then
    Exit Sub
End If
'    If StartMeUp(ListView2.SelectedItem.SubItems(2)) <= 32 Then
'        MsgBox "File Tidak Ada / File Tidak Bisa Dibuka"
'    End If
'     With MDIForm1.CommonDialog1
 '         .HelpFile = ListView2.SelectedItem.SubItems(2)
 '         .DialogTitle = ListView2.SelectedItem.SubItems(1)
 '         .HelpCommand = cdlHelpContents
 '         .ShowHelp
 '    End With
End Sub

Private Sub Form_Load()
    Dim m_objrs As ADODB.Recordset
    Dim LISTITEM As LISTITEM
    Call header
    Call header1
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from BuletinBoard", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not m_objrs.EOF
    Set LISTITEM = ListView1.ListItems.ADD(, , m_objrs("Keterangan"))
        LISTITEM.SubItems(1) = m_objrs("LokasiFile")
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from BulLetinBoard1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not m_objrs.EOF
    Set LISTITEM = ListView2.ListItems.ADD(, , m_objrs("Tanggal"))
        LISTITEM.SubItems(1) = m_objrs("Subject")
        LISTITEM.SubItems(2) = m_objrs("LokasiFile")
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing

End Sub

Private Sub header1()
    ListView2.ColumnHeaders.ADD 1, , "Tanggal", 10 * TXT
    ListView2.ColumnHeaders.ADD 2, , "Subject", 40 * TXT
    ListView2.ColumnHeaders.ADD 3, , "Lokasi File", 1 * 1
End Sub

