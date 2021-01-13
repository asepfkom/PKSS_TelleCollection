VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_unread_msg 
   BackColor       =   &H008389EC&
   BorderStyle     =   0  'None
   Caption         =   "NEW MESSAGE"
   ClientHeight    =   5235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10590
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   345
      Left            =   8400
      TabIndex        =   4
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Frame FrameInboxOutbox 
      BackColor       =   &H008389EC&
      Caption         =   "UNREAD MESSAGE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2835
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   10155
      Begin MSComctlLib.ListView LvInboxOutbox 
         Height          =   2400
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   4233
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H008389EC&
      Caption         =   "DETAIL MESSAGE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1635
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   10155
      Begin VB.TextBox TxtDetailMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   9855
      End
   End
End
Attribute VB_Name = "frm_unread_msg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs As ADODB.Recordset

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call HeaderInbox
    Call create_connection
    'Call show_inbox
End Sub
Private Sub HeaderInbox()
    LvInboxOutbox.ColumnHeaders.ADD , , "Custid", 1700
    LvInboxOutbox.ColumnHeaders.ADD , , "Name", 2000
    LvInboxOutbox.ColumnHeaders.ADD , , "No.Handphone", 2000
    LvInboxOutbox.ColumnHeaders.ADD , , "Date Time", 1500
    LvInboxOutbox.ColumnHeaders.ADD , , "Message", 3000
    LvInboxOutbox.ColumnHeaders.ADD , , "Id", 0
    LvInboxOutbox.ColumnHeaders.ADD , , "Status", 1000
End Sub
Private Sub show_inbox()
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM tbl_notif_sms WHERE agent='" & MDIForm1.TxtUsername.Text & "' and f_read='0'"
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            Set lst = LvInboxOutbox.ListItems.ADD(, , cnull(rs!CustId)) 'custid
            lst.SubItems(1) = Trim(isiname)  'Isi nama
            lst.SubItems(2) = cnull(rs!sender_number) 'Telepon
            lst.SubItems(3) = cnull(rs!received_sms_Date) 'Receivingdatetime
            lst.SubItems(4) = cnull(rs!text_sms) 'Textsms
            lst.SubItems(5) = cnull(rs!id_sms)
            rs.MoveNext
        Wend
        ' HAPUS ALL NOTIF AGENT // STOP REMINDER
        M_OBJCONN.Execute "update tbl_notif_sms set f_read='1' WHERE agent='" & MDIForm1.TxtUsername.Text & "' and f_read='0'"
        'M_OBJCONN.Execute "DELETE FROM tbl_notif_sms WHERE agent='" & MDIForm1.TxtUsername.Text & "'"
    End If
End Sub

Private Sub create_header()
    LvInboxOutbox.ColumnHeaders.ADD , , "Custid", 1700
    LvInboxOutbox.ColumnHeaders.ADD , , "Name", 2000
    LvInboxOutbox.ColumnHeaders.ADD , , "No.Handphone", 2000
    LvInboxOutbox.ColumnHeaders.ADD , , "Date Time", 2000
    LvInboxOutbox.ColumnHeaders.ADD , , "Message", 3000
    LvInboxOutbox.ColumnHeaders.ADD , , "Id", 0
    LvInboxOutbox.ColumnHeaders.ADD , , "Status", 1000
End Sub

Private Sub create_connection()
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    rs.ActiveConnection = M_OBJCONN1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
End Sub

Private Sub LvInboxOutbox_Click()
    Dim SisaSms As Integer
    Dim ListItem As ListItem
    
    If LvInboxOutbox.ListItems.Count <> 0 Then
        TxtDetailMsg.Text = LvInboxOutbox.SelectedItem.SubItems(4)
        'Update status
        CMDSQL = "UPDATE inbox SET `Processed`='' WHERE `ID`='"
        CMDSQL = CMDSQL + Trim(LvInboxOutbox.SelectedItem.SubItems(5)) + "'"
        LvInboxOutbox.SelectedItem.ForeColor = vbBlack
        M_OBJCONN1.Execute CMDSQL
        'ssql = " UPDATE inbox SET `SenderNumber`='0'||substr(trim(`SenderNumber`),4) WHERE `SenderNumber` like '+62%' AND length(`SenderNumber`)>10"
        'ssql = " UPDATE inbox SET `SenderNumber` = replace(`SenderNumber`,'+62','0')"
'        M_OBJCONN.Execute (ssql)
    End If
End Sub

Private Sub TxtDetailMsg_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
