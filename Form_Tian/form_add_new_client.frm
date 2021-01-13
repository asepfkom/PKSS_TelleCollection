VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form form_add_new_client 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Client"
   ClientHeight    =   3870
   ClientLeft      =   6300
   ClientTop       =   3030
   ClientWidth     =   4080
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4080
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   220
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin MSComctlLib.ListView Lvclient 
         Height          =   3060
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   5398
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
      Begin VB.Label Label2 
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   3360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "New Client"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Menu mnfile 
      Caption         =   "File"
      Begin VB.Menu mnbaca 
         Caption         =   "Baca"
      End
      Begin VB.Menu mnexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "form_add_new_client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
    Lvclient.ColumnHeaders.CLEAR
    With Lvclient.ColumnHeaders
        .ADD 1, , "User Name", 0 * TXT
        .ADD 2, , "DAFTAR CLIENT", 30 * TXT
    End With
End Sub

Private Sub create()
    qs = "select * from information_schema.columns where table_name = 'tbl_list_client_indium'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount = 0 Then
        M_OBJCONN.Execute "Create table tbl_list_client_indium (id serial not null, client varchar);"
        aa = Array("BCA", "BRI", "HCI", "MANDIRI", "MAYBANK", "GLOBAL", "COURT", "DANAMON")
        For i = 1 To 7
            M_OBJCONN.Execute "insert into tbl_list_client_indium (client) values ('" & aa(i) & "')"
        Next
    End If
End Sub

Private Sub selects()
    qs = "select * from tbl_list_client_indium order by 1"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    Lvclient.ListItems.CLEAR
    If M_objrs.RecordCount > 0 Then
        For i = 1 To M_objrs.RecordCount
            Set ListItem = Lvclient.ListItems.ADD(, , cnull(M_objrs("id")))
                ListItem.SubItems(1) = cnull(M_objrs("client"))
            M_objrs.MoveNext
        Next i
    End If
End Sub

Private Sub insert()
    M_OBJCONN.Execute "insert into tbl_list_client_indium (client) values ('" & Text1.text & "')"
    MsgBox "Client :" & Text1.text & ", berhasil diAdd"
    clear_load
End Sub

Private Sub DELETE()
    M_OBJCONN.Execute "delete from tbl_list_client_indium where id = " & Label2.Caption & " and client = '" & Text1.text & "'"
    MsgBox "Client :" & Text1.text & ", berhasil diDelete"
    clear_load
End Sub

Private Sub clear_load()
    Label2.Caption = ""
    Text1.text = ""
    Call selects
End Sub

Private Sub Command1_Click()
    Call insert
End Sub

Private Sub Command2_Click()
    Call DELETE
End Sub

Private Sub Form_Load()
    Call header
    Call create
    Call selects
End Sub

Private Sub Lvclient_DblClick()
    Text1.text = Lvclient.SelectedItem.SubItems(1)
    Label2.Caption = Lvclient.SelectedItem.text
End Sub

Private Sub how_to_use()
    cara_pakai = "Cara Pakai:" & vbCrLf
    cara_pakai = cara_pakai & "1.Untuk Save, cukup ketikan nama dari client yang akan di Add." & vbCrLf
    cara_pakai = cara_pakai & "2.Untuk Delete, double klik pada list data yang akan di delete." & vbCrLf & vbCrLf & vbCrLf
    cara_pakai = cara_pakai & "Tips Menggunakan:" & vbCrLf
    cara_pakai = cara_pakai & "Pastikan nama client adalah bagian dari Campaign Code." & vbCrLf
    cara_pakai = cara_pakai & "Contoh: Nama Client BRI, Campaign Code abcd_BRI_12345."

    MsgBox cara_pakai
End Sub

Private Sub mnbaca_Click()
    how_to_use
End Sub

Private Sub mnexit_Click()
    Unload Me
End Sub

Private Sub Text1_Change()
    Label2.Caption = ""
End Sub
