VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_List_Phone_Review 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form List Review Number"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10425
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "Search . . ."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   780
      Width           =   975
   End
   Begin VB.TextBox txt_cust 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   6
      Top             =   800
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CheckBox chk_all 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   130
      TabIndex        =   1
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmd_release 
      BackColor       =   &H0080FF80&
      Caption         =   "Release"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   130
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   1815
   End
   Begin MSComctlLib.ListView LvPhoneReview 
      Height          =   5100
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   8996
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
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
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer ID :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   130
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "List Phone Review"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   630
      TabIndex        =   4
      Top             =   0
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   4
      Left            =   120
      Picture         =   "Form_List_Phone_Review.frx":0000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "Form_List_Phone_Review.frx":0B0A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10440
   End
End
Attribute VB_Name = "Form_List_Phone_Review"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_release_Click()
    Call ReleaseReview
End Sub

Private Sub ReleaseReview()
    Dim w As Integer
    Dim a As String
    Dim cek As Integer
    Dim K As Integer
    Dim cmdsql As String
    
    If LvPhoneReview.ListItems.Count = 0 Then
        MsgBox "Data Is Empty!", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    cek = 0
    
    For K = 1 To LvPhoneReview.ListItems.Count
        If LvPhoneReview.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
    a = MsgBox("Are You Sure To Release This Number?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Canceled!", vbOKOnly + vbInformation, "Perhatian"
        Exit Sub
    End If
    
    If cek = 0 Then
        MsgBox "You Must Select a Phone Number!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For w = 1 To LvPhoneReview.ListItems.Count
        If LvPhoneReview.ListItems(w).Checked = True Then
            cmdsql = "DELETE FROM tbl_temp_telfon_review WHERE id ='"
            cmdsql = cmdsql + Trim(LvPhoneReview.ListItems(w).SubItems(6)) + "'"
            M_OBJCONN.Execute cmdsql
        End If
    Next w
    
    txt_cust.Text = ""
    Call Isilv
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Call Isilv
End Sub

Private Sub Form_Load()
    Call HeaderLv
    Call Isilv
End Sub

Private Sub HeaderLv()
    LvPhoneReview.ColumnHeaders.ADD , , "No", 600
    LvPhoneReview.ColumnHeaders.ADD , , "Agent", 1100
    LvPhoneReview.ColumnHeaders.ADD , , "Customer ID", 3300
    LvPhoneReview.ColumnHeaders.ADD , , "Phone Number", 2400
    LvPhoneReview.ColumnHeaders.ADD , , "Call Date", 2000
    LvPhoneReview.ColumnHeaders.ADD , , "Count", 700
    LvPhoneReview.ColumnHeaders.ADD , , "ID", 0
End Sub

Private Sub Isilv()
    Dim custid, sQuery, where, tgl_telfon As String
    Dim RS_Lv As ADODB.Recordset
    Dim num As Integer
    
    custid = txt_cust.Text
    
    If txt_cust.Text <> "" Then
        where = " AND custid = '" & custid & "' "
    End If
    
    sQuery = "SELECT * FROM tbl_temp_telfon_review WHERE jumlah_call >= 5 AND date(tanggal_telfon) = '" & Format(waktu_server_sekarang, "yyyy-mm-dd") & "'  "
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery + where, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    LvPhoneReview.ListItems.CLEAR
    If RS_Lv.RecordCount > 0 Then
        num = 0
        Do Until RS_Lv.EOF
            num = num + 1
            tgl_telfon = Format(RS_Lv("tanggal_telfon"), "dd-mm-yyyy")
            Set listItem = LvPhoneReview.ListItems.ADD(, , num)
            listItem.SubItems(1) = Trim(cnull(RS_Lv("agent")))
            listItem.SubItems(2) = Trim(cnull(RS_Lv("custid")))
            listItem.SubItems(3) = Trim(cnull(RS_Lv("no_telfon")))
            listItem.SubItems(4) = tgl_telfon
            listItem.SubItems(5) = Trim(cnull(RS_Lv("jumlah_call")))
            listItem.SubItems(6) = Trim(cnull(RS_Lv("id")))
            RS_Lv.MoveNext
        Loop
    Else
        MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
    End If
End Sub

Private Sub LvPhoneReview_DblClick()
    Form_detail_phone_review.Show vbModal
End Sub
