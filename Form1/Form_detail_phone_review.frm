VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_detail_phone_review 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "History Call"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4755
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Close"
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin MSComctlLib.ListView LvDetailReview 
      Height          =   5100
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
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
End
Attribute VB_Name = "Form_detail_phone_review"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call HeaderLv
    Call Isilv
End Sub

Private Sub HeaderLv()
    LvDetailReview.ColumnHeaders.ADD , , "No", 600
    LvDetailReview.ColumnHeaders.ADD , , "Call Date", 2000
    LvDetailReview.ColumnHeaders.ADD , , "Durasi", 1700
End Sub

Private Sub Isilv()
    Dim custid, sQuery, tgl_telfon, menit, detik, total_durasi As String
    Dim RS_Lv As ADODB.Recordset
    Dim num As Integer
    
    custid = Form_List_Phone_Review.LvPhoneReview.SelectedItem.ListSubItems(2)
    
    sQuery = "SELECT a.userid, a.custid, a. telpno, a.tgl, a.durasi, b.jumlah_call FROM tblphonemonitorhst a "
    sQuery = sQuery + " Left Join"
    sQuery = sQuery + " tbl_temp_telfon_review B"
    sQuery = sQuery + " ON"
    sQuery = sQuery + " a.custid = B.custid"
    sQuery = sQuery + " WHERE date(a.tgl) = '2016-02-12' AND a.custid = '" & custid & "'"
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    LvDetailReview.ListItems.CLEAR
    If RS_Lv.RecordCount > 0 Then
        num = 0
        Do Until RS_Lv.EOF
            num = num + 1
            tgl_telfon = Format(RS_Lv("tgl"), "dd-mm-yyyy hh:nn:ss")
            menit = RS_Lv("durasi") / 60
            detik = Split(menit, ",")(1)
            total_durasi = Split(menit, ",")(0) & ":" & Mid(detik, 1, 2)
            Set listItem = LvDetailReview.ListItems.ADD(, , num)
            listItem.SubItems(1) = tgl_telfon
            listItem.SubItems(2) = total_durasi
            RS_Lv.MoveNext
        Loop
    Else
        MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
    End If
End Sub
