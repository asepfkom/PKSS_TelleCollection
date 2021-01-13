VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
    Call isilv
End Sub

Private Sub HeaderLv()
    LvDetailReview.ColumnHeaders.ADD , , "No", 600
    LvDetailReview.ColumnHeaders.ADD , , "Call Date", 2000
    LvDetailReview.ColumnHeaders.ADD , , "Durasi", 1700
End Sub

Private Sub isilv()
    Dim CustId, notlp, sQuery, tgl_telfon, total_durasi, durasi As String
    Dim a, B, JAM, Menit, Detik As Long
    Dim RS_Lv As ADODB.Recordset
    Dim num As Integer
    
    CustId = Form_List_Phone_Review.LvPhoneReview.SelectedItem.ListSubItems(2)
    notlp = Form_List_Phone_Review.LvPhoneReview.SelectedItem.ListSubItems(3)
    
'    sQuery = "SELECT a.userid, a.custid, a. telpno, a.tgl, a.durasi :: numeric, b.jumlah_call "
'    sQuery = sQuery + " FROM tblphonemonitorhst a  "
'    sQuery = sQuery + " LEFT JOIN"
'    sQuery = sQuery + " (SELECT custid, jumlah_call, no_telfon, tanggal_telfon FROM tbl_temp_telfon_review "
'    sQuery = sQuery + " WHERE custid = '" & CustId & "' "
'    sQuery = sQuery + " AND date(tanggal_telfon) = '" & Format(waktu_server_sekarang, "yyyy-mm-dd") & "'"
'    sQuery = sQuery + " AND no_telfon = '" & notlp & "') B "
'    sQuery = sQuery + " ON"
'    sQuery = sQuery + " a.custid = B.custid"
'    sQuery = sQuery + " WHERE date(a.tgl) = '" & Format(waktu_server_sekarang, "yyyy-mm-dd") & "' AND a.custid = '" & CustId & "'"
'    sQuery = sQuery + " AND a.telpno = '" & notlp & "' ORDER BY a.tgl "

    sQuery = "SELECT * from mandiri.tblphonemonitorhst where date(tgl)= '" & Format(waktu_server_sekarang, "yyyy-mm-dd") & "' AND telpno = '" & notlp & "' AND custid = '" & CustId & "' ORDER BY tgl limit 5"
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    LvDetailReview.ListItems.CLEAR
    If RS_Lv.RecordCount > 0 Then
        num = 0
        Do Until RS_Lv.EOF
            num = num + 1
            tgl_telfon = Format(RS_Lv("tgl"), "dd-mm-yyyy hh:nn:ss")
            a = IIf(IsNull(RS_Lv("durasi")), "0", RS_Lv("durasi"))
            JAM = Int(a / 3600)
            B = JAM * 3600
            Menit = Int((a - B) / 60)
            Detik = a - B - (Menit * 60)
            'menit = RS_Lv("durasi") / 60
            'detik = Split(menit, ".")(1)
            'total_durasi = Split(menit, ".")(0) & ":" & Mid(detik, 1, 2)
            total_durasi = JAM & ":" & Menit & ":" & Detik
            Set ListItem = LvDetailReview.ListItems.ADD(, , num)
            ListItem.SubItems(1) = tgl_telfon
            ListItem.SubItems(2) = total_durasi
            RS_Lv.MoveNext
        Loop
    Else
        MsgBox "Data Not Found !", vbOKOnly + vbInformation, "Info"
    End If
End Sub

