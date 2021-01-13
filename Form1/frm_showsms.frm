VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_showsms 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11400
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   9600
      TabIndex        =   5
      Top             =   6120
      Width           =   1695
   End
   Begin MSComctlLib.ListView LstSMS 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   11175
      _ExtentX        =   19711
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
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView LstSMS1 
      Height          =   5055
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   11175
      _ExtentX        =   19711
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
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   6120
      Width           =   10095
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   5880
      Width           =   10095
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   5640
      Width           =   10095
   End
   Begin VB.Label Label6 
      Caption         =   "CH : "
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "ISI PESAN : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "NO TELP : "
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "INBOX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10080
      TabIndex        =   2
      Top             =   75
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SMS LAMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SMS BARU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      Height          =   6615
      Left            =   0
      Top             =   0
      Width           =   11400
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   6360
      Width           =   10095
   End
   Begin VB.Label Label7 
      Caption         =   "NAMA CH : "
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6360
      Width           =   1215
   End
End
Attribute VB_Name = "frm_showsms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HEADER_SendSMS()
LstSMS.ColumnHeaders.ADD 1, , "Tanggal Terima", 25 * TXT
LstSMS.ColumnHeaders.ADD 2, , "No Telp", 25 * TXT
LstSMS.ColumnHeaders.ADD 3, , "Pesan", 25 * TXT
LstSMS.ColumnHeaders.ADD 4, , "Custid", 25 * TXT
LstSMS.ColumnHeaders.ADD 5, , "Nama", 25 * TXT

LstSMS1.ColumnHeaders.ADD 1, , "Tanggal Terima", 25 * TXT
LstSMS1.ColumnHeaders.ADD 2, , "No Telp", 25 * TXT
LstSMS1.ColumnHeaders.ADD 3, , "Pesan", 25 * TXT
LstSMS1.ColumnHeaders.ADD 4, , "Custid", 25 * TXT
LstSMS1.ColumnHeaders.ADD 5, , "Nama", 25 * TXT
End Sub

Function ReplaceFirstInstance(SourceString, _
Searchstring, Replacestring)
Dim StartLoc
Dim FoundLoc
  If StartLoc = 0 Then StartLoc = 1
  FoundLoc = InStr(StartLoc, SourceString, Searchstring) '*
  If FoundLoc <> 0 And FoundLoc < 2 Then
     ReplaceFirstInstance = Left(SourceString, FoundLoc - 1) & Replacestring & Right(SourceString, Len(SourceString) - (FoundLoc - 1) - Len(Searchstring))
     StartLoc = FoundLoc + Len(Replacestring)
  ElseIf FoundLoc > 1 Then
  
      ReplaceFirstInstance = Replacestring & "21" & SourceString

  Else
     StartLoc = 1

    ReplaceFirstInstance = SourceString
  End If
End Function

Function FindReplace(SourceString, Searchstring, Replacestring) As String
  Dim tmpString1
  Dim tmpString2
  tmpString1 = SourceString
 
      tmpString2 = tmpString1
      tmpString1 = ReplaceFirstInstance(tmpString1, _
                   Searchstring, Replacestring)
      
      FindReplace = tmpString1
End Function

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call HEADER_SendSMS
Call isisms

Label1 = MDIForm1.Label9
Label2 = MDIForm1.Label10

End Sub

Private Sub Label1_Click()
LstSMS1.Visible = False
LstSMS.Visible = True
Label1.BackColor = vbBlack
Label2.BackColor = vbWhite
Label1.ForeColor = vbWhite
Label2.ForeColor = vbBlack

End Sub

Private Sub Label2_Click()
LstSMS1.Visible = True
LstSMS.Visible = False
Label1.BackColor = vbWhite
Label2.BackColor = vbBlack
Label1.ForeColor = vbBlack
Label2.ForeColor = vbWhite

End Sub
Private Sub isisms()

Dim satu As String
Dim dua As String
Dim tiga As String
Dim empat As String

'On Error Resume Next
'Dim ConnPTP As New ADODB.Connection
Dim M_objrs As New ADODB.Recordset
Dim cmdsql34 As String
Dim TELPo As String

'
TELPo = "Select receivingdatetime, sendernumber, textdecoded  from mandiri.inbox where sendernumber in ("

Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mandiri.mgm where agent = '" + MDIForm1.TxtUsername.Text + "'"
M_objrs.Open cmdsql34, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_objrs.EOF

If Len(M_objrs("mobileno")) <> 0 Then
satu = FindReplace(M_objrs("mobileno"), "0", "+62")
TELPo = TELPo + "'" + satu + "',"
Else
TELPo = TELPo
End If

If Len(M_objrs("mobileno2")) <> 0 Then
dua = FindReplace(M_objrs("mobileno2"), "0", "+62")
TELPo = TELPo + "'" + dua + "',"
Else
TELPo = TELPo
End If

If Len(M_objrs("mobilenoadd1")) <> 0 Then
tiga = FindReplace(M_objrs("mobilenoadd1"), "0", "+62")
TELPo = TELPo + "'" + tiga + "',"
Else
TELPo = TELPo
End If

If Len(M_objrs("mobilenoadd2")) <> 0 Then
empat = FindReplace(M_objrs("mobilenoadd2"), "0", "+62")
TELPo = TELPo + "'" + empat + "',"
Else
TELPo = TELPo
End If

M_objrs.MoveNext
Wend

Set M_objrs = Nothing

TELPo = Left(TELPo, Len(TELPo) - 1)
Dim TELPo1
Dim TELPo2

TELPo1 = TELPo + ") and processed='f'"
TELPo2 = TELPo + ") and processed='t'"

'cmdsql34 = "select mobileno from mgm where agent = '" + Text1 + "'"
M_objrs.Open TELPo1, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_objrs.EOF
 S = Format(M_objrs!receivingdatetime, "yyyy-mm-dd hh:mm:ss")
t = Trim(M_objrs!sendernumber)
u = M_objrs!textdecoded
v = FindReplace(t, "+62", "0")

If (Left(v, 3) = "031") Then
    v = Mid(v, 4, 20)
End If

'----------------------------------
Dim showlist As New ADODB.Recordset
Dim TOTPTP As Currency
Dim ssql As String
ssql = "SELECT custid, name FROM mandiri.mgm WHERE mobileno='" & v & "'  or mobileno2='" & v & "'  or mobilenoadd1='" & v & "'  or mobilenoadd2='" & v & "'"
showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
isicustid = showlist!CustId
isiname = showlist!Name
Set showlist = Nothing

'LstSMS.ColumnHeaders.ADD 1, , "Tanggal Terima", 25 * TXT
'LstSMS.ColumnHeaders.ADD 2, , "No Telp", 25 * TXT
'LstSMS.ColumnHeaders.ADD 3, , "Pesan", 25 * TXT
'LstSMS.ColumnHeaders.ADD 4, , "Custid", 25 * TXT
'LstSMS.ColumnHeaders.ADD 5, , "Nama", 25 * TXT
'
'

Set lst = LstSMS.ListItems.ADD(, , S)
        lst.SubItems(1) = v
        lst.SubItems(2) = IIf(IsNull(M_objrs("textdecoded")), "", M_objrs("textdecoded"))
        lst.SubItems(3) = isicustid
        lst.SubItems(4) = isiname
       ' Label1 = "SMS BARU (" & m_objrs("semua") & ")"
        M_objrs.MoveNext
Wend
Set M_objrs = Nothing
'-----------------------------
M_objrs.Open TELPo2, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_objrs.EOF

S = Format(M_objrs!receivingdatetime, "yyyy-mm-dd hh:mm:ss")
t = Trim(M_objrs!sendernumber)
u = M_objrs!textdecoded
v = FindReplace(t, "+62", "0")

'----------------------------
'Dim showlist As New ADODB.Recordset
'Dim TOTPTP As Currency
'Dim ssql As String

If (Left(v, 3) = "031") Then
    v = Mid(v, 4, 20)
End If

ssql = "SELECT custid, name FROM mandiri.mgm WHERE mobileno='" & v & "'  or mobileno2='" & v & "'  or mobilenoadd1='" & v & "'  or mobilenoadd2='" & v & "'"
showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
isicustid = showlist!CustId
isiname = showlist!Name
Set showlist = Nothing

Set LST1 = LstSMS1.ListItems.ADD(, , S)
        LST1.SubItems(1) = v
        LST1.SubItems(2) = IIf(IsNull(M_objrs("textdecoded")), "", M_objrs("textdecoded"))
        LST1.SubItems(3) = isicustid
        LST1.SubItems(4) = isiname
'Label2 = "SMS LAMA (" & m_objrs("semua") & ")"
        
        M_objrs.MoveNext
Wend
Set M_objrs = Nothing





'=======================================
'Dim RSsms As ADODB.Recordset
'Set RSsms = New ADODB.Recordset
'Dim Lst As listitem
'RSsms.CursorLocation = adUseClient
'If Left(txtMobileNo1, 1) <> "0" And txtMobileNo1 <> "" Then
'satua = "021" & txtMobileNo1
'Else
'satua = txtMobileNo1
'End If
'
'If Left(txtMobileNo2, 1) <> "0" And txtMobileNo2 <> "" Then
'duaa = "021" & txtMobileNo2
'Else
'duaa = txtMobileNo2
'End If
'
'If Left(txtMobileAdd1, 1) <> "0" And txtMobileAdd1 <> "" Then
'tigaa = "021" & txtMobileAdd1
'Else
'tigaa = txtMobileAdd1
'End If
'
'If Left(txtMobileAdd2, 1) <> "0" And txtMobileAdd2 <> "" Then
'empata = "021" & txtMobileAdd2
'Else
'empata = txtMobileAdd2
'End If
'
'
'cmdsql = "Select a.*, b.custid from receive_sms a, mgm b where (a.notelp='" + satua + "' or a.notelp='" + duaa + "' or a.notelp='" + tigaa + "' or a.notelp='" + empata + "') and b.custid='" + lblCustId + "'"
'RSsms.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not RSsms.EOF
'    Set Lst = LstSMS.ListItems.ADD(, , IIf(IsNull(RSsms("notelp")), "", RSsms("notelp")))
'         Lst.SubItems(1) = lblNama
'         Lst.SubItems(2) = IIf(IsNull(RSsms("custid")), "", RSsms("custid"))
'         Lst.SubItems(3) = IIf(IsNull(RSsms("pesan")), "", RSsms("pesan"))
'         Lst.SubItems(4) = IIf(IsNull(RSsms("tgl_terima")), "", RSsms("tgl_terima"))
'
'RSsms.MoveNext
'Wend
'Set RSsms = Nothing
'Text3 = LstSMS.ListItems.Count
'
''--------------------------------
'If Text4.Text <> "0" Then
'If Int(Text3) > Int(Text2) Then
'
'Dim RSsms_cek As ADODB.Recordset
'Set RSsms_cek = New ADODB.Recordset
'
'RSsms_cek.CursorLocation = adUseClient
'cmdsql_cek = "select * from receive_sms order by tgl_terima desc limit 1"
'RSsms_cek.Open cmdsql_cek, M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not RSsms_cek.EOF
'MsgBox "Anda mendapatkan satu SMS baru" & vbCrLf & "No Telepon : " & RSsms_cek("notelp") & vbCrLf & "Isi Pesan : " & Trim(RSsms_cek("pesan"))
'RSsms_cek.MoveNext
'Wend
'Set RSsms_cek = Nothing
'End If
'End If
'
'Text4.Text = "1"

End Sub


Private Sub LstSMS1_Click()
On Error Resume Next
If Len(LstSMS1.SelectedItem.SubItems(1)) <> 0 Then
no_telp = LstSMS1.SelectedItem.SubItems(1)
Else
no_telp = ""
End If

If Len(LstSMS1.SelectedItem.SubItems(3)) <> 0 Then
isi_Custid = LstSMS1.SelectedItem.SubItems(3)
Else
isi_Custid = ""
End If

If Len(LstSMS1.SelectedItem.SubItems(4)) <> 0 Then
isi_Nama = LstSMS1.SelectedItem.SubItems(4)
Else
isi_Nama = ""
End If

If Len(LstSMS1.SelectedItem.SubItems(2)) <> 0 Then
isi_Pesan = LstSMS1.SelectedItem.SubItems(2)
Else
isi_Pesan = ""
End If

If Len(LstSMS1.SelectedItem.Text) <> 0 Then
isi_tgl = LstSMS1.SelectedItem.Text
Else
isi_tgl = ""
End If


Label8 = no_telp & "(" & isi_tgl & ")"
Label11 = isi_Custid
Label10 = isi_Nama
Label9 = isi_Pesan



End Sub
Private Sub LstSMS_Click()
On Error Resume Next
If Len(LstSMS.SelectedItem.SubItems(1)) <> 0 Then
no_telp = LstSMS.SelectedItem.SubItems(1)
Else
no_telp = ""
End If

If Len(LstSMS.SelectedItem.SubItems(3)) <> 0 Then
isi_Custid = LstSMS.SelectedItem.SubItems(3)
Else
isi_Custid = ""
End If

If Len(LstSMS.SelectedItem.SubItems(4)) <> 0 Then
isi_Nama = LstSMS.SelectedItem.SubItems(4)
Else
isi_Nama = ""
End If

If Len(LstSMS.SelectedItem.SubItems(2)) <> 0 Then
isi_Pesan = LstSMS.SelectedItem.SubItems(2)
Else
isi_Pesan = ""
End If

If Len(LstSMS.SelectedItem.Text) <> 0 Then
isi_tgl = LstSMS.SelectedItem.Text
Else
isi_tgl = ""
End If


Label8 = no_telp & "(" & isi_tgl & ")"
Label11 = isi_Custid
Label10 = isi_Nama
Label9 = isi_Pesan
'Label9 = isi_Pesan

isi_Pesana = Replace(isi_Pesan, "'", "")
notelpon = FindReplace(no_telp, "0", "+62")

CMDSQL = "INSERT INTO mandiri.receive_sms (tgl_terima, notelp, pesan) VALUES ('" & isi_tgl & "',"
            CMDSQL = CMDSQL + " '" + no_telp + "',"
            CMDSQL = CMDSQL + " '" + isi_Pesana + "')"
            M_OBJCONN.Execute CMDSQL

cmdsql_update = "update mandiri.inbox set processed='TRUE'  where sendernumber='" + Trim$(notelpon) + "' "
M_OBJCONN1.Execute cmdsql_update

End Sub

