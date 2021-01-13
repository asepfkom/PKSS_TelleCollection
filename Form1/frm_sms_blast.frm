VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_sms_blast 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5445
   ClientLeft      =   4500
   ClientTop       =   3705
   ClientWidth     =   11895
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8520
      TabIndex        =   11
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   10200
      TabIndex        =   10
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CheckBox Text18 
      BackColor       =   &H8000000D&
      Caption         =   "Kirim"
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   8040
      Width           =   1215
   End
   Begin VB.ComboBox Text1 
      Height          =   315
      ItemData        =   "frm_sms_blast.frx":0000
      Left            =   5760
      List            =   "frm_sms_blast.frx":000A
      TabIndex        =   8
      Top             =   7920
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   9600
      TabIndex        =   7
      Text            =   "0"
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   10080
      TabIndex        =   6
      Text            =   "0"
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   10560
      TabIndex        =   5
      Text            =   "0"
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   11040
      TabIndex        =   4
      Text            =   "0"
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "0"
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   3480
      Width           =   11655
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Search"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ListView Lst_SMS_blast 
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4471
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
   Begin VB.Label Label1 
      Caption         =   "Isi Pesan :"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "AGENT :"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   11610
      TabIndex        =   13
      Top             =   3240
      Width           =   45
   End
   Begin VB.Shape Shape1 
      Height          =   5415
      Left            =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frm_sms_blast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type LVITEM
mask As Long
iItem As Long
iSubItem As Long
state As Long
stateMask As Long
pszText As String
cchTextMax As Long
iImage As Long
lParam As Long
iIndent As Long
End Type

Private bDoingSetup As Boolean
Private dirty As Boolean
Private itmClicked As ListItem
Private dwLastSubitemEdited As Long

Private Const LVM_FIRST = &H1000
Private Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
Private Const LVM_GETTOPINDEX = (LVM_FIRST + 39)
Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
Private Const LVHT_ONITEMICON = &H2
Private Const LVHT_ONITEMLABEL = &H4
Private Const LVHT_ONITEMSTATEICON = &H8
Private Const LVHT_ONITEM = (LVHT_ONITEMICON Or _
LVHT_ONITEMLABEL Or _
LVHT_ONITEMSTATEICON)
Private Const LVIR_LABEL = 2

Private Type POINTAPI
X As Long
Y As Long
End Type

Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Private Type LVHITTESTINFO
pt As POINTAPI
flags As Long
iItem As Long
iSubItem As Long
End Type

Private Declare Function ScreenToClient Lib "user32" _
(ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" _
(ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long


Private Sub Command1_Click()
DoEvents
Label2 = "Sending..."
DoEvents
Dim itm As ListItem
    For i = 1 To Lst_SMS_blast.ListItems.Count
        
            If Lst_SMS_blast.ListItems(i).Checked = True Then
            isi_data = "INSERT INTO mandiri.request_sms ( agent, custid,name,notelp,pesan,status) VALUES ('" & Trim$(Lst_SMS_blast.ListItems(i).SubItems(6)) & "', '" & Trim$(Lst_SMS_blast.ListItems(i).SubItems(1)) & "',' " & Trim$(Lst_SMS_blast.ListItems(i).Text) & "','"
            
            
                If Lst_SMS_blast.ListItems(i).SubItems(2) <> "" Then
                If Left(Lst_SMS_blast.ListItems(i).SubItems(2), 1) = "-" Then
                If Left(Mid(Lst_SMS_blast.ListItems(i).SubItems(2), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(2)) - 1), 1) <> "0" Then
                notelp2 = "031" & Mid(Lst_SMS_blast.ListItems(i).SubItems(2), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(2)) - 1)
                Else
                notelp2 = Mid(Lst_SMS_blast.ListItems(i).SubItems(2), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(2)) - 1)
                End If
                isi_data1 = isi_data & Trim$(notelp2) & "', '[HSBC COLL] " & Text7 & "', '0')" & vbCrLf
                M_OBJCONN.Execute isi_data1
                End If
                End If
                
                If Lst_SMS_blast.ListItems(i).SubItems(3) <> "" Then
                If Left(Lst_SMS_blast.ListItems(i).SubItems(3), 1) = "-" Then
                If Left(Mid(Lst_SMS_blast.ListItems(i).SubItems(3), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(3)) - 1), 1) <> "0" Then
                notelp3 = "031" & Mid(Lst_SMS_blast.ListItems(i).SubItems(3), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(3)) - 1)
                Else
                notelp3 = Mid(Lst_SMS_blast.ListItems(i).SubItems(3), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(3)) - 1)
                End If
                isi_data2 = isi_data & Trim$(notelp3) & "', '[HSBC COLL] " & Text7 & "', '0')" & vbCrLf
                M_OBJCONN.Execute isi_data2
                End If
                End If
                
                If Lst_SMS_blast.ListItems(i).SubItems(4) <> "" Then
                If Left(Lst_SMS_blast.ListItems(i).SubItems(4), 1) = "-" Then
                If Left(Mid(Lst_SMS_blast.ListItems(i).SubItems(4), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(4)) - 1), 1) <> "0" Then
                notelp4 = "031" & Mid(Lst_SMS_blast.ListItems(i).SubItems(4), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(4)) - 1)
                Else
                notelp4 = Mid(Lst_SMS_blast.ListItems(i).SubItems(4), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(4)) - 1)
                End If
                isi_data3 = isi_data & Trim$(notelp4) & "', '[HSBC COLL] " & Text7 & "', '0')" & vbCrLf
                M_OBJCONN.Execute isi_data3
                End If
                End If
                
                If Lst_SMS_blast.ListItems(i).SubItems(5) <> "" Then
                If Left(Lst_SMS_blast.ListItems(i).SubItems(5), 1) = "-" Then
                If Left(Mid(Lst_SMS_blast.ListItems(i).SubItems(5), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(5)) - 1), 1) <> "0" Then
                notelp5 = "031" & Mid(Lst_SMS_blast.ListItems(i).SubItems(5), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(5)) - 1)
                Else
                notelp5 = Mid(Lst_SMS_blast.ListItems(i).SubItems(5), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(5)) - 1)
                End If
                isi_data4 = isi_data & Trim$(notelp5) & "', '[HSBC COLL] " & Text7 & "', '0')" & vbCrLf
                M_OBJCONN.Execute isi_data4
                End If
                End If
                 
Else
            End If
                  
  Next
MsgBox "Pesan terkirim ke Admin"
Label2 = ""
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Call isi_list_agent(Combo1.Text)
End Sub

Private Sub Form_Load()
Lst_SMS_blast.ColumnHeaders.ADD 1, , "Nama", 15 * TXT
Lst_SMS_blast.ColumnHeaders.ADD 2, , "Custid", 15 * TXT
Lst_SMS_blast.ColumnHeaders.ADD 3, , "HP", 15 * TXT
Lst_SMS_blast.ColumnHeaders.ADD 4, , "HP 1", 15 * TXT
Lst_SMS_blast.ColumnHeaders.ADD 5, , "HP ADD 1", 15 * TXT
Lst_SMS_blast.ColumnHeaders.ADD 6, , "HP ADD 2", 15 * TXT
Lst_SMS_blast.ColumnHeaders.ADD 7, , "AGENT", 5 * TXT
DoEvents

Call isi_list
Call isi_combo

If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
Label3.Visible = True
Combo1.Visible = True
Command3.Visible = True
End If

If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Then
Label3.Visible = False
Combo1.Visible = False
Command3.Visible = False
End If



End Sub
Sub isi_list()



Dim R_blast As ADODB.Recordset
Set R_blast = New ADODB.Recordset
Dim lst As ListItem
R_blast.CursorLocation = adUseClient


If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
CMDSQL = "select mgm.name, mgm.custid, mgm.agent, mgm.mobileno, mgm.mobileno2 , mgm.mobilenoadd1, mgm.mobilenoadd2 from mandiri.usertbl inner join mandiri.mgm on mgm.agent=usertbl.userid where usertbl.team='" & MDIForm1.TxtUsername.Text & "'"
End If

If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Then
CMDSQL = "select mgm.name, mgm.custid, mgm.agent, mgm.mobileno, mgm.mobileno2 , mgm.mobilenoadd1, mgm.mobilenoadd2 from mandiri.usertbl inner join mandiri.mgm on mgm.agent=usertbl.userid"
End If

If UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Then
CMDSQL = "select mgm.name, mgm.custid, mgm.agent, mgm.mobileno, mgm.mobileno2 , mgm.mobilenoadd1, mgm.mobilenoadd2 from mandiri.usertbl inner join mandiri.mgm on mgm.agent=usertbl.userid"
End If


R_blast.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not R_blast.EOF
    Set lst = Lst_SMS_blast.ListItems.ADD(, , IIf(IsNull(R_blast("name")), "", R_blast("name")))
         lst.SubItems(1) = IIf(IsNull(R_blast("custid")), "", R_blast("custid"))
         lst.SubItems(2) = IIf(IsNull(R_blast("mobileno")), "", R_blast("mobileno"))
         lst.SubItems(3) = IIf(IsNull(R_blast("mobileno2")), "", R_blast("mobileno2"))
         lst.SubItems(4) = IIf(IsNull(R_blast("mobilenoadd1")), "", R_blast("mobilenoadd1"))
         lst.SubItems(5) = IIf(IsNull(R_blast("mobilenoadd2")), "", R_blast("mobilenoadd2"))
         lst.SubItems(6) = IIf(IsNull(R_blast("agent")), "", R_blast("agent"))
R_blast.MoveNext
Wend
Set R_blast = Nothing
Label4 = "Jumlah Data : " & Lst_SMS_blast.ListItems.Count

MDIForm1.MousePointer = vbNormal
End Sub

Sub isi_list_agent(agentno)

Lst_SMS_blast.ListItems.CLEAR
Dim R_blast As ADODB.Recordset
Set R_blast = New ADODB.Recordset
Dim lst As ListItem
R_blast.CursorLocation = adUseClient
CMDSQL = "select mgm.name, mgm.custid, mgm.agent, mgm.mobileno, mgm.mobileno2 , mgm.mobilenoadd1, mgm.mobilenoadd2 from mandiri.usertbl inner join mandiri.mgm on mgm.agent=usertbl.userid where usertbl.team='" & MDIForm1.TxtUsername.Text & "' and mgm.agent='" & agentno & "'"
R_blast.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not R_blast.EOF
    Set lst = Lst_SMS_blast.ListItems.ADD(, , IIf(IsNull(R_blast("name")), "", R_blast("name")))
         lst.SubItems(1) = IIf(IsNull(R_blast("custid")), "", R_blast("custid"))
         lst.SubItems(2) = IIf(IsNull(R_blast("mobileno")), "", R_blast("mobileno"))
         lst.SubItems(3) = IIf(IsNull(R_blast("mobileno2")), "", R_blast("mobileno2"))
         lst.SubItems(4) = IIf(IsNull(R_blast("mobilenoadd1")), "", R_blast("mobilenoadd1"))
         lst.SubItems(5) = IIf(IsNull(R_blast("mobilenoadd2")), "", R_blast("mobilenoadd2"))
         lst.SubItems(6) = IIf(IsNull(R_blast("agent")), "", R_blast("agent"))
R_blast.MoveNext
Wend
Set R_blast = Nothing
Label4 = "Jumlah Data : " & Lst_SMS_blast.ListItems.Count

End Sub

Sub isi_combo()

Dim R_combo As ADODB.Recordset
Set R_combo = New ADODB.Recordset
Dim lst As ListItem
R_combo.CursorLocation = adUseClient
CMDSQL = "select distinct mgm.agent from mandiri.usertbl inner join mandiri.mgm on mgm.agent=usertbl.userid where usertbl.team='" & MDIForm1.TxtUsername.Text & "'"

R_combo.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not R_combo.EOF

    Combo1.AddItem R_combo("agent")
R_combo.MoveNext
Wend
Set R_combo = Nothing
End Sub
Private Sub Lst_SMS_blast_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

'--------------------------

 If ColumnHeader.Index = 3 Then
 If Text2.Text = "0" Then
 Text2.Text = "1"
 For i = 1 To Lst_SMS_blast.ListItems.Count - 1
 If Left(Lst_SMS_blast.ListItems(i).SubItems(2), 1) <> "" Then
 If Left(Lst_SMS_blast.ListItems(i).SubItems(2), 1) <> "-" Then
 Lst_SMS_blast.ListItems(i).SubItems(2) = "-" & Lst_SMS_blast.ListItems(i).SubItems(2)
 Lst_SMS_blast.ListItems(i).Checked = True
 Call cek_checka(i)

 End If
 End If
 Next
 Else
 Text2.Text = "0"
 For i = 1 To Lst_SMS_blast.ListItems.Count - 1
 If Left(Lst_SMS_blast.ListItems(i).SubItems(2), 1) <> "" Then
 If Left(Lst_SMS_blast.ListItems(i).SubItems(2), 1) = "-" Then
 Lst_SMS_blast.ListItems(i).SubItems(2) = Mid(Lst_SMS_blast.ListItems(i).SubItems(2), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(2)) - 1)
 Call cek_checka(i)

 End If
 End If
 Next
 End If
 End If
 
 '---------------------
 
 If ColumnHeader.Index = 4 Then
 If Text3.Text = "0" Then
 Text3.Text = "1"
 For i = 1 To Lst_SMS_blast.ListItems.Count - 1
 If Left(Lst_SMS_blast.ListItems(i).SubItems(3), 1) <> "" Then
 If Left(Lst_SMS_blast.ListItems(i).SubItems(3), 1) <> "-" Then
 Lst_SMS_blast.ListItems(i).SubItems(3) = "-" & Lst_SMS_blast.ListItems(i).SubItems(3)
 Lst_SMS_blast.ListItems(i).Checked = True
 Call cek_checka(i)
 End If
 End If
 Next
 Else
 Text3.Text = "0"
 For i = 1 To Lst_SMS_blast.ListItems.Count - 1
 If Left(Lst_SMS_blast.ListItems(i).SubItems(3), 1) <> "" Then
 If Left(Lst_SMS_blast.ListItems(i).SubItems(3), 1) = "-" Then
 Lst_SMS_blast.ListItems(i).SubItems(3) = Mid(Lst_SMS_blast.ListItems(i).SubItems(3), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(3)) - 1)
 Call cek_checka(i)
 End If
 End If
 Next
 End If
 End If
 
 '---------------------
 
 If ColumnHeader.Index = 5 Then
 If Text4.Text = "0" Then
 Text4.Text = "1"
 For i = 1 To Lst_SMS_blast.ListItems.Count - 1
 If Left(Lst_SMS_blast.ListItems(i).SubItems(4), 1) <> "" Then
 If Left(Lst_SMS_blast.ListItems(i).SubItems(4), 1) <> "-" Then
 Lst_SMS_blast.ListItems(i).SubItems(4) = "-" & Lst_SMS_blast.ListItems(i).SubItems(4)
 Lst_SMS_blast.ListItems(i).Checked = True
 Call cek_checka(i)
 End If
 End If
 Next
 Else
 Text4.Text = "0"
 For i = 1 To Lst_SMS_blast.ListItems.Count - 1
 If Left(Lst_SMS_blast.ListItems(i).SubItems(4), 1) <> "" Then
 If Left(Lst_SMS_blast.ListItems(i).SubItems(4), 1) = "-" Then
 Lst_SMS_blast.ListItems(i).SubItems(4) = Mid(Lst_SMS_blast.ListItems(i).SubItems(4), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(4)) - 1)
 Call cek_checka(i)
 End If
 End If
 Next
 End If
 End If
 
 '---------------------
 
 If ColumnHeader.Index = 6 Then
 If Text5.Text = "0" Then
 Text5.Text = "1"
 For i = 1 To Lst_SMS_blast.ListItems.Count - 1
 If Left(Lst_SMS_blast.ListItems(i).SubItems(5), 1) <> "" Then
 If Left(Lst_SMS_blast.ListItems(i).SubItems(5), 1) <> "-" Then
 Lst_SMS_blast.ListItems(i).SubItems(5) = "-" & Lst_SMS_blast.ListItems(i).SubItems(5)
 Lst_SMS_blast.ListItems(i).Checked = True
 Call cek_checka(i)
 End If
 End If
 Next
 Else
 Text5.Text = "0"
 For i = 1 To Lst_SMS_blast.ListItems.Count - 1
 If Left(Lst_SMS_blast.ListItems(i).SubItems(5), 1) <> "" Then
 If Left(Lst_SMS_blast.ListItems(i).SubItems(5), 1) = "-" Then
 Lst_SMS_blast.ListItems(i).SubItems(5) = Mid(Lst_SMS_blast.ListItems(i).SubItems(5), 2, Len(Lst_SMS_blast.ListItems(i).SubItems(5)) - 1)
 Call cek_checka(i)
 End If
 End If
 Next
 End If
 End If
 
End Sub

Private Sub Lst_SMS_blast_ItemCheck(ByVal Item As MSComctlLib.ListItem)
 With Item

If Item.Checked = True Then
For i = 2 To 5
If Item.SubItems(i) <> "" Then
 Item.SubItems(i) = "-" & Item.SubItems(i)
 End If
Next

End If

If Item.Checked = False Then

For i = 2 To 5
If Left(Item.SubItems(i), 1) = "-" Then
 Item.SubItems(i) = Mid(Item.SubItems(i), 2, Len(Item.SubItems(i)) - 1)
 Else
 Item.SubItems(i) = Item.SubItems(i)
 End If
Next
End If
    End With
End Sub



Private Sub Lst_SMS_blast_MouseDown(Button As Integer, _
Shift As Integer, _
X As Single, _
Y As Single)

Dim HTI As LVHITTESTINFO
Dim fpx As Single
Dim fpy As Single
Dim fpw As Single
Dim fph As Single
Dim rc As RECT
Dim topindex As Long

bDoingSetup = True


Text1.Visible = False

With HTI
.pt.X = (X / Screen.TwipsPerPixelX)
.pt.Y = (Y / Screen.TwipsPerPixelY)
.flags = LVHT_ONITEM
End With

Call SendMessage(Lst_SMS_blast.hwnd, _
LVM_SUBITEMHITTEST, _
0, HTI)

If HTI.iItem <> -1 And HTI.iSubItem > 0 Then

Lst_SMS_blast.LabelEdit = lvwManual

rc.Left = LVIR_LABEL
rc.Top = HTI.iSubItem
Call SendMessage(Lst_SMS_blast.hwnd, _
LVM_GETSUBITEMRECT, _
HTI.iItem, _
rc)


Set itmClicked = Lst_SMS_blast.ListItems(HTI.iItem + 1)
itmClicked.Selected = True
' MsgBox HTI.iSubItem
'topindex = SendMessage(Lst_SMS_blast.hwnd, _
'LVM_GETTOPINDEX, _
'0&, _
'ByVal 0&)

'fpx = Lst_SMS_blast.Left + _
'(rc.Left * Screen.TwipsPerPixelX) + 80
'
'fpy = Lst_SMS_blast.Top + _
'(HTI.iItem + 1 - topindex) + _
'(rc.Top * Screen.TwipsPerPixelY)
'
'fph = 120

'fpw = SendMessage(Lst_SMS_blast.hwnd, _
'LVM_GETCOLUMNWIDTH, _
'HTI.iSubItem, _
'ByVal 0&)
'
'
'fpw = (fpw * Screen.TwipsPerPixelX) - 40


With Text1

dwLastSubitemEdited = HTI.iSubItem

If HTI.iSubItem > 1 Then
If Len(itmClicked.SubItems(HTI.iSubItem)) <> 0 Then
If Left(itmClicked.SubItems(dwLastSubitemEdited), 1) = "-" Then
 itmClicked.SubItems(dwLastSubitemEdited) = Mid(itmClicked.SubItems(dwLastSubitemEdited), 2, Len(itmClicked.SubItems(dwLastSubitemEdited)) - 1)
 Else
  itmClicked.SubItems(dwLastSubitemEdited) = "-" & itmClicked.SubItems(dwLastSubitemEdited)
End If

Call cek_check(HTI.iItem)

Else
MsgBox "Kosong"
Call cek_check(HTI.iItem)

End If
End If
End With
End If

End Sub


Sub cek_check(subitemlist)
For i = 2 To 5
If Left(Lst_SMS_blast.ListItems(subitemlist + 1).SubItems(i), 1) = "-" Then
tt = 1
End If
Next
If tt = 1 Then
Lst_SMS_blast.ListItems(subitemlist + 1).Checked = True
Else
Lst_SMS_blast.ListItems(subitemlist + 1).Checked = False
End If
End Sub
Sub cek_checka(subitemlist)
For i = 2 To 5
If Left(Lst_SMS_blast.ListItems(subitemlist).SubItems(i), 1) = "-" Then
tt = 1
End If
Next
If tt = 1 Then
Lst_SMS_blast.ListItems(subitemlist).Checked = True
Else
Lst_SMS_blast.ListItems(subitemlist).Checked = False
End If
End Sub


Private Sub Text1_Click()
If Text1.Text = "KIRIM" Then

If Left(itmClicked.SubItems(dwLastSubitemEdited), 1) = "-" Then
 itmClicked.SubItems(dwLastSubitemEdited) = itmClicked.SubItems(dwLastSubitemEdited)
 Else
  itmClicked.SubItems(dwLastSubitemEdited) = "-" & itmClicked.SubItems(dwLastSubitemEdited)
End If
'Text1.Value = vbUnchecked
Text1.Visible = False
 
Else
If Left(itmClicked.SubItems(dwLastSubitemEdited), 1) = "-" Then
 itmClicked.SubItems(dwLastSubitemEdited) = Mid(itmClicked.SubItems(dwLastSubitemEdited), 2, Len(itmClicked.SubItems(dwLastSubitemEdited)) - 1)
 Else
  itmClicked.SubItems(dwLastSubitemEdited) = itmClicked.SubItems(dwLastSubitemEdited)
End If
'Text1.Value = vbUnchecked
Text1.Visible = False

'MsgBox itmClicked.SubItems(dwLastSubitemEdited)
End If
End Sub

Private Sub Text1_LostFocus()

If dirty And dwLastSubitemEdited > 0 Then
itmClicked.SubItems(dwLastSubitemEdited) = Text1
dirty = False

End If

End Sub



