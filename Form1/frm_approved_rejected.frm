VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_approved_rejected 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6135
   ClientLeft      =   5220
   ClientTop       =   2490
   ClientWidth     =   11430
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdLoad 
      Caption         =   "&Load data"
      Height          =   375
      Left            =   7800
      TabIndex        =   15
      Top             =   5640
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5700
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   9600
      TabIndex        =   0
      Top             =   5640
      Width           =   1695
   End
   Begin MSComctlLib.ListView LstSMS 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
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
      TabIndex        =   2
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
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   225
      Left            =   7800
      TabIndex        =   16
      Top             =   180
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Shape Shape1 
      Height          =   6135
      Left            =   0
      Top             =   0
      Width           =   11385
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "APPROVED"
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
      TabIndex        =   11
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "REJECTED"
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
      TabIndex        =   10
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SMS STATUS"
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
      Left            =   7920
      TabIndex        =   9
      Top             =   75
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "NO TELP : "
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   7440
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
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "CH : "
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   7440
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
      TabIndex        =   4
      Top             =   7680
      Width           =   10095
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   7920
      Width           =   10095
   End
   Begin VB.Label Label11 
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   8160
      Width           =   10095
   End
   Begin VB.Label Label7 
      Caption         =   "NAMA CH : "
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   8160
      Width           =   1215
   End
End
Attribute VB_Name = "frm_approved_rejected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HEADER_SendSMS()
LstSMS.ColumnHeaders.ADD 1, , "No Telpon", 25 * TXT
LstSMS.ColumnHeaders.ADD 2, , "Custid", 25 * TXT
LstSMS.ColumnHeaders.ADD 3, , "Nama", 25 * TXT
LstSMS.ColumnHeaders.ADD 4, , "Agent", 25 * TXT
LstSMS.ColumnHeaders.ADD 5, , "Pesan", 25 * TXT

LstSMS1.ColumnHeaders.ADD 1, , "Tanggal Reject", 25 * TXT
LstSMS1.ColumnHeaders.ADD 2, , "No Telp", 25 * TXT
LstSMS1.ColumnHeaders.ADD 3, , "Custid", 25 * TXT
LstSMS1.ColumnHeaders.ADD 4, , "Nama", 25 * TXT
LstSMS1.ColumnHeaders.ADD 5, , "Pesan", 25 * TXT
LstSMS1.ColumnHeaders.ADD 6, , "Agent", 25 * TXT

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

Private Sub CmdLoad_Click()
    LstSMS1.ListItems.CLEAR
    LstSMS.ListItems.CLEAR
    Call isisms
    'Label12.Caption = LstSMS.ListItems.Count
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call HEADER_SendSMS
    LstSMS1.ListItems.CLEAR
    LstSMS.ListItems.CLEAR
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

On Error Resume Next
'Dim ConnPTP As New ADODB.Connection
Dim M_objrs As New ADODB.Recordset
Dim cmdsql34 As String
Dim unrejected As String
Dim rejected As String

sqljmlrejected = "select count(*) as jmlrej from request_sms where rejected=true and status='0'"
sqljmlunrejected = "select count(*) as jmlunrej from request_sms where rejected=false and status='1'"

M_objrs.Open sqljmlrejected, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_objrs.EOF
jml_reject = M_objrs!jmlrej
M_objrs.MoveNext
Wend
Set M_objrs = Nothing


M_objrs.Open sqljmlunrejected, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_objrs.EOF
jml_un_reject = M_objrs!jmlunrej
M_objrs.MoveNext
Wend
Set M_objrs = Nothing

ProgressBar1.Max = jml_un_reject + 20


Label1 = "APPROVED (" & jml_un_reject & ")"
Label2 = "REJECTED (" & jml_reject & ")"

sqlunrejected = "select * from request_sms where id is not null"
'sqlrejected = "select * from request_sms where rejected=true"

'Set M_OBJRS = New ADODB.Recordset
M_objrs.Open sqlunrejected, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText


While Not M_objrs.EOF
    r = IIf(IsNull(M_objrs!tgl_reject), "", M_objrs!tgl_reject)
    S = M_objrs!Name
    t = M_objrs!notelp
    u = M_objrs!pesan
    v = M_objrs!CustId
    W = M_objrs!AGENT

    If M_objrs!rejected = "0" And M_objrs!STATUS = "1" Then
        'Approved
        Set lst = LstSMS.ListItems.ADD(, , t)
            lst.SubItems(1) = v
            lst.SubItems(2) = S
            lst.SubItems(3) = W
            lst.SubItems(4) = u
                'DoEvents
    End If
    
    If M_objrs!rejected = "1" And M_objrs!STATUS = "0" Then
       'Rejected
        Set LST1 = LstSMS1.ListItems.ADD(, , r)
            LST1.SubItems(1) = t
            LST1.SubItems(2) = v
            LST1.SubItems(3) = S
            LST1.SubItems(4) = u
            LST1.SubItems(5) = W
                'DoEvents
    End If
            'DoEvents
            xxx = xxx + 1
            ProgressBar1.Value = xxx
            M_objrs.MoveNext
Wend
'unrejectedpos:

Set M_objrs = Nothing
ProgressBar1.Value = 1

'-----------------------------
'M_OBJRS.Open sqlrejected, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
''If M_OBJRS.RecordCount < 1 Then
''GoTo rejectedpos
''End If
'
'While Not M_OBJRS.EOF
'r = M_OBJRS!tgl_reject
's = M_OBJRS!Name
't = M_OBJRS!notelp
'u = M_OBJRS!Pesan
'v = M_OBJRS!CustId
'w = M_OBJRS!agent
'
'
'Set Lst1 = LstSMS1.ListItems.ADD(, , r)
'        Lst1.SubItems(1) = t
'        Lst1.SubItems(2) = v
'        Lst1.SubItems(3) = s
'        Lst1.SubItems(4) = u
'        Lst1.SubItems(5) = w
'        M_OBJRS.MoveNext
'Wend
'rejectedpos:
'Set M_OBJRS = Nothing
'Exit Sub

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

CMDSQL = "INSERT INTO receive_sms (tgl_terima, notelp, pesan) VALUES ('" & isi_tgl & "',"
            CMDSQL = CMDSQL + " '" + no_telp + "',"
            CMDSQL = CMDSQL + " '" + isi_Pesana + "')"
            M_OBJCONN.Execute CMDSQL

cmdsql_update = "update inbox set processed='TRUE'  where sendernumber='" + Trim$(notelpon) + "' "
M_OBJCONN1.Execute cmdsql_update

End Sub



