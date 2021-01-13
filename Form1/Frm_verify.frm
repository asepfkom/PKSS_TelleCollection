VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Frm_verify 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15795
   LinkTopic       =   "Form2"
   ScaleHeight     =   7800
   ScaleWidth      =   15795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Reject"
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
      Left            =   12360
      TabIndex        =   6
      Top             =   7170
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   7200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
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
      Left            =   10680
      TabIndex        =   4
      Top             =   7170
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Approve"
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
      Left            =   14040
      TabIndex        =   2
      Top             =   7170
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
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
      Left            =   9000
      TabIndex        =   1
      Top             =   7170
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4560
      TabIndex        =   0
      Text            =   "0"
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ListView Lst_SMS_verify 
      Height          =   6435
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   11351
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "APPROVAL / REJECT SMS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   30
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      Height          =   7800
      Left            =   0
      Top             =   0
      Width           =   15795
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D4B9AF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15855
   End
End
Attribute VB_Name = "Frm_verify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Public M_OBJCONN1  As New ADODB.Connection
Private Sub verify_SendSMS()
Lst_SMS_verify.ListItems.CLEAR
Lst_SMS_verify.ColumnHeaders.ADD 1, , "Tanggal", 15 * TXT
Lst_SMS_verify.ColumnHeaders.ADD 2, , "No. Tujuan", 15 * TXT
Lst_SMS_verify.ColumnHeaders.ADD 3, , "Nama", 15 * TXT
Lst_SMS_verify.ColumnHeaders.ADD 4, , "Pesan", 150 * TXT
Lst_SMS_verify.ColumnHeaders.ADD 5, , "Custid", 15 * TXT
Lst_SMS_verify.ColumnHeaders.ADD 6, , "Agent", 15 * TXT
End Sub



Private Sub Command1_Click()


Dim jmlgagal As Double
Dim jmlsukses As Double
Dim itm As ListItem
jmlgagal = 0
jmlsukses = 0
    For i = 1 To Lst_SMS_verify.ListItems.Count
        
            If Lst_SMS_verify.ListItems(i).Checked = True Then
            
            If IsNumeric(Trim(Lst_SMS_verify.ListItems(i).SubItems(1))) = False Then
           
            CMDSQL = "delete from request_sms where custid='" & Trim(Lst_SMS_verify.ListItems(i).SubItems(4)) & "' and notelp='" & Trim(Lst_SMS_verify.ListItems(i).SubItems(1)) & "'"
            M_OBJCONN.Execute CMDSQL
          '  MsgBox "Pesan dihapas karena ada kesalahan dalam penulisan nomor telepon"
            jmlgagal = jmlgagal + 1
            Else
            aaa = Replace(Trim(Lst_SMS_verify.ListItems(i).SubItems(1)), ".", "")
            aaa = Replace(aaa, ",", "")
            aaa = Replace(aaa, "/", "")
            aaa = Replace(aaa, "\", "")
            aaa = Replace(aaa, "(", "")
            aaa = Replace(aaa, ")", "")
            aaa = Replace(aaa, "{", "")
            aaa = Replace(aaa, "}", "")
            aaa = Replace(aaa, "[", "")
            aaa = Replace(aaa, "]", "")
            aaa = Replace(aaa, " ", "")
            If Left(aaa, 1) <> "0" Then
            
            CMDSQL = "delete from request_sms where custid='" & Trim(Lst_SMS_verify.ListItems(i).SubItems(4)) & "' and notelp='" & Trim(Lst_SMS_verify.ListItems(i).SubItems(1)) & "'"
            M_OBJCONN.Execute CMDSQL
           ' MsgBox "Pesan dihapas karena ada kesalahan dalam penulisan nomor telepon"
            jmlgagal = jmlgagal + 1
            Else
            
            cid = Trim(Lst_SMS_verify.ListItems(i).SubItems(4)) & "-" & Trim(Lst_SMS_verify.ListItems(i).SubItems(5))
            
            '@@ 09022011 Tambahan lakukan looping untuk sms yang lebih dari 160 karakter
            Dim banyaksms As Integer
            Dim pesan As String
            banyaksms = Ceiling(Val(Len(Trim(Lst_SMS_verify.ListItems(i).SubItems(3)))) / 160)
            
            For aa = 1 To banyaksms
                awalpesan = (160 * aa) - 160
                pesan = Mid(Lst_SMS_verify.ListItems(i).SubItems(3), awalpesan + 1, 160)
                
                
                CMDSQL = "INSERT INTO outbox "
                CMDSQL = CMDSQL + " (`DestinationNumber`,"
                CMDSQL = CMDSQL + " `TextDecoded`,`CreatorID`)"
                CMDSQL = CMDSQL + " VALUES"
                CMDSQL = CMDSQL + " ( '" + aaa + "',"
                CMDSQL = CMDSQL + " '" + Trim(pesan) + "', '" & cid & "')"
                M_OBJCONN1.Execute CMDSQL
            Next aa
            
            
            cmdsqltglsys = "SELECT now() AS tglsystem"
            Set R_tglsys = New ADODB.Recordset
            R_tglsys.Open cmdsqltglsys, M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not R_tglsys.EOF
            TGLw = R_tglsys("tglsystem")
            TGLSERVERb = Format(TGLw, "yyyy-mm-dd hh:mm:ss")
            
            R_tglsys.MoveNext
            Wend
            
            Set R_tglsys = Nothing
            
            
            
            

            cmdsqla = "Update request_sms set status='1', tgl_approve= '" & TGLSERVERb & "'  where notelp='" + Trim(Lst_SMS_verify.ListItems(i).SubItems(1)) + "' and pesan='" + Trim(Lst_SMS_verify.ListItems(i).SubItems(3)) + "'"
            M_OBJCONN.Execute cmdsqla

            
            
            cmdsqltglsys = "SELECT now() AS tglsystem"
            Set R_tglsys = New ADODB.Recordset
            R_tglsys.Open cmdsqltglsys, M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not R_tglsys.EOF
            TGLw = R_tglsys("tglsystem")
            Text1 = Format(TGLw, "yyyy-mm-dd hh:mm:ss")
            
            R_tglsys.MoveNext
            Wend
            
            Set R_tglsys = Nothing
    
            cmdsqlb = "INSERT INTO mgm_hst"
            cmdsqlb = cmdsqlb + " (custid,agent,tgl,phoneno,hst)"
            cmdsqlb = cmdsqlb + " VALUES"
            cmdsqlb = cmdsqlb + " ('" & Trim(Lst_SMS_verify.ListItems(i).SubItems(4)) & "','" & Trim(Lst_SMS_verify.ListItems(i).SubItems(5)) & "','" & Text1 & "','" & aaa & "','" & Trim(Lst_SMS_verify.ListItems(i).SubItems(3)) & "')"
            
            M_OBJCONN.Execute cmdsqlb
            
            
            jmlsukses = jmlsukses + 1
            'If i = Lst_SMS_verify.ListItems.Count Then
            '    MsgBox "SMS telah terkirim"
            'End If
     

'=============================
'             cmdsql = "INSERT INTO outbox "
'            cmdsql = cmdsql + " (destinationnumber,"
'            cmdsql = cmdsql + " textdecoded,creatorid,senderid)"
'            cmdsql = cmdsql + " VALUES"
'            cmdsql = cmdsql + " ( '" + Trim(Lst_SMS_verify.ListItems(i).Text) + "',"
'            cmdsql = cmdsql + " '" + Trim(Lst_SMS_verify.ListItems(i).SubItems(2)) + "', '', 'phone1')"
'            M_OBJCONN1.Execute cmdsql
'
'            cmdsqla = "Update request_sms set status='1' where notelp='" + Lst_SMS_verify.ListItems(i).Text + "' and pesan='" + Lst_SMS_verify.SelectedItem.SubItems(2) + "'"
'            M_OBJCONN.Execute cmdsqla
            End If
            End If
            
            End If
            
        
    Next
    
    If Lst_SMS_verify.ListItems.Count <> 0 Then
        If jmlgagal = 0 And jmlsukses > 0 Then
            MsgBox "Data telah Terkirim sebanyak : " & CStr(jmlsukses)
        End If
        
        If jmlgagal > 0 And jmlsuskes = 0 Then
            MsgBox "Data Gagal sebanyak : " & CStr(jmlgagal)
        End If
        If jmlgagal > 0 And jmlsukses > 0 Then
            MsgBox "Data telah Terkirim sebanyak : " & CStr(jmlsukses) & vbCrLf & _
            "Data Yang Gagal sebanyak : " & CStr(jmlgagal)
        End If
        
    End If
    

Lst_SMS_verify.ListItems.CLEAR
'Call verify_SendSMS
Call isi_list


End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim itm As ListItem
    For i = 1 To Lst_SMS_verify.ListItems.Count
        
            If Lst_SMS_verify.ListItems(i).Checked = True Then
            
            'ddd = Trim(Lst_SMS_verify.ListItems(i).SubItems(3))
            CMDSQL = "delete from request_sms where custid='" & Trim(Lst_SMS_verify.ListItems(i).SubItems(4)) & "' and notelp='" & Trim(Lst_SMS_verify.ListItems(i).SubItems(1)) & "'"
            M_OBJCONN.Execute CMDSQL
           
            End If
        
    Next
    

Lst_SMS_verify.ListItems.CLEAR
'Call verify_SendSMS
Call isi_list
End Sub

Private Sub Command4_Click()
Dim itm As ListItem
    For i = 1 To Lst_SMS_verify.ListItems.Count
        
            If Lst_SMS_verify.ListItems(i).Checked = True Then
            
             
            cmdsqltglsys = "SELECT now() AS tglsystem"
            Set R_tglsys = New ADODB.Recordset
            R_tglsys.Open cmdsqltglsys, M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not R_tglsys.EOF
            TGLw = R_tglsys("tglsystem")
            TGLSERVERa = Format(TGLw, "yyyy-mm-dd hh:mm:ss")
            
            R_tglsys.MoveNext
            Wend
            
            Set R_tglsys = Nothing
            
            
            
            
            CMDSQL = "update request_sms set rejected = true, tgl_reject='" & TGLSERVERa & "' where custid='" & Trim(Lst_SMS_verify.ListItems(i).SubItems(4)) & "' and notelp='" & Trim(Lst_SMS_verify.ListItems(i).SubItems(1)) & "'"
            M_OBJCONN.Execute CMDSQL
           
            End If
        
    Next
    

Lst_SMS_verify.ListItems.CLEAR
'Call verify_SendSMS
Call isi_list
End Sub

Private Sub Form_Load()
'CMDSQLOPEN1 = "Provider=MSDASQL.1;Persist Security Info=False;User ID=admin;PWD=admin321;Data Source=sms"
'If M_OBJCONN1.state = 1 Then M_OBJCONN1.state = 0
'M_OBJCONN1.Open CMDSQLOPEN1
Call verify_SendSMS
Call isi_list
End Sub
Sub isi_list()

Dim R_send_verify As ADODB.Recordset
Set R_send_verify = New ADODB.Recordset
Dim lst As ListItem
R_send_verify.CursorLocation = adUseClient
'cmdsql = "select a.notelp, a.pesan, a.agent, b.name, b.custid from request_sms a, mgm b where (b.mobileno= a.notelp or b.mobileno2= a.notelp or b.mobilenoadd1= a.notelp or b.mobilenoadd2= a.notelp) and a.status='0' and  a.custid=b.custid"
CMDSQL = "select * from request_sms where status ='0' and rejected = false"

R_send_verify.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not R_send_verify.EOF
       Set lst = Lst_SMS_verify.ListItems.ADD(, , IIf(IsNull(R_send_verify("tgl_kirim")), "", R_send_verify("tgl_kirim")))
   
         'Set Lst = Lst_SMS_verify.ListItems.ADD(, , Left(R_send_verify("tgl_kirim"), 18))
         lst.SubItems(1) = R_send_verify("notelp")
         lst.SubItems(2) = R_send_verify("name")
         lst.SubItems(3) = R_send_verify("pesan")
         lst.SubItems(4) = R_send_verify("custid")
         lst.SubItems(5) = R_send_verify("agent")
R_send_verify.MoveNext
Wend
Set R_send_verify = Nothing
End Sub

Private Sub Lst_SMS_verify_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If ColumnHeader.Index = 1 Then
 If Text2.Text = "0" Then
 Text2.Text = "1"
 For i = 1 To Lst_SMS_verify.ListItems.Count
 Lst_SMS_verify.ListItems(i).Checked = True
 Next
 Else
 Text2.Text = "0"
 For i = 1 To Lst_SMS_verify.ListItems.Count
 Lst_SMS_verify.ListItems(i).Checked = False
 Next
 End If
End If
 
End Sub

Private Sub Lst_SMS_verify_DblClick()
If Lst_SMS_verify.ListItems.Count > 0 Then
Load frmeditsms
no_telp = Lst_SMS_verify.SelectedItem.SubItems(1)
isi_Custid = Lst_SMS_verify.SelectedItem.SubItems(4)
isi_Pesan = Lst_SMS_verify.SelectedItem.SubItems(3)
frmeditsms.Text5 = no_telp
frmeditsms.Text1 = Trim$(isi_Pesan)
frmeditsms.Text3 = isi_Custid

frmeditsms.Show vbModal
'MsgBox "No Telepon : " & no_telp & vbCrLf & "Isi Pesan : " & Trim(isi_Pesan) & vbCrLf & "custid : " & Trim(isi_CustId)

    Else
    Exit Sub
 End If
End Sub
'@@09022011 Fungsi buat membulatkan desimal
Private Function Ceiling(number As Double) As Long
    Ceiling = -Int(-number)
End Function
