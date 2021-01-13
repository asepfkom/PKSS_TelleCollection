VERSION 5.00
Begin VB.Form FrmSendSMS 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5145
   ClientLeft      =   7155
   ClientTop       =   3705
   ClientWidth     =   5025
   LinkTopic       =   "Form2"
   ScaleHeight     =   5145
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "FrmSendSMS.frx":0000
      Left            =   1320
      List            =   "FrmSendSMS.frx":000D
      TabIndex        =   20
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   345
      Left            =   2700
      TabIndex        =   18
      Top             =   6300
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.ComboBox CmbSubOption 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   6780
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.ComboBox CmbOption 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "FrmSendSMS.frx":002A
      Left            =   1320
      List            =   "FrmSendSMS.frx":002C
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   6300
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1320
      MaxLength       =   160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2520
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
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
      Left            =   3600
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3255
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
      Left            =   2400
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2100
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1365
      TabIndex        =   6
      Top             =   2100
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category :"
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
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub option:"
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
      Left            =   240
      TabIndex        =   17
      Top             =   6750
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Option:"
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
      Left            =   240
      TabIndex        =   16
      Top             =   6390
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah :"
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
      Left            =   360
      TabIndex        =   13
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No :"
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
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Text :"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   5235
      Left            =   0
      Top             =   -120
      Width           =   5010
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent :"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Custid :"
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
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama :"
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
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "SEND SMS"
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
      TabIndex        =   19
      Top             =   45
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D4B9AF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "FrmSendSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public awal As String
Public btsakhir As Integer
Dim d As Integer
Dim awalk As Integer
Dim akhirk As Integer
Dim AvgMarks(50, 50) As Double
Dim rowaray As Integer

Private Sub CmbOption_Click()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    Text1.text = ""
    CmbSubOption.CLEAR
    
    CMDSQL = "select * from tblscriptsms where option='"
    CMDSQL = CMDSQL + CmbOption.text + "'"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    
    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    CmbSubOption.CLEAR
    While Not M_objrs.EOF
        CmbSubOption.AddItem M_objrs("suboption")
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End Sub

Private Sub CmbSubOption_Click()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    Text1.text = ""
    CMDSQL = "select * from tblscriptsms where option='"
    CMDSQL = CMDSQL + CmbOption.text + "' and suboption='"
    CMDSQL = CMDSQL + CmbSubOption.text + "'"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    Text1.text = Trim(M_objrs("scriptsms"))
    
    rowaray = 0
    For i = 1 To Len(Text1.text)
    If Mid(Text1.text, i, 1) = "[" Then
            awalk = i
            AvgMarks(0, rowaray) = i
            
    ElseIf Mid(Text1.text, i, 1) = "]" Then
        akhirk = i
         AvgMarks(1, rowaray) = i
         rowaray = rowaray + 1
    End If
    Next i
    Set M_objrs = Nothing
End Sub

Private Sub Combo1_Click()
If Text5 = "" Then
If Left(Combo1, 1) <> "0" Then
Text5.text = Text5.text & "031" & Combo1.text
Else
Text5.text = Text5.text & Combo1.text
End If
Else
If Left(Combo1, 1) <> "0" Then
Text5.text = Text5.text & ",031" & Combo1.text
Else
Text5.text = Text5.text & "," & Combo1.text
End If
End If
End Sub

Private Sub Combo2_DropDown()
    Dim RSsms_send As ADODB.Recordset
    Dim lst As ListItem
    
    Set RSsms_send = New ADODB.Recordset
    
    RSsms_send.CursorLocation = adUseClient
    CMDSQL = "Select contact1,contact2,mobileno FROM tbl_address WHERE custid='" + FrmCC_Colection.lblCustId + "' AND adr_type='" & Combo2.text & "'"
    RSsms_send.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
     
    Combo1.CLEAR
    'While Not RSsms_send.EOF
    If RSsms_send("contact1") <> "" Then
        Combo1.AddItem Replace(Trim(RSsms_send("contact1")), " ", "")
    End If
    If RSsms_send("contact2") <> "" Then
        Combo1.AddItem Replace(Trim(RSsms_send("contact2")), " ", "")
    End If
    If RSsms_send("mobileno") <> "" Then
        Combo1.AddItem Replace(Trim(RSsms_send("mobileno")), " ", "")
    End If
    '    RSsms_send.MoveNext
    'Wend
    Set RSsms_send = Nothing
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
 Dim teks1 As String
    Dim teks2 As String
    
    teks1 = Replace(Text1.text, "[", "")
    teks2 = Replace(teks1, "]", "")
    teks2 = Replace(teks2, "'", "")

'cek data udah di simpen ke tabel receive apa belum??
Dim fields() As String
fields() = Split(Text5.text, ",")
''MsgBox (UBound(fields) + 1)
For i = 0 To UBound(fields)
'    List1.AddItem Trim$(Fields(i))
'Next

'isi disini


            cmdsqltglsys = "SELECT now() AS tglsystem"
            Set R_tglsys = New ADODB.Recordset
            R_tglsys.Open cmdsqltglsys, M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not R_tglsys.EOF
            TGLw = R_tglsys("tglsystem")
            TGLSERVERc = Format(TGLw, "yyyy-mm-dd hh:mm:ss")
            
            R_tglsys.MoveNext
            Wend
            
            Set R_tglsys = Nothing
            

       CMDSQL = "select * from request_sms where agent='" & Trim$(Text2) & "' and custid='" & Trim$(Text3) & "' and notelp='" & Trim$(fields(i)) & "' and status='0'"
       Set M_objrs = New ADODB.Recordset
       If M_objrs.State = 1 Then M_objrs.Close
              
 M_objrs.CursorLocation = adUseClient
      M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

       If M_objrs.RecordCount = 0 Then
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'If Text1 <> "" Or Text5 <> "" Then



                CMDSQL = "INSERT INTO request_sms "
                CMDSQL = CMDSQL + " ( agent, custid,name,notelp,pesan,status,tgl_kirim)"
                CMDSQL = CMDSQL + " VALUES"
                CMDSQL = CMDSQL + " ( '" + Trim$(Text2) + "', '" + Trim$(Text3) + "', '" + Trim$(Text4) + "', '" + Trim$(fields(i)) + "', '" + Trim$(teks2) + "', '0', '" + TGLSERVERc + "')"
                M_OBJCONN.Execute CMDSQL
'Unload Me
'Else
'End If
End If
Next
'MsgBox "SMS terkirim"
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
    Dim teks1 As String
    Dim teks2 As String
    
    teks1 = Replace(Text1.text, "[", "")
    teks2 = Replace(teks1, "]", "")
    
    MsgBox teks2
End Sub

Private Sub Form_Load()
    Text3 = FrmCC_Colection.lblCustId
    Text4 = Replace(FrmCC_Colection.lblNama, "'", "")
    Text2 = MDIForm1.TxtUsername.text
    
    Load_Data_Option_SMSScript
End Sub


Private Sub Text1_Change()
Label6 = "Jumlah : " & Len(Text1)

If Len(Text1) > 160 Then
MsgBox "Hanya dapat mengirim sms sebanyak 160 Karakter"
End If
End Sub

Private Sub Load_Data_Option_SMSScript()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    CMDSQL = "select distinct option from tblscriptsms"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    CmbOption.CLEAR
    While Not M_objrs.EOF
        CmbOption.AddItem M_objrs("option")
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim cek As Boolean
    cek = False
    For K = 0 To rowaray - 1
        Debug.Print Text1.SelStart
        update
        If Text1.SelStart >= AvgMarks(0, K) And Text1.SelStart < AvgMarks(1, K) Then
  
            If KeyAscii = vbKeyBack Then
                a = Mid(Text1.text, Text1.SelStart, 1)
                If a = "[" Or a = "]" Then
                    KeyAscii = 0
                End If
            End If
            cek = True
            Exit For
        End If
    Next K

    If cek = False Then
        KeyAscii = 0
    End If
End Sub

Public Sub update()
    Dim i As Integer
    rowaray = 0
    For i = 1 To Len(Text1.text)
        If Mid(Text1.text, i, 1) = "[" Then
            awalk = i
            AvgMarks(0, rowaray) = i
        ElseIf Mid(Text1.text, i, 1) = "]" Then
            akhirk = i
            AvgMarks(1, rowaray) = i
            rowaray = rowaray + 1
        End If
    Next i
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        MsgBox "Anda tidak dapat menggunakan klik kanan!", vbCritical + vbOKOnly, "Peringatan"
        Text1.text = ""
    End If
End Sub
