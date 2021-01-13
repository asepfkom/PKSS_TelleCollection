VERSION 5.00
Begin VB.Form FrmSendSmsNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send SMS"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1320
      Width           =   2950
   End
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
      ItemData        =   "FrmSendSmsNew.frx":0000
      Left            =   6000
      List            =   "FrmSendSmsNew.frx":0002
      TabIndex        =   22
      Top             =   1395
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtnm_agent 
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
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   5760
      Visible         =   0   'False
      Width           =   3255
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
      ItemData        =   "FrmSendSmsNew.frx":0004
      Left            =   1200
      List            =   "FrmSendSmsNew.frx":0006
      TabIndex        =   9
      Top             =   1320
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
      Left            =   2280
      TabIndex        =   8
      Top             =   4515
      Width           =   1095
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   240
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   3255
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   3255
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
      Left            =   3480
      TabIndex        =   4
      Top             =   4515
      Width           =   1095
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
      Left            =   1200
      MaxLength       =   160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2625
      Width           =   3495
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
      ItemData        =   "FrmSendSmsNew.frx":0008
      Left            =   1200
      List            =   "FrmSendSmsNew.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1740
      Width           =   2535
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
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2175
      Width           =   3405
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
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
      Left            =   1260
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label Label2 
      Caption         =   "Type SMS"
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
      Left            =   120
      TabIndex        =   25
      Top             =   1785
      Width           =   975
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
      Left            =   4920
      TabIndex        =   23
      Top             =   1395
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblid 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   20
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label LblLayer 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   930
      TabIndex        =   19
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Layer:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   18
      Top             =   4695
      Width           =   735
   End
   Begin VB.Label Label5 
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
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
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
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   5655
      Left            =   0
      Top             =   0
      Width           =   4890
   End
   Begin VB.Label Label2 
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
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   2745
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   13
      Top             =   1395
      Width           =   975
   End
   Begin VB.Label Label6 
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
      Left            =   240
      TabIndex        =   12
      Top             =   4395
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Option:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6270
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Sub option:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6630
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FrmSendSmsNew"
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

Private Sub CmbOption_Change()
     '@@Jika sms free style
    If CmbOption.text = "Free SMS Style" Then
        CmbSubOption.text = "Free SMS Style"
        Text1.text = "[]"
    End If
End Sub

Private Sub get_optionid()
    Dim M_objrs     As ADODB.Recordset
    Dim CMDSQL      As String
    
    CMDSQL = "SELECT id FROM tblscriptsms WHERE option='" & Trim(CmbOption.text) & "' " & _
            "AND suboption='" & Trim(CmbSubOption.text) & "'"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        lblID.Caption = M_objrs!ID
    End If
  
    Set M_objrs = Nothing
End Sub

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
        '@@Jika sms free style
        If CmbOption.text = "Free SMS Style" Then
            Text1.Visible = False
            Text1.text = ""
            TxtSmsFreeStyle.Visible = True
            TxtSmsFreeStyle.text = ""
            TxtSmsFreeStyle.SetFocus
        End If
        Set M_objrs = Nothing
        Exit Sub
    Else
        Text1.Visible = True
'        TxtSmsFreeStyle.Visible = False
'        TxtSmsFreeStyle.Text = ""
    End If
    
    CmbSubOption.CLEAR
    While Not M_objrs.EOF
        CmbSubOption.AddItem M_objrs("suboption")
        M_objrs.MoveNext
    Wend
    
  
    Set M_objrs = Nothing
End Sub

Private Sub CmbOption_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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
    If CmbOption.text = "Free SMS Style" Then
        Text1.text = "[]"
        lblID.Caption = 0
    Else
        Call get_optionid
    End If
    
    Text1.text = Replace(Text1.text, "*agent*", txtnm_agent.text)
    Text1.text = Replace(Text1.text, "*cust*", Text4.text)
    
    Set M_objrs = Nothing
End Sub

Private Sub CmbSubOption_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo1_Click()
If Text5 = "" Then
If Left(Combo1, 1) <> "0" Then
Text5.text = Text5.text & "021" & Combo1.text
Else
Text5.text = Text5.text & Combo1.text
End If
Else
If Left(Combo1, 1) <> "0" Then
Text5.text = Text5.text & ",021" & Combo1.text
Else
Text5.text = Text5.text & "," & Combo1.text
End If
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo2_Click()
    Dim RSsms_send As ADODB.Recordset
    Dim lst As ListItem
    
    Set RSsms_send = New ADODB.Recordset
    
    RSsms_send.CursorLocation = adUseClient
    CMDSQL = "Select homeno,officeno,mobileno FROM mgm WHERE custid='" + FrmCC_Colection.lblCustId.text + "'"
    RSsms_send.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
     
    Combo1.CLEAR
    'While Not RSsms_send.EOF
    If RSsms_send.RecordCount > 0 Then
        If RSsms_send("homeno") <> "" Then
            Combo1.AddItem Replace(Trim(RSsms_send("homeno")), " ", "")
        End If
        If RSsms_send("officeno") <> "" Then
            Combo1.AddItem Replace(Trim(RSsms_send("officeno")), " ", "")
        End If
        If RSsms_send("mobileno") <> "" Then
            Combo1.AddItem Replace(Trim(RSsms_send("mobileno")), " ", "")
        End If
        '    RSsms_send.MoveNext
        'Wend
    End If
    Set RSsms_send = Nothing
End Sub

Private Sub Combo2_DropDown()
    Dim RSsms_send As ADODB.Recordset
    Set RSsms_send = New ADODB.Recordset
    
    RSsms_send.CursorLocation = adUseClient
    CMDSQL = "Select distinct adr_type FROM tbl_address WHERE custid='" + FrmCC_Colection.lblCustId.text + "' "
    RSsms_send.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
     Combo2.CLEAR
        While Not RSsms_send.EOF
            Combo2.AddItem IIf(IsNull(RSsms_send!adr_type), "", RSsms_send!adr_type)
            RSsms_send.MoveNext
        Wend
     Set RSsms_send = Nothing
End Sub

Private Sub Command1_Click()
    Dim teks1 As String
    Dim teks2 As String
    Dim fields() As String
    Dim banyaksms As Integer
    Dim pesan As String
    Dim aa As Integer
    Dim m_objrscekkuota As ADODB.Recordset
    Dim SisaSms As Integer
    Dim SisaSmsSekrg As Integer
    Dim remarks_x As String
    Dim M_objrs As ADODB.Recordset

    Set M_objrs = New ADODB.Recordset
    
    teks2 = Text1.text
    
    'cek data udah di simpen ke tabel receive apa belum??
    fields() = Split(Text5.text, ",")
    For i = 0 To UBound(fields)
        '@@ 09022011 - Ambil tanggal system
        cmdsqltglsys = "SELECT now() AS tglsystem"
        Set R_tglsys = New ADODB.Recordset
        R_tglsys.Open cmdsqltglsys, M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not R_tglsys.EOF
            TGLw = R_tglsys("tglsystem")
            TGLSERVERc = Format(TGLw, "yyyy-mm-dd hh:mm:ss")
            R_tglsys.MoveNext
        Wend
        Set R_tglsys = Nothing
                
        CMDSQL = "SELECT * FROM request_sms WHERE custid='" & Trim$(Text3) & "' " & _
                " AND notelp='" & Trim$(fields(i)) & "' AND date(tgl_kirim)=date(now()) AND id_option=" & Val(lblID.Caption) & ""
           
        If M_objrs.State = 1 Then M_objrs.Close
     
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
'        If M_objrs.RecordCount = 0 Then
            CMDSQL = "INSERT INTO request_sms "
            CMDSQL = CMDSQL + " ( agent, custid,name,notelp,pesan,status,tgl_kirim,id_option)"
            CMDSQL = CMDSQL + " VALUES"
            CMDSQL = CMDSQL + " ( '" + Trim$(Text2) + "', '" + Trim$(Text3) + "', '" + Trim$(Text4) + "', '" + Trim$(fields(i)) + "', '" + Trim$(teks2) + "', '0', '" + TGLSERVERc + "'," & Val(lblID.Caption) & ")"
            M_OBJCONN.Execute CMDSQL
'        Else
'            MsgBox "Anda sudah mengirim sms ke no:" & Trim$(fields(i)) & ". Sebelumnya. SMS gagal dikirim!", vbOKOnly + vbInformation, "Informasi"
'            Exit Sub
'        End If
    Next
    
    CMDSQL = "INSERT INTO tblnotif_info "
    CMDSQL = CMDSQL & "( type_notif,notif_from) values ('"
    CMDSQL = CMDSQL & "sms','" & Trim$(Text2) & "')"
    M_OBJCONN.Execute CMDSQL
    
    MsgBox "Sms berhasil disimpan! Akan dikirim setelah di approve oleh SPV!", vbOKOnly + vbInformation, "Informasi"
    
    Set M_objrs = Nothing
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
Dim RSsms_send As ADODB.Recordset
Set RSsms_send = New ADODB.Recordset
Dim lst As ListItem

On Error Resume Next

RSsms_send.CursorLocation = adUseClient
CMDSQL = "SELECT btrim as no_tlp FROM ("
CMDSQL = CMDSQL + "    SELECT trim(mobileno) FROM mgm WHERE trim(mobileno) not in (select no_telp from tblblacklist) and custid        = '" + FrmCC_Colection.lblCustId + "' "
CMDSQL = CMDSQL + "    Union All"
CMDSQL = CMDSQL + "    SELECT trim(mobileno2) FROM mgm WHERE trim(mobileno2) not in (select no_telp from tblblacklist) and            custid = '" + FrmCC_Colection.lblCustId + "' "
CMDSQL = CMDSQL + "    Union All"
CMDSQL = CMDSQL + "    SELECT trim(mobilenoadd1) FROM mgm WHERE trim(mobilenoadd1) not in (select no_telp from tblblacklist) and       custid = '" + FrmCC_Colection.lblCustId + "' "
CMDSQL = CMDSQL + "    Union All"
CMDSQL = CMDSQL + "    SELECT trim(mobilenoadd2) FROM mgm WHERE trim(mobilenoadd2) not in (select no_telp from tblblacklist) and       custid = '" + FrmCC_Colection.lblCustId + "') a "
RSsms_send.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

While Not RSsms_send.EOF
    Combo1.AddItem Replace(Trim(RSsms_send("no_tlp")), " ", "")
    RSsms_send.MoveNext
Wend




'RSsms_send.CursorLocation = adUseClient
'cmdsql = "Select * from mgm where custid='" + FrmCC_Colection.lblCustId + "'"
'RSsms_send.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'While Not RSsms_send.EOF
'
'    If (IsNull(RSsms_send("mobileno"))) Or RSsms_send("mobileno") = "" Then
'    Else
'        '@@281010 Cek apakah no telepon masuk dalam blacklist??
'        If Trim(RSsms_send("f_mobileno")) = "0" Then
'            Combo1.AddItem Replace(Trim(RSsms_send("mobileno")), " ", "")
'        End If
'    End If
'
'    If (IsNull(RSsms_send("mobileno2"))) Or RSsms_send("mobileno2") = "" Then
'    Else
'        '@@281010 Cek apakah no telepon masuk dalam blacklist??
'        If Trim(RSsms_send("f_mobileno2")) = "0" Then
'            Combo1.AddItem Replace(Trim(RSsms_send("mobileno2")), " ", "")
'        End If
'    End If
'
'    If (IsNull(RSsms_send("mobilenoadd1"))) Or RSsms_send("mobilenoadd1") = "" Then
'    Else
'        '@@281010 Cek apakah no telepon masuk dalam blacklist??
'        If Trim(RSsms_send("f_mobilenoadd1")) = "0" Then
'            Combo1.AddItem Replace(Trim(RSsms_send("mobilenoadd1")), " ", "")
'        End If
'    End If
'
'    If (IsNull(RSsms_send("mobilenoadd2"))) Or RSsms_send("mobilenoadd2") = "" Then
'    Else
'        '@@281010 Cek apakah no telepon masuk dalam blacklist??
'        If Trim(RSsms_send("f_mobilenoadd2")) = "0" Then
'            Combo1.AddItem Replace(Trim(RSsms_send("mobilenoadd2")), " ", "")
'        End If
'    End If


Set RSsms_send = Nothing

Text3 = FrmCC_Colection.lblCustId
Text4 = Replace(FrmCC_Colection.lblNama, "'", "")
Text2 = MDIForm1.TxtUsername.text
txtnm_agent.text = MDIForm1.txtnama.text

Load_Data_Option_SMSScript
End Sub


Private Sub Text1_Change()
Label6 = "Jumlah : " & Len(Text1)

LblLayer.Caption = Ceiling(Val(Len(Trim(Text1.text))) / 160)


If Len(Text1) > 320 Then
    MsgBox "Hanya dapat mengirim sms sebanyak 320 Karakter"
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
    'CmbOption.AddItem "Free SMS Style"
    While Not M_objrs.EOF
        CmbOption.AddItem M_objrs("option")
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'    Dim cek As Boolean
'    cek = False
'    For K = 0 To rowaray - 1
'        Debug.Print Text1.SelStart
'        update
'        If Text1.SelStart >= AvgMarks(0, K) And Text1.SelStart < AvgMarks(1, K) Then
'
'            If KeyAscii = vbKeyBack Then
'                a = Mid(Text1.Text, Text1.SelStart, 1)
'                If a = "[" Or a = "]" Then
'                    KeyAscii = 0
'                End If
'            End If
'            cek = True
'            Exit For
'        End If
'    Next K
'
'    If cek = False Then
'        KeyAscii = 0
'    End If
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

'Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 Then
'        MsgBox "Anda tidak dapat menggunakan klik kanan!", vbCritical + vbOKOnly, "Peringatan"
'        Text1.Text = ""
'    End If
'End Sub

Private Sub TxtSmsFreeStyle_Change()
    Label6 = "Jumlah : " & Len(TxtSmsFreeStyle.text)
    
    LblLayer.Caption = Ceiling(Val(Len(Trim(TxtSmsFreeStyle.text))) / 160)

    If Len(TxtSmsFreeStyle.text) > 320 Then
        MsgBox "Hanya dapat mengirim sms sebanyak 320 Karakter"
    End If
End Sub
'@@09022011 Fungsi buat membulatkan desimal
Private Function Ceiling(number As Double) As Long
    Ceiling = -Int(-number)
End Function

