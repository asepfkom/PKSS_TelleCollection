VERSION 5.00
Begin VB.Form FrmSendSmsNew2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send SMS - Reply"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.ComboBox CmbSubOption 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2100
      Width           =   3405
   End
   Begin VB.ComboBox CmbOption 
      Height          =   315
      ItemData        =   "FrmSendSmsNew2.frx":0000
      Left            =   1200
      List            =   "FrmSendSmsNew2.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1740
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   1260
      MaxLength       =   160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   3540
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2340
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      Top             =   1320
      Width           =   3315
   End
   Begin VB.TextBox TxtSmsFreeStyle 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   1260
      MaxLength       =   320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label8 
      Caption         =   "Sub option:"
      Height          =   255
      Left            =   180
      TabIndex        =   21
      Top             =   2070
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Option:"
      Height          =   255
      Left            =   180
      TabIndex        =   20
      Top             =   1710
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Jumlah :"
      Height          =   255
      Left            =   300
      TabIndex        =   19
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Mobile No :"
      Height          =   255
      Left            =   180
      TabIndex        =   18
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Text :"
      Height          =   255
      Left            =   180
      TabIndex        =   17
      Top             =   2520
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   5175
      Left            =   60
      Top             =   0
      Width           =   4890
   End
   Begin VB.Label Label3 
      Caption         =   "Agent :"
      Height          =   255
      Left            =   180
      TabIndex        =   16
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Custid :"
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Nama :"
      Height          =   255
      Left            =   180
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Layer:"
      Height          =   315
      Left            =   300
      TabIndex        =   13
      Top             =   4620
      Width           =   735
   End
   Begin VB.Label LblLayer 
      Caption         =   "1"
      Height          =   315
      Left            =   1080
      TabIndex        =   12
      Top             =   4620
      Width           =   1095
   End
End
Attribute VB_Name = "FrmSendSmsNew2"
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
    If CmbOption.Text = "Free SMS Style" Then
        CmbSubOption.Text = "Free SMS Style"
        Text1.Text = "[]"
    End If
End Sub

Private Sub CmbOption_Click()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    Text1.Text = ""
    CmbSubOption.CLEAR
    
    CMDSQL = "select * from tblscriptsms where option='"
    CMDSQL = CMDSQL + CmbOption.Text + "'"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    
    If M_objrs.RecordCount = 0 Then
        '@@Jika sms free style
        If CmbOption.Text = "Free SMS Style" Then
            Text1.Visible = False
            Text1.Text = ""
            TxtSmsFreeStyle.Visible = True
            TxtSmsFreeStyle.Text = ""
            TxtSmsFreeStyle.SetFocus
        End If
        Set M_objrs = Nothing
        Exit Sub
    Else
        Text1.Visible = True
        TxtSmsFreeStyle.Visible = False
        TxtSmsFreeStyle.Text = ""
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
    
    Text1.Text = ""
    CMDSQL = "select * from tblscriptsms where option='"
    CMDSQL = CMDSQL + CmbOption.Text + "' and suboption='"
    CMDSQL = CMDSQL + CmbSubOption.Text + "'"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    

    
    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    Text1.Text = Trim(M_objrs("scriptsms"))
    
    rowaray = 0
    For i = 1 To Len(Text1.Text)
    If Mid(Text1.Text, i, 1) = "[" Then
            awalk = i
            AvgMarks(0, rowaray) = i
            
    ElseIf Mid(Text1.Text, i, 1) = "]" Then
        akhirk = i
         AvgMarks(1, rowaray) = i
         rowaray = rowaray + 1
    End If
    Next i
    If CmbOption.Text = "Free SMS Style" Then
        Text1.Text = "[]"
    End If
    Set M_objrs = Nothing
End Sub

Private Sub CmbSubOption_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo1_Click()
If Text5 = "" Then
If Left(Combo1, 1) <> "0" Then
Text5.Text = Text5.Text & "031" & Combo1.Text
Else
Text5.Text = Text5.Text & Combo1.Text
End If
Else
If Left(Combo1, 1) <> "0" Then
Text5.Text = Text5.Text & ",031" & Combo1.Text
Else
Text5.Text = Text5.Text & "," & Combo1.Text
End If
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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
    
   

    '@@ 09022011 Ini jika  free sms style
    'If CmbOption.Text = "Free SMS Style" Then
        
        teks1 = Replace(TxtSmsFreeStyle.Text, "[", "")
        
        teks1 = Replace(teks1, "'", "")
        teks1 = Replace(teks1, "\", "")
        
        teks2 = Replace(teks1, "]", "")
        teks2 = Replace(teks2, "'", "")
        
        'cek data udah di simpen ke tabel receive apa belum??
        fields() = Split(Text5.Text, ",")
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
                
            CMDSQL = "select * from request_sms where agent='" & Trim$(Text2) & "' and custid='" & Trim$(Text3) & "' and notelp='" & Trim$(fields(i)) & "' and status='0'"
            Set M_objrs = New ADODB.Recordset
           
            If M_objrs.State = 1 Then M_objrs.Close
     
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
            If M_objrs.RecordCount = 0 Then
                
                'Jika yang login admin/spv tidak usah hitung kuota sms
                If UCase(Trim(MDIForm1.txtlevel.Text)) = "ADMIN" Or UCase(Trim(MDIForm1.txtlevel.Text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.txtlevel.Text)) = "SUPERVISOR" Or UCase(Trim(MDIForm1.txtlevel.Text)) = "AGENT" Then
                    GoTo lanjut
                End If
                'Cek dulu deh udah berapa kali kirim
'                Set m_objrscekkuota = New ADODB.Recordset
'                m_objrscekkuota.CursorLocation = adUseClient
'                CMDSQL = "select kuota_sms from usertbl where userid='"
'                CMDSQL = CMDSQL + Trim$(Text2) + "'"
'                m_objrscekkuota.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'                SisaSms = Val(MDIForm1.KuotaSms) - Val(m_objrscekkuota("kuota_sms"))
'                If Val(m_objrscekkuota("kuota_sms")) >= Val(MDIForm1.KuotaSms) Then
'                    MsgBox "Kuota sms anda hari ini sudah habis! Max.sms/hari:" & MDIForm1.KuotaSms & ". Silahkan coba esok hari!", vbOKOnly + vbInformation, "Informasi"
'                    Exit Sub
'                    Set m_objrscekkuota = Nothing
'                End If
'
'                'Jika sisa sms lebih kecil dari sisa kuota maka batalkan
'                If SisaSms < Val(LblLayer.Caption) Then
'                    MsgBox "SMS yang anda kirim melebihi kuota! Sisa sms yang diperbolehkan hari ini:" & SisaSms & " layer(160 karakter)", vbOKOnly + vbInformation, "Informasi"
'                    Set m_objrscekkuota = Nothing
'                    Exit Sub
'                End If
'
'                'Nah di sini jika kuoat masih tersedia, kurangi deh kuota dari agent tersebut
'                SisaSmsSekrg = Val(LblLayer.Caption) + Val(m_objrscekkuota("kuota_sms"))
'                CMDSQL = "update usertbl set kuota_sms='"
'                CMDSQL = CMDSQL + CStr(SisaSmsSekrg) + "' where userid='"
'                CMDSQL = CMDSQL + Trim$(Text2) + "'"
'                M_OBJCONN.Execute CMDSQL
'                Set m_objrscekkuota = Nothing
lanjut:

                CMDSQL = "INSERT INTO request_sms "
                CMDSQL = CMDSQL + " ( agent, custid,name,notelp,pesan,status,tgl_kirim)"
                CMDSQL = CMDSQL + " VALUES"
                CMDSQL = CMDSQL + " ( '" + Trim$(Text2) + "', '" + Trim$(Text3) + "', '" + Trim$(Text4) + "', '" + Trim$(fields(i)) + "', '" + Text1.Text + "', '0', '" + TGLSERVERc + "')"
                M_OBJCONN.Execute CMDSQL
                
            MsgBox "Sms berhasil disimpan! Akan dikirim setelah di approve oleh SPV!", vbOKOnly + vbInformation, "Informasi"
            Else
                MsgBox "Anda sudah mengirim sms ke no:" & Trim$(fields(i)) & ". Sebelumnya sms ke no ini belum di approve oleh SPV! SMS gagal dikirim!", vbOKOnly + vbInformation, "Informasi"
            End If
        Next
        
    
'    Else
'        '@@ Ini jika sms yang sesuai dengan format yang sudah ditentukan
'        teks1 = Replace(Text1.Text, "[", "")
'        teks2 = Replace(teks1, "]", "")
'        teks2 = Replace(teks2, "'", "")
'
'        'cek data udah di simpen ke tabel receive apa belum??
'        fields() = Split(Text5.Text, ",")
'        For i = 0 To UBound(fields)
'            '@@ 09022011 - Ambil tanggal system
'            cmdsqltglsys = "SELECT now() AS tglsystem"
'            Set R_tglsys = New ADODB.Recordset
'            R_tglsys.Open cmdsqltglsys, M_OBJCONN, adOpenDynamic, adLockOptimistic
'            While Not R_tglsys.EOF
'                TGLw = R_tglsys("tglsystem")
'                TGLSERVERc = Format(TGLw, "yyyy-mm-dd hh:mm:ss")
'                R_tglsys.MoveNext
'            Wend
'            Set R_tglsys = Nothing
'
'            'CMDSQL = "select * from request_sms_log where agent='" & Trim$(Text2) & "' and custid='" & Trim$(Text3) & "' and notelp='" & Trim$(fields(i)) & "' and status='0'"
'            'Set M_OBJRS = New ADODB.Recordset
'
'            'If M_OBJRS.state = 1 Then M_OBJRS.Close
'
'            'M_OBJRS.CursorLocation = adUseClient
'            'M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'            'If M_OBJRS.RecordCount = 0 Then
'
'
'                CMDSQL = "INSERT INTO request_sms_log "
'                CMDSQL = CMDSQL + " ( agent, custid,name,notelp,pesan,status,tgl_kirim)"
'                CMDSQL = CMDSQL + " VALUES"
'                CMDSQL = CMDSQL + " ( '" + Trim$(Text2) + "', '" + Trim$(Text3) + "', '" + Trim$(Text4) + "', '" + Trim$(fields(i)) + "', '" + Trim$(teks2) + "', '0', '" + TGLSERVERc + "')"
'                M_OBJCONN.Execute CMDSQL
'
'                '@@ 09022011 Tambahan jika pesan lebih dari 160 karakter
'                banyaksms = Ceiling(Val(Len(Trim$(fields(i)))) / 160)
'                For aa = 1 To banyaksms
'                    awalpesan = (160 * aa) - 160
'                    pesan = Mid(Trim$(fields(i)), awalpesan + 1, 160)
'                    CMDSQL = "insert into outbox (`DestinationNumber`,`TextDecoded`,`SenderID`,`CreatorID`) values ('"
'                    CMDSQL = CMDSQL + Trim(pesan) + "','"
'                    CMDSQL = CMDSQL + Trim$(teks2) + "','phone2','"
'                    CMDSQL = CMDSQL + Trim$(Text3) + "-" + Trim$(Text2) + "')"
'                    M_OBJCONN1.Execute CMDSQL
'                Next aa
'                MsgBox "SMS anda sudah terkirim!Sebanyak :" & banyaksms, vbOKOnly + vbInformation, "Informasi"
'
'            'End If
'        Next
'    End If
    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
    Dim teks1 As String
    Dim teks2 As String
    
    teks1 = Replace(Text1.Text, "[", "")
    teks2 = Replace(teks1, "]", "")
    
    MsgBox teks2
End Sub

Private Sub Form_Load()
    Load_Data_Option_SMSScript
End Sub

Private Sub Text1_Change()
Label6 = "Jumlah : " & Len(Text1)

LblLayer.Caption = Ceiling(Val(Len(Trim(Text1.Text))) / 160)


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
    Dim cek As Boolean
    cek = False
    For K = 0 To rowaray - 1
        Debug.Print Text1.SelStart
        update
        If Text1.SelStart >= AvgMarks(0, K) And Text1.SelStart < AvgMarks(1, K) Then
  
            If KeyAscii = vbKeyBack Then
                a = Mid(Text1.Text, Text1.SelStart, 1)
                If a = "[" Or a = "]" Then
                    KeyAscii = 0
                End If
            End If
            cek = True
            Exit For
        End If
    Next K

'    If cek = False Then
'        KeyAscii = 0
'    End If
End Sub

Public Sub update()
    Dim i As Integer
    rowaray = 0
    For i = 1 To Len(Text1.Text)
        If Mid(Text1.Text, i, 1) = "[" Then
            awalk = i
            AvgMarks(0, rowaray) = i
        ElseIf Mid(Text1.Text, i, 1) = "]" Then
            akhirk = i
            AvgMarks(1, rowaray) = i
            rowaray = rowaray + 1
        End If
    Next i
End Sub
'
'Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 Then
'        MsgBox "Anda tidak dapat menggunakan klik kanan!", vbCritical + vbOKOnly, "Peringatan"
'        Text1.Text = ""
'    End If
'End Sub

Private Sub TxtSmsFreeStyle_Change()
    Label6 = "Jumlah : " & Len(TxtSmsFreeStyle.Text)
    
    LblLayer.Caption = Ceiling(Val(Len(Trim(TxtSmsFreeStyle.Text))) / 160)

    If Len(TxtSmsFreeStyle.Text) > 320 Then
        MsgBox "Hanya dapat mengirim sms sebanyak 320 Karakter"
    End If
End Sub
'@@09022011 Fungsi buat membulatkan desimal
Private Function Ceiling(number As Double) As Long
    Ceiling = -Int(-number)
End Function


