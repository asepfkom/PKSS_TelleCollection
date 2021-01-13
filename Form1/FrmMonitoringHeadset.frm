VERSION 5.00
Begin VB.Form FrmMonitoringHeadset 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Monitoring Penggunaan Headset"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CMbTipeHeadset 
      Height          =   315
      ItemData        =   "FrmMonitoringHeadset.frx":0000
      Left            =   1140
      List            =   "FrmMonitoringHeadset.frx":000A
      TabIndex        =   18
      Top             =   7200
      Width           =   2835
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   9060
      TabIndex        =   16
      Top             =   8700
      Width           =   1635
   End
   Begin VB.TextBox TxtSaran 
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   8040
      Width           =   7815
   End
   Begin VB.TextBox TxtKetFisik 
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   5940
      Width           =   7815
   End
   Begin VB.ComboBox CmbKondisiFisik 
      Height          =   315
      ItemData        =   "FrmMonitoringHeadset.frx":001C
      Left            =   4140
      List            =   "FrmMonitoringHeadset.frx":0029
      TabIndex        =   11
      Top             =   5280
      Width           =   2835
   End
   Begin VB.TextBox TxtKetMic 
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   4440
      Width           =   7815
   End
   Begin VB.ComboBox CmbKondisiMic 
      Height          =   315
      ItemData        =   "FrmMonitoringHeadset.frx":0047
      Left            =   5940
      List            =   "FrmMonitoringHeadset.frx":0057
      TabIndex        =   7
      Top             =   3780
      Width           =   2835
   End
   Begin VB.TextBox TxtKetKOndisiHeadset 
      Height          =   735
      Left            =   540
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2760
      Width           =   7815
   End
   Begin VB.ComboBox CmbKondisiSuara 
      Height          =   315
      ItemData        =   "FrmMonitoringHeadset.frx":007C
      Left            =   5700
      List            =   "FrmMonitoringHeadset.frx":008C
      TabIndex        =   3
      Top             =   1980
      Width           =   2835
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Perhatian"
      ForeColor       =   &H0000FFFF&
      Height          =   1275
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10635
      Begin VB.Label Label1 
         BackColor       =   &H0000C000&
         Caption         =   $"FrmMonitoringHeadset.frx":00B1
         Height          =   735
         Left            =   300
         TabIndex        =   1
         Top             =   360
         Width           =   10155
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Anda lebih nyaman menggunakan headset Stereo (2 speaker) atau Mono (hanya 1 speaker)?"
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   300
      TabIndex        =   17
      Top             =   6900
      Width           =   6135
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Apa saran anda bagi kami, agar kontrol headset dan pelayanan penggunaan headset lebih baik? "
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   7740
      Width           =   9495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Beri penjelasan mengenai kondisi fisik headset yang anda gunakan:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   5700
      Width           =   5175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Bagaimana kondisi fisik headset yang anda gunakan?"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   5340
      Width           =   3975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Beri komentar anda terhadap kondisi mic headset yang anda gunakan pada kolom di bawah ini:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4200
      Width           =   9495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Bagaimana kondisi mic anda? (suara anda yang terdengar di CH saat menelpon) "
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   9495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   $"FrmMonitoringHeadset.frx":0151
      ForeColor       =   &H000000FF&
      Height          =   915
      Left            =   300
      TabIndex        =   5
      Top             =   2520
      Width           =   9495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   795
      Left            =   10200
      Shape           =   1  'Square
      Top             =   1020
      Width           =   555
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   $"FrmMonitoringHeadset.frx":01D9
      ForeColor       =   &H000000FF&
      Height          =   915
      Left            =   300
      TabIndex        =   2
      Top             =   1500
      Width           =   9495
   End
End
Attribute VB_Name = "FrmMonitoringHeadset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VSAVE As Boolean
Dim TglMonitor As String
Dim TglSekrg As String





Private Sub CmbKondisiFisik_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmbKondisiMic_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmbKondisiSuara_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub CMbTipeHeadset_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdOK_Click()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    Call CekValid
    
    If VSAVE = False Then
        Exit Sub
    End If
    
    CMDSQL = "insert into monitoring_headset (agent,nama,"
    CMDSQL = CMDSQL + "tglmonitoring,tglisibyagent,kondisi_suara,"
    CMDSQL = CMDSQL + "ket_kondisisuara,kondisi_mic,ket_kondisimic,"
    CMDSQL = CMDSQL + "kondisi_fisik,ket_kondisifisik,saran,tipe_headset) values ('"
    CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.Text) + "','"
    CMDSQL = CMDSQL + Trim(MDIForm1.txtNama.Text) + "','"
    CMDSQL = CMDSQL + TglMonitor + "','"
    CMDSQL = CMDSQL + TglSekrg + "','"
    CMDSQL = CMDSQL + Trim(CmbKondisiSuara.Text) + "','"
    CMDSQL = CMDSQL + Trim(TxtKetKOndisiHeadset.Text) + "','"
    CMDSQL = CMDSQL + Trim(CmbKondisiMic.Text) + "','"
    CMDSQL = CMDSQL + Trim(TxtKetMic.Text) + "','"
    CMDSQL = CMDSQL + Trim(CmbKondisiFisik.Text) + "','"
    CMDSQL = CMDSQL + Trim(TxtKetFisik.Text) + "','"
    CMDSQL = CMDSQL + Trim(TxtSaran.Text) + "','"
    CMDSQL = CMDSQL + Trim(CMbTipeHeadset.Text) + "')"
    
    M_OBJCONN.Execute CMDSQL
    
    'Update status di usertbl, bahwa user tersebut sudah mengisi form evalusai headset
    CMDSQL = "update usertbl set monitoring_headset='0' where userid='"
    CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.Text) + "'"
    M_OBJCONN.Execute CMDSQL
    
    MsgBox "Terima kasih atas partisipasi anda mengisi form evaluasi headset! Form ini secara berkala akan muncul kembali, untuk mengevaluasi headset! :)", vbOKOnly + vbInformation, "Informasi"
    
    Me.Hide
    
End Sub

Private Sub CekValid()
    If CmbKondisiFisik.Text = "" Or _
        CmbKondisiMic.Text = "" Or _
        CmbKondisiSuara.Text = "" Or _
        TxtKetFisik.Text = "" Or _
        TxtKetKOndisiHeadset.Text = "" Or _
        TxtKetMic.Text = "" Or _
        TxtSaran.Text = "" Or _
        CMbTipeHeadset.Text = "" Then
        
        MsgBox "Mohon isi semua form isian Headset!", vbOKOnly + vbInformation, "Informasi"
        VSAVE = False
     Else
        VSAVE = True
     End If
End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    CMDSQL = "select * from manajemen_site where nama_monitoring='HEADSET'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        TglMonitor = Format(M_objrs("tgl_monitoring"), "yyyy-mm-dd")
    Set M_objrs = Nothing
    
    CMDSQL = "select now()"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
        TglSekrg = Format(M_objrs(0), "yyyy-mm-dd hh:mm:ss")
    Set M_objrs = Nothing
End Sub
