VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FRMSENDMSG 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5805
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "..."
      Height          =   330
      Left            =   4710
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   975
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   615
      Width           =   3585
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   3420
      TabIndex        =   4
      Top             =   5250
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   4335
      TabIndex        =   5
      Top             =   5250
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Kirim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2505
      TabIndex        =   3
      Top             =   5250
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   4980
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   4065
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   7170
         _Version        =   393217
         ScrollBars      =   1
         TextRTF         =   $"FRMSENDMSG.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MESSAGE"
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
      TabIndex        =   9
      Top             =   30
      Width           =   1935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "TO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   375
      TabIndex        =   6
      Top             =   675
      Width           =   585
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D4B9AF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "FRMSENDMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
Dim CMDSQL As String
Dim NAMA_MSG As String
Dim i As Integer
Dim NAMA_AM As String
Dim NAMA_SPVCODE As String
Dim M_OBJspv As ADODB.Recordset

On Error Resume Next
Select Case Index
Case 0
    If Text1.Text = Empty Then
        MsgBox "Penerima Message Harus Diisi...!!", vbInformation + vbOKOnly, "Aplikasi"
        Exit Sub
    End If
    If RichTextBox1.Text = Empty Then
        MsgBox "Isi Message Tidak Boleh Kosong...!!", vbInformation + vbOKOnly, "Aplikasi"
        Exit Sub
    End If
    
    ' INI CC NYA''' KE SUPERVISOR TIAP ANAK DAN KE AM
  '  If UCase(Trim(mdiform1.txtlevel.text)) = "AGENT" Then
        Set M_OBJspv = New ADODB.Recordset
        M_OBJspv.CursorLocation = adUseClient
        M_OBJspv.Open "Select AM,SPVCODE FROM usertbl WHERE USERID ='" + Trim(MDIForm1.TxtUsername.Text) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        While Not M_OBJspv.EOF
'            'NAMA_AM = IIf(IsNull(M_OBJspv!AM), "", M_OBJspv!AM)
'            NAMA_SPVCODE = IIf(IsNull(M_OBJspv!SPVCODE), "", M_OBJspv!SPVCODE)
'                CMDSQL = "INSERT INTO MSGTBL "
'                CMDSQL = CMDSQL + " ( RECIPIENT,"
'                CMDSQL = CMDSQL + " DATETIME,"
'                CMDSQL = CMDSQL + " SENDER,"
'                CMDSQL = CMDSQL + " SENTFROM,"
'                CMDSQL = CMDSQL + " MSG)"
'                CMDSQL = CMDSQL + " VALUES"
'                CMDSQL = CMDSQL + " ( '" + Trim(NAMA_AM) + "',"
'                CMDSQL = CMDSQL + " '" + Format(Date, "yyyymmdd") + "',"
'                CMDSQL = CMDSQL + " '" + Trim(MDIForm1.TxtUsername.Text) + "',"
'                CMDSQL = CMDSQL + " '" + CStr(MDIForm1.Winsock1.LocalIP) + "',"
'                CMDSQL = CMDSQL + " '" + RichTextBox1.Text & vbCr & " Dikirim Ke (" & Text1.Text & ")" & "')"
'                M_OBJCONN.Execute CMDSQL
'
''                CMDSQL = "INSERT INTO MSGTBL "
''                CMDSQL = CMDSQL + " ( RECIPIENT,"
''                CMDSQL = CMDSQL + " DATETIME,"
''                CMDSQL = CMDSQL + " SENDER,"
''                CMDSQL = CMDSQL + " SENTFROM,"
''                CMDSQL = CMDSQL + " MSG)"
''                CMDSQL = CMDSQL + " VALUES"
''                CMDSQL = CMDSQL + " ( '" + Trim(NAMA_SPVCODE) + "',"
''                CMDSQL = CMDSQL + " '" + Format(Date, "yyyymmdd") + "',"
''                CMDSQL = CMDSQL + " '" + mdiform1.txtusername.text + "',"
''                CMDSQL = CMDSQL + " '" + CStr(MDIForm1.Winsock1.LocalIP) + "',"
''                CMDSQL = CMDSQL + " '" + RichTextBox1.Text & vbCr & " Dikirim Ke (" & Text1.Text & ")" & "')"
''                'cmdsql = cmdsql + " '" + RichTextBox1.Text + "')"
''                M_OBJCONN.Execute CMDSQL
'            M_OBJspv.MoveNext
'        Wend
'        Set M_OBJspv = Nothing
   ' End If
        For i = 1 To Len(Text1.Text)
        Select Case Mid(Text1.Text, i, 1)
        Case ";"
            CMDSQL = "INSERT INTO msgtbl "
            CMDSQL = CMDSQL + " ( RECIPIENT,"
            CMDSQL = CMDSQL + " DATETIME,"
            CMDSQL = CMDSQL + " SENDER,"
            CMDSQL = CMDSQL + " SENTFROM,"
            CMDSQL = CMDSQL + " MSG)"
            CMDSQL = CMDSQL + " VALUES"
            CMDSQL = CMDSQL + " ( '" + Trim(Left(NAMA_MSG, 10)) + "',"
            CMDSQL = CMDSQL + " '" + Format(Date, "yyyymmdd") + "',"
            CMDSQL = CMDSQL + " '" + Trim(MDIForm1.TxtUsername.Text) + "',"
            CMDSQL = CMDSQL + " '" + CStr(MDIForm1.Winsock1.LocalIP) + "',"
            CMDSQL = CMDSQL + " '" + RichTextBox1.Text + "')"
            M_OBJCONN.Execute CMDSQL
            NAMA_MSG = ""
        Case Else
            NAMA_MSG = NAMA_MSG + Mid(Text1.Text, i, 1) 'add to txt
        End Select
        Next i
        Unload Me
'        If MDIForm1.Winsock1.State <> 7 Then
'            MDIForm1.Winsock1.Close
'            MDIForm1.Winsock1.RemoteHost = IPSERVER
'            MDIForm1.Winsock1.Connect
'            MDIForm1.Winsock1.SendData "MESSAGE" + "/" + Text1.Text + "^_^" & MDIForm1.Winsock1.LocalIP & "~!" & RichTextBox1.Text
'            WaitSecs (2)
'            MDIForm1.Winsock1.SendData "MESSAGE" + "/" + Text1.Text + "^_^" & MDIForm1.Winsock1.LocalIP & "~!" & RichTextBox1.Text
'        Else
'            MDIForm1.Winsock1.SendData "MESSAGE" + "/" + Text1.Text + "^_^" & MDIForm1.Winsock1.LocalIP & "~!" & RichTextBox1.Text
'            Unload Me
'        End If
Case 1
    Unload Me
Case 2
    Text1.Text = Empty
    RichTextBox1.Text = Empty
End Select
End Sub

Private Sub Command2_Click()
    FRMUNTUK.Show vbModal
End Sub



Private Sub Form_Load()
RichTextBox1.SelColor = &HC00000
End Sub

Private Sub RichTextBox1_Change()
Label2.Caption = 255 - Len(RichTextBox1.Text)
End Sub
