VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FRMSENDMSG 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "..."
      Height          =   330
      Left            =   4350
      TabIndex        =   1
      Top             =   720
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
      Left            =   735
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   735
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
      Left            =   3180
      TabIndex        =   4
      Top             =   5370
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
      Left            =   4095
      TabIndex        =   5
      Top             =   5370
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
      Left            =   2265
      TabIndex        =   3
      Top             =   5370
      Width           =   855
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   1164
      _Version        =   196610
      Font3D          =   5
      ForeColor       =   0
      BackColor       =   -2147483643
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FRMSENDMSG.frx":0000
      Caption         =   "Send Message"
      BevelWidth      =   2
      BorderWidth     =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   0
      TabIndex        =   7
      Top             =   1080
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
         TextRTF         =   $"FRMSENDMSG.frx":031A
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
   Begin VB.Label Label1 
      Caption         =   "To :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   135
      TabIndex        =   6
      Top             =   795
      Width           =   585
   End
End
Attribute VB_Name = "FRMSENDMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
Dim cmdsql As String
Dim NAMA_MSG As String
Dim i As Integer
Dim NAMA_AM As String
Dim NAMA_SPVCODE As String
Dim M_OBJspv As ADODB.Recordset

On Error Resume Next
Select Case Index
Case 0
    If Text1.Text = Empty Then
        MsgBox "Penerima Message Harus Diisi...!!", vbInformation + vbOKOnly, "TeleGrandi"
        Exit Sub
    End If
    If RichTextBox1.Text = Empty Then
        MsgBox "Isi Message Tidak Boleh Kosong...!!", vbInformation + vbOKOnly, "TeleGrandi"
        Exit Sub
    End If
    
    ' INI CC NYA''' KE SUPERVISOR TIAP ANAK DAN KE AM
  '  If UCase(Trim(MDIForm1.Text2.Text)) = "AGENT" Then
        Set M_OBJspv = New ADODB.Recordset
        M_OBJspv.CursorLocation = adUseClient
        M_OBJspv.Open "Select AM,SPVCODE FROM USERTBL WHERE USERID ='" + MDIForm1.Text1.Text + "'", m_objconn, adOpenDynamic, adLockOptimistic, adCmdText
        While Not M_OBJspv.EOF
            NAMA_AM = IIf(IsNull(M_OBJspv!AM), "", M_OBJspv!AM)
            NAMA_SPVCODE = IIf(IsNull(M_OBJspv!SPVCODE), "", M_OBJspv!SPVCODE)
                cmdsql = "INSERT INTO MSGTBL "
                cmdsql = cmdsql + " ( RECIPIENT,"
                cmdsql = cmdsql + " DATETIME,"
                cmdsql = cmdsql + " SENDER,"
                cmdsql = cmdsql + " SENTFROM,"
                cmdsql = cmdsql + " MSG)"
                cmdsql = cmdsql + " VALUES"
                cmdsql = cmdsql + " ( '" + Trim(NAMA_AM) + "',"
                cmdsql = cmdsql + " '" + Format(Date, "yyyymmdd") + "',"
                cmdsql = cmdsql + " '" + MDIForm1.Text1.Text + "',"
                cmdsql = cmdsql + " '" + CStr(MDIForm1.Winsock1.LocalIP) + "',"
                cmdsql = cmdsql + " '" + RichTextBox1.Text & vbCr & "    (" & Text1.Text & ")" & "')"
                m_objconn.Execute cmdsql
                
                cmdsql = "INSERT INTO MSGTBL "
                cmdsql = cmdsql + " ( RECIPIENT,"
                cmdsql = cmdsql + " DATETIME,"
                cmdsql = cmdsql + " SENDER,"
                cmdsql = cmdsql + " SENTFROM,"
                cmdsql = cmdsql + " MSG)"
                cmdsql = cmdsql + " VALUES"
                cmdsql = cmdsql + " ( '" + Trim(NAMA_SPVCODE) + "',"
                cmdsql = cmdsql + " '" + Format(Date, "yyyymmdd") + "',"
                cmdsql = cmdsql + " '" + MDIForm1.Text1.Text + "',"
                cmdsql = cmdsql + " '" + CStr(MDIForm1.Winsock1.LocalIP) + "',"
                cmdsql = cmdsql + " '" + RichTextBox1.Text & vbCr & "    (" & Text1.Text & ")" & "')"
                'cmdsql = cmdsql + " '" + RichTextBox1.Text + "')"
                m_objconn.Execute cmdsql
            M_OBJspv.MoveNext
        Wend
        Set M_OBJspv = Nothing
   ' End If
        For i = 1 To Len(Text1.Text)
        Select Case Mid(Text1.Text, i, 1)
        Case ";"
            cmdsql = "INSERT INTO MSGTBL "
            cmdsql = cmdsql + " ( RECIPIENT,"
            cmdsql = cmdsql + " DATETIME,"
            cmdsql = cmdsql + " SENDER,"
            cmdsql = cmdsql + " SENTFROM,"
            cmdsql = cmdsql + " MSG)"
            cmdsql = cmdsql + " VALUES"
            cmdsql = cmdsql + " ( '" + Trim(Left(NAMA_MSG, 10)) + "',"
            cmdsql = cmdsql + " '" + Format(Date, "yyyymmdd") + "',"
            cmdsql = cmdsql + " '" + MDIForm1.Text1.Text + "',"
            cmdsql = cmdsql + " '" + CStr(MDIForm1.Winsock1.LocalIP) + "',"
            cmdsql = cmdsql + " '" + RichTextBox1.Text + "')"
            m_objconn.Execute cmdsql
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
