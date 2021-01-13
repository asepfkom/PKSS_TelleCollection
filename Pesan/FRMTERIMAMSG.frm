VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FRMTERIMAMSG 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   75
      TabIndex        =   8
      Top             =   5295
      Width           =   795
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
      Left            =   660
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   795
      Width           =   3525
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   4020
      Left            =   0
      TabIndex        =   5
      Top             =   1125
      Width           =   4500
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3840
         Left            =   30
         TabIndex        =   4
         Top             =   135
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   6773
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"FRMTERIMAMSG.frx":0000
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
      Height          =   405
      Index           =   2
      Left            =   3645
      TabIndex        =   2
      Top             =   5295
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Teruskan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   2655
      TabIndex        =   1
      Top             =   5295
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Balas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   1860
      TabIndex        =   0
      Top             =   5295
      Width           =   780
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   1164
      _Version        =   196610
      Font3D          =   5
      ForeColor       =   0
      BackColor       =   -2147483644
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
      Picture         =   "FRMTERIMAMSG.frx":007D
      Caption         =   "Message"
      BevelWidth      =   2
      BorderWidth     =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Dari :"
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
      Left            =   195
      TabIndex        =   6
      Top             =   855
      Width           =   465
   End
End
Attribute VB_Name = "FRMTERIMAMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click(Index As Integer)
Dim cmdsql As String

Select Case Index
    Case 0
        'FRMSENDMSG.Text1.Text = Text1.Text + ";"
        'FRMSENDMSG.Command2.Enabled = False
        Unload Me
        'FRMSENDMSG.Show vbModal
    Case 1
        'FRMSENDMSG.RichTextBox1.Text = RichTextBox1.Text
        Unload Me
        'FRMSENDMSG.Show vbModal
    Case 2
        Unload Me
    Case 3
        'cmdsql = "Insert into         "
End Select
End Sub

Private Sub Form_Load()
Dim m_objrs As New ADODB.Recordset
Dim PENERIMA As String

PENERIMA = BUKA_FILE_KONEKSI("D:\MSG.TXT")
m_objrs.CursorLocation = adUseClient
m_objrs.Open "select * from MSGTBL, usertbl  where MSGTBL.SENDER = USERTBL.USERID AND RECIPIENT ='" + PENERIMA + "' AND STS =0", m_objconn, adOpenDynamic, adLockOptimistic, adCmdText
On Error Resume Next
While Not m_objrs.EOF
    RichTextBox1.SelColor = &HC00000
    Text1.Text = IIf(IsNull(m_objrs!SENDER), "", m_objrs!SENDER)
    RichTextBox1.Text = RichTextBox1.Text + "Dari :" + IIf(IsNull(m_objrs!SENDER), "", m_objrs!SENDER) + " - " + IIf(IsNull(m_objrs!agent), "", m_objrs!agent) + vbCrLf
    RichTextBox1.Text = RichTextBox1.Text + "Kepada :" + IIf(IsNull(m_objrs!RECIPIENT), "", m_objrs!RECIPIENT) + vbCrLf
    RichTextBox1.Text = RichTextBox1.Text + "Tanggal :" + IIf(IsNull(m_objrs!DateTime), "", m_objrs!DateTime) + vbCrLf
    RichTextBox1.Text = RichTextBox1.Text + "Isi Pesan :" + vbCrLf
    RichTextBox1.Text = RichTextBox1.Text + IIf(IsNull(m_objrs!MSG), "", m_objrs!MSG)
    RichTextBox1.Text = RichTextBox1.Text + " " + vbCrLf
    RichTextBox1.Text = RichTextBox1.Text & vbCrLf
    m_objrs.MoveNext
Wend
m_objconn.Execute "UPDATE MSGTBL SET STS =1 WHERE RECIPIENT ='" + PENERIMA + "'"
Set m_objrs = Nothing
End Sub
