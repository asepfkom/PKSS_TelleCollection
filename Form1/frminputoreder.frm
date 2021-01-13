VERSION 5.00
Begin VB.Form frminputoreder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FORM DATA OFFER"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtkey 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5340
      TabIndex        =   14
      Top             =   480
      Width           =   3285
   End
   Begin VB.ComboBox cbooperator 
      Height          =   315
      ItemData        =   "frminputoreder.frx":0000
      Left            =   3090
      List            =   "frminputoreder.frx":000A
      TabIndex        =   12
      Top             =   1230
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ok"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5820
      UseMaskColor    =   -1  'True
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Batal"
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
      Left            =   1050
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5820
      UseMaskColor    =   -1  'True
      Width           =   810
   End
   Begin VB.TextBox txtremarks 
      Height          =   3075
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1950
      Width           =   11175
   End
   Begin VB.ComboBox cbopersentase 
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   1260
      Width           =   1125
   End
   Begin VB.ComboBox cbomap 
      Height          =   315
      Left            =   1110
      TabIndex        =   4
      Top             =   870
      Width           =   3225
   End
   Begin VB.TextBox txtketerangan 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1110
      TabIndex        =   2
      Top             =   510
      Width           =   3285
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Key Rumus"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   510
      Width           =   915
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Operator"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1290
      Width           =   1005
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Persentase"
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Field Map"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   930
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   540
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   5
      Left            =   60
      Picture         =   "frminputoreder.frx":0014
      Stretch         =   -1  'True
      Top             =   0
      Width           =   420
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Input Data Offer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   540
      TabIndex        =   0
      Top             =   30
      Width           =   3405
   End
   Begin VB.Image Image2 
      Height          =   435
      Index           =   8
      Left            =   0
      Picture         =   "frminputoreder.frx":0B1E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20700
   End
End
Attribute VB_Name = "frminputoreder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ok As Boolean
Private Sub Command1_Click(Index As Integer)
Dim VSAVE As Boolean
VSAVE = True
Select Case Index
    Case 0
        VSAVE = VSAVE And txtkey.Text <> Empty
        VSAVE = VSAVE And txtketerangan.Text <> Empty
        VSAVE = VSAVE And cbomap.Text <> Empty
        VSAVE = VSAVE And cbopersentase.Text <> Empty
        VSAVE = VSAVE And cbooperator.Text <> Empty
        If VSAVE Then
            ok = True
            Me.Hide
           frmheaderoffeer.ListView1.SetFocus
        Else
            MsgBox "Data Yang Anda Masukan Tidak Lengkap", vbInformation, "Informasi"
        End If
    Case 1
        ok = False
        Unload Me
        frmheaderoffeer.ListView1.SetFocus
End Select
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim STRSQL As String
For i = 1 To 100
  cbopersentase.AddItem i
  
Next i

STRSQL = "SELECT column_name From information_schema.Columns WHERE table_name='mgm' ORDER BY ordinal_position"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not rs.EOF
cbomap.AddItem rs!column_name
rs.MoveNext
Wend

End Sub

Private Sub Text2_Change()

End Sub
