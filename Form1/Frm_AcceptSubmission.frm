VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Frm_AcceptSubmission 
   Caption         =   "Accept-Decline Submission"
   ClientHeight    =   930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4380
   Icon            =   "Frm_AcceptSubmission.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   930
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand SSCommand1 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   0
      Top             =   465
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   661
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
      PictureMaskColor=   -2147483644
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Frm_AcceptSubmission.frx":000C
      Caption         =   "&Exit"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   1
      Left            =   1185
      TabIndex        =   1
      Top             =   465
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   661
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
      PictureMaskColor=   -2147483644
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Frm_AcceptSubmission.frx":0166
      Caption         =   "&Accept"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   465
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   661
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
      PictureMaskColor=   -2147483644
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Frm_AcceptSubmission.frx":0488
      Caption         =   "&Decline"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   3
      Left            =   2265
      TabIndex        =   3
      Top             =   465
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   661
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
      PictureMaskColor=   -2147483644
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Frm_AcceptSubmission.frx":07AA
      Caption         =   "&Cancel"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin VB.Label Label2 
      Height          =   315
      Left            =   165
      TabIndex        =   5
      Top             =   75
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2790
      TabIndex        =   4
      Top             =   75
      Width           =   1455
   End
End
Attribute VB_Name = "Frm_AcceptSubmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sts As Boolean

Private Sub Form_Load()
    SSCommand1(3).Enabled = False
End Sub

Private Sub SSCommand1_Click(Index As Integer)
    Select Case Index
        Case 0
            SSCommand1(0).Enabled = False
            SSCommand1(1).Enabled = False
            SSCommand1(2).Enabled = False
            SSCommand1(3).Enabled = True
            Call decline
            SSCommand1(0).Enabled = True
            SSCommand1(1).Enabled = True
            SSCommand1(2).Enabled = True
            SSCommand1(3).Enabled = False
        Case 1
            SSCommand1(0).Enabled = False
            SSCommand1(1).Enabled = False
            SSCommand1(2).Enabled = False
            SSCommand1(3).Enabled = True
            Call Accept
            SSCommand1(0).Enabled = True
            SSCommand1(1).Enabled = True
            SSCommand1(2).Enabled = True
            SSCommand1(3).Enabled = False
        Case 2
            Unload Me
        Case 3
            sts = True
            SSCommand1(0).Enabled = True
            SSCommand1(1).Enabled = True
            SSCommand1(2).Enabled = True
            SSCommand1(3).Enabled = False
    End Select
End Sub

Private Sub Accept()
Dim direktorifile As String
Dim count As Double
Dim cif As String
Dim nobilyet As String
Dim cabang As String
Dim amount As String
Dim cmdsql As String
On Error GoTo IsiPortoFolioErr
'porto folio deposito
direktorifile = App.Path & "\CheckSubmission\Accept.txt"
        Open direktorifile For Input As #1
        M_OBJCONN.BeginTrans
        count = 0
        Do Until EOF(1)
            count = count + 1
            Line Input #1, lineoftext$
            DataText = lineoftext$
            cif = Trim(Left(DataText, 25))
            Label1.Caption = count
            DoEvents
            If sts = True Then
                M_OBJCONN.RollbackTrans
                Exit Do
            End If
            cmdsql = " Update Cc_Custtbl set F_CEK= 'Accept' WHERE CUSTID ='" + cif + "'"
            M_OBJCONN.Execute cmdsql
        Loop
        M_OBJCONN.CommitTrans
        Close
        MsgBox "Done"
    Exit Sub
IsiPortoFolioErr:
    MsgBox Err.Description
    If Err.Number = 53 Then
    Else
        M_OBJCONN.RollbackTrans
    End If
  '  Resume
End Sub

Private Sub decline()
Dim direktorifile As String
Dim count As Double
Dim cif As String
Dim nobilyet As String
Dim cabang As String
Dim amount As String
On Error GoTo IsiPortoFolioErr
'porto folio deposito
direktorifile = App.Path & "\CheckSubmission\RETURN.txt"
        Open direktorifile For Input As #1
        M_OBJCONN.BeginTrans
        isi = ""
        count = 0
        Do Until EOF(1)
            count = count + 1
            Line Input #1, lineoftext$
            DataText = lineoftext$
            cif = Trim(Left(DataText, 25))
            Label1.Caption = count
            DoEvents
            If sts = True Then
                M_OBJCONN.RollbackTrans
                Exit Do
            End If
            cmdsql = " Update Cc_Custtbl set F_CEK= 'RETURN', KETHSLKERJA ='', DAY_DOB='',MONTH_DOB='',YEAR_DOB='',NOAREATELP='',NOTELP='',CARDED=0,BT=0,SEGMENTED=0 WHERE CUSTID ='" + cif + "'"
            M_OBJCONN.Execute cmdsql
        Loop
        M_OBJCONN.CommitTrans
        Close
        MsgBox "Done"
        Exit Sub
IsiPortoFolioErr:
    MsgBox Err.Description
    If Err.Number = 53 Then
    Else
        M_OBJCONN.RollbackTrans
    End If
End Sub
