VERSION 5.00
Begin VB.Form Formapptransferdata 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "APPROVAL TRANSFER DATA"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdapprove 
      Caption         =   "Approve"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblpemohon 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Akan dilakukan transfer data oleh :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Formapptransferdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApprove_Click()
    Dim QUERY As String
    
    QUERY = "update tampungtransferdata set y_n = 1"
    QUERY = QUERY + " where pengupload = '" + Label2.Caption + "' and tujapproval = '" + Label3.Caption + "'"
    M_OBJCONN.Execute QUERY
    
    MsgBox "Approved", vbOKOnly, "Approved"
    Unload Me
    Exit Sub
End Sub
