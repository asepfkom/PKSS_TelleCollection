VERSION 5.00
Begin VB.Form frmeditsms 
   BorderStyle     =   0  'None
   Caption         =   "Edit SMS"
   ClientHeight    =   4575
   ClientLeft      =   7395
   ClientTop       =   4035
   ClientWidth     =   9480
   LinkTopic       =   "Form2"
   ScaleHeight     =   4575
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1230
      TabIndex        =   3
      Top             =   4020
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   2430
      TabIndex        =   1
      Top             =   4020
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   3105
      Left            =   1200
      MaxLength       =   160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   840
      Width           =   7275
   End
   Begin VB.Label Label4 
      Caption         =   "Custid :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   4545
      Left            =   0
      Top             =   0
      Width           =   9480
   End
   Begin VB.Label Label2 
      Caption         =   "Text :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Mobile No :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmeditsms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

                CMDSQL = "update request_sms set pesan='" & Text1.Text & "' where custid='" & Text3.Text & "' and notelp='" & Text5.Text & "'"
                M_OBJCONN.Execute CMDSQL
                Unload Me
                Unload Frm_verify
                Load Frm_verify
                Frm_verify.Show vbModal
                

End Sub

Private Sub Command2_Click()
Unload Me
End Sub



