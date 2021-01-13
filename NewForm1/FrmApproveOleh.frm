VERSION 5.00
Begin VB.Form FrmApproveOleh 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Approve Oleh:"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmApproveOleh.frx":0000
      Left            =   240
      List            =   "FrmApproveOleh.frx":0002
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   900
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Pembuatan PTP dan CPA di Approve Oleh:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4275
   End
End
Attribute VB_Name = "FrmApproveOleh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
