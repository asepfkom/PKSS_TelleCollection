VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FRM_UpdateVisit 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2640
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   5085
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3345
      TabIndex        =   4
      Top             =   2130
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   390
      Width           =   2295
   End
   Begin RichTextLib.RichTextBox TxtDetails 
      Height          =   1305
      Left            =   1020
      TabIndex        =   3
      Top             =   765
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2302
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"UpdateVisit.frx":0000
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   661
      _Version        =   196610
      Font3D          =   4
      ForeColor       =   12582912
      Caption         =   "Update Visit"
      BevelWidth      =   2
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Hasil Visit :"
      Height          =   255
      Left            =   45
      TabIndex        =   1
      Top             =   810
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "VisitNo :"
      Height          =   255
      Left            =   45
      TabIndex        =   0
      Top             =   435
      Width           =   930
   End
End
Attribute VB_Name = "FRM_UpdateVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim CMDSQL As String
'Dim M_DATA As New CLS_FRMSEARCH
'CMDSQL = "UPDATE TblVisit set DetailsV = '" + Trim(TxtDetails.Text) + "'"
'CMDSQL = CMDSQL + " WHERE id= '" + FrmCC_Colection.LstVisit.SelectedItem.SubItems(6) + "'"
'M_OBJCONN.Execute CMDSQL
'
'CMDSQL = "INSERT INTO mgm_hst(CUSTID,AGENT,HST,KODEDS,TGL,F_CEK)VALUES ('" + FrmCC_Colection.lblCustId.Caption + "','" + VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(9) + "','" + "FIELD : " & Trim(TxtDetails.Text) + "','" + FrmCC_Colection.LstVisit.SelectedItem.SubItems(7) + "','" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd hh:mm:ss") + "','" + FrmCC_Colection.LstVisit.SelectedItem.SubItems(7) + "')"
'M_OBJCONN.Execute CMDSQL
'
'MsgBox "Update Done..."
'FrmCC_Colection.LstVisit.SelectedItem.SubItems(4) = TxtDetails.Text
'
'
'Unload Me
End Sub
