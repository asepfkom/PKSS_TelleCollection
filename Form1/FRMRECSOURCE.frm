VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmhasildistribusi 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6210
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   6525
   ControlBox      =   0   'False
   Icon            =   "FRMRECSOURCE.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cari"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   4005
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   825
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
      Height          =   315
      Index           =   1
      Left            =   4965
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   810
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1590
      MaxLength       =   20
      TabIndex        =   0
      Top             =   105
      Width           =   2190
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5475
      Left            =   0
      TabIndex        =   3
      Top             =   630
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   9657
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Batch :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   315
      TabIndex        =   4
      Top             =   135
      Width           =   1170
   End
End
Attribute VB_Name = "frmhasildistribusi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click(Index As Integer)
Dim m_objcari As ADODB.Recordset
Dim ListItem As ListItem
Select Case Index
    Case 0
        If Text1.Text = Empty Then
            MsgBox "Data Source Harus Diisi", vbInformation + vbOKOnly, "Informasi"
            Text1.SetFocus
            Exit Sub
        End If
        Set m_objcari = New ADODB.Recordset
        m_objcari.CursorLocation = adUseClient
        m_objcari.Open "select agent, count(custid) as jumlah from mandiri.cc_custtbl where recsource ='" + Text1.Text + "' group by agent", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        ListView1.ListItems.CLEAR
        While Not m_objcari.EOF
            Set ListItem = ListView1.ListItems.ADD(, , m_objcari("agent"))
              ListItem.SubItems(1) = m_objcari("jumlah")
              m_objcari.MoveNext
        Wend
        Set m_objcari = Nothing
    Case 1
        Unload Me
End Select
End Sub

Private Sub Form_Load()
Dim ListItem As ListItem
Call header
Text1.Text = Empty
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "User Id", 10 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Jumlah Data", 15 * TXT
End Sub

