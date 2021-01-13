VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRM_SET_PWD 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9675
   ControlBox      =   0   'False
   Icon            =   "FRM_SET_PWD.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   5505
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   9660
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000004&
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
         Height          =   720
         Index           =   3
         Left            =   8625
         Picture         =   "FRM_SET_PWD.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1125
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000004&
         Caption         =   "&Ubah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   1
         Left            =   8625
         Picture         =   "FRM_SET_PWD.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   390
         Width           =   885
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5325
         Left            =   30
         TabIndex        =   3
         Top             =   135
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   9393
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
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
   End
End
Attribute VB_Name = "FRM_SET_PWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Sub Command1_Click(Index As Integer)
'Dim m_data As New CLSSPV_AGENT
'Select Case Index
'        Case 1
'        If ListView1.ListItems.Count = 0 Then
'            Exit Sub
'        End If
'            With FRM_SET_PWD_0
'                .Caption = "Ubah Password"
'                .Text1.Text = ListView1.SelectedItem.Text
'                .Text2.Text = ListView1.SelectedItem.SubItems(2)
'                .Text3.Text = ListView1.SelectedItem.SubItems(3)
'                .Text1.Locked = True
'                .Text1.TabStop = False
'                .Text1.BackColor = &H8000000F
'                .Text1.Appearance = 0
'                .Show vbModal
'                If .ok Then
'                    m_data.UPDATE_Password M_OBJCONN, .Text1.Text, .Text2.Text, .Text3.Text
'                    On Error GoTo add_error
'                    If m_data.ADD_OK Then
'                        ListView1.SelectedItem.SubItems(2) = .Text2.Text
'                        ListView1.SelectedItem.SubItems(3) = .Text3.Text
'                    On Error GoTo 0
'                    End If
'                End If
'                Unload FRM_SET_PWD_0
'            End With
'        Exit Sub
'    Case 3
'        Unload Me
'        Exit Sub
'End Select
'add_error:
'End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim m_data As New CLSSPV_AGENT
    Dim listitem As listitem
    Dim LS As listitem
    Dim cek As Integer
    Call HEADER
    Set M_objrs = m_data.QUERY_SET_PWDAGENT(M_OBJCONN, " USERTYPE=1 ")
    While Not M_objrs.EOF
    Set LS = ListView1.ListItems.ADD(, , Trim(M_objrs("USERID")))
         Set listitem = ListView1.ListItems.ADD(, , M_objrs("USERID"))
             listitem.SubItems(1) = M_objrs("AGENT")
             listitem.SubItems(2) = IIf(IsNull(M_objrs("ACCREC")), "", M_objrs("ACCREC"))
             listitem.SubItems(3) = IIf(IsNull(M_objrs("AUTH")), "", M_objrs("AUTH"))
         M_objrs.MoveNext
    Wend
    M_objrs.Close
    Set M_objrs = Nothing
End Sub

Private Sub HEADER()
    ListView1.ColumnHeaders.ADD 1, , "User Name", 10 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Nama Agent", 20 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Password", 15 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Login EXT", 15 * TXT
End Sub

Private Sub ListView1_DblClick()
    'Call Command1_Click(1)
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub

