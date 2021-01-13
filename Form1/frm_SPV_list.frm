VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_SPV_LIST 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   5475
      Left            =   -15
      TabIndex        =   1
      Top             =   -30
      Width           =   9015
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   7995
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   8010
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   570
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Del"
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
         Left            =   8010
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "E&xit"
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
         Index           =   3
         Left            =   8010
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1365
         Width           =   885
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5295
         Left            =   30
         TabIndex        =   0
         Top             =   135
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   9340
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
Attribute VB_Name = "FRM_SPV_LIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Kode Team Leader", 10 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Nama Team Leader", 50 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Unit", 15 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Team", 15 * TXT
End Sub

Private Sub Form_Load()
    Dim M_OBJRS As ADODB.Recordset
    Dim M_DATA As New CLSSPV_AGENT
    Dim listitem As listitem
    Call header
    Set M_OBJRS = M_DATA.QUERY(M_OBJCONN, "")
    While Not M_OBJRS.EOF
         Set listitem = ListView1.ListItems.ADD(, , M_OBJRS("SPVCODE"))
             listitem.SubItems(1) = IIf(IsNull(M_OBJRS("SPVNAME")), "", M_OBJRS("SPVNAME"))
             listitem.SubItems(2) = IIf(IsNull(M_OBJRS("UNIT")), "", M_OBJRS("UNIT"))
             listitem.SubItems(3) = IIf(IsNull(M_OBJRS("TEAM")), "", M_OBJRS("TEAM"))
         M_OBJRS.MoveNext
    Wend
    M_OBJRS.Close
    Set M_OBJRS = Nothing
End Sub


Private Sub Command1_Click(Index As Integer)
Dim m_msgbox As Variant
Dim listitem As listitem
Dim L_JABATAN As String
Dim M_DATA As New CLSSPV_AGENT
Select Case Index
    Case 0
            With frm_SPV
                .Caption = "Tambah Data Team Leader"
                .Show vbModal
                If .ok Then
                    If .Option1(2).Value Then
                        L_JABATAN = UCase(.Option1(2).Caption)
                    Else
                        L_JABATAN = UCase(.Option1(3).Caption)
                    End If
                    M_DATA.ADD M_OBJCONN, .Text1.Text, .Text2.Text, .Text3.Text, .Combo1.Text, .TDBNumber1.Text, L_JABATAN
                    On Error GoTo add_error
                    If M_DATA.ADD_OK Then
                        Set listitem = ListView1.ListItems.ADD(, , .Text1.Text)
                            listitem.SubItems(1) = .Text2.Text
                            listitem.SubItems(2) = .Combo1.Text
                            listitem.SubItems(3) = .Text3.Text
                     On Error GoTo 0
                    End If
                End If
                Unload frm_SPV
            End With
        Exit Sub
    Case 1
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
            With frm_SPV
                .Caption = "Ubah Data Team Leader"
                .Text1.Text = ListView1.SelectedItem.Text
                .Text2.Text = ListView1.SelectedItem.SubItems(1)
                .Combo1.Text = ListView1.SelectedItem.SubItems(2)
                .Text3.Text = ListView1.SelectedItem.SubItems(3)
                .Text1.Locked = True
                .Text1.TabStop = False
                .Text1.Appearance = 0
                .Text1.BackColor = &H8000000F
                .Show vbModal
                If .ok Then
                    M_DATA.UPDATE M_OBJCONN, .Text1.Text, .Text2.Text, .Text3.Text, .Combo1.Text, .TDBNumber1.Text
                    On Error GoTo add_error
                    If M_DATA.ADD_OK Then
                        ListView1.SelectedItem.SubItems(1) = .Text2.Text
                        ListView1.SelectedItem.SubItems(2) = .Combo1.Text
                        ListView1.SelectedItem.SubItems(3) = .Text3.Text
                      On Error GoTo 0
                    End If
                End If
                Unload frm_SPV
            End With
        Exit Sub
    Case 2
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
        m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
        If m_msgbox = 1 Then
            M_DATA.DELETE M_OBJCONN, ListView1.SelectedItem.Text
            M_DATA.DELETE_AGENT M_OBJCONN, ListView1.SelectedItem.Text
            If M_DATA.ADD_OK Then
                ListView1.ListItems.Remove ListView1.SelectedItem.Index
            End If
        End If
        Exit Sub
    Case 3
        Unload Me
        Exit Sub
End Select
add_error:
End Sub
Private Sub ListView1_DblClick()
    Call Command1_Click(1)
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click(1)
    End If
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub
