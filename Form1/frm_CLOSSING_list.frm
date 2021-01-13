VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_CLOSSING_LIST 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   5505
      Left            =   -15
      TabIndex        =   1
      Top             =   -15
      Width           =   9660
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
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
         Left            =   8535
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1425
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
         Left            =   8535
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1020
         Width           =   885
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
         Left            =   8535
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   630
         Width           =   885
      End
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
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5325
         Left            =   30
         TabIndex        =   0
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
Attribute VB_Name = "FRM_CLOSSING_LIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Kode", 10 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Keterangan", 30 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Untuk DataBase", 15 * TXT
End Sub

Private Sub Form_Load()
    Dim M_OBJRS As ADODB.Recordset
    Dim M_DATA As New CLSSPV_AGENT
    Dim listitem As listitem
    Call header
    
    Set M_OBJRS = M_DATA.QUERY_CLOSSING(M_OBJCONN, "")
    While Not M_OBJRS.EOF
         Set listitem = ListView1.ListItems.ADD(, , M_OBJRS("KDCLS"))
             listitem.SubItems(1) = M_OBJRS("KETCLS")
             listitem.SubItems(2) = M_OBJRS("Jenis")
        M_OBJRS.MoveNext
    Wend
        M_OBJRS.Close
        Set M_OBJRS = Nothing
End Sub


Private Sub Command1_Click(Index As Integer)
Dim m_msgbox As Variant
Dim STATUS As String
Dim listitem As listitem
Dim M_DATA As New CLSSPV_AGENT
Select Case Index
    Case 0
            With frm_CLOSSING
                .Caption = "Tambah Data "
                .Show vbModal
                If .ok Then
                    M_DATA.ADD_CLOSSING M_OBJCONN, .Text1.Text, .Text2.Text, .Combo1.Text
                    
                    On Error GoTo add_error
                    
                    If M_DATA.ADD_OK Then
                        Set listitem = ListView1.ListItems.ADD(, , .Text1.Text)
                            listitem.SubItems(1) = .Text2.Text
                            listitem.SubItems(2) = .Combo1.Text
                    On Error GoTo 0
                    End If
                End If
                Unload frm_CLOSSING
            End With
        Exit Sub
    Case 1
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
            With frm_CLOSSING
                .Caption = "Ubah Data "
                .Text1.Text = ListView1.SelectedItem.Text
                .Text2.Text = ListView1.SelectedItem.SubItems(1)
                .Combo1.Text = ListView1.SelectedItem.SubItems(2)
                .Text1.Locked = True
                .Text1.TabStop = False
                .Text1.BackColor = &H8000000F
                .Text1.Appearance = 0
                .Show vbModal
                If .ok Then
                    M_DATA.UPDATE_CLOSSING M_OBJCONN, .Text1.Text, .Text2.Text, .Combo1.Text
                    On Error GoTo add_error
                    If M_DATA.ADD_OK Then
                        ListView1.SelectedItem.SubItems(1) = .Text2.Text
                        ListView1.SelectedItem.SubItems(2) = .Combo1.Text
                    On Error GoTo 0
                    End If
                End If
                Unload frm_CLOSSING
            End With
        Exit Sub
    Case 2
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
        m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
        If m_msgbox = 1 Then
            M_DATA.DELETE_CLOSSING M_OBJCONN, ListView1.SelectedItem.Text
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


