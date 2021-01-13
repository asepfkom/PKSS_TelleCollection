VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_DATASOURCE_LIST 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   5505
      Left            =   -15
      TabIndex        =   1
      Top             =   -30
      Width           =   9660
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
         Left            =   8655
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
         Left            =   8655
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
         Left            =   8655
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
         Left            =   8640
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
Attribute VB_Name = "FRM_DATASOURCE_LIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Kode Data Source", 15 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Keterangan", 50 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Jenis Call", 11 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Program", 11 * TXT
End Sub

Private Sub Form_Load()
    Dim M_OBJRS As ADODB.Recordset
    Dim M_DATA As New CLSSPV_AGENT
    Dim listitem As listitem
    Call header
    Set M_OBJRS = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "")
    While Not M_OBJRS.EOF
         Set listitem = ListView1.ListItems.ADD(, , M_OBJRS("KODEDS"))
             listitem.SubItems(1) = IIf(IsNull(M_OBJRS("KETERANGAN")), "", M_OBJRS("KETERANGAN"))
             If M_OBJRS("STATUS") = "O" Then
                listitem.SubItems(2) = "Outbounds"
             Else
                listitem.SubItems(2) = "Inbounds"
             End If
             listitem.SubItems(3) = IIf(IsNull(M_OBJRS("KDPROGRAM")), "", M_OBJRS("KDPROGRAM"))
        M_OBJRS.MoveNext
    Wend
        M_OBJRS.Close
        Set M_OBJRS = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim m_msgbox As Variant
Dim M_DATA As New CLSSPV_AGENT
Dim listitem As listitem
Dim STATUS As String
Select Case Index
    Case 0
           With frm_DATASOURCE
                .Caption = "Tambah Data Source"
                .Option1(0).Value = True
                .Show vbModal
                If .ok Then
                If .Option1(0).Value Then
                    STATUS = "I"
                Else
                    STATUS = "O"
                End If
                    M_DATA.ADD_DATASOURCE M_OBJCONN, .Text1.Text, STATUS, .Text2.Text, .Combo1.Text
                    On Error GoTo add_error
                    If M_DATA.ADD_OK Then
                        Set listitem = ListView1.ListItems.ADD(, , .Text1.Text)
                            listitem.SubItems(1) = .Text2.Text
                            If .Option1(0).Value Then
                                listitem.SubItems(2) = "Inbounds"
                            Else
                                listitem.SubItems(2) = "OutBounds"
                            End If
                                listitem.SubItems(3) = .Combo1.Text
                    On Error GoTo 0
                    End If
                End If
                Unload frm_DATASOURCE
            End With
        Exit Sub
    Case 1
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
            With frm_DATASOURCE
                .Caption = "Ubah"
                .Text1.Text = ListView1.SelectedItem.Text
                .Text2.Text = ListView1.SelectedItem.SubItems(1)
                If UCase(ListView1.SelectedItem.SubItems(2)) = "OUTBOUNDS" Then
                    .Option1(1).Value = True
                Else
                    .Option1(0).Value = True
                End If
                .Combo1.Text = ListView1.SelectedItem.SubItems(3)
                .Text1.Locked = True
                .Text1.TabStop = False
                .Text1.BackColor = &H8000000F
                .Text1.Appearance = 0
                .Show vbModal
                If .ok Then
                    If .Option1(0).Value Then
                        STATUS = "I"
                    Else
                        STATUS = "O"
                    End If
                    M_DATA.UPDATE_DATASOURCE M_OBJCONN, .Text1.Text, STATUS, .Text2.Text, .Combo1.Text
                    On Error GoTo add_error
                    If M_DATA.ADD_OK Then
                        ListView1.SelectedItem.SubItems(1) = .Text2.Text
                        If STATUS = "O" Then
                            ListView1.SelectedItem.SubItems(2) = "OutBounds"
                        Else
                            ListView1.SelectedItem.SubItems(2) = "InBounds"
                        End If
                        ListView1.SelectedItem.SubItems(3) = .Combo1.Text
                    On Error GoTo 0
                    End If
                End If
                Unload frm_DATASOURCE
            End With
        Exit Sub
    Case 2
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
        m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
        If m_msgbox = 1 Then
'            If UCase(ListView1.SelectedItem.SubItems(2)) = "OUTBOUNDS" Then
'                MsgBox "Data Outbounds Tidak Dapat Dihapus", vbInformation + vbOKOnly, "Informasi"
'                Exit Sub
'            End If
            M_DATA.DELETE_DATASOURCE M_OBJCONN, ListView1.SelectedItem.Text
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
