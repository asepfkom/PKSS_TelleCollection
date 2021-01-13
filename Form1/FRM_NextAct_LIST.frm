VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_NextAct_LIST 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5415
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   5505
      Left            =   -15
      TabIndex        =   5
      Top             =   -75
      Width           =   9660
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
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
         Left            =   8610
         Picture         =   "FRM_NextAct_LIST.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1110
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Cancel          =   -1  'True
         Caption         =   "&Tutup"
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
         Left            =   8610
         Picture         =   "FRM_NextAct_LIST.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2625
         Width           =   885
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
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Hapus"
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
         Index           =   2
         Left            =   8610
         Picture         =   "FRM_NextAct_LIST.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1860
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Tambah"
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
         Index           =   0
         Left            =   8610
         Picture         =   "FRM_NextAct_LIST.frx":0B5E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   900
      End
   End
End
Attribute VB_Name = "FRM_NextAct_LIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Code", 15 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Next Action Desc", 50 * TXT
End Sub

Private Sub Form_Load()
    Dim M_OBJRS As ADODB.Recordset
    Dim listitem As listitem
    Call header
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open "Select * from StsNextAct", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not M_OBJRS.EOF
         Set listitem = ListView1.ListItems.ADD(, , M_OBJRS("KdStsNextAct"))
             listitem.SubItems(1) = M_OBJRS("NmStsNextAct")
        M_OBJRS.MoveNext
    Wend
        M_OBJRS.Close
        Set M_OBJRS = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim m_msgbox As Variant
Dim listitem As listitem
Dim cmdsql As String
Select Case Index
    Case 0
           With FRM_NextAct
                .Caption = "Tambah"
                .Show vbModal
                
                If .ok Then
                    cmdsql = "Insert into StsNextAct "
                    cmdsql = cmdsql + "  (KdStsNextAct,NmStsNextAct)"
                    cmdsql = cmdsql + " values"
                    cmdsql = cmdsql + " ('" + .Text1.Text + "',"
                    cmdsql = cmdsql + " '" + .Text2.Text + "')"
                    On Error GoTo add_error
                    M_OBJCONN.Execute cmdsql
                        Set listitem = ListView1.ListItems.ADD(, , .Text1.Text)
                            listitem.SubItems(1) = .Text2.Text
                    On Error GoTo 0
                End If
                Unload FRM_NextAct
            End With
        Exit Sub
    Case 1
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
        
            With FRM_NextAct
                .Caption = "Ubah"
                .Text1.Text = ListView1.SelectedItem.Text
                .Text2.Text = ListView1.SelectedItem.SubItems(1)
                .Text1.Locked = True
                .Text1.TabStop = False
                .Text1.BackColor = &H8000000F
                .Text1.Appearance = 0
                .Show vbModal
                
                If .ok Then
                    cmdsql = "Update StsNextAct "
                    cmdsql = cmdsql + " Set NmStsNextAct= '" + .Text2.Text + "'"
                    cmdsql = cmdsql + " Where KdStsNextAct= '" + .Text1.Text + "'"
                    On Error GoTo add_error
                    M_OBJCONN.Execute cmdsql
                    ListView1.SelectedItem.SubItems(1) = .Text2.Text
                    On Error GoTo 0
                End If
                Unload FRM_NextAct
            End With
        Exit Sub
        
    Case 2
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
        m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
        If m_msgbox = 1 Then
            M_OBJCONN.Execute "Delete From StsNextAct where KdStsNextAct  ='" + ListView1.SelectedItem.Text + "'"
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
        End If
        Exit Sub
    Case 3
        Unload Me
        Exit Sub
End Select
add_error:
    MsgBox Err.Description
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
