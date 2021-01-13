VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Form_filter_hide 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter Hide System"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   5505
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   9710
      _Version        =   196610
      BackColor       =   14737632
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "HIDE COLUMN LIST ACCOUNT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2925
         Left            =   180
         TabIndex        =   1
         Top             =   210
         Width           =   7605
         Begin VB.CommandButton cmdbatal 
            BackColor       =   &H00C0C0C0&
            Cancel          =   -1  'True
            Caption         =   "&BATAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6495
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2400
            UseMaskColor    =   -1  'True
            Width           =   930
         End
         Begin VB.CommandButton cmdsimpan 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&SIMPAN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2400
            UseMaskColor    =   -1  'True
            Width           =   945
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">"
            Height          =   375
            Index           =   0
            Left            =   3480
            TabIndex        =   5
            Top             =   450
            Width           =   675
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<"
            Height          =   375
            Index           =   1
            Left            =   3480
            TabIndex        =   4
            Top             =   840
            Width           =   675
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">>"
            Height          =   375
            Index           =   2
            Left            =   3480
            TabIndex        =   3
            Top             =   1230
            Width           =   675
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<<"
            Height          =   375
            Index           =   3
            Left            =   3480
            TabIndex        =   2
            Top             =   1620
            Width           =   675
         End
         Begin MSComctlLib.ListView lv1 
            Height          =   1935
            Left            =   120
            TabIndex        =   6
            Top             =   330
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   3413
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
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
         Begin MSComctlLib.ListView lv2 
            Height          =   1935
            Left            =   4230
            TabIndex        =   7
            Top             =   330
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
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
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Mark Up"
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   720
            Width           =   1185
         End
      End
   End
End
Attribute VB_Name = "Form_filter_hide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdSimpan_Click()
    On Error GoTo ke
    
    If MsgBox("Apakah anda yakin merubah config menu", vbQuestion + vbYesNo, "Question") = vbYes Then
        If lv2.ListItems.Count <> 0 Then
            'M_OBJCONN.BeginTrans
            
            For i = 1 To lv2.ListItems.Count
                sStrsql = ""
                sStrsql = "update tblheader_hide set tblheader_hide_status = 1, tblheader_hide_kduser='" + MDIForm1.TxtUsername.Text + "' "
                sStrsql = sStrsql + "where tblheader_hide_key_menu ='" + lv2.ListItems(i).Text + "'"
                M_OBJCONN.Execute (sStrsql)
            Next
            
            For i = 1 To lv1.ListItems.Count
                sStrsql = ""
                sStrsql = "update tblheader_hide set tblheader_hide_status = 0, tblheader_hide_kduser='" + MDIForm1.TxtUsername.Text + "' "
                sStrsql = sStrsql + "where tblheader_hide_key_menu ='" + lv1.ListItems(i).Text + "'"
                M_OBJCONN.Execute (sStrsql)
            Next
            
            'M_OBJCONN.CommitTrans
            lv2.ListItems.CLEAR
            lv1.ListItems.CLEAR
            MsgBox "Has Been reload menu", vbInformation + vbOKOnly, App.Title
        
            Tampil_Data_lv1
            Tampil_Data_lv2
        
        End If
        
    End If
    Exit Sub
ke:
    M_OBJCONN.RollbackTrans
End Sub

Private Sub Form_Load()
    Call create_header_menu
    Call Tampil_Data_lv1
    Call Tampil_Data_lv2
End Sub

Public Sub create_header_menu()
    Dim list As ListItems

    With lv1
        .ColumnHeaders.ADD , , "                      Keterangan", 100 * TXT
    End With

    With lv2
        .ColumnHeaders.ADD , , "                      Keterangan", 100 * TXT
    End With
End Sub

Private Sub Tampil_Data_lv1()
    Dim M_objrs As ADODB.Recordset
    Dim strsql As String
    Dim list As ListItem
    
    Set M_objrs = New ADODB.Recordset
    
    M_objrs.CursorLocation = adUseClient
    strsql = "select tblheader_hide_key_menu from tblheader_hide where tblheader_hide_status=0 order by tblheader_hide_index asc"
    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    lv1.ListItems.CLEAR
    While Not M_objrs.EOF
        Set list = lv1.ListItems.ADD(, , IIf(IsNull(M_objrs!tblheader_hide_key_menu), "", M_objrs!tblheader_hide_key_menu))
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub Tampil_Data_lv2()
    Dim M_objrs As ADODB.Recordset
    Dim strsql As String
    Dim list As ListItem
    
    Set M_objrs = New ADODB.Recordset
    
    M_objrs.CursorLocation = adUseClient
    strsql = "select tblheader_hide_key_menu from tblheader_hide where tblheader_hide_status=1 order by tblheader_hide_id asc"
    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    lv2.ListItems.CLEAR
    
    While Not M_objrs.EOF
        Set list = lv2.ListItems.ADD(, , IIf(IsNull(M_objrs!tblheader_hide_key_menu), "", M_objrs!tblheader_hide_key_menu))
        M_objrs.MoveNext
    Wend

    Set M_objrs = Nothing
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim lList As ListItem
    Dim n As Integer
    Select Case Index
    Case 1
        If lv2.ListItems.Count <> 0 Then
            Set lList = lv1.ListItems.ADD(, , lv2.SelectedItem.Text)
            lv2.ListItems.Remove lv2.SelectedItem.Index
        End If
    Case 3
        n = lv2.ListItems.Count
        For i = 1 To lv2.ListItems.Count
           Set lList = lv1.ListItems.ADD(, , lv2.ListItems(n).Text)
            lv2.ListItems.Remove n
            n = n - 1
        Next
    Case 0
        If lv1.ListItems.Count <> 0 Then
            Set lList = lv2.ListItems.ADD(, , lv1.SelectedItem.Text)
            lv1.ListItems.Remove lv1.SelectedItem.Index
        End If
    Case 2
        n = lv1.ListItems.Count
        For i = 1 To lv1.ListItems.Count
            Set lList = lv2.ListItems.ADD(, , lv1.ListItems(n).Text)
            
            lv1.ListItems.Remove n
            n = n - 1
        Next
    End Select
End Sub

