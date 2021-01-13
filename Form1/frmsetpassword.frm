VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmsetpassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrator Module"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10290
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7470
      Left            =   30
      TabIndex        =   1
      Top             =   510
      Width           =   10260
      Begin VB.TextBox txtjml 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Height          =   285
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   7020
         Width           =   1785
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   60
         TabIndex        =   8
         Top             =   150
         Width           =   10155
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "&Search"
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
            Index           =   5
            Left            =   8190
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   180
            Width           =   1590
         End
         Begin VB.TextBox txtnama 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4650
            TabIndex        =   10
            Top             =   210
            Width           =   2820
         End
         Begin VB.TextBox txtuserid 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   9
            Top             =   210
            Width           =   2250
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Officer ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   150
            TabIndex        =   13
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Officer Name "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   3390
            TabIndex        =   12
            Top             =   240
            Width           =   1185
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000004&
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
         Height          =   450
         Index           =   3
         Left            =   8730
         Picture         =   "frmsetpassword.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2175
         Width           =   1470
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000004&
         Caption         =   "&Reset Pwd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   825
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000004&
         Caption         =   "&User Id Activation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1245
         Width           =   1440
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000004&
         Caption         =   "&Unlock User Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   4
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1680
         Width           =   1440
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6150
         Left            =   30
         TabIndex        =   6
         Top             =   810
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   10848
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
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Rows :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7635
         TabIndex        =   16
         Top             =   7065
         Width           =   1185
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   8490
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5565
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Set Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   1
      Left            =   510
      TabIndex        =   14
      Top             =   60
      Width           =   1650
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   1
      Left            =   15
      Picture         =   "frmsetpassword.frx":014A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   -2100
      Picture         =   "frmsetpassword.frx":0C54
      Stretch         =   -1  'True
      Top             =   -300
      Width           =   19980
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   570
      TabIndex        =   7
      Top             =   60
      Width           =   2745
   End
End
Attribute VB_Name = "frmsetpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
Dim strlevel As String
Dim ListItem As ListItem
Dim m_data As New CLSSPV_AGENT
Dim m_msgbox As Variant
Dim CMDSQL As String

    If MDIForm1.txtlevel = "Agent" Then
            strlevel = "1"
    ElseIf MDIForm1.txtlevel = "Supervisor" Then
             strlevel = "2"
    ElseIf MDIForm1.txtlevel = "Manager" Then
           strlevel = "3"
    End If

'On Error GoTo add_error
Select Case Index
        Case 0
            If ListView1.ListItems.Count = 0 Then
                Exit Sub
            End If
                M_OBJCONN.Execute "Update Usertbl set ACCREC ='" + Encrypt(Len(ListView1.SelectedItem.text), "PASS12345") + "' WHERE USERID = '" + ListView1.SelectedItem.text + "'"
                
                CMDSQL = "Insert Into TblLogUserAdm ( UserId, Keterangan, UserType) VALUES ( '" + MDIForm1.TxtUsername.text + "','Reset User Password','" + strlevel + "') "
                M_OBJCONN.Execute CMDSQL
                MsgBox "Password Has Been Reset ", vbInformation + vbOKOnly, "TNIS"
                ListView1.SelectedItem.SubItems(2) = "PASS12345"
        Case 1
        Case 2
            M_OBJCONN.Execute "Update Usertbl set ADMINSERVER ='" + ListView1.SelectedItem.text + "' WHERE USERID = '" + ListView1.SelectedItem.text + "'"
            MsgBox "Done...", vbInformation + vbOKOnly, "Telegrandi"
            CMDSQL = "Insert Into TblLogUserAdm ( UserId, Keterangan, UserType) VALUES ( '" + MDIForm1.TxtUsername.text + "','Activasi User','" + strlevel + "') "
            M_OBJCONN.Execute CMDSQL
    Case 3
        Unload Me
    Case 4
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
            M_OBJCONN.Execute "Update Usertbl set f_status_login = 0 WHERE USERID = '" + ListView1.SelectedItem.text + "'"
            
            CMDSQL = "Insert Into TblLogUserAdm ( UserId, Keterangan, UserType) VALUES ( '" + MDIForm1.TxtUsername.text + "','UnLock User Login','" + strlevel + "') "
            M_OBJCONN.Execute CMDSQL
            MsgBox "Done", vbInformation + vbOKOnly, "Telegrandi"
            Exit Sub
    Case 5
        
    Dim M_objrs As ADODB.Recordset
    Dim mwhere As String
    CMDSQL = "SELECT * FROM usertbl"
    mwhere = ""
    
    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        If Len(mwhere) = 0 Then
            mwhere = " where  spvcode = '" & UCase(MDIForm1.TxtUsername.text) & "'"
        Else
             mwhere = mwhere + " and spvcode = '" & UCase(MDIForm1.TxtUsername.text) & "'"
        End If
    End If
    
    If txtuserid.text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = " where  userid  like '%" + txtuserid.text + "%'"
        Else
             mwhere = mwhere + " and  userid  like '%" + txtuserid.text + "%'"
        End If
    End If
    
    If txtnama.text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = " where  agent  like '%" + txtnama.text + "%'"
        Else
             mwhere = mwhere + " and  agent  like '%" + txtnama.text + "%'"
        End If
    End If
    CMDSQL = CMDSQL + mwhere + " ORDER BY USERID"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ListView1.ListItems.CLEAR
    While Not M_objrs.EOF
         Set ListItem = ListView1.ListItems.ADD(, , M_objrs("USERID"))
             ListItem.SubItems(1) = IIf(IsNull(M_objrs("AGENT")), "", M_objrs("AGENT"))
             ListItem.SubItems(2) = Decrypt(Len(M_objrs("USERID")), IIf(IsNull(M_objrs("ACCREC")), "", M_objrs("ACCREC")))
         M_objrs.MoveNext
    Wend
    txtjml.text = M_objrs.RecordCount
End Select
Exit Sub
add_error:
MsgBox err.Description
'Resume
End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim m_data As New CLSSPV_AGENT
    Dim ListItem As ListItem
    Dim cek As Integer
    Call header
        Set M_objrs = m_data.QUERY_SET_PWDAGENT(M_OBJCONN, "")
    While Not M_objrs.EOF
         Set ListItem = ListView1.ListItems.ADD(, , M_objrs("USERID"))
             ListItem.SubItems(1) = IIf(IsNull(M_objrs("AGENT")), "", M_objrs("AGENT"))
             ListItem.SubItems(2) = Decrypt(Len(M_objrs("USERID")), IIf(IsNull(M_objrs("ACCREC")), "", M_objrs("ACCREC")))
         M_objrs.MoveNext
    Wend
    txtjml.text = M_objrs.RecordCount
    M_objrs.Close
    Set M_objrs = Nothing
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "User Name", 10 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Nama AGENT", 20 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Password", 10 * TXT
End Sub

Private Sub ListView1_Click()
'Text1.Text = ListView1.SelectedItem.SubItems(6)
End Sub

Private Sub ListView1_DblClick()
   Call Command1_Click(1)
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub


