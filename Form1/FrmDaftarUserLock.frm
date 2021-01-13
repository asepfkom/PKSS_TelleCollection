VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDaftarUserLock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daftar User Lock"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3420
      TabIndex        =   2
      Top             =   5340
      Width           =   1275
   End
   Begin VB.CommandButton CMdTambah 
      Caption         =   "&Add"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   5340
      Width           =   1275
   End
   Begin MSComctlLib.ListView LvUser 
      Height          =   5055
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmDaftarUserLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Call headerUser
    Call IsiUser
End Sub

Private Sub headerUser()
    LvUser.ColumnHeaders.ADD , , "User", 1000
    LvUser.ColumnHeaders.ADD , , "Name", 2000
    LvUser.ColumnHeaders.ADD , , "SPV Code", 1000
End Sub

Private Sub IsiUser()
    Dim listitem As listitem
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        CMDSQL = "select * from usertbl where usertype='1' and spvcode='"
        CMDSQL = CMDSQL + Trim(Replace(MDIForm1.TxtUsername.Text, "TL", "SPV")) + "'"
        CMDSQL = CMDSQL + " order by spvcode,userid asc"
    End If
    If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Or _
        UCase(MDIForm1.txtlevel.Text) = "ADMIN" Or _
        UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Then
        
        CMDSQL = "Select * from usertbl where usertype='1'"
        CMDSQL = CMDSQL + " order by spvcode,userid asc"
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvUser.ListItems.CLEAR
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set listitem = LvUser.ListItems.ADD(, , Trim(M_objrs("userid")))
                listitem.SubItems(1) = Trim(M_objrs("agent"))
                listitem.SubItems(2) = Trim(M_objrs("spvcode"))
            M_objrs.MoveNext
        Wend
    End If
    
    Set M_objrs = Nothing
    
End Sub
