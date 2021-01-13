VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frmsettarget 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entry Or Update Target"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   2460
      TabIndex        =   8
      Top             =   300
      Width           =   2145
   End
   Begin VB.CommandButton cmdexec 
      Caption         =   "exit"
      Height          =   495
      Index           =   1
      Left            =   2400
      TabIndex        =   7
      Top             =   3210
      Width           =   1215
   End
   Begin VB.CommandButton cmdexec 
      Caption         =   "Save"
      Height          =   495
      Index           =   0
      Left            =   1050
      TabIndex        =   6
      Top             =   3210
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2145
      Left            =   1020
      TabIndex        =   5
      Top             =   690
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   1020
      TabIndex        =   1
      Top             =   300
      Width           =   1455
   End
   Begin TDBNumber6Ctl.TDBNumber lblAmount 
      Height          =   255
      Left            =   1050
      TabIndex        =   2
      Top             =   750
      Visible         =   0   'False
      Width           =   3585
      _Version        =   65536
      _ExtentX        =   6324
      _ExtentY        =   450
      Calculator      =   "FrmSettarget.frx":0000
      Caption         =   "FrmSettarget.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmSettarget.frx":008C
      Keys            =   "FrmSettarget.frx":00AA
      Spin            =   "FrmSettarget.frx":00F4
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999999999999
      MinValue        =   -99999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   1701642245
      MinValueVT      =   3801093
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Note :"
      Height          =   405
      Left            =   210
      TabIndex        =   4
      Top             =   690
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Target :"
      Height          =   405
      Left            =   210
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SPV :"
      Height          =   225
      Left            =   300
      TabIndex        =   0
      Top             =   360
      Width           =   705
   End
End
Attribute VB_Name = "frmsettarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexec_Click(Index As Integer)
Select Case Index
    Case 0
    If UCase(Combo1(0).Text) = "ALL" Then
        strsql = "update usertbl set  ntargetspv='" + CStr(lblAmount.Value) + "' ,note ='" + Text1.Text + "'"
         strsql = strsql + " where (usertype =1 or usertype =6 )"
         M_OBJCONN.Execute (strsql)
    Else
    
         strsql = "update usertbl set  ntargetspv='" + CStr(lblAmount.Value) + "' ,note ='" + Text1.Text + "'"
         strsql = strsql + " where spvcode='" + Combo1(0).Text + "' and (usertype =1 or usertype =6 )"
         M_OBJCONN.Execute (strsql)
    End If
    
    Case 1
        Unload Me
End Select
End Sub
Private Sub Combo1_Click(Index As Integer)
Dim M_objrs As New ADODB.Recordset
Select Case Index
Case 0
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "select * from spvtbl where spvcode='" + Combo1(0).Text + "' ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
        Combo1(1).Text = M_objrs!SPVNAME
    End If
    Set M_objrs = Nothing
Case 1
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "select * from spvtbl where spvname='" + Combo1(1).Text + "' ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    If Not M_objrs.EOF Then
        Combo1(1).Text = M_objrs!SPVCODE
    End If
    Set M_objrs = Nothing
End Select
End Sub

Private Sub Form_Load()
Dim M_objrs As New ADODB.Recordset
Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "select * from SPVTBL ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    Combo1(0).AddItem "ALL"
    Combo1(1).AddItem "ALL"
    While Not M_objrs.EOF
          Combo1(0).AddItem M_objrs!SPVCODE
          Combo1(1).AddItem M_objrs!SPVNAME
       M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    
End Sub
