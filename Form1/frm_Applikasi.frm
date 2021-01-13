VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRMCUST_CC_Applikasi 
   Caption         =   "Application Form"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   Icon            =   "frm_Applikasi.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   5400
   StartUpPosition =   1  'CenterOwner
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   330
      Left            =   1425
      TabIndex        =   13
      Top             =   165
      Width           =   1440
      _Version        =   65536
      _ExtentX        =   2540
      _ExtentY        =   582
      Calendar        =   "frm_Applikasi.frx":000C
      Caption         =   "frm_Applikasi.frx":0124
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_Applikasi.frx":0190
      Keys            =   "frm_Applikasi.frx":01AE
      Spin            =   "frm_Applikasi.frx":020C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "12/05/2004"
      ValidateMode    =   0
      ValueVT         =   6815751
      Value           =   38119
      CenturyMode     =   0
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   1425
      TabIndex        =   10
      Top             =   870
      Width           =   2310
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   1425
      TabIndex        =   8
      Top             =   510
      Width           =   2310
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4470
      TabIndex        =   5
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3630
      TabIndex        =   4
      Top             =   3840
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add On"
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   1740
      Width           =   5100
      Begin VB.CommandButton Command3 
         Caption         =   "&Del"
         Height          =   330
         Index           =   1
         Left            =   4425
         TabIndex        =   7
         Top             =   615
         Width           =   570
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Add"
         Height          =   330
         Index           =   0
         Left            =   4425
         TabIndex        =   6
         Top             =   255
         Width           =   570
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1680
         Left            =   45
         TabIndex        =   3
         Top             =   180
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   2963
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   1425
      TabIndex        =   0
      Top             =   1230
      Width           =   3360
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Date :"
      Height          =   330
      Left            =   120
      TabIndex        =   12
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Agent Name :"
      Height          =   330
      Left            =   120
      TabIndex        =   11
      Top             =   930
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Agent Code :"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer Name:"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "FRMCUST_CC_Applikasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private add_new1 As Boolean

Private Sub Command1_Click()
Dim m_objrs As New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from frmApplikasi where custid ='" + FRMCUST_CC.Text1(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_objrs.CursorLocation <> 0 Then
    m_objrs.AddNew
    m_objrs!TANGGAL = CStr(Format(TDBDate1.Text, "mm/dd/yyyy"))
    m_objrs!AgentCode = Text2.Text
    m_objrs!AgentName = Text3.Text
    m_objrs!CustomerName = Text1.Text
    m_objrs!CUSTID = FRMCUST_CC.Text1(1).Text
    m_objrs.UPDATE
Else
    m_objrs!TANGGAL = CStr(Format(TDBDate1.Text, "mm/dd/yyyy"))
    m_objrs!AgentCode = Text2.Text
    m_objrs!AgentName = Text3.Text
    m_objrs!CustomerName = Text1.Text
    m_objrs!CUSTID = FRMCUST_CC.Text1(1).Text
    m_objrs.UPDATE
End If
Unload FRMCUST_CC
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click(Index As Integer)
Dim cmdsql As String
Dim M_MSGBOX As Variant
Select Case Index
    Case 0
        If FRMCUST_CC.Text1(1).Text = Empty Then
            MsgBox "Save Data Terlebih dahulu....", vbOKOnly + vbCritical, "Telegrandi"
            Exit Sub
        End If
        With FRMCUST_CC_Applikasi1
            .Show vbModal
            cmdsql = "Insert Into FrmApplikasi1 "
            cmdsql = cmdsql + " (Custid,AddOnName,DateOfBirth)"
            cmdsql = cmdsql + " Values"
            cmdsql = cmdsql + " ('" + FRMCUST_CC.Text1(1).Text + "' ,"
            cmdsql = cmdsql + " '" + .Text1.Text + "' ,"
            cmdsql = cmdsql + " '" + Format(.TDBDate1.Text, "mm/dd/yyyy") + "')"
            M_OBJCONN.Execute cmdsql
        End With
        Call detail
        Unload FRMCUST_CC_Applikasi1
    Case 1
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
        M_MSGBOX = MsgBox("Yakin Akan Dihapus?", vbYesNo + vbQuestion, "Telegrandi")
        If M_MSGBOX = vbYes Then
            M_OBJCONN.Execute "Delete From FrmApplikasi1 where Id ='" + ListView1.SelectedItem.SubItems(2) + "'"
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
        End If
End Select
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Dim LISTITEM As LISTITEM
Call header

Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from FrmApplikasi where custid ='" + FRMCUST_CC.Text1(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_objrs.RecordCount <> 0 Then
    add_new1 = False
    Text2.Text = IIf(IsNull(m_objrs!AgentCode), "", m_objrs!AgentCode)
    Text3.Text = IIf(IsNull(m_objrs!AgentName), "", m_objrs!AgentName)
    Text1.Text = IIf(IsNull(m_objrs!CustomerName), "", m_objrs!CustomerName)
    Call detail
Else
    add_new1 = True
    Text2.Text = Empty
    Text3.Text = Empty
    Text1.Text = Empty
End If

End Sub

Private Sub detail()
Dim LISTITEM As LISTITEM
ListView1.ListItems.Clear
Dim m_objrs As New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from FrmApplikasi1 where custid ='" + FRMCUST_CC.Text1(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    Set LISTITEM = ListView1.ListItems.ADD(, , m_objrs("AddOnName"))
        LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("DateOfBirth")), "", Format(m_objrs("DateOfBirth"), "dd/mm/yyyy"))
        LISTITEM.SubItems(2) = m_objrs("Id")
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Add On Name", 15 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Date Of Birth", 50 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Id", 1 * TXT
End Sub

