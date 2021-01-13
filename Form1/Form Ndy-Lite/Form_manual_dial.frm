VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Form_manual_dial 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5415
   ClientLeft      =   15705
   ClientTop       =   1440
   ClientWidth     =   3270
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Report"
      Height          =   435
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txt_no 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "081212319921"
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton cmd_redial 
      BackColor       =   &H0080FF80&
      Caption         =   "Redial"
      Height          =   435
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmd_angka 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear"
      Height          =   495
      Index           =   11
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmd_angka 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Del"
      Height          =   495
      Index           =   10
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmd_angka 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   495
      Index           =   9
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmd_angka 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      Height          =   495
      Index           =   8
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmd_angka 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      Height          =   495
      Index           =   7
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmd_angka 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      Height          =   495
      Index           =   6
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmd_angka 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      Height          =   495
      Index           =   5
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmd_angka 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      Height          =   495
      Index           =   4
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmd_angka 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      Height          =   495
      Index           =   3
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmd_angka 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      Height          =   495
      Index           =   2
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmd_angka 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      Height          =   495
      Index           =   1
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmd_angka 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      Height          =   495
      Index           =   0
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmd_exit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Exit"
      Height          =   435
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   2055
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   720
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1270
      _Version        =   196610
      Font3D          =   4
      MousePointer    =   16
      ForeColor       =   12582912
      PictureMaskColor=   -2147483644
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form_manual_dial.frx":0000
      AutoSize        =   1
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Default         =   -1  'True
      Height          =   720
      Index           =   0
      Left            =   720
      TabIndex        =   17
      Top             =   1080
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1270
      _Version        =   196610
      Font3D          =   4
      MousePointer    =   16
      ForeColor       =   12582912
      PictureMaskColor=   -2147483644
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form_manual_dial.frx":051C
      AutoSize        =   1
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   4920
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   0
      X2              =   3240
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   0
      X2              =   3240
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   3240
      X2              =   3240
      Y1              =   0
      Y2              =   5400
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   5400
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   ". . : Ndy - Lite : . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H003F9E0C&
      Caption         =   "Hang Up"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H003F9E0C&
      Caption         =   "Call"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   900
   End
End
Attribute VB_Name = "Form_manual_dial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim F_Call As Boolean


Private Declare Function GetSystemMenu Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal bRevert As Long) As Long

Private Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, _
     ByVal nPosition As Long, _
     ByVal wFlags As Long) As Long
     
Private Const MF_BYPOSITION = &H400&

Public Function DisableCloseButton(frm As Form) As Boolean
'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu
    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long

    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)

    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)

   'Remove seperator bar
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    
    'Return True if both calls were successful
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)
End Function

Private Sub cmd_angka_Click(Index As Integer)
    Dim nomor As String
    
    Select Case Index
        Case 0
            nomor = txt_no.Text
            txt_no.Text = nomor & "1"
        Case 1
            nomor = txt_no.Text
            txt_no.Text = nomor & "2"
        Case 2
            nomor = txt_no.Text
            txt_no.Text = nomor & "3"
        Case 3
            nomor = txt_no.Text
            txt_no.Text = nomor & "4"
        Case 4
            nomor = txt_no.Text
            txt_no.Text = nomor & "5"
        Case 5
            nomor = txt_no.Text
            txt_no.Text = nomor & "6"
        Case 6
            nomor = txt_no.Text
            txt_no.Text = nomor & "7"
        Case 7
            nomor = txt_no.Text
            txt_no.Text = nomor & "8"
        Case 8
            nomor = txt_no.Text
            txt_no.Text = nomor & "9"
        Case 9
            nomor = txt_no.Text
            txt_no.Text = nomor & "0"
        Case 10
            If Len(txt_no.Text) = 0 Then Exit Sub
            txt_no.Text = Mid(txt_no.Text, 1, Len(txt_no.Text) - 1)
        Case 11
            txt_no.Text = ""
    End Select
End Sub

Private Sub cmd_exit_Click()
    If F_Call = True Then
        MsgBox "Hang-Up Terlebih Dahulu Sebelum Exit !", vbOKOnly + vbInformation, "Info"
        Exit Sub
    Else
        'Label2.Caption = 0
        Unload Me
    End If
End Sub

Private Sub cmd_redial_Click()
    Dim sQuery As String
    Dim rs As ADODB.Recordset
    
    sQuery = "SELECT phone_number FROM mandiri.tbl_manual_dial "
    sQuery = sQuery + "WHERE tgl_call = (SELECT max(tgl_call) FROM mandiri.tbl_manual_dial WHERE agent = '" & MDIForm1.TxtUsername.Text & "')"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        txt_no.Text = rs!phone_number
    End If
End Sub

Private Sub Command1_Click()
    'formhistoryhp.Show vbModal
End Sub

Private Sub Form_Activate()
    MakeTopMost hwnd
    'Label2.Caption = 1
End Sub

Private Sub Form_Load()
    DisableCloseButton Me
    txt_no.Text = ""
    F_Call = False
    SSCommand1(1).Enabled = False
    'Label2.Caption = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKey1 Or KeyCode = vbKeyNumpad1 Then
        txt_no.Text = txt_no.Text & "1"
    ElseIf KeyCode = vbKey2 Or KeyCode = vbKeyNumpad2 Then
        txt_no.Text = txt_no.Text & "2"
    ElseIf KeyCode = vbKey3 Or KeyCode = vbKeyNumpad3 Then
        txt_no.Text = txt_no.Text & "3"
    ElseIf KeyCode = vbKey4 Or KeyCode = vbKeyNumpad4 Then
        txt_no.Text = txt_no.Text & "4"
    ElseIf KeyCode = vbKey5 Or KeyCode = vbKeyNumpad5 Then
        txt_no.Text = txt_no.Text & "5"
    ElseIf KeyCode = vbKey6 Or KeyCode = vbKeyNumpad6 Then
        txt_no.Text = txt_no.Text & "6"
    ElseIf KeyCode = vbKey7 Or KeyCode = vbKeyNumpad7 Then
        txt_no.Text = txt_no.Text & "7"
    ElseIf KeyCode = vbKey8 Or KeyCode = vbKeyNumpad8 Then
        txt_no.Text = txt_no.Text & "8"
    ElseIf KeyCode = vbKey9 Or KeyCode = vbKeyNumpad9 Then
        txt_no.Text = txt_no.Text & "9"
    ElseIf KeyCode = vbKey0 Or KeyCode = vbKeyNumpad0 Then
        txt_no.Text = txt_no.Text & "0"
    ElseIf KeyCode = vbKeyBack Then
        If Len(txt_no.Text) = 0 Then Exit Sub
        txt_no.Text = Mid(txt_no.Text, 1, Len(txt_no.Text) - 1)
    End If
    
End Sub

Private Sub Label3_Click()
    Form_manual_dial.WindowState = vbMinimized
End Sub

Private Sub SSCommand1_Click(Index As Integer)
    Select Case Index
        Case 0
            If txt_no.Text = "" Then
                MsgBox "Input Nomor Telefon Terlebih Dahulu !", vbOKOnly + vbInformation, "Info"
                Exit Sub
            Else
                MDIForm1.Enabled = False
                Call Dial
            End If
        Case 1
            MDIForm1.Enabled = True
            Call Hangup
    End Select
End Sub

Private Sub Dial()
    Dim nomordial, iQuery As String
    Dim rs As ADODB.Recordset

    nomordial = Replace(txt_no.Text, " ", "")
    
    If Obelisk = False Then
        'UNTUK ORANGE CLIENT
        MDIForm1.ActionCTI ("DIAL|020892" & GetNumber(CStr(Replace(nomordial, " ", ""))) & "-" & MDIForm1.TxtUsername.Text)
    Else
        'UNTUK OBELISK
        MDIForm1.ActionCTI ("DIAL|" & GetNumber(CStr(Replace(nomordial, " ", ""))) & "-" & MDIForm1.TxtUsername.Text)
    End If
    
    MDIForm1.ActionCTI ("DIAL|020892" & GetNumber(CStr(Replace(nomordial, " ", ""))) & "-" & MDIForm1.TxtUsername.Text)
    
    iQuery = "select distinct phone_number from ("
    iQuery = iQuery + " select ahomeno as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  homeno as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  ahomeno2 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  homeno2 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  mobileno as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  mobileno2 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  aofficeno as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  officeno as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  aofficeno2 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  officeno2 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  homenoadd1 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  ahomenoadd1 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  homenoadd2 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  ahomenoadd2 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  officenoadd1 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  aofficenoadd1 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  officenoadd2 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  aofficenoadd2 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  mobilenoadd1 as phone_number from mandiri.mgm union all"
    iQuery = iQuery + " select  mobilenoadd2 as phone_number from mandiri.mgm) a where phone_number is not null and phone_number = '" & nomordial & "'"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open iQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If rs.RecordCount = 0 Then
        iQuery = "INSERT INTO mandiri.tbl_manual_dial(agent, phone_number, tgl_call, exists_number)"
        iQuery = iQuery + " VALUES ('" & MDIForm1.TxtUsername.Text & "', '" & nomordial & "', '" & waktu_server_sekarang & "', 0 )"
    Else
        iQuery = "INSERT INTO mandiri.tbl_manual_dial(agent, phone_number, tgl_call, exists_number)"
        iQuery = iQuery + " VALUES ('" & MDIForm1.TxtUsername.Text & "', '" & nomordial & "', '" & waktu_server_sekarang & "', 1 )"
    End If
    M_OBJCONN.Execute iQuery
    
    F_Call = True
    SSCommand1(0).Enabled = False
    SSCommand1(1).Enabled = True
    txt_no.Text = "    D i a l i n g . . ."
End Sub

Private Sub Hangup()
    DoEvents
    MDIForm1.ActionCTI ("HANGUP")
    txt_no.Text = ""
    F_Call = False
    SSCommand1(0).Enabled = True
    SSCommand1(1).Enabled = False
End Sub

