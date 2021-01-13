VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Begin VB.Form FRM_LOGIN_LOGOUT 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "FRM_LOGIN_LOGOUT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   7035
      Left            =   45
      TabIndex        =   1
      Top             =   600
      Width           =   11820
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2070
         Top             =   795
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6855
         Left            =   30
         TabIndex        =   2
         Top             =   135
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   12091
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
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
         Picture         =   "FRM_LOGIN_LOGOUT.frx":030A
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   1058
      _Version        =   196610
      Font3D          =   5
      ForeColor       =   4194368
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Status Agent"
      BevelWidth      =   2
      BorderWidth     =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Enter or Double Click -->> Untuk Lihat Layar"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   1
         Left            =   540
         TabIndex        =   4
         Top             =   225
         Width           =   3465
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Tekan ESC or Alt + F4 -->> Untuk Keluar"
         ForeColor       =   &H000000C0&
         Height          =   165
         Index           =   0
         Left            =   8640
         TabIndex        =   3
         Top             =   240
         Width           =   2910
      End
   End
End
Attribute VB_Name = "FRM_LOGIN_LOGOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim m_objrs As New ADODB.Recordset
Dim M_LOGIN As ADODB.Recordset
Dim imgX As ListImage
Dim itmX As LISTITEM
Dim CMDSQL As String
Dim M_AGENT As String
    m_objrs.CursorLocation = adUseClient
    If Len(TL) > 0 Then
        m_objrs.Open "SELECT * FROM USERTBL WHERE UNIT = '" + UNIT + "' AND SPVCODE = '" + TL + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    Else
        m_objrs.Open "SELECT * FROM USERTBL WHERE UNIT = '" + UNIT + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
    TL = ""
    UNIT = ""
    While Not m_objrs.EOF
        M_AGENT = IIf(IsNull(m_objrs("USERID")), "", m_objrs("USERID"))
       
       Set M_LOGIN = New ADODB.Recordset
        M_LOGIN.CursorLocation = adUseClient
        CMDSQL = "SELECT ACTIVITY FROM USERLOG WHERE AGENT = '" + M_AGENT + "' AND DATETIME >= '" + Format(MDIForm1.TDBDate1.Text, "mm/dd/yyyy") & "-00:00" + "' AND DATETIME <= '" + Format(MDIForm1.TDBDate1.Text, "mm/dd/yyyy") & "-23:59" + "' "
        M_LOGIN.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
         If M_LOGIN.RecordCount <> 0 Then
            If M_LOGIN("ACTIVITY") = "LOGIN" Then
                Set imgX = ImageList1. _
                    ListImages.ADD(, M_AGENT, LoadPicture(App.Path + "\icon\Drivenet.ico"))
                ListView1.Icons = ImageList1
                Set itmX = ListView1.ListItems.ADD()
                    itmX.Icon = M_AGENT
                    itmX.Text = M_AGENT
            Else
                Set imgX = ImageList1. _
                    ListImages.ADD(, M_AGENT, LoadPicture(App.Path + "\icon\Drivedsc.ico"))
                ListView1.Icons = ImageList1
                Set itmX = ListView1.ListItems.ADD()
                    itmX.Icon = M_AGENT
                    itmX.Text = M_AGENT   ' Set Text string.
            End If
        Else
            Set imgX = ImageList1. _
                ListImages.ADD(, M_AGENT, LoadPicture(App.Path + "\icon\Drivedsc.ico"))
            ListView1.Icons = ImageList1
            Set itmX = ListView1.ListItems.ADD()
                itmX.Icon = M_AGENT
                itmX.Text = M_AGENT   ' Set Text string.
        End If
        Set M_LOGIN = Nothing
    m_objrs.MoveNext
    Wend
    
   Set M_LOGIN = Nothing
     ' Set Icons property.
End Sub

Private Sub ListView1_DblClick()
'    WaitSecs (2)
    Call SEND_DATA_CAPTURE
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
If KeyAscii = 13 Then
    Call SEND_DATA_CAPTURE
End If
End Sub

Private Sub SEND_DATA_CAPTURE()
Dim m_objrs As ADODB.Recordset
Dim TETS As Object
On Error Resume Next
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient

m_objrs.Open "SELECT ACTIVITY FROM USERLOG WHERE AGENT = '" + ListView1.SelectedItem.Text + "' AND DATETIME >= '" + Format(MDIForm1.TDBDate1.Text, "mm/dd/yyyy") & "-00:00" + "' AND DATETIME <= '" + Format(MDIForm1.TDBDate1.Text, "mm/dd/yyyy") & "-23:59" + "' ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

If m_objrs.RecordCount > 0 Then
    If m_objrs("ACTIVITY") = "LOGIN" Then
        Call SEND_CAPTURE_GAMBAR("CAPTURE" + "/" + FRM_LOGIN_LOGOUT.ListView1.SelectedItem.Text + "^_^")
        'Call SEND_CAPTURE_GAMBAR("CAPTURE" + "/" + NAMA + "^_^")
        'WaitSecs (2)

    Else
        MsgBox "Tidak Dapat Di Capture... Agent Tidak Aktif...!!", vbInformation + vbOKOnly, "TeleGrandi"
        Exit Sub
    End If
Else
    MsgBox "Tidak Dapat Di Capture... Agent Tidak Aktif...!!", vbInformation + vbOKOnly, "TeleGrandi"
    Exit Sub
End If
End Sub

Private Sub Option1_Click(Index As Integer)
    ListView1.VIEW = Index
End Sub

