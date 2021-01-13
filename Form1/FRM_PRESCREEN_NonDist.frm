VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRM_PRESCREEN_NonDist 
   BackColor       =   &H80000004&
   Caption         =   "MGM Data"
   ClientHeight    =   9780
   ClientLeft      =   -3345
   ClientTop       =   450
   ClientWidth     =   13185
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "FRM_PRESCREEN_NonDist.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   13185
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   10875
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   10440
      Width           =   3045
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Close"
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
      Left            =   13965
      TabIndex        =   2
      Top             =   10395
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   10305
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   15210
      Begin MSComctlLib.ListView ListView1 
         Height          =   10140
         Left            =   45
         TabIndex        =   1
         Top             =   135
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   17886
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
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   300
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   765
      Visible         =   0   'False
      Width           =   9060
   End
   Begin VB.Menu MnFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnclaim 
         Caption         =   "Claim"
      End
   End
End
Attribute VB_Name = "FRM_PRESCREEN_NonDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HEADER_VIEW_ALL()
    ListView1.ColumnHeaders.ADD 1, , "No", 3 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Priority", 5 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Tgl Schedule", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Nama Customer", 25 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Next Action", 17 * TXT
'    ListView1.ColumnHeaders.ADD 7, , "Tlp. Rumah", 10 * TXT
'    ListView1.ColumnHeaders.ADD 8, , "Tlp. Kantor", 10 * TXT
'    ListView1.ColumnHeaders.ADD 9, , "Tlp. Hp", 10 * TXT
'    ListView1.ColumnHeaders.ADD 7, , "Team Leader", 15 * TXT
    ListView1.ColumnHeaders.ADD 7, , "SalesCode", 10 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Agent", 10 * TXT
    ListView1.ColumnHeaders.ADD 9, , "DataBase", 10 * TXT
    ListView1.ColumnHeaders.ADD 10, , "LastCall Date", 10 * TXT
'    ListView1.ColumnHeaders.ADD 11, , "Sts LastCall", 10 * TXT
'    ListView1.ColumnHeaders.ADD 12, , "Code", 5 * TXT
'    ListView1.ColumnHeaders.ADD 13, , "Complaint Note", 15 * TXT
End Sub
  
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Dim listitem As listitem
Dim M_AGENT As String
Dim M_DATAS As String
Dim M_SPV As String

Dim NAMACUST As String
Dim NAMAAGENT As String
Dim DATASOURCE As String
Dim TGLLAHIR As String
Dim OFFPHONE As String
Dim OFFPHONE2 As String
Dim HOMEPHONE As String
Dim HOMEPHONE2 As String
Dim MOBILEPHONE As String
Dim MOBILEPHONE2 As String
Dim M_DATA As New CLS_FRMSEARCH
Dim FAXPHONE As String
Dim FAXPHONE2 As String

Dim i As Integer
i = 1
On Error GoTo HELL
Me.MousePointer = vbHourglass
Call HEADER_VIEW_ALL
    Text2.Text = "View All"
    
    With FRM_SEARCH_NonDist
    .Height = 4815
    .Frame1.Visible = True
        If .Text1(0).Text <> Empty Then
            NAMACUST = "NAME LIKE " + "'%" + UBAH_QUOTE(.Text1(0).Text) + "%'"
        End If
        If .Combo1(0).Text <> Empty Then
            NAMAAGENT = "AGENT = '" + .Combo1(0).Text + "'"
        End If
        If .Combo1(2).Text <> Empty Then
            DATASOURCE = "RECSOURCE = '" + .Combo1(2).Text + "'"
        End If
        If .TDBDate1.ValueIsNull Then
        Else
            TGLLAHIR = "BIRTHD = '" + Format(.TDBDate1.Value, "mm/dd/yyyy") + "'"
        End If
        If Len(.TDBMask1.Value) > 4 Then
            OFFPHONE = "OFFICENO = '" + .TDBMask1.Value + "'"
            OFFPHONE2 = "OFFICENO2 = '" + .TDBMask1.Value + "'"
            HOMEPHONE = "HOMENO = '" + .TDBMask1.Value + "'"
            HOMEPHONE2 = "HOMENO2 = '" + .TDBMask1.Value + "'"
            FAXPHONE = "FAXNO = '" + .TDBMask1.Value + "'"
            FAXPHONE2 = "FAXNO2 = '" + .TDBMask1.Value + "'"
        End If
        If Len(.TDBMask2.Value) > 4 Then
            MOBILEPHONE = "MOBILENO = '" + .TDBMask2.Value + "'"
            MOBILEPHONE2 = "MOBILENO2 = '" + .TDBMask2.Value + "'"
        End If
        Set m_objrs = M_DATA.QUERY_SEARCH_nonDist(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
                                                OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text)
    End With
    FRM_SEARCH_NonDist.ProgressBar1.Max = m_objrs.RecordCount + 1
    While Not m_objrs.EOF
    FRM_SEARCH_NonDist.ProgressBar1.Value = m_objrs.Bookmark
        Set listitem = ListView1.ListItems.ADD(, , m_objrs.Bookmark)
        listitem.SubItems(1) = IIf(IsNull(m_objrs("CUSTID")), "", m_objrs("CUSTID"))
        listitem.SubItems(2) = "Available"
        listitem.SubItems(3) = IIf(IsNull(m_objrs("NEXTACTDATE")), "", Format(m_objrs("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
        listitem.SubItems(4) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
        listitem.SubItems(5) = IIf(IsNull(m_objrs("NEXTACT")), "", m_objrs("NEXTACT"))
   '     LISTITEM.SubItems(6) = IIf(IsNull(M_OBJRS("SPVcode")), "", M_OBJRS("SPVcode"))
        listitem.SubItems(6) = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
        listitem.SubItems(7) = IIf(IsNull(m_objrs("NamaAGENT")), "", m_objrs("NamaAGENT"))
        listitem.SubItems(8) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
        listitem.SubItems(9) = IIf(IsNull(m_objrs("TGLSTATUS")), "", Format(m_objrs("TGLSTATUS"), "DD/MM/YYYY"))
 '       listitem.SubItems(10) = IIf(IsNull(m_objrs("StsLastCall")), "", m_objrs("StsLastCall"))
 '       listitem.SubItems(11) = IIf(IsNull(m_objrs("KdComplaint")), "", m_objrs("KdComplaint"))
 '       listitem.SubItems(12) = IIf(IsNull(m_objrs("RemarkComplaint")), "", m_objrs("RemarkComplaint"))
        m_objrs.MoveNext
    Wend
    
    If ListView1.ListItems.Count = 0 Then
        Text1.Text = "Tidak Ada Data"
    Else
        Text1.Text = "Total " + CStr(m_objrs.RecordCount) + " Records"
    End If
ListView1.SortKey = 2
ListView1.Sorted = True
FRM_SEARCH_NonDist.ProgressBar1.Value = 0
FRM_SEARCH_NonDist.ProgressBar1.Visible = False
Set m_objrs = Nothing
Me.MousePointer = vbNormal
Unload FRM_SEARCH_NonDist
Exit Sub
HELL:
    Me.MousePointer = vbNormal
    MsgBox Err.Description
    Set m_objrs = Nothing
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If UCase(MDIForm1.Text1.Text) = UCase(ListView1.SelectedItem.SubItems(6)) Then
    Exit Sub
End If
If Button = 2 Then
  Dim hItem As Long, nod As Node, xRect As RECT, xPop As Integer, yPop As Integer
  Set nod = ListView1.HitTest(x, y)

  If ListView1.ListItems.Count <> 0 Then
   TreeView_GetItemRect ListView1.hWnd, hItem, xRect, CTrue
 '  MsgBox xRect.Left & "," & xRect.Top & "," & xRect.Right & "," & xRect.Bottom
   xPop = ListView1.SelectedItem.Left + ScaleX(xRect.Left, vbPixels, vbTwips)
   yPop = ListView1.SelectedItem.Top + ScaleY(xRect.Bottom, vbPixels, vbTwips)
   PopupMenu MnFile, , xPop, yPop
  End If
End If
End Sub

Private Sub ListView1_DblClick()
If ListView1.ListItems.Count = 0 Then
    Exit Sub
End If

Dim m_objrs As ADODB.Recordset
If ListView1.ListItems.Count = 0 Then
    Exit Sub
End If
    If UCase(Me.Caption) = "SEARCH" Then
        If UCase(MDIForm1.Text2.Text) = "AGENT" Then
            If UCase(MDIForm1.Text1.Text) <> UCase(ListView1.SelectedItem.SubItems(6)) Then
                MsgBox "Anda Tidak Berhak Untuk Mengedit Data Ini", vbCritical + vbOKOnly, "TeleGrandi"
                Exit Sub
            End If
        End If
        If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
            Set m_objrs = New ADODB.Recordset
            m_objrs.CursorLocation = adUseClient
            If UCase(Me.Caption) = "SEARCH" Then
                m_objrs.Open "SELECT USERID FROM USERTBL WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(6) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            Else
                m_objrs.Open "SELECT USERID FROM USERTBL WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(6) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            End If
            If m_objrs.RecordCount <> 0 Then
            Else
                MsgBox "Data Ini Milik TSE Team Leader Yang Lain", vbCritical + vbOKOnly, "TeleGrandi"
                Set m_objrs = Nothing
                Exit Sub
            End If
            Set m_objrs = Nothing
        End If
    Else
        If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
            Set m_objrs = New ADODB.Recordset
            m_objrs.CursorLocation = adUseClient
            If UCase(Me.Caption) = "SEARCH" Then
                m_objrs.Open "SELECT USERID FROM USERTBL WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(6) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            Else
                m_objrs.Open "SELECT USERID FROM USERTBL WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(6) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            End If
            If m_objrs.RecordCount <> 0 Then
            Else
                MsgBox "Data Ini Milik TSE Team Leader Yang Lain", vbCritical + vbOKOnly, "TeleGrandi"
                Set m_objrs = Nothing
                Exit Sub
            End If
            Set m_objrs = Nothing
        End If
    End If
    SCREENER = False
    FRMCUST_CC_MGM.Show vbModal
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call ListView1_DblClick
End If
End Sub

Private Sub mnclaim_Click()
    With FRMCLAIM
        .Text1.Text = ListView1.SelectedItem.SubItems(1)
        .Text5.Text = ListView1.SelectedItem.SubItems(4)
        .Text6.Text = ListView1.SelectedItem.SubItems(6)
        .Text7.Text = ListView1.SelectedItem.SubItems(7)
        .Show vbModal
    End With
End Sub
