VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VIEWCUSTINCOMING 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9870
   ClientLeft      =   -3360
   ClientTop       =   435
   ClientWidth     =   11925
   Icon            =   "VIEWCUSTINCOMING.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9870
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   345
      Left            =   7860
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   9450
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
      Height          =   360
      Left            =   10980
      TabIndex        =   2
      Top             =   9450
      Width           =   900
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   660
      Width           =   11880
      Begin MSComctlLib.ListView ListView1 
         Height          =   8460
         Left            =   30
         TabIndex        =   1
         Top             =   150
         Width           =   11805
         _ExtentX        =   20823
         _ExtentY        =   14923
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   1164
      _Version        =   196610
      Font3D          =   5
      ForeColor       =   0
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
      Caption         =   "Tampilkan Data Incoming"
      BevelWidth      =   2
      BorderWidth     =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
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
End
Attribute VB_Name = "VIEWCUSTINCOMING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub HEADER_VIEW_BANYAK()
    ListView1.ColumnHeaders.ADD 1, , "No.", 4 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Customers Id", 15 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Customers Name", 40 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Alamat", 15 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Tanggal Lahir", 15 * TXT
    ListView1.ColumnHeaders.ADD 6, , "No. Telephone", 15 * TXT
    ListView1.ColumnHeaders.ADD 7, , "No. Tlp. Kantor", 15 * TXT
    ListView1.ColumnHeaders.ADD 8, , "No. Mobile", 15 * TXT
    ListView1.ColumnHeaders.ADD 9, , "TeamLeader Name", 50 * TXT
    ListView1.ColumnHeaders.ADD 10, , "Agent Name", 50 * TXT
End Sub

Private Sub HEADER_VIEW_ALL()
    ListView1.ColumnHeaders.ADD 1, , "No.", 4 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Customers Id", 15 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Customers Name", 40 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Alamat", 50 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Tanggal Lahir", 15 * TXT
    ListView1.ColumnHeaders.ADD 6, , "No. Telephone", 15 * TXT
    ListView1.ColumnHeaders.ADD 7, , "No. Telp. Kantor", 15 * TXT
    ListView1.ColumnHeaders.ADD 8, , "No. Mobile", 18 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Data Source", 15 * TXT
    ListView1.ColumnHeaders.ADD 10, , "TeamLeader Name", 50 * TXT
    ListView1.ColumnHeaders.ADD 11, , "Agent Name", 50 * TXT
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New VIEW
Dim LISTITEM As LISTITEM
Dim M_AGENT As String
Dim M_DATAS As String
Dim M_SPV As String

Dim i As Integer
i = 1
On Error GoTo HELL

With frmVIEW_INCOMING
    If .Option1(1).Value Then
    Call HEADER_VIEW_BANYAK
    Text2.Text = "Berdasarkan " + "'" + frmVIEW_INCOMING.Combo1(1).Text + "'" + " -- " + "'" + frmVIEW_INCOMING.Combo1(3).Text + "'" + " -- " + "'" + frmVIEW_INCOMING.Combo1(5).Text + "'"
    Select Case UCase(MDIForm1.Text3.Text)
        Case "KTA"
            If .Combo1(1).Text <> Empty Then
                M_DATAS = " KTA_CUSTTBL.RECSOURCE = '" + .Combo1(0).Text + "'"
            End If
            If .Combo1(2).Text <> Empty Then
                M_AGENT = " KTA_CUSTTBL.AGENT = '" + .Combo1(2).Text + "'"
            End If
        Case "CREDIT CARD"
            If .Combo1(1).Text <> Empty Then
                M_DATAS = " CC_CUSTTBL.RECSOURCE = '" + .Combo1(0).Text + "'"
            End If
            If .Combo1(2).Text <> Empty Then
                M_AGENT = " CC_CUSTTBL.AGENT = '" + .Combo1(2).Text + "'"
            End If
        Case "KTA - CROSS SELL"
            If .Combo1(1).Text <> Empty Then
                M_DATAS = " CS_CUSTTBL.RECSOURCE = '" + .Combo1(0).Text + "'"
            End If
            If .Combo1(2).Text <> Empty Then
                M_AGENT = " CS_CUSTTBL.AGENT = '" + .Combo1(2).Text + "'"
            End If
        Case "CC - CROSS SELL"
            If .Combo1(1).Text <> Empty Then
                M_DATAS = " CCCS_CUSTTBL.RECSOURCE = '" + .Combo1(0).Text + "'"
            End If
            If .Combo1(2).Text <> Empty Then
                M_AGENT = " CCCS_CUSTTBL.AGENT = '" + .Combo1(2).Text + "'"
            End If
        Case Else
                M_DATAS = Empty
                M_AGENT = Empty
    End Select
        If .Combo1(4).Text <> Empty Then
            M_SPV = " USERTBL.SPVCODE = '" + .Combo1(4).Text + "'"
        End If
        MDIForm1.ProgressBar1.Visible = True
        Set m_objrs = M_DATA.QUERY_INCOMING_NEW(M_OBJCONN, M_DATAS, M_AGENT, M_SPV, "NAME", MDIForm1.Text3.Text)
        MDIForm1.ProgressBar1.Max = m_objrs.RecordCount + 2
            While Not m_objrs.EOF
                MDIForm1.ProgressBar1.Value = m_objrs.Bookmark
            Set LISTITEM = ListView1.ListItems.ADD(, , CStr(i))
                LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("CUSTID")), "", m_objrs("CUSTID"))
                LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
                LISTITEM.SubItems(3) = IIf(IsNull(m_objrs("ADDRNOW")), "", m_objrs("ADDRNOW"))
                If IsNull(m_objrs("BIRTHD")) Then
                    LISTITEM.SubItems(4) = " "
                Else
                    LISTITEM.SubItems(4) = Right(m_objrs("BIRTHD"), 2) + "/" + Mid(m_objrs("BIRTHD"), 5, 2) + "/" + Left(m_objrs("BIRTHD"), 4)
                End If
                LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("HOMENO")), "", m_objrs("HOMENO"))
                LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("OFFICENO")), "", m_objrs("OFFICENO"))
                LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("MOBILENO")), "", m_objrs("MOBILENO"))
                LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("SPVNAME")), "", m_objrs("SPVNAME"))
                LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("NAMAAGENT")), "", m_objrs("NAMAAGENT"))
                i = CCur(i) + 1
            m_objrs.MoveNext
            Wend
        If ListView1.ListItems.Count = 0 Then
            Text1.Text = "Tidak Ada Data Yang Incoming"
        Else
            Text1.Text = "Total " + CStr(m_objrs.RecordCount) + " Records"
            ListView1.SortKey = 1
            ListView1.Sorted = True
        End If
        MDIForm1.ProgressBar1.Value = 0
        MDIForm1.ProgressBar1.Visible = False
        Unload frmVIEW_INCOMING
    Exit Sub
End If
End With
Select Case UCase(Trim(frmVIEW_INCOMING.HEADER_JUDUL))
        Case "TAMPILKAN"
            Call HEADER_VIEW_ALL
            Text2.Text = "View All"
            MDIForm1.ProgressBar1.Visible = True
            Set m_objrs = M_DATA.QUERY_INCOMING(M_OBJCONN, "USERTBL.SPVCODE ='" + MDIForm1.Text1.Text + "'", " NAME", MDIForm1.Text3.Text)
            MDIForm1.ProgressBar1.Max = m_objrs.RecordCount + 2
            While Not m_objrs.EOF
                MDIForm1.ProgressBar1.Value = m_objrs.Bookmark
                Set LISTITEM = ListView1.ListItems.ADD(, , CStr(i))
                    LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("CUSTID")), "", m_objrs("CUSTID"))
                    LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
                    LISTITEM.SubItems(3) = IIf(IsNull(m_objrs("ADDRNOW")), "", m_objrs("ADDRNOW"))
                    If IsNull(m_objrs("BIRTHD")) Then
                        LISTITEM.SubItems(4) = " "
                    Else
                        LISTITEM.SubItems(4) = Right(m_objrs("BIRTHD"), 2) + "/" + Mid(m_objrs("BIRTHD"), 5, 2) + "/" + Left(m_objrs("BIRTHD"), 4)
                    End If
                        LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("HOMENO")), "", m_objrs("HOMENO"))
                        LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("OFFICENO")), "", m_objrs("OFFICENO"))
                        LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("MOBILENO")), "", m_objrs("MOBILENO"))
                        LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
                        LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("SPVNAME")), "", m_objrs("SPVNAME"))
                        LISTITEM.SubItems(10) = IIf(IsNull(m_objrs("NAMAAGENT")), "", m_objrs("NAMAAGENT"))
                    i = CCur(i) + 1
                m_objrs.MoveNext
            Wend
    End Select
    If ListView1.ListItems.Count = 0 Then
        Text1.Text = "Tidak Ada Data Yang Incoming"
    Else
        Text1.Text = "Total " + CStr(m_objrs.RecordCount) + " Records"
        ListView1.SortKey = 1
        ListView1.Sorted = True
    End If
    MDIForm1.ProgressBar1.Value = 0
    MDIForm1.ProgressBar1.Visible = False
    Unload frmVIEW_INCOMING
ListView1.SortKey = 2
ListView1.Sorted = True
Set M_DATA = Nothing
Set m_objrs = Nothing
Exit Sub
HELL:
    Set M_DATA = Nothing
    Set m_objrs = Nothing
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub

