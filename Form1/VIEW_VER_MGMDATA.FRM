VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_VER_MGMDATA 
   BackColor       =   &H80000004&
   Caption         =   "MGM Data"
   ClientHeight    =   9780
   ClientLeft      =   -3345
   ClientTop       =   450
   ClientWidth     =   13185
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "VIEW_VER_MGMDATA.frx":0000
   LinkTopic       =   "Form1"
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
      Left            =   10995
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   10770
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
      Left            =   14085
      TabIndex        =   2
      Top             =   10725
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   10710
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   15210
      Begin MSComctlLib.ListView ListView1 
         Height          =   10530
         Left            =   15
         TabIndex        =   1
         Top             =   135
         Width           =   15150
         _ExtentX        =   26723
         _ExtentY        =   18574
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
            Weight          =   400
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
End
Attribute VB_Name = "FRM_VER_MGMDATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HEADER_VIEW_ALL()
    ListView1.ColumnHeaders.ADD 1, , "No", 5 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Customers Id", 10 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Agent", 10 * TXT
    ListView1.ColumnHeaders.ADD 4, , "NamaAgent", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Customers Name", 20 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Alamat Rumah", 15 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Kantor", 15 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Alamat Kantor", 15 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Home Telp", 10 * TXT
    ListView1.ColumnHeaders.ADD 10, , "Home Telp2", 10 * TXT
    ListView1.ColumnHeaders.ADD 11, , "Office Telp", 10 * TXT
    ListView1.ColumnHeaders.ADD 12, , "Office Telp2", 10 * TXT
    ListView1.ColumnHeaders.ADD 13, , "Fax", 10 * TXT
    ListView1.ColumnHeaders.ADD 14, , "Fax2", 10 * TXT
    ListView1.ColumnHeaders.ADD 15, , "Hp", 10 * TXT
    ListView1.ColumnHeaders.ADD 16, , "Hp2", 10 * TXT
    ListView1.ColumnHeaders.ADD 17, , "Database", 10 * TXT
    ListView1.ColumnHeaders.ADD 18, , "Tgl Entry", 10 * TXT
    ListView1.ColumnHeaders.ADD 19, , "Tgl Schedule", 10 * TXT
    ListView1.ColumnHeaders.ADD 20, , "Next Action", 17 * TXT
    ListView1.ColumnHeaders.ADD 21, , "LastCall Date", 10 * TXT
    ListView1.ColumnHeaders.ADD 22, , "Sts LastCall", 10 * TXT
End Sub
  
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Dim LISTITEM As LISTITEM
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

Dim I As Integer
I = 1
On Error GoTo HELL
Me.MousePointer = vbHourglass
Call HEADER_VIEW_ALL
    Text2.Text = "View All"
    
    With FRM_VER_SEARCH
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
            OFFPHONE = "OFFICENO Like '%" + .TDBMask1.Value + "%'"
            OFFPHONE2 = "OFFICENO2 Like  '%" + .TDBMask1.Value + "%'"
            HOMEPHONE = "HOMENO Like  '%" + .TDBMask1.Value + "%'"
            HOMEPHONE2 = "HOMENO2 Like  '%" + .TDBMask1.Value + "%'"
            FAXPHONE = "FAXNO Like  '%" + .TDBMask1.Value + "%'"
            FAXPHONE2 = "FAXNO2 Like  '%" + .TDBMask1.Value + "%'"
        End If
        If Len(.TDBMask2.Value) > 4 Then
            MOBILEPHONE = "MOBILENO Like '%" + .TDBMask2.Value + "%'"
            MOBILEPHONE2 = "MOBILENO2 Like '%" + .TDBMask2.Value + "%'"
        End If
        Set m_objrs = M_DATA.QUERY_SEARCH_CONDITION_mgm(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
                                                OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text)
    End With
    FRM_VER_SEARCH.ProgressBar1.Max = m_objrs.RecordCount + 1
    While Not m_objrs.EOF
    FRM_VER_SEARCH.ProgressBar1.Value = m_objrs.Bookmark
        Set LISTITEM = ListView1.ListItems.ADD(, , m_objrs.Bookmark)
            LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("custid")), "", JADI_QUOTE(m_objrs("custid")))
            LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
            LISTITEM.SubItems(3) = IIf(IsNull(m_objrs("NamaAGENT")), "", m_objrs("NamaAGENT"))
            LISTITEM.SubItems(4) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
            LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("ADDRNOW")), "", m_objrs("ADDRNOW"))
            LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("NAMAPT")), "", m_objrs("NAMAPT"))
            LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("ADDRPT")), "", m_objrs("ADDRPT"))
            LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("HOMENO")), "", m_objrs("HOMENO"))
            LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("HOMENO2")), "", m_objrs("HOMENO2"))
            LISTITEM.SubItems(10) = IIf(IsNull(m_objrs("OFFICENO")), "", m_objrs("OFFICENO"))
            LISTITEM.SubItems(11) = IIf(IsNull(m_objrs("OFFICENO2")), "", m_objrs("OFFICENO2"))
            LISTITEM.SubItems(12) = IIf(IsNull(m_objrs("FAXNO")), "", m_objrs("FAXNO"))
            LISTITEM.SubItems(13) = IIf(IsNull(m_objrs("FAXNO2")), "", m_objrs("FAXNO2"))
            LISTITEM.SubItems(14) = IIf(IsNull(m_objrs("MOBILENO")), "", m_objrs("MOBILENO"))
            LISTITEM.SubItems(15) = IIf(IsNull(m_objrs("MOBILENO2")), "", m_objrs("MOBILENO2"))
            LISTITEM.SubItems(16) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
            LISTITEM.SubItems(17) = IIf(IsNull(m_objrs("TGLSOURCE")), "", Format(m_objrs("TGLSOURCE"), "YYYY/MM/DD"))
            LISTITEM.SubItems(18) = IIf(IsNull(m_objrs("NEXTACTDATE")), "", Format(m_objrs("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
            LISTITEM.SubItems(19) = IIf(IsNull(m_objrs("NEXTACT")), "", m_objrs("NEXTACT"))
            LISTITEM.SubItems(20) = IIf(IsNull(m_objrs("TglStatus")), "", Format(m_objrs("TglStatus"), "YYYY/MM/DD"))
            LISTITEM.SubItems(21) = IIf(IsNull(m_objrs("StsLastCall")), "", m_objrs("StsLastCall"))
    m_objrs.MoveNext
    Wend
    If ListView1.ListItems.Count = 0 Then
        Text1.Text = "Tidak Ada Data"
    Else
        Text1.Text = "Total " + CStr(m_objrs.RecordCount) + " Records"
    End If
ListView1.SortKey = 2
ListView1.Sorted = True
FRM_VER_SEARCH.ProgressBar1.Value = 0
FRM_VER_SEARCH.ProgressBar1.Visible = False
Set m_objrs = Nothing
Me.MousePointer = vbNormal
Unload FRM_VER_SEARCH
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

