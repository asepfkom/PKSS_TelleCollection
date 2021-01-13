VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmUserTarget_list 
   Caption         =   "Target"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "FrmUserTarget_list.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   7590
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   11295
      Begin VB.ComboBox CmbSpv 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   165
         Width           =   2130
      End
      Begin VB.ComboBox CmbTahun 
         Height          =   315
         Left            =   4305
         TabIndex        =   3
         Top             =   465
         Width           =   945
      End
      Begin VB.ComboBox CmbBulan 
         Height          =   315
         Left            =   4305
         TabIndex        =   2
         Top             =   150
         Width           =   810
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6660
         Left            =   30
         TabIndex        =   1
         Top             =   885
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   11748
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
      Begin Threed.SSCommand Command1 
         Height          =   330
         Index           =   0
         Left            =   5400
         TabIndex        =   8
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Search"
      End
      Begin Threed.SSCommand Command1 
         Height          =   360
         Index           =   1
         Left            =   10470
         TabIndex        =   9
         Top             =   900
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Add"
      End
      Begin Threed.SSCommand Command1 
         Height          =   360
         Index           =   2
         Left            =   10470
         TabIndex        =   10
         Top             =   1275
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Update"
      End
      Begin Threed.SSCommand Command1 
         Height          =   360
         Index           =   3
         Left            =   10470
         TabIndex        =   11
         Top             =   1665
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Delete"
      End
      Begin Threed.SSCommand Command1 
         Cancel          =   -1  'True
         Height          =   360
         Index           =   4
         Left            =   10470
         TabIndex        =   12
         Top             =   2055
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   635
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Close"
      End
      Begin Threed.SSCommand Command1 
         Height          =   330
         Index           =   5
         Left            =   5415
         TabIndex        =   13
         Top             =   465
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Transfer"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Supervisor :"
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Tahun :"
         Height          =   270
         Index           =   0
         Left            =   3630
         TabIndex        =   6
         Top             =   510
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Bulan :"
         Height          =   270
         Left            =   3630
         TabIndex        =   5
         Top             =   180
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmUserTarget_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Id", 1 * 120
    ListView1.ColumnHeaders.ADD 2, , "Agent Code", 10 * 120
    ListView1.ColumnHeaders.ADD 3, , "Agent Name", 15 * 120
    ListView1.ColumnHeaders.ADD 4, , "Supervisor", 10 * 120 '
    ListView1.ColumnHeaders.ADD 5, , "Bulan", 5 * 120
    ListView1.ColumnHeaders.ADD 6, , "Tahun", 5 * 120
    ListView1.ColumnHeaders.ADD 7, , "Hadir1", 5 * 120
    ListView1.ColumnHeaders.ADD 8, , "Target1", 5 * 120
    ListView1.ColumnHeaders.ADD 9, , "Hadir2", 5 * 120
    ListView1.ColumnHeaders.ADD 10, , "Target2", 5 * 120
    ListView1.ColumnHeaders.ADD 11, , "Hadir3", 5 * 120
    ListView1.ColumnHeaders.ADD 12, , "Target3", 5 * 120
    ListView1.ColumnHeaders.ADD 13, , "Hadir4", 5 * 120
    ListView1.ColumnHeaders.ADD 14, , "Target4", 5 * 120
    ListView1.ColumnHeaders.ADD 15, , "Hadir5", 5 * 120
    ListView1.ColumnHeaders.ADD 16, , "Target5", 5 * 120
End Sub


Private Sub Command1_Click(Index As Integer)
Dim cmdsql As String
Dim m_obj1 As ADODB.Recordset
Dim m_objrs As ADODB.Recordset
Dim m_msgbox As Variant
On Error GoTo ADD
Select Case Index
    Case 0
        Call cari_data
    Case 1
        With FrmUserTarget
                .Caption = "Tambah Data "
                .Show vbModal
                If .ok Then
                    cmdsql = " Insert Into UserTblTarget "
                    cmdsql = cmdsql + " (UserId,NamaAgent,SpvCode,Bulan,tahun,Absent1,Absent2,Absent3,Absent4,absent5, target1,target2,target3,target4,target5)"
                    cmdsql = cmdsql + " Values"
                    cmdsql = cmdsql + " ('" + .Combo1(0).Text + "',"
                    cmdsql = cmdsql + " '" + .Combo1(1).Text + "',"
                    cmdsql = cmdsql + " '" + CmbSpv.Text + "',"
                    cmdsql = cmdsql + " " + .CmbBulan.Text + ","
                    cmdsql = cmdsql + " " + .CmbTahun.Text + ","
                    cmdsql = cmdsql + " " + .TxtAbsen(0).Text + ","
                    cmdsql = cmdsql + " " + .TxtAbsen(1).Text + ","
                    cmdsql = cmdsql + " " + .TxtAbsen(2).Text + ","
                    cmdsql = cmdsql + " " + .TxtAbsen(3).Text + ","
                    cmdsql = cmdsql + " " + .TxtAbsen(4).Text + ","
                    cmdsql = cmdsql + " " + .TxtTarget(0).Text + ","
                    cmdsql = cmdsql + " " + .TxtTarget(1).Text + ","
                    cmdsql = cmdsql + " " + .TxtTarget(2).Text + ","
                    cmdsql = cmdsql + " " + .TxtTarget(3).Text + ","
                    cmdsql = cmdsql + " " + .TxtTarget(4).Text + ")"
                    If .Combo1(0).Text = "" Then
                        Exit Sub
                    End If
                    M_OBJCONN.Execute cmdsql
                    Call cari_data
                End If
          Unload FrmUserTarget
        End With
    Case 2
        If ListView1.ListItems.count = 0 Then
            Exit Sub
        End If
        With FrmUserTarget
                .Caption = "Tambah Data "
                .Text1.Text = ListView1.SelectedItem.Text
                .Combo1(0).Text = ListView1.SelectedItem.SubItems(1)
                .Combo1(1).Text = ListView1.SelectedItem.SubItems(2)
                .CmbBulan.Text = ListView1.SelectedItem.SubItems(4)
                .CmbTahun.Text = ListView1.SelectedItem.SubItems(5)
                .TxtAbsen(0).Text = ListView1.SelectedItem.SubItems(6)
                .TxtTarget(0).Text = ListView1.SelectedItem.SubItems(7)
                .TxtAbsen(1).Text = ListView1.SelectedItem.SubItems(8)
                .TxtTarget(1).Text = ListView1.SelectedItem.SubItems(9)
                .TxtAbsen(2).Text = ListView1.SelectedItem.SubItems(10)
                .TxtTarget(2).Text = ListView1.SelectedItem.SubItems(11)
                .TxtAbsen(3).Text = ListView1.SelectedItem.SubItems(12)
                .TxtTarget(3).Text = ListView1.SelectedItem.SubItems(13)
                .TxtAbsen(4).Text = ListView1.SelectedItem.SubItems(14)
                .TxtTarget(4).Text = ListView1.SelectedItem.SubItems(15)
                .Show vbModal
                If .ok Then
                    cmdsql = " Update UserTblTarget "
                    cmdsql = cmdsql + " set UserId='" + .Combo1(0).Text + "',"
                    cmdsql = cmdsql + " NamaAgent='" + .Combo1(1).Text + "',"
                    cmdsql = cmdsql + " Bulan=" + .CmbBulan.Text + ","
                    cmdsql = cmdsql + " tahun=" + .CmbTahun.Text + ","
                    cmdsql = cmdsql + " Absent1=" + .TxtAbsen(0).Text + ","
                    cmdsql = cmdsql + " Absent2=" + .TxtAbsen(1).Text + ","
                    cmdsql = cmdsql + " Absent3=" + .TxtAbsen(2).Text + ","
                    cmdsql = cmdsql + " Absent4=" + .TxtAbsen(3).Text + ","
                    cmdsql = cmdsql + " Absent5=" + .TxtAbsen(4).Text + ","
                    cmdsql = cmdsql + " target1=" + .TxtTarget(0).Text + ","
                    cmdsql = cmdsql + " target2=" + .TxtTarget(1).Text + ","
                    cmdsql = cmdsql + " target3=" + .TxtTarget(2).Text + ","
                    cmdsql = cmdsql + " target4=" + .TxtTarget(3).Text + ","
                    cmdsql = cmdsql + " target5=" + .TxtTarget(4).Text + ""
                    cmdsql = cmdsql + " where Id = " + .Text1.Text + ""
                    M_OBJCONN.Execute cmdsql
                    
                    ListView1.SelectedItem.Text = .Text1.Text
                    ListView1.SelectedItem.SubItems(1) = .Combo1(0).Text
                    ListView1.SelectedItem.SubItems(2) = .Combo1(1).Text
                    ListView1.SelectedItem.SubItems(4) = .CmbBulan.Text
                    ListView1.SelectedItem.SubItems(5) = .CmbTahun.Text
                    ListView1.SelectedItem.SubItems(6) = .TxtAbsen(0).Text
                    ListView1.SelectedItem.SubItems(7) = .TxtTarget(0).Text
                    ListView1.SelectedItem.SubItems(8) = .TxtAbsen(1).Text
                    ListView1.SelectedItem.SubItems(9) = .TxtTarget(1).Text
                    ListView1.SelectedItem.SubItems(10) = .TxtAbsen(2).Text
                    ListView1.SelectedItem.SubItems(11) = .TxtTarget(2).Text
                    ListView1.SelectedItem.SubItems(12) = .TxtAbsen(3).Text
                    ListView1.SelectedItem.SubItems(13) = .TxtTarget(3).Text
                    ListView1.SelectedItem.SubItems(14) = .TxtAbsen(4).Text
                    ListView1.SelectedItem.SubItems(15) = .TxtTarget(4).Text
                
                End If
          Unload FrmUserTarget
        End With
    Case 3
        If ListView1.ListItems.count = 0 Then
            Exit Sub
        End If
        m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
        If m_msgbox = 1 Then
            cmdsql = "Delete From UserTblTarget where Id =" + ListView1.SelectedItem.Text + ""
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
        End If
    Case 4
        Unload Me
    Case 5
        If CmbSpv.Text = "" Or CmbBulan.Text = "" Or CmbTahun.Text = "" Then
            MsgBox "Data Tidak Lengkap", vbCritical + vbOKOnly, "Informasi"
            Exit Sub
        End If
        Set m_obj1 = New ADODB.Recordset
        Set m_objrs = New ADODB.Recordset
        m_obj1.CursorLocation = adUseClient
        m_objrs.CursorLocation = adUseClient
        m_obj1.Open "Select * from UserTblTarget where spvcode = '" + CmbSpv.Text + "' and  Bulan ='" + CmbBulan.Text + "' and tahun = '" + CmbTahun.Text + "' ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If m_obj1.RecordCount <> 0 Then
            m_msgbox = MsgBox("Data sudah pernah ada.. Teruskan Proses??..", vbYesNo, "Informasi")
            If m_msgbox = vbNo Then
                Call cari_data
                Exit Sub
            End If
        End If
        M_OBJCONN.Execute "Delete from UserTblTarget where spvcode = '" + CmbSpv.Text + "' and  Bulan ='" + CmbBulan.Text + "' and tahun = '" + CmbTahun.Text + "'"
        m_objrs.Open "Select * from usertbl where usertype = 1 and aktif = 0 and spvcode ='" + CmbSpv.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        While Not m_objrs.EOF
            m_obj1.AddNew
            m_obj1!USERID = m_objrs!USERID
            m_obj1!NAMAAGENT = m_objrs!agent
            m_obj1!SPVCODE = m_objrs!SPVCODE
            m_obj1!Bulan = CmbBulan.Text
            m_obj1!tahun = CmbTahun.Text
            m_obj1!Absent1 = 0
            m_obj1!Absent2 = 0
            m_obj1!Absent3 = 0
            m_obj1!Absent4 = 0
            m_obj1!Absent5 = 0
            m_obj1!target1 = 0
            m_obj1!target2 = 0
            m_obj1!target3 = 0
            m_obj1!target4 = 0
            m_obj1!target5 = 0
            m_obj1.UPDATE
            m_objrs.MoveNext
        Wend
        Set m_obj1 = Nothing
        Set m_objrs = Nothing
        Call cari_data
End Select
Exit Sub
ADD:
MsgBox Err.Description
'Resume
End Sub

Private Sub cari_data()
Dim LISTITEM As LISTITEM
Dim m_objrs As New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
    cmdsql = "Select * from USERTBL "
    cmdsql = cmdsql + " INNER JOIN"
    cmdsql = cmdsql + " UserTblTarget ON USERTBL.USERID = UserTblTarget.UserId"
    cmdsql = cmdsql + " Where USERTBL.SPVCODE ='" + CmbSpv.Text + "'"
    If Trim(CmbBulan.Text) <> "" Then
        cmdsql = cmdsql + " And UserTblTarget.Bulan =" + CmbBulan.Text + " "
    End If
    If Trim(CmbTahun.Text) <> "" Then
        cmdsql = cmdsql + " and UserTblTarget.Tahun =" + CmbTahun.Text + " "
    End If
    cmdsql = cmdsql + " and usertbl.aktif = 0 "
    
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

ListView1.ListItems.Clear
While Not m_objrs.EOF
    Set LISTITEM = ListView1.ListItems.ADD(, , m_objrs!Id)
        LISTITEM.SubItems(1) = IIf(IsNull(m_objrs!USERID), "", m_objrs!USERID)
        LISTITEM.SubItems(2) = IIf(IsNull(m_objrs!agent), "", m_objrs!agent)
        LISTITEM.SubItems(3) = IIf(IsNull(m_objrs!SPVCODE), "", m_objrs!SPVCODE)
        LISTITEM.SubItems(4) = IIf(IsNull(m_objrs!Bulan), "0", m_objrs!Bulan)
        LISTITEM.SubItems(5) = IIf(IsNull(m_objrs!tahun), "0", m_objrs!tahun)
        LISTITEM.SubItems(6) = IIf(IsNull(m_objrs!Absent1), "0", m_objrs!Absent1)
        LISTITEM.SubItems(7) = IIf(IsNull(m_objrs!target1), "0", m_objrs!target1)
        LISTITEM.SubItems(8) = IIf(IsNull(m_objrs!Absent2), "0", m_objrs!Absent2)
        LISTITEM.SubItems(9) = IIf(IsNull(m_objrs!target2), "0", m_objrs!target2)
        LISTITEM.SubItems(10) = IIf(IsNull(m_objrs!Absent3), "0", m_objrs!Absent3)
        LISTITEM.SubItems(11) = IIf(IsNull(m_objrs!target3), "0", m_objrs!target3)
        LISTITEM.SubItems(12) = IIf(IsNull(m_objrs!Absent4), "0", m_objrs!Absent4)
        LISTITEM.SubItems(13) = IIf(IsNull(m_objrs!target4), "0", m_objrs!target4)
        LISTITEM.SubItems(14) = IIf(IsNull(m_objrs!Absent5), "0", m_objrs!Absent5)
        LISTITEM.SubItems(15) = IIf(IsNull(m_objrs!target5), "0", m_objrs!target5)
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
End Sub

    
Private Sub Form_Load()
Dim LISTITEM As LISTITEM
Dim m_objrs As ADODB.Recordset
CmbSpv.Text = MDIForm1.Text1.Text
Call header
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
'm_objrs.Open "SELECT * FROM SPVTBL ORDER BY SPVCODE", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
m_objrs.Open "select distinct SPVTBL.SPVCODE from SPVTBL, USERTBL where SPVTBL.SPVCODE = USERTBL.SPVCODE AND USERTYPE = '6'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

While Not m_objrs.EOF
    CmbSpv.AddItem m_objrs!SPVCODE
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing

CmbTahun.AddItem 2005
CmbTahun.AddItem 2006
CmbTahun.AddItem 2007
CmbTahun.AddItem 2008
CmbTahun.AddItem 2009
CmbTahun.AddItem 2010
For i = 1 To 12
    CmbBulan.AddItem i
Next i
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
    Call Command1_Click(2)
End Sub
