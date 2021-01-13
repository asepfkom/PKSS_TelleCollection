VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCustIdReview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Review"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8265
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "List Review"
      TabPicture(0)   =   "FrmCustIdReview.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LvTransfer"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chk_all"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Log Release Review"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(0)"
      Tab(1).Control(1)=   "Label2(1)"
      Tab(1).Control(2)=   "TDBDate2"
      Tab(1).Control(3)=   "ListView1"
      Tab(1).Control(4)=   "TDBDate1"
      Tab(1).Control(5)=   "Command3"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   375
         Left            =   -70920
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   255
         Left            =   -73920
         TabIndex        =   8
         Top             =   600
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   450
         Calendar        =   "FrmCustIdReview.frx":001C
         Caption         =   "FrmCustIdReview.frx":0134
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmCustIdReview.frx":01A0
         Keys            =   "FrmCustIdReview.frx":01BE
         Spin            =   "FrmCustIdReview.frx":021C
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm/dd/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/dd/yyyy"
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
         Text            =   "07/03/2014"
         ValidateMode    =   0
         ValueVT         =   6815751
         Value           =   41823
         CenturyMode     =   0
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "SET TO OLD AGENT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6000
         Width           =   1815
      End
      Begin VB.CheckBox chk_all 
         Caption         =   "Check All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H008080FF&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6000
         Width           =   1815
      End
      Begin MSComctlLib.ListView LvTransfer 
         Height          =   5100
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   8996
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5100
         Left            =   -74760
         TabIndex        =   6
         Top             =   1080
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   8996
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin TDBDate6Ctl.TDBDate TDBDate2 
         Height          =   255
         Left            =   -72240
         TabIndex        =   10
         Top             =   600
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   450
         Calendar        =   "FrmCustIdReview.frx":0244
         Caption         =   "FrmCustIdReview.frx":035C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmCustIdReview.frx":03C8
         Keys            =   "FrmCustIdReview.frx":03E6
         Spin            =   "FrmCustIdReview.frx":0444
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm/dd/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/dd/yyyy"
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
         Text            =   "07/03/2014"
         ValidateMode    =   0
         ValueVT         =   6815751
         Value           =   41823
         CenturyMode     =   0
      End
      Begin VB.Label Label2 
         Caption         =   "to"
         Height          =   255
         Index           =   1
         Left            =   -72600
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         Height          =   255
         Index           =   0
         Left            =   -74400
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   1560
         X2              =   7440
         Y1              =   5760
         Y2              =   5760
      End
   End
End
Attribute VB_Name = "FrmCustIdReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs_list As ADODB.Recordset

Private Sub Command1_Click()
    Dim xxx As Integer
    Dim bceklist As Boolean
    
    bceklist = False
    konfirmasi_pesan = MsgBox("Data yang diceklist akan dikembalikan ke agent asli / lama ???", vbYesNo + vbQuestion, "Konfirmasi")
    If konfirmasi_pesan = vbYes Then
        For xxx = 1 To LvTransfer.ListItems.Count
            If LvTransfer.ListItems(xxx).Checked = True Then
                bceklist = True
                Exit For
            End If
        Next xxx
        
        If bceklist = True Then
            For xxx = 1 To LvTransfer.ListItems.Count
                If LvTransfer.ListItems(xxx).Checked = True Then
                    M_OBJCONN.Execute "UPDATE mgm SET agent='" & LvTransfer.ListItems(xxx).SubItems(3) & "', spv_allow = now() WHERE " & _
                                "custid='" & LvTransfer.ListItems(xxx).Text & "'"
                    ' Isi Log, user yang membalikan data - update 2014-07-02
                    M_OBJCONN.Execute "insert into tbllogreview_hst(custid, agentlama, agentbaru,lastupdateuser)values('" + LvTransfer.ListItems(xxx).Text + "','REVIEW','" & LvTransfer.ListItems(xxx).SubItems(3) & "','" + MDIForm1.TxtUsername.Text + "') "
                    ' Hapus log 5x Call diblock - Update 2013-04-25 By Izuddin
                    M_OBJCONN.Execute "DELETE FROM user_phone_log WHERE custid='" & Trim(LvTransfer.ListItems(xxx).Text) & "' "
                End If
            Next xxx
            Call IsiCustidOtomatis
        Else
            MsgBox "Anda belum mencentang ceklist data yang dipilih!!", vbOKOnly
            Exit Sub
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
Dim rsreview As New ADODB.Recordset
Dim ListItem As ListItem

CMDSQL = "select * from tbllogreview_hst where date(input_date) between '" + Format(TDBDate1.Value, "yyyy-mm-dd") + "' "
CMDSQL = CMDSQL + " and '" + Format(TDBDate2.Value, "yyyy-mm-dd") + "'"
Set rsreview = New ADODB.Recordset
ListView1.ListItems.CLEAR
rsreview.CursorLocation = adUseClient
rsreview.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not rsreview.EOF
            Set ListItem = ListView1.ListItems.ADD(, , IIf(IsNull(rsreview!CustId), "", rsreview!CustId))
                            ListItem.SubItems(1) = IIf(IsNull(rsreview!agentlama), "", rsreview!agentlama)
                            ListItem.SubItems(2) = IIf(IsNull(rsreview!agentbaru), "", rsreview!agentbaru)
                            ListItem.SubItems(3) = IIf(IsNull(rsreview!lastupdateuser), "", rsreview!lastupdateuser)
                            ListItem.SubItems(4) = IIf(IsNull(rsreview!input_date), "", rsreview!input_date)
rsreview.MoveNext
Wend
 

End Sub

Private Sub Form_Load()
    'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    Call koneksi
    Call HeaderListTransfer
    Call IsiCustidOtomatis
    TDBDate1.Value = Now
    TDBDate2.Value = Now
End Sub

Private Sub chk_all_Click()
    Dim W As Integer
    If LvTransfer.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvTransfer.ListItems.Count
        If chk_all.Value = 1 Then
            LvTransfer.ListItems(W).Checked = True
        Else
            LvTransfer.ListItems(W).Checked = False
        End If
    Next W
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Rs_list = Nothing
    VIEW_MGMDATA.WindowState = 2
End Sub

Private Sub HeaderListTransfer()
    LvTransfer.ColumnHeaders.ADD , , "Custid", 1000
    LvTransfer.ColumnHeaders.ADD , , "Nama", 3000
    LvTransfer.ColumnHeaders.ADD , , "Telp", 1500
    LvTransfer.ColumnHeaders.ADD , , "Old Agent", 1500
    LvTransfer.ColumnHeaders.ADD , , "Nama Agent", 1500
    
    ListView1.ColumnHeaders.ADD , , "Custid", 2000
    ListView1.ColumnHeaders.ADD , , "agent lama", 1500
    ListView1.ColumnHeaders.ADD , , "agent baru", 1500
    ListView1.ColumnHeaders.ADD , , "LastUserUpdate", 1500
    ListView1.ColumnHeaders.ADD , , "Inputdate", 1500
    
End Sub

Private Sub IsiCustidOtomatis()
    Dim ListItem As ListItem
    
    LvTransfer.ListItems.CLEAR
    If Rs_list.State = 1 Then Rs_list.Close
    ' SET AGENT ASLI di a.agent_asli 19 Agustus 2014
    ' REVISI LAGI
    Rs_list.Open "SELECT distinct(b.custid),a.name,b.agent,b.telp,c.agent as nama_agent FROM mgm a, " & _
                "(SELECT x.custid,x.telp,x.agent FROM tbl_log_acc_review x,(SELECT custid,max(tgl) as Tgl_terakhir " & _
                "FROM tbl_log_acc_review GROUP BY custid) y WHERE x.custid=y.custid AND x.tgl=y.Tgl_terakhir )b,usertbl c " & _
                "WHERE a.custid=b.custid AND b.agent=c.userid AND lower(a.agent) LIKE '%review%' order by b.agent;"

'    Rs_list.Open "SELECT distinct(b.custid),a.name,b.agent,b.telp,c.agent as nama_agent FROM mgm a, " & _
'                "(SELECT custid,agent,telp,max(tgl) FROM tbl_log_acc_review  GROUP BY custid,agent,telp) b,usertbl c WHERE a.custid=b.custid" & _
'                " AND b.agent=c.userid AND lower(a.agent) LIKE '%review%' order by b.agent"
    If Rs_list.RecordCount > 0 Then
        Do Until Rs_list.EOF
            Set ListItem = LvTransfer.ListItems.ADD(, , IIf(IsNull(Rs_list!CustId), "", Rs_list!CustId))
                            ListItem.SubItems(1) = IIf(IsNull(Rs_list!Name), "", Rs_list!Name)
                            ListItem.SubItems(2) = IIf(IsNull(Rs_list!TELP), "", Rs_list!TELP)
                            ListItem.SubItems(3) = IIf(IsNull(Rs_list!AGENT), "", Rs_list!AGENT)
                            ListItem.SubItems(4) = IIf(IsNull(Rs_list!nama_agent), "", Rs_list!nama_agent)
            Rs_list.MoveNext
        Loop
        Label1.Caption = "Rows : " & Rs_list.RecordCount
    Else
        MsgBox "Data customer REVIEW tidak ada / kosong", vbOKOnly + vbInformation, "Info"
        Label1.Caption = "Rows : 0"
    End If
End Sub

Private Sub LvTransfer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvTransfer.SortKey = ColumnHeader.Index - 1
    LvTransfer.Sorted = True
End Sub

Private Sub koneksi()
    Set Rs_list = New ADODB.Recordset
    Rs_list.CursorLocation = adUseClient
    Rs_list.ActiveConnection = M_OBJCONN
    Rs_list.CursorType = adOpenDynamic
    Rs_list.LockType = adLockOptimistic
End Sub

Private Sub LvTransfer_DblClick()
    If LvTransfer.ListItems.Count > 0 Then
        sReminder_CUST_ID = LvTransfer.SelectedItem.Text
        If bAktif_form_customer Then
            Unload FrmCC_Colection
        End If
        bAktif_Cust_Review = True
        FrmCC_Colection.Show vbModal
    End If
End Sub

