VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_map_lock_acc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mapping Lock Account"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10245
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Antrian lock account"
      TabPicture(0)   =   "frm_map_lock_acc.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CmdRefreshLock"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmdDelLock"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdAddLock"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LvLockAcc"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Lock account current"
      TabPicture(1)   =   "frm_map_lock_acc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmdRelease"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "TxtStatusLocked"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TxtLockby"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "TxtAccountLock"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "TxtEndLock"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "TxtStartLock"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "TxtDateLock"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label2(4)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label2(3)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label2(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label2(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label2(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label1"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Log lock account"
      TabPicture(2)   =   "frm_map_lock_acc.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "EndDate"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "StartDate"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "LvLockAccLog"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CommandButton Command1 
         Caption         =   "&Report Log Lock"
         Height          =   435
         Left            =   3780
         TabIndex        =   19
         Top             =   420
         Width           =   1800
      End
      Begin VB.CommandButton CmdRelease 
         Caption         =   "&Release.."
         Height          =   435
         Left            =   -66495
         TabIndex        =   17
         Top             =   630
         Width           =   1275
      End
      Begin VB.TextBox TxtStatusLocked 
         Enabled         =   0   'False
         Height          =   330
         Left            =   -73110
         TabIndex        =   16
         Top             =   2205
         Width           =   4845
      End
      Begin VB.TextBox TxtLockby 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73110
         TabIndex        =   15
         Top             =   1890
         Width           =   2430
      End
      Begin VB.TextBox TxtAccountLock 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73110
         TabIndex        =   14
         Top             =   1575
         Width           =   2430
      End
      Begin VB.TextBox TxtEndLock 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73110
         TabIndex        =   13
         Top             =   1260
         Width           =   2430
      End
      Begin VB.TextBox TxtStartLock 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73110
         TabIndex        =   12
         Top             =   945
         Width           =   2430
      End
      Begin VB.TextBox TxtDateLock 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73110
         TabIndex        =   7
         Top             =   630
         Width           =   2430
      End
      Begin VB.CommandButton CmdRefreshLock 
         Caption         =   "&Refresh"
         Height          =   495
         Left            =   -66525
         TabIndex        =   4
         Top             =   1785
         Width           =   1650
      End
      Begin VB.CommandButton CmdDelLock 
         Caption         =   "&Del Schedule lock"
         Height          =   495
         Left            =   -66525
         TabIndex        =   2
         Top             =   1215
         Width           =   1650
      End
      Begin VB.CommandButton CmdAddLock 
         Caption         =   "&Add Schedule Lock"
         Height          =   495
         Left            =   -66495
         TabIndex        =   1
         Top             =   615
         Width           =   1650
      End
      Begin MSComctlLib.ListView LvLockAcc 
         Height          =   3960
         Left            =   -74790
         TabIndex        =   3
         Top             =   630
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   6985
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
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LvLockAccLog 
         Height          =   3645
         Left            =   210
         TabIndex        =   18
         Top             =   945
         Width           =   9720
         _ExtentX        =   17145
         _ExtentY        =   6429
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
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin TDBDate6Ctl.TDBDate StartDate 
         Height          =   315
         Left            =   210
         TabIndex        =   20
         Top             =   525
         Width           =   1560
         _Version        =   65536
         _ExtentX        =   2752
         _ExtentY        =   556
         Calendar        =   "frm_map_lock_acc.frx":0054
         Caption         =   "frm_map_lock_acc.frx":016C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_map_lock_acc.frx":01D8
         Keys            =   "frm_map_lock_acc.frx":01F6
         Spin            =   "frm_map_lock_acc.frx":0254
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   1.12794198814265E-317
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate EndDate 
         Height          =   315
         Left            =   2100
         TabIndex        =   21
         Top             =   525
         Width           =   1560
         _Version        =   65536
         _ExtentX        =   2752
         _ExtentY        =   556
         Calendar        =   "frm_map_lock_acc.frx":027C
         Caption         =   "frm_map_lock_acc.frx":0394
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm_map_lock_acc.frx":0400
         Keys            =   "frm_map_lock_acc.frx":041E
         Spin            =   "frm_map_lock_acc.frx":047C
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   1.12794198814265E-317
         CenturyMode     =   0
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   225
         Left            =   1785
         TabIndex        =   22
         Top             =   525
         Width           =   330
      End
      Begin VB.Label Label2 
         Caption         =   "Status Locked:"
         Height          =   330
         Index           =   4
         Left            =   -74685
         TabIndex        =   11
         Top             =   2205
         Width           =   1170
      End
      Begin VB.Label Label2 
         Caption         =   "Lock by:"
         Height          =   330
         Index           =   3
         Left            =   -74685
         TabIndex        =   10
         Top             =   1890
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Account Lock:"
         Height          =   330
         Index           =   2
         Left            =   -74685
         TabIndex        =   9
         Top             =   1575
         Width           =   1170
      End
      Begin VB.Label Label2 
         Caption         =   "End Lock:"
         Height          =   330
         Index           =   1
         Left            =   -74685
         TabIndex        =   8
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Start Lock:"
         Height          =   330
         Index           =   0
         Left            =   -74685
         TabIndex        =   6
         Top             =   945
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Date Lock:"
         Height          =   330
         Left            =   -74685
         TabIndex        =   5
         Top             =   630
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm_map_lock_acc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAddLock_Click()
    frmlockaccountfromspv.Show 1
End Sub

Private Sub CmdDelLock_Click()
    Dim M_OBJRS As ADODB.Recordset
    Dim CMDSQL As String
    Dim a As String
    
    If LvLockAcc.ListItems.Count <> 0 Then
        
        a = MsgBox("Yakin data akan dihapus?", vbYesNo + vbQuestion, "Informasi")
        If a = vbYes Then
            CMDSQL = "delete from tbltemplockacc where id='"
            CMDSQL = CMDSQL + Trim(LvLockAcc.SelectedItem.SubItems(5)) + "'"
            Set M_OBJRS = New ADODB.Recordset
            M_OBJRS.CursorLocation = adUseClient
            M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            Set M_OBJRS = Nothing
            LvLockAcc.ListItems.Remove LvLockAcc.SelectedItem.Index
        End If
        
    End If
End Sub

Private Sub CmdRefreshLock_Click()
    Call IsiMapLock
End Sub

Private Sub CmdRelease_Click()
    Dim M_OBJRS As ADODB.Recordset
    Dim cmdsqlserver As String
    Dim a As String
    
    If IsNull(TxtDateLock.Text) = True Then
        MsgBox "Tidak ada data yang di release!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Apakah anda yakin data akan di release?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        Exit Sub
    End If
    
    
            'Clear lock data yang sedang berjalan sesuai dengan agent yang di lock
            cmdsqlserver = "update usertbl set dilockoleh='ClearByAutomatic',"
            cmdsqlserver = cmdsqlserver + "lockdarispv=null,lock_entry_lpd=null,fromaccount=null,"
            cmdsqlserver = cmdsqlserver + "lockmarkup=null,lockdarispvbuattl=null"
            'Buat ambil kondisi agent yang sedang di lock
            If Trim(TxtAccountLock.Text) = "ALL" Then
                cmdsqlserver = cmdsqlserver + " "
            ElseIf Left(Trim(TxtAccountLock.Text), 3) = "SPV" Then
                cmdsqlserver = cmdsqlserver + " where spvcode='"
                cmdsqlserver = cmdsqlserver + Trim(TxtAccountLock.Text) + "'"
            Else
                cmdsqlserver = cmdsqlserver + " where userid='"
                cmdsqlserver = cmdsqlserver + Trim(TxtAccountLock.Text) + "'"
            End If
            M_OBJCONN.Execute cmdsqlserver
            
            'Update status pesan ke nilai 1,untuk menampilkan pesan ke agent
            cmdsqlserver = "update usertbl set f_pesanresetauto='1' "
            'Buat mengupdate pesan kondisi agent yang di lock
            If Trim(TxtAccountLock.Text) = "ALL" Then
                cmdsqlserver = cmdsqlserver + " "
            ElseIf Left(Trim(TxtAccountLock.Text), 3) = "SPV" Then
                cmdsqlserver = cmdsqlserver + " where spvcode='"
                cmdsqlserver = cmdsqlserver + Trim(TxtAccountLock.Text) + "'"
            Else
                cmdsqlserver = cmdsqlserver + " where userid='"
                cmdsqlserver = cmdsqlserver + Trim(TxtAccountLock.Text) + "'"
            End If
            M_OBJCONN.Execute cmdsqlserver
            
            'Pindahkan data lock account current ke tabel data log tbltemplockacc_log
            cmdsqlserver = "insert into tbltemplockacc_log select * from tbltemplockacc_current"
            M_OBJCONN.Execute cmdsqlserver
            
            'Hapus data di tabel locktemp current
            cmdsqlserver = "delete from tbltemplockacc_current"
            M_OBJCONN.Execute cmdsqlserver
            
       
    
End Sub

Private Sub Form_Activate()
    CmdRefreshLock_Click
End Sub

Private Sub Form_Load()
    Call HeaderMapLock
    Call IsiMapLock
    Call HeaderMapLockLog
    Call IsiLockLog
    Call IsiLockCurrent
End Sub

Private Sub HeaderMapLock()

    LvLockAcc.ColumnHeaders.ADD 1, , "Date Lock", 2000
    LvLockAcc.ColumnHeaders.ADD 2, , "Start Lock", 2000
    LvLockAcc.ColumnHeaders.ADD 3, , "End Lock", 2000
    LvLockAcc.ColumnHeaders.ADD 4, , "Account Lock", 1500
    LvLockAcc.ColumnHeaders.ADD 5, , "Lock By", 1500
    LvLockAcc.ColumnHeaders.ADD 6, , "Id", 0
    LvLockAcc.ColumnHeaders.ADD 7, , "Status Locked", 4000

End Sub

Private Sub HeaderMapLockLog()

    LvLockAccLog.ColumnHeaders.ADD 1, , "Date Lock", 2000
    LvLockAccLog.ColumnHeaders.ADD 2, , "Start Lock", 2000
    LvLockAccLog.ColumnHeaders.ADD 3, , "End Lock", 2000
    LvLockAccLog.ColumnHeaders.ADD 4, , "Account Lock", 1500
    LvLockAccLog.ColumnHeaders.ADD 5, , "Lock By", 1500
    LvLockAccLog.ColumnHeaders.ADD 6, , "Id", 0
    LvLockAccLog.ColumnHeaders.ADD 7, , "Status Locked", 4000

End Sub

Private Sub IsiMapLock()
    Dim M_OBJRS As ADODB.Recordset
    Dim CMDSQL As String
    Dim listitem As listitem
    
    CMDSQL = "select * from tbltemplockacc order by start_lock asc"
    
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvLockAcc.ListItems.CLEAR
    
    While Not M_OBJRS.EOF
        Set listitem = LvLockAcc.ListItems.ADD(, , Format(M_OBJRS("date_lock"), "dd-mm-yyyy hh:mm:ss"))
            listitem.SubItems(1) = Format(M_OBJRS("start_lock"), "dd-mm-yyyy hh:mm:ss")
            listitem.SubItems(2) = Format(M_OBJRS("end_lock"), "dd-mm-yyyy hh:mm:ss")
            listitem.SubItems(3) = Trim(M_OBJRS("account_lock"))
            listitem.SubItems(4) = Trim(M_OBJRS("lock_by"))
            listitem.SubItems(5) = Trim(M_OBJRS("id"))
            listitem.SubItems(6) = Replace(IIf(IsNull(M_OBJRS("status_lock")), "", M_OBJRS("status_lock")), "@", "")
        M_OBJRS.MoveNext
    Wend
    
    
End Sub

Private Sub IsiLockLog()
    Dim M_OBJRS As ADODB.Recordset
    Dim CMDSQL As String
    Dim listitem As listitem
    
    CMDSQL = "select * from tbltemplockacc_log order by start_lock asc"
    
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvLockAccLog.ListItems.CLEAR
    
    While Not M_OBJRS.EOF
        Set listitem = LvLockAccLog.ListItems.ADD(, , Format(M_OBJRS("date_lock"), "dd-mm-yyyy hh:mm:ss"))
            listitem.SubItems(1) = Format(M_OBJRS("start_lock"), "dd-mm-yyyy hh:mm:ss")
            listitem.SubItems(2) = Format(M_OBJRS("end_lock"), "dd-mm-yyyy hh:mm:ss")
            listitem.SubItems(3) = Trim(M_OBJRS("account_lock"))
            listitem.SubItems(4) = Trim(M_OBJRS("lock_by"))
            listitem.SubItems(5) = Trim(M_OBJRS("id"))
            listitem.SubItems(6) = Replace(IIf(IsNull(M_OBJRS("status_lock")), "", M_OBJRS("status_lock")), "@", "")
        M_OBJRS.MoveNext
    Wend
    
    
End Sub


Private Sub IsiLockCurrent()
    Dim M_OBJRS As ADODB.Recordset
    Dim CMDSQL As String
    Dim listitem As listitem
    On Error Resume Next
    
    CMDSQL = "select * from tbltemplockacc_current"
    
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    
      TxtDateLock.Text = IIf(IsNull(M_OBJRS("date_lock")), "", M_OBJRS("date_lock"))
      TxtStartLock.Text = IIf(IsNull(M_OBJRS("start_lock")), "", Format(M_OBJRS("start_lock"), "dd-mm-yyyy hh:mm:ss"))
      TxtEndLock.Text = IIf(IsNull(M_OBJRS("end_lock")), "", Format(M_OBJRS("end_lock"), "dd-mm-yyyy hh:mm:ss"))
      TxtAccountLock.Text = IIf(IsNull(M_OBJRS("account_lock")), "", Trim(M_OBJRS("account_lock")))
      TxtLockby.Text = IIf(IsNull(M_OBJRS("lock_by")), "", Trim(M_OBJRS("lock_by")))
      TxtStatusLocked.Text = IIf(IsNull(M_OBJRS("status_lock")), "", Trim(M_OBJRS("status_lock")))
       
    Set M_OBJRS = Nothing
    
End Sub



