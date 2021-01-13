VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form frm_list_schedule_tl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Schedule TL"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12135
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   9315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   16431
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Antrian lock account"
      TabPicture(0)   =   "Frm_List_Schedule_TL.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LvLockAcc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdRefreshLock"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdDelLock"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdAddLock"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtJmlDataAntrian"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CmdCekAllLockAcc"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CmdUnCekAllLockAcc"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Lock account current"
      TabPicture(1)   =   "Frm_List_Schedule_TL.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "LvLockAccCurrent"
      Tab(1).Control(2)=   "CmdRelease"
      Tab(1).Control(3)=   "CmdRefreshCurrent"
      Tab(1).Control(4)=   "CmdCekAll"
      Tab(1).Control(5)=   "CmdUncekAll"
      Tab(1).Control(6)=   "TxtJmlDataCurrent"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Log lock account"
      TabPicture(2)   =   "Frm_List_Schedule_TL.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(1)=   "LvLockAccLog"
      Tab(2).Control(2)=   "CmdRefreshLog"
      Tab(2).Control(3)=   "TxtJmlDataLog"
      Tab(2).ControlCount=   4
      Begin VB.CommandButton CmdUnCekAllLockAcc 
         Caption         =   "UnCek All"
         Height          =   495
         Left            =   10380
         TabIndex        =   28
         Top             =   1560
         Width           =   1650
      End
      Begin VB.CommandButton CmdCekAllLockAcc 
         Caption         =   "Cek All"
         Height          =   495
         Left            =   10380
         TabIndex        =   27
         Top             =   1080
         Width           =   1650
      End
      Begin VB.TextBox TxtJmlDataLog 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   -64260
         TabIndex        =   26
         Text            =   "0"
         Top             =   1620
         Width           =   1335
      End
      Begin VB.TextBox TxtJmlDataAntrian 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   10500
         TabIndex        =   24
         Text            =   "0"
         Top             =   3660
         Width           =   1335
      End
      Begin VB.TextBox TxtJmlDataCurrent 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   -64260
         TabIndex        =   22
         Text            =   "0"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton CmdUncekAll 
         Caption         =   "UnCek All"
         Height          =   435
         Left            =   -64260
         TabIndex        =   20
         Top             =   1620
         Width           =   1275
      End
      Begin VB.CommandButton CmdCekAll 
         Caption         =   "Cek All"
         Height          =   435
         Left            =   -64260
         TabIndex        =   19
         Top             =   1140
         Width           =   1275
      End
      Begin VB.CommandButton CmdRefreshLog 
         Caption         =   "&Refresh"
         Height          =   435
         Left            =   -64200
         TabIndex        =   18
         Top             =   660
         Width           =   1275
      End
      Begin VB.CommandButton CmdRefreshCurrent 
         Caption         =   "&Refresh"
         Height          =   435
         Left            =   -64260
         TabIndex        =   17
         Top             =   2700
         Width           =   1275
      End
      Begin VB.CommandButton CmdAddLock 
         Caption         =   "&Add Schedule Lock"
         Height          =   495
         Left            =   10365
         TabIndex        =   4
         Top             =   495
         Width           =   1650
      End
      Begin VB.CommandButton CmdDelLock 
         Caption         =   "&Del Schedule lock"
         Height          =   495
         Left            =   10380
         TabIndex        =   3
         Top             =   2100
         Width           =   1650
      End
      Begin VB.CommandButton CmdRefreshLock 
         Caption         =   "&Refresh"
         Height          =   495
         Left            =   10395
         TabIndex        =   2
         Top             =   2685
         Width           =   1650
      End
      Begin VB.CommandButton CmdRelease 
         Caption         =   "R&elease.."
         Height          =   435
         Left            =   -64260
         TabIndex        =   1
         Top             =   660
         Width           =   1275
      End
      Begin MSComctlLib.ListView LvLockAcc 
         Height          =   8820
         Left            =   180
         TabIndex        =   5
         Top             =   420
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   15558
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
         Height          =   8505
         Left            =   -74790
         TabIndex        =   6
         Top             =   630
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   15002
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
         Left            =   -74790
         TabIndex        =   7
         Top             =   525
         Width           =   1560
         _Version        =   65536
         _ExtentX        =   2752
         _ExtentY        =   556
         Calendar        =   "Frm_List_Schedule_TL.frx":0054
         Caption         =   "Frm_List_Schedule_TL.frx":016C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_List_Schedule_TL.frx":01D8
         Keys            =   "Frm_List_Schedule_TL.frx":01F6
         Spin            =   "Frm_List_Schedule_TL.frx":0254
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
         Value           =   40505
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate EndDate 
         Height          =   315
         Left            =   -72900
         TabIndex        =   8
         Top             =   525
         Width           =   1560
         _Version        =   65536
         _ExtentX        =   2752
         _ExtentY        =   556
         Calendar        =   "Frm_List_Schedule_TL.frx":027C
         Caption         =   "Frm_List_Schedule_TL.frx":0394
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_List_Schedule_TL.frx":0400
         Keys            =   "Frm_List_Schedule_TL.frx":041E
         Spin            =   "Frm_List_Schedule_TL.frx":047C
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
         Value           =   40505
         CenturyMode     =   0
      End
      Begin MSComctlLib.ListView LvLockAccCurrent 
         Height          =   8535
         Left            =   -74790
         TabIndex        =   16
         Top             =   525
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   15055
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Jumlah Data"
         Height          =   255
         Left            =   -64140
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Jumlah Data"
         Height          =   255
         Left            =   10500
         TabIndex        =   23
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Jumlah Data"
         Height          =   255
         Left            =   -64260
         TabIndex        =   21
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Date Lock:"
         Height          =   330
         Index           =   0
         Left            =   -74685
         TabIndex        =   15
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Start Lock:"
         Height          =   330
         Index           =   0
         Left            =   -74685
         TabIndex        =   14
         Top             =   945
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "End Lock:"
         Height          =   330
         Index           =   1
         Left            =   -74685
         TabIndex        =   13
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Account Lock:"
         Height          =   330
         Index           =   2
         Left            =   -74685
         TabIndex        =   12
         Top             =   1575
         Width           =   1170
      End
      Begin VB.Label Label2 
         Caption         =   "Lock by:"
         Height          =   330
         Index           =   3
         Left            =   -74685
         TabIndex        =   11
         Top             =   1890
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Status Locked:"
         Height          =   330
         Index           =   4
         Left            =   -74685
         TabIndex        =   10
         Top             =   2205
         Width           =   1170
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   225
         Left            =   -73215
         TabIndex        =   9
         Top             =   525
         Width           =   330
      End
   End
End
Attribute VB_Name = "frm_list_schedule_tl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private first_load As Boolean

Private Sub CmdAddLock_Click()
    'frmlockaccountfromspv.Show 1
    'frm_add_schedule_tl.Show 1
    FrmLockNew.Show 1
End Sub

Private Sub CmdCekAll_Click()
    Dim z As Integer
    
    If LvLockAccCurrent.ListItems.Count = 0 Then
        MsgBox "Tidak ada data yang tersedia!", vbOKOnly + vbOKOnly, "Informasi"
        Exit Sub
    End If
    
    For z = 1 To LvLockAccCurrent.ListItems.Count
        LvLockAccCurrent.ListItems(z).Checked = True
    Next z
End Sub

Private Sub CmdCekAllLockAcc_Click()
    Dim W As Integer
    
    If LvLockAcc.ListItems.Count = 0 Then
        MsgBox "Tidak ada data yang tersedia!", vbOKOnly + vbOKOnly, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvLockAcc.ListItems.Count
        LvLockAcc.ListItems(W).Checked = True
    Next W
End Sub

Private Sub CmdDelLock_Click()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim a As String
    Dim W As Integer
    
    If LvLockAcc.ListItems.Count <> 0 Then
        
        a = MsgBox("Yakin data akan dihapus?", vbYesNo + vbQuestion, "Informasi")
        If a = vbYes Then
                
            For W = 1 To LvLockAcc.ListItems.Count
                If LvLockAcc.ListItems(W).Checked = True Then
                    CMDSQL = "delete from tbltemplockacc where id='"
                    CMDSQL = CMDSQL + Trim(LvLockAcc.ListItems(W).SubItems(5)) + "'"
                    Set M_objrs = New ADODB.Recordset
                    M_objrs.CursorLocation = adUseClient
                    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    Set M_objrs = Nothing
                    'LvLockAcc.ListItems.Remove LvLockAcc.SelectedItem.Index
                End If
            Next W
            CmdRefreshLock_Click
        End If
        
    End If
    TxtJmlDataAntrian.Text = LvLockAcc.ListItems.Count
End Sub

Private Sub CmdRefreshCurrent_Click()
    Call IsiLockCurrent
    TxtJmlDataCurrent.Text = LvLockAccCurrent.ListItems.Count
End Sub

Private Sub CmdRefreshLock_Click()
    Call IsiMapLock
    TxtJmlDataAntrian.Text = LvLockAcc.ListItems.Count
End Sub

Private Sub CmdRefreshLog_Click()
    Call IsiLockLog
    TxtJmlDataLog.Text = LvLockAccLog.ListItems.Count
End Sub

'@@ 07-03-2011 Release Tanpa perulangan
'Private Sub CmdRelease_Click()
'    Dim M_OBJRS As ADODB.Recordset
'    Dim cmdsqlserver As String
'    Dim a As String
'
'    If LvLockAccCurrent.ListItems.Count = 0 Then
'        MsgBox "Tidak ada data yang di release!", vbOKOnly + vbInformation, "Informasi"
'        Exit Sub
'    End If
'
'    a = MsgBox("Apakah anda yakin data akan di release?", vbYesNo + vbQuestion, "Konfirmasi")
'    If a = vbNo Then
'        Exit Sub
'    End If
'
'           If Trim(LvLockAccCurrent.SelectedItem.SubItems(4)) = "SEPTIAN" Then
'                If Trim(UCase(mdiform1.txtusername.text)) = "WULAN" Or Trim(UCase(mdiform1.txtusername.text)) = "JOKO" Then
'                    MsgBox "Data di blok oleh Pak Septian! Harap hubungi Pak Septian untuk merelease data!", vbOKOnly + vbExclamation, "Peringatan"
'                    Exit Sub
'                End If
'           End If
'
'            'Clear lock data yang sedang berjalan sesuai dengan agent yang di lock
'            cmdsqlserver = "update usertbl set dilockoleh='" + Trim(mdiform1.txtlevel.text) + "',"
'            cmdsqlserver = cmdsqlserver + "lockdarispv=null,lock_entry_lpd=null,fromaccount=null,"
'            cmdsqlserver = cmdsqlserver + "lockmarkup=null,lockdarispvbuattl=null"
'            'Buat ambil kondisi agent yang sedang di lock
'            If Trim(LvLockAccCurrent.SelectedItem.SubItems(3)) = "ALL" Then
'                cmdsqlserver = cmdsqlserver + " where usertype='1' "
'            ElseIf Left(Trim(LvLockAccCurrent.SelectedItem.SubItems(3)), 3) = "SPV" Then
'                cmdsqlserver = cmdsqlserver + " where spvcode='"
'                cmdsqlserver = cmdsqlserver + Trim(LvLockAccCurrent.SelectedItem.SubItems(3)) + "'"
'            Else
'                cmdsqlserver = cmdsqlserver + " where userid='"
'                cmdsqlserver = cmdsqlserver + Trim(LvLockAccCurrent.SelectedItem.SubItems(3)) + "'"
'            End If
'            M_OBJCONN.Execute cmdsqlserver
'
'            'Update status pesan ke nilai 1,untuk menampilkan pesan ke agent
'            cmdsqlserver = "update usertbl set f_pesanresetauto='1',f_idsessend=null,f_pesanlockauto=null,f_idsessstart=null "
'            'Buat mengupdate pesan kondisi agent yang di lock
'            If Trim(LvLockAccCurrent.SelectedItem.SubItems(3)) = "ALL" Then
'                cmdsqlserver = cmdsqlserver + " where usertype='1' "
'            ElseIf Left(Trim(LvLockAccCurrent.SelectedItem.SubItems(3)), 3) = "SPV" Then
'                cmdsqlserver = cmdsqlserver + " where spvcode='"
'                cmdsqlserver = cmdsqlserver + Trim(LvLockAccCurrent.SelectedItem.SubItems(3)) + "'"
'            Else
'                cmdsqlserver = cmdsqlserver + " where userid='"
'                cmdsqlserver = cmdsqlserver + Trim(LvLockAccCurrent.SelectedItem.SubItems(3)) + "'"
'            End If
'            M_OBJCONN.Execute cmdsqlserver
'
'
'                            'Clossing Session
'                            Dim UpdateDtCloseSession As String
'                            Dim m_ObjWktSrv As ADODB.Recordset
'                            Dim CmdsqlWktSrv As String
'                            Dim WaktuServer As Date
'
'                            CmdsqlWktSrv = "select now()"
'                            Set m_ObjWktSrv = New ADODB.Recordset
'                            m_ObjWktSrv.Open CmdsqlWktSrv, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'                            WaktuServer = Format(m_ObjWktSrv(0), "yyyy-mm-dd hh:mm:ss")
'                            Set m_ObjWktSrv = Nothing
'
'                            UpdateDtCloseSession = "update tblperformpersessionlock set f_ceksekrg=a.f_cek_akhir ,"
'                            UpdateDtCloseSession = UpdateDtCloseSession + " tgl_f_ceksekrg=a.tgl_akhir,endlock='" + CStr(Format(WaktuServer, "yyyy-mm-dd hh:mm:ss")) + "' from "
'                            UpdateDtCloseSession = UpdateDtCloseSession + " (select mgm.custid as custid_mgm,mgm.agent as agent_mgm,"
'                            UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.custid as custid_lock,"
'                            UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.agent as agent_lock,"
'                            UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.idlock as id_lock,"
'                            UpdateDtCloseSession = UpdateDtCloseSession + " mgm.f_cek_new as f_cek_akhir, mgm.tglcall as tgl_akhir"
'                            UpdateDtCloseSession = UpdateDtCloseSession + " from tblperformpersessionlock inner join mgm "
'                            UpdateDtCloseSession = UpdateDtCloseSession + " on mgm.custid=tblperformpersessionlock.custid "
'                            UpdateDtCloseSession = UpdateDtCloseSession + " and mgm.agent=tblperformpersessionlock.agent) as a "
'                            UpdateDtCloseSession = UpdateDtCloseSession + " where tblperformpersessionlock.custid=a.custid_mgm and tblperformpersessionlock.agent=a.agent_mgm and a.id_lock='"
'                            UpdateDtCloseSession = UpdateDtCloseSession + Trim(LvLockAccCurrent.SelectedItem.SubItems(5)) + "'"
'                            M_OBJCONN.Execute UpdateDtCloseSession
'                            'Akhir dari closing session
'
'
'            'Pindahkan data lock account current ke tabel data log tbltemplockacc_log
'            cmdsqlserver = "insert into tbltemplockacc_log select * from tbltemplockacc_current where "
'            cmdsqlserver = cmdsqlserver + " id='"
'            cmdsqlserver = cmdsqlserver + Trim(LvLockAccCurrent.SelectedItem.SubItems(5)) + "'"
'            M_OBJCONN.Execute cmdsqlserver
'
'            'Hapus data di tabel locktemp current
'            cmdsqlserver = "delete from tbltemplockacc_current where id='"
'            cmdsqlserver = cmdsqlserver + Trim(LvLockAccCurrent.SelectedItem.SubItems(5)) + "'"
'            M_OBJCONN.Execute cmdsqlserver
'
'            LvLockAccCurrent.ListItems.Remove LvLockAccCurrent.SelectedItem.Index
'            MsgBox "Lock data berhasil di release!", vbOKOnly + vbInformation, "Informasi"
'
'
'End Sub

'@@ 07-03-2011 Release dengan perulangan
Private Sub CmdRelease_Click()
    Dim M_objrs As ADODB.Recordset
    Dim cmdsqlserver As String
    Dim a As String
    Dim F As Integer
    
    If LvLockAccCurrent.ListItems.Count = 0 Then
        MsgBox "Tidak ada data yang di release!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Apakah anda yakin data akan di release?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        Exit Sub
    End If
                
    For F = 1 To LvLockAccCurrent.ListItems.Count
        If LvLockAccCurrent.ListItems(F).Checked = True Then
           
           If Trim(UCase(LvLockAccCurrent.ListItems(F).SubItems(4))) = "SEPTIAN" Then
                If Trim(UCase(MDIForm1.TxtUsername.Text)) = "WULAN" Or Trim(UCase(MDIForm1.TxtUsername.Text)) = "JOKO" Then
                    MsgBox "Data di blok oleh Pak Septian! Harap hubungi Pak Septian untuk merelease data!", vbOKOnly + vbExclamation, "Peringatan"
                    Exit Sub
                End If
           End If
    
            'Clear lock data yang sedang berjalan sesuai dengan agent yang di lock
            cmdsqlserver = "update usertbl set dilockoleh='" + Trim(MDIForm1.txtlevel.Text) + "',"
            cmdsqlserver = cmdsqlserver + "lockdarispv=null,lock_entry_lpd=null,fromaccount=null,"
            cmdsqlserver = cmdsqlserver + "lockmarkup=null,lockdarispvbuattl=null,lockpayment=null "
            'Buat ambil kondisi agent yang sedang di lock
            If Trim(LvLockAccCurrent.ListItems(F).SubItems(3)) = "ALL" Then
                cmdsqlserver = cmdsqlserver + " where usertype='1' "
            ElseIf Left(Trim(LvLockAccCurrent.ListItems(F).SubItems(3)), 3) = "SPV" Then
                cmdsqlserver = cmdsqlserver + " where spvcode='"
                cmdsqlserver = cmdsqlserver + Trim(LvLockAccCurrent.ListItems(F).SubItems(3)) + "'"
            Else
                cmdsqlserver = cmdsqlserver + " where userid='"
                cmdsqlserver = cmdsqlserver + Trim(LvLockAccCurrent.ListItems(F).SubItems(3)) + "'"
            End If
            M_OBJCONN.Execute cmdsqlserver
            
            'Update status pesan ke nilai 1,untuk menampilkan pesan ke agent
            cmdsqlserver = "update usertbl set f_pesanresetauto='1',f_idsessend=null,f_pesanlockauto=null,f_idsessstart=null "
            'Buat mengupdate pesan kondisi agent yang di lock
            If Trim(LvLockAccCurrent.ListItems(F).SubItems(3)) = "ALL" Then
                cmdsqlserver = cmdsqlserver + " where usertype='1' "
            ElseIf Left(Trim(LvLockAccCurrent.ListItems(F).SubItems(3)), 3) = "SPV" Then
                cmdsqlserver = cmdsqlserver + " where spvcode='"
                cmdsqlserver = cmdsqlserver + Trim(LvLockAccCurrent.ListItems(F).SubItems(3)) + "'"
            Else
                cmdsqlserver = cmdsqlserver + " where userid='"
                cmdsqlserver = cmdsqlserver + Trim(LvLockAccCurrent.ListItems(F).SubItems(3)) + "'"
            End If
            M_OBJCONN.Execute cmdsqlserver
            
            
                            'Clossing Session
                            Dim UpdateDtCloseSession As String
                            Dim m_ObjWktSrv As ADODB.Recordset
                            Dim CmdsqlWktSrv As String
                            Dim WaktuServer As Date
                            
                            CmdsqlWktSrv = "select now()"
                            Set m_ObjWktSrv = New ADODB.Recordset
                            m_ObjWktSrv.Open CmdsqlWktSrv, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                            WaktuServer = Format(m_ObjWktSrv(0), "yyyy-mm-dd hh:mm:ss")
                            Set m_ObjWktSrv = Nothing
                            
                            UpdateDtCloseSession = "update tblperformpersessionlock set f_ceksekrg=a.f_cek_akhir ,"
                            UpdateDtCloseSession = UpdateDtCloseSession + " tgl_f_ceksekrg=a.tgl_akhir,endlock='" + CStr(Format(WaktuServer, "yyyy-mm-dd hh:mm:ss")) + "' from "
                            UpdateDtCloseSession = UpdateDtCloseSession + " (select mgm.custid as custid_mgm,mgm.agent as agent_mgm,"
                            UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.custid as custid_lock,"
                            UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.agent as agent_lock,"
                            UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.idlock as id_lock,"
                            UpdateDtCloseSession = UpdateDtCloseSession + " mgm.f_cek_new as f_cek_akhir, mgm.tglcall as tgl_akhir"
                            UpdateDtCloseSession = UpdateDtCloseSession + " from tblperformpersessionlock inner join mgm "
                            UpdateDtCloseSession = UpdateDtCloseSession + " on mgm.custid=tblperformpersessionlock.custid "
                            UpdateDtCloseSession = UpdateDtCloseSession + " and mgm.agent=tblperformpersessionlock.agent) as a "
                            UpdateDtCloseSession = UpdateDtCloseSession + " where tblperformpersessionlock.custid=a.custid_mgm and tblperformpersessionlock.agent=a.agent_mgm and a.id_lock='"
                            UpdateDtCloseSession = UpdateDtCloseSession + Trim(LvLockAccCurrent.ListItems(F).SubItems(5)) + "'"
                            M_OBJCONN.Execute UpdateDtCloseSession
                            'Akhir dari closing session
            
            
            'Pindahkan data lock account current ke tabel data log tbltemplockacc_log
            cmdsqlserver = "insert into tbltemplockacc_log select * from tbltemplockacc_current where "
            cmdsqlserver = cmdsqlserver + " id='"
            cmdsqlserver = cmdsqlserver + Trim(LvLockAccCurrent.ListItems(F).SubItems(5)) + "'"
            M_OBJCONN.Execute cmdsqlserver
            
            'Hapus data di tabel locktemp current
            cmdsqlserver = "delete from tbltemplockacc_current where id='"
            cmdsqlserver = cmdsqlserver + Trim(LvLockAccCurrent.ListItems(F).SubItems(5)) + "'"
            M_OBJCONN.Execute cmdsqlserver
            
            'LvLockAccCurrent.ListItems(f).Remove LvLockAccCurrent.SelectedItem.Index
            'LvLockAccCurrent.ListItems.Remove LvLockAccCurrent.ListItems(F).Index
        End If
  Next F
            CmdRefreshCurrent_Click
            MsgBox "Lock data berhasil di release!", vbOKOnly + vbInformation, "Informasi"
            
       
End Sub

Private Sub CmdUnCekAll_Click()
       Dim z As Integer
    
    If LvLockAccCurrent.ListItems.Count = 0 Then
        MsgBox "Tidak ada data yang tersedia!", vbOKOnly + vbOKOnly, "Informasi"
        Exit Sub
    End If
    
    For z = 1 To LvLockAccCurrent.ListItems.Count
        LvLockAccCurrent.ListItems(z).Checked = False
    Next z
End Sub

Private Sub CmdUnCekAllLockAcc_Click()
    Dim K As Integer
    
    If LvLockAcc.ListItems.Count = 0 Then
        MsgBox "Tidak ada data yang tersedia!", vbOKOnly + vbOKOnly, "Informasi"
        Exit Sub
    End If
    
    For K = 1 To LvLockAcc.ListItems.Count
        LvLockAcc.ListItems(K).Checked = False
    Next K
End Sub

Private Sub Form_Activate()
    If Not first_load Then
        CmdRefreshLock_Click
    End If
    first_load = False
End Sub

Private Sub Form_Load()
    first_load = True
    Call HeaderMapLock
    Call IsiMapLock
    Call HeaderMapLockLog
    Call IsiLockLog
    Call HeaderCurrentLock
    Call IsiLockCurrent
    TxtJmlDataCurrent.Text = LvLockAccCurrent.ListItems.Count
    TxtJmlDataAntrian.Text = LvLockAcc.ListItems.Count
    TxtJmlDataLog.Text = LvLockAccLog.ListItems.Count
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
Private Sub HeaderCurrentLock()

    LvLockAccCurrent.ColumnHeaders.ADD 1, , "Date Lock", 2000
    LvLockAccCurrent.ColumnHeaders.ADD 2, , "Start Lock", 2000
    LvLockAccCurrent.ColumnHeaders.ADD 3, , "End Lock", 2000
    LvLockAccCurrent.ColumnHeaders.ADD 4, , "Account Lock", 1500
    LvLockAccCurrent.ColumnHeaders.ADD 5, , "Lock By", 1500
    LvLockAccCurrent.ColumnHeaders.ADD 6, , "Id", 0
    LvLockAccCurrent.ColumnHeaders.ADD 7, , "Status Locked", 4000

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
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim listitem As listitem
    
    '@@ 11-11-10 jika yang loginnya tl
    If Left(Trim(MDIForm1.TxtUsername.Text), 2) = "TL" Then
        CMDSQL = "select * from tbltemplockacc where lock_by='"
        CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.Text) + "' order by start_lock asc"
    Else
        CMDSQL = "select * from tbltemplockacc order by start_lock asc"
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvLockAcc.ListItems.CLEAR
    
    While Not M_objrs.EOF
        Set listitem = LvLockAcc.ListItems.ADD(, , Format(M_objrs("date_lock"), "dd-mm-yyyy hh:mm:ss"))
            listitem.SubItems(1) = Format(M_objrs("start_lock"), "dd-mm-yyyy hh:mm:ss")
            listitem.SubItems(2) = Format(M_objrs("end_lock"), "dd-mm-yyyy hh:mm:ss")
            listitem.SubItems(3) = Trim(M_objrs("account_lock"))
            listitem.SubItems(4) = Trim(M_objrs("lock_by"))
            listitem.SubItems(5) = Trim(M_objrs("id"))
            listitem.SubItems(6) = Replace(IIf(IsNull(M_objrs("status_lock")), "", M_objrs("status_lock")), "@", "")
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub IsiLockLog()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim listitem As listitem
    
    '@@ 11-11-10 jika yang loginnya tl
    If Left(Trim(MDIForm1.TxtUsername.Text), 2) = "TL" Then
        CMDSQL = "select * from tbltemplockacc_log where lock_by='"
        CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.Text) + "' order by start_lock desc"
    Else
        CMDSQL = "select * from tbltemplockacc_log order by start_lock desc"
    End If
        
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvLockAccLog.ListItems.CLEAR
    
    While Not M_objrs.EOF
        Set listitem = LvLockAccLog.ListItems.ADD(, , Format(M_objrs("date_lock"), "dd-mm-yyyy hh:mm:ss"))
            listitem.SubItems(1) = Format(M_objrs("start_lock"), "dd-mm-yyyy hh:mm:ss")
            listitem.SubItems(2) = Format(M_objrs("end_lock"), "dd-mm-yyyy hh:mm:ss")
            listitem.SubItems(3) = Trim(M_objrs("account_lock"))
            listitem.SubItems(4) = Trim(M_objrs("lock_by"))
            listitem.SubItems(5) = Trim(M_objrs("id"))
            listitem.SubItems(6) = Replace(IIf(IsNull(M_objrs("status_lock")), "", M_objrs("status_lock")), "@", "")
        M_objrs.MoveNext
    Wend
    
    
End Sub


Private Sub IsiLockCurrent()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim listitem As listitem
    
    '@@ 11-11-10 jika yang loginnya tl
    If Left(Trim(MDIForm1.TxtUsername.Text), 2) = "TL" Then
        CMDSQL = "select * from tbltemplockacc_current where lock_by='"
        CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.Text) + "' order by account_lock,start_lock asc"
    Else
        CMDSQL = "select * from tbltemplockacc_current order by account_lock,start_lock asc"
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvLockAccCurrent.ListItems.CLEAR
    
    While Not M_objrs.EOF
        Set listitem = LvLockAccCurrent.ListItems.ADD(, , Format(M_objrs("date_lock"), "dd-mm-yyyy hh:mm:ss"))
            listitem.SubItems(1) = Format(M_objrs("start_lock"), "dd-mm-yyyy hh:mm:ss")
            listitem.SubItems(2) = Format(M_objrs("end_lock"), "dd-mm-yyyy hh:mm:ss")
            listitem.SubItems(3) = Trim(M_objrs("account_lock"))
            listitem.SubItems(4) = Trim(M_objrs("lock_by"))
            listitem.SubItems(5) = Trim(M_objrs("id"))
            listitem.SubItems(6) = Replace(IIf(IsNull(M_objrs("status_lock")), "", M_objrs("status_lock")), "@", "")
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    
End Sub




