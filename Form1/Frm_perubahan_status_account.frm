VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Frm_Cek_status_acc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Perubahan Status Account"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12795
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8835
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   15584
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "STATUS ACCOUNT GLOBAL"
      TabPicture(0)   =   "Frm_perubahan_status_account.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "OptAgent"
      Tab(0).Control(1)=   "OptTL"
      Tab(0).Control(2)=   "CmbUser"
      Tab(0).Control(3)=   "CmdTampilkan"
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(5)=   "LvStsAcc"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "STATUS ACCOUNT PER SESSION"
      TabPicture(1)   =   "Frm_perubahan_status_account.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "LblStartLock"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label7"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "LblEndLock"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label8"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "LblLockby"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label9"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "LblAccLock"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label2"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "LblStsLock"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label10"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "RPT"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "LvLogSession"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "LvLogLock"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "CmdTampilLogSess"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "TxtJmlSessDt"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "TxtJmlBerubah"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "TxtJmlTetap"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "TxtLmKerja"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "CmdReport"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).ControlCount=   24
      Begin VB.CommandButton CmdReport 
         Caption         =   "&Lihat Report"
         Height          =   540
         Left            =   11025
         TabIndex        =   29
         Top             =   1470
         Width           =   1485
      End
      Begin VB.TextBox TxtLmKerja 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         TabIndex        =   28
         Text            =   "0"
         Top             =   8400
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox TxtJmlTetap 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4515
         TabIndex        =   16
         Text            =   "0"
         Top             =   8400
         Width           =   540
      End
      Begin VB.TextBox TxtJmlBerubah 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1995
         TabIndex        =   14
         Text            =   "0"
         Top             =   8400
         Width           =   540
      End
      Begin VB.TextBox TxtJmlSessDt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   11865
         TabIndex        =   12
         Text            =   "0"
         Top             =   8400
         Width           =   540
      End
      Begin VB.CommandButton CmdTampilLogSess 
         Caption         =   "&Tampilkan"
         Height          =   540
         Left            =   11025
         TabIndex        =   9
         Top             =   840
         Width           =   1485
      End
      Begin VB.OptionButton OptAgent 
         Caption         =   "Pilih berdasarkan nama agent"
         Height          =   330
         Left            =   -74790
         TabIndex        =   5
         Top             =   525
         Width           =   2430
      End
      Begin VB.OptionButton OptTL 
         Caption         =   "Pilih berdasarkan kelompok TL"
         Height          =   225
         Left            =   -72165
         TabIndex        =   4
         Top             =   630
         Width           =   3060
      End
      Begin VB.ComboBox CmbUser 
         Height          =   315
         Left            =   -74790
         TabIndex        =   3
         Top             =   1050
         Width           =   2745
      End
      Begin VB.CommandButton CmdTampilkan 
         Caption         =   "&Tampilkan"
         Height          =   435
         Left            =   -71850
         TabIndex        =   2
         Top             =   945
         Width           =   1065
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Keluar"
         Height          =   435
         Left            =   -70695
         TabIndex        =   1
         Top             =   945
         Width           =   1065
      End
      Begin MSComctlLib.ListView LvStsAcc 
         Height          =   7320
         Left            =   -74790
         TabIndex        =   6
         Top             =   1470
         Width           =   12345
         _ExtentX        =   21775
         _ExtentY        =   12912
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
      Begin MSComctlLib.ListView LvLogLock 
         Height          =   1860
         Left            =   105
         TabIndex        =   7
         Top             =   840
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   3281
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
      Begin MSComctlLib.ListView LvLogSession 
         Height          =   4905
         Left            =   105
         TabIndex        =   10
         Top             =   3360
         Width           =   12345
         _ExtentX        =   21775
         _ExtentY        =   8652
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
      Begin Crystal.CrystalReport RPT 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label10 
         Caption         =   "Lama Pengerjaan:"
         Height          =   225
         Left            =   5250
         TabIndex        =   27
         Top             =   8400
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label LblStsLock 
         Caption         =   "[Not Selected]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1575
         TabIndex        =   26
         Top             =   2730
         Width           =   8415
      End
      Begin VB.Label Label2 
         Caption         =   "Status Locked:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   25
         Top             =   2730
         Width           =   2115
      End
      Begin VB.Label LblAccLock 
         Caption         =   "[Not Selected]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10815
         TabIndex        =   24
         Top             =   3045
         Width           =   1800
      End
      Begin VB.Label Label9 
         Caption         =   "Acc.Lock:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9765
         TabIndex        =   23
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label LblLockby 
         Caption         =   "[Not Selected]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8295
         TabIndex        =   22
         Top             =   3045
         Width           =   1800
      End
      Begin VB.Label Label8 
         Caption         =   "Lock By:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7455
         TabIndex        =   21
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label LblEndLock 
         Caption         =   "[Not Selected]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4620
         TabIndex        =   20
         Top             =   3045
         Width           =   2325
      End
      Begin VB.Label Label7 
         Caption         =   "End Lock:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3570
         TabIndex        =   19
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label LblStartLock 
         Caption         =   "[Not Selected]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         TabIndex        =   18
         Top             =   3045
         Width           =   2325
      End
      Begin VB.Label Label6 
         Caption         =   "Start Lock:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   17
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "Status Account Tetap:"
         Height          =   225
         Left            =   2730
         TabIndex        =   15
         Top             =   8400
         Width           =   2010
      End
      Begin VB.Label Label4 
         Caption         =   "Status Account Berubah:"
         Height          =   225
         Left            =   105
         TabIndex        =   13
         Top             =   8400
         Width           =   2010
      End
      Begin VB.Label Label3 
         Caption         =   "Jumlah Data:"
         Height          =   225
         Left            =   10605
         TabIndex        =   11
         Top             =   8400
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Log Session Lock Data:"
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   525
         Width           =   2220
      End
   End
End
Attribute VB_Name = "Frm_Cek_status_acc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdReport_Click()
    WaitSecs (2)
    RPT.Reset
    RPT.Formulas(1) = "@StartLock = totext('" + CStr(LblStartLock.Caption) + "')"
    RPT.Formulas(2) = "@EndLock = totext('" + CStr(LblEndLock.Caption) + "')"
    RPT.Formulas(3) = "@StatusLocked = totext('" + CStr(LblStsLock.Caption) + "')"
    RPT.Formulas(4) = "@LockBy = totext('" + CStr(LblLockby.Caption) + "')"
    RPT.Formulas(5) = "@AccLock = totext('" + CStr(LblAccLock.Caption) + "')"
    RPT.Formulas(6) = "@StatusLocked = totext('" + CStr(LblStsLock.Caption) + "')"
    RPT.Formulas(7) = "@TotalData = totext('" + CStr(TxtJmlSessDt.Text) + "')"
    RPT.Formulas(8) = "@DataBerubah = totext('" + CStr(TxtJmlBerubah.Text) + "')"
    RPT.Formulas(9) = "@DataTetap= totext('" + CStr(TxtJmlTetap.Text) + "')"
    RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptPerformPerSession.rpt"
  Call SHOW_PRN
End Sub

Private Sub CmdTampilkan_Click()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim i As Integer
    
    
    If CmbUser.Text = "" Then
        MsgBox "Pilih userid!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    CMDSQL = "select f_cek_new,f_cekhst,agent,custid,name from mandiri.mgm where "
    If OptAgent.Value Then
        CMDSQL = CMDSQL + "agent='" + Trim(CmbUser.Text) + "'"
    End If
    If OptTL.Value Then
        CMDSQL = CMDSQL + "agent in ("
        CMDSQL = CMDSQL + "select userid from mandiri.usertbl where spvcode='"
        CMDSQL = CMDSQL + CmbUser.Text + "')"
    End If
    CMDSQL = CMDSQL + "   order by agent,f_cek_new asc"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvStsAcc.ListItems.CLEAR
    While Not M_objrs.EOF
       
        Set ListItem = LvStsAcc.ListItems.ADD(, , M_objrs.Bookmark)
            ListItem.SubItems(1) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("name")), "", M_objrs("name"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
            If IsNull(M_objrs("f_cekhst")) = False Then
                fcekhst = Split(M_objrs("f_cekhst"), ">")
                ListItem.SubItems(4) = fcekhst(UBound(fcekhst))
            Else
                 ListItem.SubItems(4) = ""
            End If
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("f_cek_new")), "", M_objrs("f_cek_new"))
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub CmdTampilLogSess_Click()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim m_objrsAccBerubah As ADODB.Recordset
    Dim ListItem As ListItem
    
    Dim CustId As String
    Dim nama As String
    Dim AGENT As String
    Dim fceklalu As String
    Dim tglfceklalu As String
    Dim fceknow As String
    Dim tglfceknow As String
    
    
    If LvLogLock.ListItems.Count = 0 Then
        MsgBox "Tidak ada session yang ditampilkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    LblStsLock.Caption = LvLogLock.SelectedItem.SubItems(6)
    LblEndLock.Caption = LvLogLock.SelectedItem.SubItems(2)
    LblLockby.Caption = LvLogLock.SelectedItem.SubItems(4)
    LblStartLock.Caption = LvLogLock.SelectedItem.SubItems(1)
    LblAccLock.Caption = LvLogLock.SelectedItem.SubItems(3)
    
    CMDSQL = "select *,(endlock-startlock) as Selisih from mandiri.tblperformpersessionlock where idlock='"
    CMDSQL = CMDSQL + Trim(LvLogLock.SelectedItem.SubItems(5)) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    LvLogSession.ListItems.CLEAR
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia! Kemungkinan data ini di release sebelum waktunya, atau agent yang di lock tidak memiliki data tersebut!", vbOKOnly + vbInformation, "Informasi"
    Else
         TxtLmKerja.Text = IIf(IsNull(M_objrs("selisih")), "0", M_objrs("selisih"))
         M_RPTCONN.Execute "delete from RptPerformPerSession "
        While Not M_objrs.EOF
            
            CustId = Trim(IIf(IsNull(M_objrs("custid")), "", M_objrs("custid")))
            nama = Trim(IIf(IsNull(M_objrs("name")), "", M_objrs("name")))
            AGENT = Trim(IIf(IsNull(M_objrs("agent")), "", M_objrs("agent")))
            
            If IsNull(M_objrs("f_ceklalu")) Then
               fceklalu = ""
            Else
                fceklalu = Trim(M_objrs("f_ceklalu"))
            End If
            
            If IsNull(M_objrs("tgl_f_ceklalu")) Then
                tglfceklalu = "null"
            Else
                tglfceklalu = "'" + Format(M_objrs("tgl_f_ceklalu"), "yyyy-mm-dd hh:mm:ss") + "'"
            End If
            
            If IsNull(M_objrs("f_ceksekrg")) Then
                fceknow = ""
            Else
               fceknow = Trim(M_objrs("f_ceksekrg"))
            End If
            
            If IsNull(M_objrs("tgl_f_ceksekrg")) Then
                tglfceknow = "null"
            Else
                tglfceknow = "'" + Format(M_objrs("tgl_f_ceksekrg"), "yyyy-mm-dd hh:mm:ss") + "'"
            End If
            
        
            Set ListItem = LvLogSession.ListItems.ADD(, , M_objrs.Bookmark)
            ListItem.SubItems(1) = CustId
            ListItem.SubItems(2) = Trim(IIf(IsNull(M_objrs("name")), "", M_objrs("name")))
            ListItem.SubItems(3) = AGENT
            ListItem.SubItems(4) = fceklalu '& " | " & tglfceklalu
            ListItem.SubItems(5) = fceknow '& " |" & tglfceknow
            
            'Update ke access buat report
            CMDSQL = "insert into mandiri.RptPerformPerSession (custid,nama,"
            CMDSQL = CMDSQL + "agent,f_ceklalu,tgl_f_ceklalu,f_ceksekrg,tgl_f_ceksekrg) values ('"
            CMDSQL = CMDSQL + CustId + "','"
            CMDSQL = CMDSQL + nama + "','"
            CMDSQL = CMDSQL + AGENT + "','"
            CMDSQL = CMDSQL + fceklalu + "',"
            CMDSQL = CMDSQL + CStr(tglfceklalu) + ",'"
            CMDSQL = CMDSQL + fceknow + "',"
            CMDSQL = CMDSQL + CStr(tglfceknow) + ")"
            M_RPTCONN.Execute CMDSQL
            
       
            M_objrs.MoveNext
        Wend
    End If
    TxtJmlSessDt.Text = M_objrs.RecordCount
    Set M_objrs = Nothing
    
    'Cek jumlah account yang berubah
    CMDSQL = "select * from tblperformpersessionlock where idlock='"
    CMDSQL = CMDSQL + Trim(LvLogLock.SelectedItem.SubItems(5)) + "' and f_ceklalu<>f_ceksekrg "
    Set m_objrsAccBerubah = New ADODB.Recordset
    m_objrsAccBerubah.CursorLocation = adUseClient
    m_objrsAccBerubah.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    TxtJmlBerubah.Text = m_objrsAccBerubah.RecordCount
    TxtJmlTetap.Text = Val(TxtJmlSessDt.Text) - m_objrsAccBerubah.RecordCount
    Set m_objrsAccBerubah = Nothing
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    
    Call HeaderList
    Call IsiAgent
    
    Call HeaderLockLog
    Call IsiLogLock
    Call HeaderLogSess
  
End Sub

Private Sub HeaderList()
    LvStsAcc.ColumnHeaders.ADD 1, , "No.", 500
    LvStsAcc.ColumnHeaders.ADD 2, , "Custid", 2000
    LvStsAcc.ColumnHeaders.ADD 3, , "Name", 3500
    LvStsAcc.ColumnHeaders.ADD 4, , "Agent", 900
    LvStsAcc.ColumnHeaders.ADD 5, , "Status Acc. Lalu", 3000
    LvStsAcc.ColumnHeaders.ADD 6, , "Status Acc. Sekarang", 1500
End Sub
Private Sub HeaderLogSess()
    LvLogSession.ColumnHeaders.ADD 1, , "No.", 500
    LvLogSession.ColumnHeaders.ADD 2, , "Custid", 2000
    LvLogSession.ColumnHeaders.ADD 3, , "Name", 3000
    LvLogSession.ColumnHeaders.ADD 4, , "Agent", 900
    LvLogSession.ColumnHeaders.ADD 5, , "Status Acc. Lalu", 2500
    LvLogSession.ColumnHeaders.ADD 6, , "Status Acc. Sekarang", 2500
End Sub

Private Sub IsiAgent()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    CMDSQL = "select userid  from usertbl where usertype='1' order by userid asc"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    CmbUser.CLEAR
    If M_objrs.RecordCount <> 0 Then
        While Not M_objrs.EOF
            CmbUser.AddItem Trim(M_objrs("userid"))
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub


Private Sub IsiTL()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    CMDSQL = "select spvcode  from usertbl where usertype='6' order by spvcode asc"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    CmbUser.CLEAR
    If M_objrs.RecordCount <> 0 Then
        While Not M_objrs.EOF
            CmbUser.AddItem Trim(M_objrs("spvcode"))
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub




Private Sub LvLogLock_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   LvLogLock.SortKey = ColumnHeader.Index - 1
   LvLogLock.Sorted = True
End Sub

Private Sub LvLogSession_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   LvLogSession.SortKey = ColumnHeader.Index - 1
   LvLogSession.Sorted = True
End Sub



Private Sub LvStsAcc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   LvStsAcc.SortKey = ColumnHeader.Index - 1
   LvStsAcc.Sorted = True
End Sub

Private Sub OptAgent_Click()
    If OptAgent.Value Then
        Call IsiAgent
    Else
        Call IsiTL
    End If
End Sub

Private Sub OptTL_Click()
    If OptTL.Value Then
        Call IsiTL
    Else
        Call IsiAgent
    End If
End Sub

Private Sub HeaderLockLog()

    LvLogLock.ColumnHeaders.ADD 1, , "Date Lock", 2000
    LvLogLock.ColumnHeaders.ADD 2, , "Start Lock", 2000
    LvLogLock.ColumnHeaders.ADD 3, , "End Lock", 2000
    LvLogLock.ColumnHeaders.ADD 4, , "Account Lock", 1500
    LvLogLock.ColumnHeaders.ADD 5, , "Lock By", 1500
    LvLogLock.ColumnHeaders.ADD 6, , "Id", 0
    LvLogLock.ColumnHeaders.ADD 7, , "Status Locked", 4000

End Sub



Private Sub IsiLogLock()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    
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
    
    LvLogLock.ListItems.CLEAR
    
    While Not M_objrs.EOF
        Set ListItem = LvLogLock.ListItems.ADD(, , Format(M_objrs("date_lock"), "dd-mm-yyyy hh:mm:ss"))
            ListItem.SubItems(1) = Format(M_objrs("start_lock"), "dd-mm-yyyy hh:mm:ss")
            ListItem.SubItems(2) = Format(M_objrs("end_lock"), "dd-mm-yyyy hh:mm:ss")
            ListItem.SubItems(3) = Trim(M_objrs("account_lock"))
            ListItem.SubItems(4) = Trim(M_objrs("lock_by"))
            ListItem.SubItems(5) = Trim(M_objrs("id"))
            ListItem.SubItems(6) = Replace(IIf(IsNull(M_objrs("status_lock")), "", M_objrs("status_lock")), "@", "")
        M_objrs.MoveNext
    Wend
    
    
End Sub

Private Sub SHOW_PRN()
    RPT.RetrieveDataFiles
    RPT.WindowLeft = 0
    RPT.WindowTop = 0
    RPT.WindowState = crptMaximized
    RPT.WindowShowPrintBtn = True
    RPT.WindowShowRefreshBtn = True
    RPT.WindowShowSearchBtn = True
    RPT.WindowShowPrintSetupBtn = True
    RPT.WindowControls = True
    RPT.PrintReport
    'RPT.Action = 1
    'RPT.Reset
End Sub

