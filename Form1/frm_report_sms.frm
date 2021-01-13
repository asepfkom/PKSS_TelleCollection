VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm_report_sms 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report SMS"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   Icon            =   "frm_report_sms.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtTarifOpLain 
      Height          =   285
      Left            =   3660
      TabIndex        =   9
      Text            =   "150"
      Top             =   1620
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6750
      TabIndex        =   6
      Top             =   1590
      Width           =   1125
   End
   Begin VB.CommandButton CmdPreview 
      Caption         =   "&Preview"
      Height          =   375
      Left            =   5580
      TabIndex        =   5
      Top             =   1620
      Width           =   1125
   End
   Begin Crystal.CrystalReport RPT 
      Left            =   3450
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   225
      Left            =   3510
      TabIndex        =   0
      Top             =   1170
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin TDBDate6Ctl.TDBDate Tgl2 
      Height          =   315
      Left            =   6390
      TabIndex        =   1
      Top             =   600
      Width           =   1440
      _Version        =   65536
      _ExtentX        =   2540
      _ExtentY        =   556
      Calendar        =   "frm_report_sms.frx":058A
      Caption         =   "frm_report_sms.frx":06A2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_report_sms.frx":070E
      Keys            =   "frm_report_sms.frx":072C
      Spin            =   "frm_report_sms.frx":078A
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
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
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate Tgl1 
      Height          =   315
      Left            =   4500
      TabIndex        =   2
      Top             =   600
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "frm_report_sms.frx":07B2
      Caption         =   "frm_report_sms.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_report_sms.frx":0936
      Keys            =   "frm_report_sms.frx":0954
      Spin            =   "frm_report_sms.frx":09B2
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
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
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin MSComctlLib.ListView LvReportSMS 
      Height          =   1875
      Left            =   60
      TabIndex        =   7
      Top             =   30
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   3307
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
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
   Begin VB.Label LblReport 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4140
      TabIndex        =   8
      Top             =   150
      Width           =   3165
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "To"
      Height          =   300
      Index           =   0
      Left            =   5880
      TabIndex        =   4
      Top             =   630
      Width           =   390
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Date :"
      Height          =   300
      Index           =   5
      Left            =   3510
      TabIndex        =   3
      Top             =   600
      Width           =   930
   End
End
Attribute VB_Name = "frm_report_sms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
    Me.Hide
End Sub

Private Sub CmdPreview_Click()
    
    If Tgl1.Text = "__-___-____" Or Tgl2.Text = "__-___-____" Then
        Tgl1.Value = "01-05-2010"
        Tgl2.Value = "01-05-2020"
    End If
    
    Select Case LvReportSMS.SelectedItem.Text
    
    
    
    Case 1
        Call ambil_data_tarif
        Call ambil_data_lainnya
    
        WaitSecs (2)
        RPT.Reset
        RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.TxtUsername.Text) + "')"
        RPT.Formulas(2) = "@TglShow = totext('" + CStr(Tgl1.Text) + "')"
        RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(Tgl2.Text) + "')"
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ReportGlobalSMS.rpt"
        Call SHOW_PRN
    
    Case 2
        Call ambil_data_tarif
        Call ambil_data_lainnya
    
        WaitSecs (2)
        RPT.Reset
        RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.TxtUsername.Text) + "')"
        RPT.Formulas(2) = "@TglShow = totext('" + CStr(Tgl1.Text) + "')"
        RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(Tgl2.Text) + "')"
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ReportDetailSMS.rpt"
        Call SHOW_PRN
    Case 3
       Call update_custid
       Call ambil_data_inbox
       WaitSecs (2)
       RPT.Reset
       RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.TxtUsername.Text) + "')"
       RPT.Formulas(2) = "@TglShow = totext('" + CStr(Tgl1.Text) + "')"
       RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(Tgl2.Text) + "')"
       RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ReportDetailSMSInbox.rpt"
       Call SHOW_PRN
       
    'Report Reject SMS
    Case 4
        Call input_reject_sms
        WaitSecs (2)
        RPT.Reset
        RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.TxtUsername.Text) + "')"
        RPT.Formulas(2) = "@TglShow = totext('" + CStr(Tgl1.Text) + "')"
        RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(Tgl2.Text) + "')"
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ReportSmSReject.rpt"
        Call SHOW_PRN
    
    'Report sentitems card yang ok!
    Case 5
        Call tarif_card_ok
        Call no_tarif_card_ok
        WaitSecs (2)
        RPT.Reset
        RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.TxtUsername.Text) + "')"
        RPT.Formulas(2) = "@TglShow = totext('" + CStr(Tgl1.Text) + "')"
        RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(Tgl2.Text) + "')"
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ReportGlobalSMS.rpt"
        Call SHOW_PRN
    'Report sentitems card yang error!
    Case 6
        Call tarif_card_error
        Call no_tarif_card_error
        WaitSecs (2)
        RPT.Reset
        RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.TxtUsername.Text) + "')"
        RPT.Formulas(2) = "@TglShow = totext('" + CStr(Tgl1.Text) + "')"
        RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(Tgl2.Text) + "')"
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ReportGlobalSMS.rpt"
        Call SHOW_PRN
    
    'Report sentitems aw yang ok
    Case 7
        Call tarif_aw_ok
        Call no_tarif_aw_ok
        WaitSecs (2)
        RPT.Reset
        RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.TxtUsername.Text) + "')"
        RPT.Formulas(2) = "@TglShow = totext('" + CStr(Tgl1.Text) + "')"
        RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(Tgl2.Text) + "')"
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ReportGlobalSMS.rpt"
        Call SHOW_PRN
    'Report sentitems aw error
    Case 8
        Call tarif_aw_error
        Call no_tarif_aw_error
        WaitSecs (2)
        RPT.Reset
        RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.TxtUsername.Text) + "')"
        RPT.Formulas(2) = "@TglShow = totext('" + CStr(Tgl1.Text) + "')"
        RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(Tgl2.Text) + "')"
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ReportGlobalSMS.rpt"
        Call SHOW_PRN
    'Report sentitems pil ok
    Case 9
        Call tarif_pil_ok
        Call no_tarif_pil_ok
        WaitSecs (2)
        RPT.Reset
        RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.TxtUsername.Text) + "')"
        RPT.Formulas(2) = "@TglShow = totext('" + CStr(Tgl1.Text) + "')"
        RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(Tgl2.Text) + "')"
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ReportGlobalSMS.rpt"
        Call SHOW_PRN
    'Report sentitems pil error
    Case 10
        Call tarif_pil_error
        Call no_tarif_pil_error
        WaitSecs (2)
        RPT.Reset
        RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.TxtUsername.Text) + "')"
        RPT.Formulas(2) = "@TglShow = totext('" + CStr(Tgl1.Text) + "')"
        RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(Tgl2.Text) + "')"
        RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\ReportGlobalSMS.rpt"
        Call SHOW_PRN
    End Select
End Sub
Private Sub ambil_data_tarif()
    Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update As String
    Dim CustId As String
    Dim agent As String
    
    M_RPTCONN.Execute "delete from reportsmssentitems"
    
    'Ini buat yang all
'    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,tbl_tarif.operator,tbl_tarif.biaya,"
'    CMDSQL = CMDSQL + "sentitems.sendingdatetime,sentitems.tarif,sentitems.creatorid from sentitems,tbl_tarif where date(sentitems.sendingdatetime) between '"
'    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
'    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
'    CMDSQL = CMDSQL + "substring(sentitems.destinationnumber,1,4)=tbl_tarif.no and "
'    CMDSQL = CMDSQL + "sentitems.creatorid is not null"

    'Ini buat yang card yang ok!
    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,tbl_tarif.operator,tbl_tarif.biaya,"
    CMDSQL = CMDSQL + "sentitems.sendingdatetime,sentitems.tarif,sentitems.creatorid from sentitems,tbl_tarif where date(sentitems.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(sentitems.destinationnumber,1,4)=tbl_tarif.no and "
    CMDSQL = CMDSQL + "sentitems.creatorid is not null and sentitems.creatorid  not like 'AW%' and "
    CMDSQL = CMDSQL + "sentitems.status<>'SendingError' and sentitems.creatorid not like 'PIL%' "
    CMDSQL = CMDSQL + "sentitems.creatorid not like '%BlastExcelPil%'"
    
'    'Ini buat yang AW yang ok!
'    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,tbl_tarif.operator,tbl_tarif.biaya,"
'    CMDSQL = CMDSQL + "sentitems.sendingdatetime,sentitems.tarif,sentitems.creatorid from sentitems,tbl_tarif where date(sentitems.sendingdatetime) between '"
'    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
'    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
'    CMDSQL = CMDSQL + "substring(sentitems.destinationnumber,1,4)=tbl_tarif.no and "
'    CMDSQL = CMDSQL + "sentitems.creatorid is not null and sentitems.creatorid  like 'AW%'"

' 'Ini buat yang AW yang senderror
'    CMDSQL = "select sentitems_senderror.destinationnumber,sentitems_senderror.textdecoded,tbl_tarif.operator,tbl_tarif.biaya,"
'    CMDSQL = CMDSQL + "sentitems_senderror.sendingdatetime,sentitems_senderror.tarif,sentitems_senderror.creatorid from sentitems_senderror,tbl_tarif where date(sentitems_senderror.sendingdatetime) between '"
'    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
'    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
'    CMDSQL = CMDSQL + "substring(sentitems_senderror.destinationnumber,1,4)=tbl_tarif.no and "
'    CMDSQL = CMDSQL + "sentitems_senderror.creatorid is not null and sentitems_senderror.creatorid  like 'AW%'"
    
    'Ini buat yang Card yang senderror
'    CMDSQL = "select sentitems_senderror.destinationnumber,sentitems_senderror.textdecoded,tbl_tarif.operator,tbl_tarif.biaya,"
'    CMDSQL = CMDSQL + "sentitems_senderror.sendingdatetime,sentitems_senderror.tarif,sentitems_senderror.creatorid from sentitems_senderror,tbl_tarif where date(sentitems_senderror.sendingdatetime) between '"
'    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
'    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
'    CMDSQL = CMDSQL + "substring(sentitems_senderror.destinationnumber,1,4)=tbl_tarif.no and "
'    CMDSQL = CMDSQL + "sentitems_senderror.creatorid is not null and "
'    CMDSQL = CMDSQL + "sentitems_senderror.creatorid  not like 'AW%' "

    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    PB1.Max = M_objrs.RecordCount
    
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
            
        End If
       
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("operator")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
        'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("biaya"))) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("biaya"))) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        M_objrs.MoveNext
    Wend
    
End Sub

Private Sub ambil_data_lainnya()
    Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update
    Dim CustId As String
    Dim agent As String
    
    'Ini buat yang all
'    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,sendingdatetime,tarif,creatorid "
'    CMDSQL = CMDSQL + "from sentitems where date(sentitems.sendingdatetime) between '"
'    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
'    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
'    CMDSQL = CMDSQL + "substring(destinationnumber,1,4) not in (select no from tbl_tarif) and "
'    CMDSQL = CMDSQL + "(sentitems.creatorid is not null or sentitems.creatorid <>'')"
    
    'Ini buat yang card
'    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,sendingdatetime,tarif,creatorid "
'    CMDSQL = CMDSQL + "from sentitems where date(sentitems.sendingdatetime) between '"
'    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
'    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
'    CMDSQL = CMDSQL + "substring(destinationnumber,1,4) not in (select no from tbl_tarif) and "
'    CMDSQL = CMDSQL + "(sentitems.creatorid is not null or sentitems.creatorid <>'') and sentitems.creatorid not like 'AW%' and sentitems.status<>'SendingError'"

    'Ini buat yang AW
'    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,sendingdatetime,tarif,creatorid "
'    CMDSQL = CMDSQL + "from sentitems where date(sentitems.sendingdatetime) between '"
'    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
'    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
'    CMDSQL = CMDSQL + "substring(destinationnumber,1,4) not in (select no from tbl_tarif) and "
'    CMDSQL = CMDSQL + "(sentitems.creatorid is not null or sentitems.creatorid <>'') and sentitems.creatorid  like 'AW%'"

'    'Ini buat yang AW senderror
    CMDSQL = "select sentitems_senderror.destinationnumber,sentitems_senderror.textdecoded,sendingdatetime,tarif,creatorid "
    CMDSQL = CMDSQL + "from sentitems_senderror where date(sentitems_senderror.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(destinationnumber,1,4) not in (select no from tbl_tarif) and "
    CMDSQL = CMDSQL + "(sentitems_senderror.creatorid is not null or sentitems_senderror.creatorid <>'') and sentitems_senderror.creatorid  like 'AW%'"

  'Ini buat yang Card senderror
'    CMDSQL = "select sentitems_senderror.destinationnumber,sentitems_senderror.textdecoded,sendingdatetime,tarif,creatorid "
'    CMDSQL = CMDSQL + "from sentitems_senderror where date(sentitems_senderror.sendingdatetime) between '"
'    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
'    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
'    CMDSQL = CMDSQL + "substring(destinationnumber,1,4) not in (select no from tbl_tarif) and "
'    CMDSQL = CMDSQL + "(sentitems_senderror.creatorid is not null or sentitems_senderror.creatorid <>'') and sentitems_senderror.creatorid not like 'AW%'"

    
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
        End If
    
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim("LAINNYA") + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
         'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(TxtTarifOpLain.Text)) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(TxtTarifOpLain.Text)) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        
        
        M_objrs.MoveNext
    Wend
    
End Sub

Private Sub ambil_data_inbox()
    Dim CMDSQL As String
    Dim cmdsql_insert As String
    Dim cmdsql_cari As String
    Dim cmdsql_update As String
    Dim M_objrs As ADODB.Recordset
    Dim m_objrs_cari As ADODB.Recordset
    Dim m_objrs_update As ADODB.Recordset
    Dim isi As String
    
    
    CMDSQL = "select * from inbox where date(receivingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "'"
    

    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        isi = Replace(M_objrs("textdecoded"), "'", "''")
        
        cmdsql_insert = "insert into ReportSmsInbox (no_telepon,"
        cmdsql_insert = cmdsql_insert + "detail_sms,tgl_terima,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("sendernumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(isi) + "','"
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("receivingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
       
            
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

Private Sub Form_Load()
    Call header_report
End Sub
Private Sub header_report()
    Dim listitem As listitem
    
    LvReportSMS.ColumnHeaders.ADD , , "No.", "500"
    LvReportSMS.ColumnHeaders.ADD , , "Nama Report", "10000"
    
    Set listitem = LvReportSMS.ListItems.ADD(, , "1")
        listitem.SubItems(1) = "Report Global SMS Sentitems"
        
    Set listitem = LvReportSMS.ListItems.ADD(, , "2")
        listitem.SubItems(1) = "Report Sentitems Detail"
        
    'Set listitem = LvReportSMS.ListItems.ADD(, , "3")
        'listitem.SubItems(1) = "Report Inbox Detail"
        
    Set listitem = LvReportSMS.ListItems.ADD(, , "4")
         listitem.SubItems(1) = "Report SMS Rejected"
         
    Set listitem = LvReportSMS.ListItems.ADD(, , "5")
        listitem.SubItems(1) = "Report Sentitems CARD OK"
    
    Set listitem = LvReportSMS.ListItems.ADD(, , "6")
        listitem.SubItems(1) = "Report Sentitems CARD Error"
        
    Set listitem = LvReportSMS.ListItems.ADD(, , "7")
        listitem.SubItems(1) = "Report Sentitems AW OK"
        
    Set listitem = LvReportSMS.ListItems.ADD(, , "8")
        listitem.SubItems(1) = "Report Sentitems AW Error"
    
    Set listitem = LvReportSMS.ListItems.ADD(, , "9")
        listitem.SubItems(1) = "Report Sentitems PIL OK"
        
    Set listitem = LvReportSMS.ListItems.ADD(, , "10")
        listitem.SubItems(1) = "Report Sentitems PIL Error"
    
End Sub

Private Sub LvReportSMS_Click()
        LblReport.Caption = LvReportSMS.SelectedItem.SubItems(1)
End Sub

Private Sub update_custid()
    Dim CMDSQL As String
    Dim cmdsql_cari As String
    Dim cmdsql_update As String
    Dim M_objrs As ADODB.Recordset
    Dim m_objrs_cari As ADODB.Recordset
    
    M_RPTCONN.Execute "delete from ReportSmsInbox"
    CMDSQL = "select * from inbox where date(receivingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "custid is null and agent is null"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        '@@ 200610 buat update custid dan agentnya
        cmdsql_cari = "select custid,agent from mgm where "
        cmdsql_cari = "mobileno like '%" + M_objrs("sendernumber") + "' or mobilenoadd1 like='%"
        cmdsql_cari = M_objrs("sendernumber") + "' or mobilenoadd2 like '%"
        cmdsql_cari = M_objrs("sendernumber") + "' or ec_telp like '%"
        cmdsql_cari = M_objrs("sendernumber") + "'"
            
        Set m_objrs_cari = New ADODB.Recordset
        m_objrs_cari.CursorLocation = adUseClient
        m_objrs_cari.Open cmdsql_cari, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If m_objrs_cari.RecordCount > 0 Then
            cmdsql_update = "update inbox set custid='"
            cmdsql_update = cmdsql_update + m_objrs_cari("custid") + "',agent='"
            cmdsql_update = cmdsql_update + m_objrs_cari("agent") + "' where sendernumber='"
            cmdsql_update = cmdsql_update + M_objrs("sendernumber") + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        Set m_objrs_cari = Nothing
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub input_reject_sms()
    Dim CMDSQL As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    
    M_RPTCONN.Execute "delete from reportsmsreject"
    
    CMDSQL = "select * from request_sms where date(tgl_reject) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '" + Format(Tgl2.Value, "yyyy-mm-dd") + "' and rejected='t'"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    
    While Not M_objrs.EOF
        cmdsql_insert = "insert into ReportSmsReject (tgl_reject,custid,agent,no_telp) values ('"
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tgl_reject"))) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("custid")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("agent")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("notelp")) + "')"
        
        M_RPTCONN.Execute cmdsql_insert
        M_objrs.MoveNext
    Wend
    
End Sub

'Report sentitems ----------------

'Report sentitems card ok,, yang ada di daftar tarif
Private Sub tarif_card_ok()
    Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update As String
    Dim CustId As String
    Dim agent As String
    
    M_RPTCONN.Execute "delete from reportsmssentitems"
    
    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,tbl_tarif.operator,tbl_tarif.biaya,"
    CMDSQL = CMDSQL + "sentitems.sendingdatetime,sentitems.tarif,sentitems.creatorid from sentitems,tbl_tarif where date(sentitems.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(sentitems.destinationnumber,1,4)=tbl_tarif.no and "
    CMDSQL = CMDSQL + "sentitems.creatorid is not null and sentitems.creatorid  not like 'AW%' and "
    CMDSQL = CMDSQL + "sentitems.status<>'SendingError' and sentitems.creatorid not like 'PIL%' "
    CMDSQL = CMDSQL + "and sentitems.creatorid not like '%BlastExcelPil%'"
    
       Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    PB1.Max = M_objrs.RecordCount
    
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
            
        End If
       
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("operator")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
        'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("biaya"))) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("biaya"))) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        M_objrs.MoveNext
    Wend
End Sub
'Ini buat sentitems card yang ok,, tapi ga ada di tabel tarif
Private Sub no_tarif_card_ok()
    Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update
    Dim CustId As String
    Dim agent As String
    
    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,sendingdatetime,tarif,creatorid "
    CMDSQL = CMDSQL + "from sentitems where date(sentitems.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(destinationnumber,1,4) not in (select no from tbl_tarif) and "
    CMDSQL = CMDSQL + "(sentitems.creatorid is not null or sentitems.creatorid <>'') and "
    CMDSQL = CMDSQL + "sentitems.creatorid not like 'AW%' and sentitems.status<>'SendingError' "
    CMDSQL = CMDSQL + "and sentitems.creatorid not like '%BlastExcelPil%' and sentitems.creatorid not like 'PIL%'"
    
     Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
        End If
    
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim("LAINNYA") + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
         'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(TxtTarifOpLain.Text)) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(TxtTarifOpLain.Text)) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        
        
        M_objrs.MoveNext
    Wend
End Sub

'Report berdasarkan tarif card yang error
Private Sub tarif_card_error()
     Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update As String
    Dim CustId As String
    Dim agent As String
    
    M_RPTCONN.Execute "delete from reportsmssentitems"
    
    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,tbl_tarif.operator,tbl_tarif.biaya,"
    CMDSQL = CMDSQL + "sentitems.sendingdatetime,sentitems.tarif,sentitems.creatorid from sentitems,tbl_tarif where date(sentitems.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(sentitems.destinationnumber,1,4)=tbl_tarif.no and "
    CMDSQL = CMDSQL + "sentitems.creatorid is not null and sentitems.creatorid  not like 'AW%' and "
    CMDSQL = CMDSQL + "sentitems.status='SendingError' and sentitems.creatorid not like 'PIL%' "
    CMDSQL = CMDSQL + "and sentitems.creatorid not like '%BlastExcelPil%'"
    
       Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    PB1.Max = M_objrs.RecordCount
    
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
            
        End If
       
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("operator")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
        'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("biaya"))) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("biaya"))) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        M_objrs.MoveNext
    Wend
End Sub
'Report card yang error (no tarif)
Private Sub no_tarif_card_error()
    Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update
    Dim CustId As String
    Dim agent As String
    
    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,sendingdatetime,tarif,creatorid "
    CMDSQL = CMDSQL + "from sentitems where date(sentitems.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(destinationnumber,1,4) not in (select no from tbl_tarif) and "
    CMDSQL = CMDSQL + "(sentitems.creatorid is not null or sentitems.creatorid <>'') and "
    CMDSQL = CMDSQL + "sentitems.creatorid not like 'AW%' and sentitems.status='SendingError' and  sentitems.creatorid not like 'PIL%'"
    CMDSQL = CMDSQL + "and sentitems.creatorid not like '%BlastExcelPil%' "
    
     Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
        End If
    
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim("LAINNYA") + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
         'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(TxtTarifOpLain.Text)) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(TxtTarifOpLain.Text)) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        
        
        M_objrs.MoveNext
    Wend
End Sub


'Report sentitems aw tarif ok
Private Sub tarif_aw_ok()
    Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update As String
    Dim CustId As String
    Dim agent As String
    
    M_RPTCONN.Execute "delete from reportsmssentitems"
    
    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,tbl_tarif.operator,tbl_tarif.biaya,"
    CMDSQL = CMDSQL + "sentitems.sendingdatetime,sentitems.tarif,sentitems.creatorid from sentitems,tbl_tarif where date(sentitems.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(sentitems.destinationnumber,1,4)=tbl_tarif.no and "
    CMDSQL = CMDSQL + "sentitems.creatorid is not null and sentitems.creatorid like 'AW%' and "
    CMDSQL = CMDSQL + "sentitems.status<>'SendingError' "
   
    
       Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    PB1.Max = M_objrs.RecordCount
    
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
            
        End If
       
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("operator")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
        'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("biaya"))) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("biaya"))) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        M_objrs.MoveNext
    Wend
End Sub
'No tarif aw ok
Private Sub no_tarif_aw_ok()
    Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update
    Dim CustId As String
    Dim agent As String
    
    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,sendingdatetime,tarif,creatorid "
    CMDSQL = CMDSQL + "from sentitems where date(sentitems.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(destinationnumber,1,4) not in (select no from tbl_tarif) and "
    CMDSQL = CMDSQL + "(sentitems.creatorid is not null or sentitems.creatorid <>'') and "
    CMDSQL = CMDSQL + "sentitems.creatorid like 'AW%' and sentitems.status<>'SendingError' "
   
     Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
        End If
    
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim("LAINNYA") + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
         'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(TxtTarifOpLain.Text)) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(TxtTarifOpLain.Text)) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        
        
        M_objrs.MoveNext
    Wend
End Sub
'report sentitems aw tarif yang error
Private Sub tarif_aw_error()
    Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update As String
    Dim CustId As String
    Dim agent As String
    
    M_RPTCONN.Execute "delete from reportsmssentitems"
    
    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,tbl_tarif.operator,tbl_tarif.biaya,"
    CMDSQL = CMDSQL + "sentitems.sendingdatetime,sentitems.tarif,sentitems.creatorid from sentitems,tbl_tarif where date(sentitems.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(sentitems.destinationnumber,1,4)=tbl_tarif.no and "
    CMDSQL = CMDSQL + "sentitems.creatorid is not null and sentitems.creatorid like 'AW%' and "
    CMDSQL = CMDSQL + "sentitems.status='SendingError' "
   
    
       Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    PB1.Max = M_objrs.RecordCount
    
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
            
        End If
       
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("operator")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
        'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("biaya"))) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("biaya"))) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        M_objrs.MoveNext
    Wend
End Sub
'report sentitems aw error no tarif
Private Sub no_tarif_aw_error()
    Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update
    Dim CustId As String
    Dim agent As String
    
    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,sendingdatetime,tarif,creatorid "
    CMDSQL = CMDSQL + "from sentitems where date(sentitems.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(destinationnumber,1,4) not in (select no from tbl_tarif) and "
    CMDSQL = CMDSQL + "(sentitems.creatorid is not null or sentitems.creatorid <>'') and "
    CMDSQL = CMDSQL + "sentitems.creatorid like 'AW%' and sentitems.status='SendingError' "
   
     Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
        End If
    
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim("LAINNYA") + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
         'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(TxtTarifOpLain.Text)) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(TxtTarifOpLain.Text)) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        
        
        M_objrs.MoveNext
    Wend
End Sub

'Report sentitems tarif pil ok
Private Sub tarif_pil_ok()
    Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update As String
    Dim CustId As String
    Dim agent As String
    
    M_RPTCONN.Execute "delete from reportsmssentitems"
    
    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,tbl_tarif.operator,tbl_tarif.biaya,"
    CMDSQL = CMDSQL + "sentitems.sendingdatetime,sentitems.tarif,sentitems.creatorid from sentitems,tbl_tarif where date(sentitems.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(sentitems.destinationnumber,1,4)=tbl_tarif.no and "
    CMDSQL = CMDSQL + "sentitems.creatorid is not null  and "
    CMDSQL = CMDSQL + "sentitems.status<>'SendingError' and (sentitems.creatorid  like 'PIL%' "
    CMDSQL = CMDSQL + "or sentitems.creatorid  like '%BlastExcelPil%')"
    
       Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    PB1.Max = M_objrs.RecordCount
    
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
            
        End If
       
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("operator")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
        'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("biaya"))) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("biaya"))) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        M_objrs.MoveNext
    Wend
End Sub
'report sentitems pil ok no tarif
Private Sub no_tarif_pil_ok()
    Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update
    Dim CustId As String
    Dim agent As String
    
    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,sendingdatetime,tarif,creatorid "
    CMDSQL = CMDSQL + "from sentitems where date(sentitems.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(destinationnumber,1,4) not in (select no from tbl_tarif) and "
    CMDSQL = CMDSQL + "(sentitems.creatorid is not null or sentitems.creatorid <>'') and "
    CMDSQL = CMDSQL + "sentitems.status<>'SendingError' and (sentitems.creatorid like 'PIL%' or  "
    CMDSQL = CMDSQL + "sentitems.creatorid  like '%BlastExcelPil%') "
    
     Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
        End If
    
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim("LAINNYA") + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
         'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(TxtTarifOpLain.Text)) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(TxtTarifOpLain.Text)) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        
        
        M_objrs.MoveNext
    Wend
End Sub

'report sendsms pil error tarif
Private Sub tarif_pil_error()
    Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update As String
    Dim CustId As String
    Dim agent As String
    
    M_RPTCONN.Execute "delete from reportsmssentitems"
    
    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,tbl_tarif.operator,tbl_tarif.biaya,"
    CMDSQL = CMDSQL + "sentitems.sendingdatetime,sentitems.tarif,sentitems.creatorid from sentitems,tbl_tarif where date(sentitems.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(sentitems.destinationnumber,1,4)=tbl_tarif.no and "
    CMDSQL = CMDSQL + "sentitems.creatorid is not null  and "
    CMDSQL = CMDSQL + "sentitems.status='SendingError' and (sentitems.creatorid  like 'PIL%' "
    CMDSQL = CMDSQL + "or sentitems.creatorid  like '%BlastExcelPil%')"
    
       Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    PB1.Max = M_objrs.RecordCount
    
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
            
        End If
       
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("operator")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
        'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("biaya"))) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("biaya"))) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        M_objrs.MoveNext
    Wend
End Sub
'report sentitems pil error no tarif
Private Sub no_tarif_pil_error()
    Dim CMDSQL, nama_operator As String
    Dim cmdsql_insert As String
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_update
    Dim CustId As String
    Dim agent As String
    
    CMDSQL = "select sentitems.destinationnumber,sentitems.textdecoded,sendingdatetime,tarif,creatorid "
    CMDSQL = CMDSQL + "from sentitems where date(sentitems.sendingdatetime) between '"
    CMDSQL = CMDSQL + Format(Tgl1.Value, "yyyy-mm-dd") + "' and '"
    CMDSQL = CMDSQL + Format(Tgl2.Value, "yyyy-mm-dd") + "' and "
    CMDSQL = CMDSQL + "substring(destinationnumber,1,4) not in (select no from tbl_tarif) and "
    CMDSQL = CMDSQL + "(sentitems.creatorid is not null or sentitems.creatorid <>'') and "
    CMDSQL = CMDSQL + "sentitems.status='SendingError' and (sentitems.creatorid like 'PIL%' or  "
    CMDSQL = CMDSQL + "sentitems.creatorid  like '%BlastExcelPil%') "
    
     Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount = 0 Then
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        
        If IsNull(M_objrs("creatorid")) Or M_objrs("creatorid") = "''" Or M_objrs("creatorid") = "" Then
            CustId = ""
            agent = ""
        Else
            a = Split(M_objrs("creatorid"), "-")
            CustId = a(0)
            agent = a(1)
        End If
    
        cmdsql_insert = "insert into reportsmssentitems (operator,"
        cmdsql_insert = cmdsql_insert + "no_telepon,detail_sms,biaya,tgl_sending,custid,agent) values ('"
        cmdsql_insert = cmdsql_insert + Trim("LAINNYA") + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("destinationnumber")) + "','"
        cmdsql_insert = cmdsql_insert + Trim(M_objrs("textdecoded")) + "','"
         'Cek apakah tarif di tabel sentitems sudah ada?
        If IsNull(M_objrs("tarif")) Then
            'jika tarif di tabel sentitems masih kosong maka ambil biaya dari tabel tarif
            cmdsql_insert = cmdsql_insert + Trim(CStr(TxtTarifOpLain.Text)) + "','"
        Else
            'Jika tarif sudah ada di tabel sentitems maka ambil biaya yang suda terekan dalam tabel sentitems
            cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("tarif"))) + "','"
        End If
        cmdsql_insert = cmdsql_insert + Trim(CStr(M_objrs("sendingdatetime"))) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(CustId), "", Trim(CustId)) + "','"
        cmdsql_insert = cmdsql_insert + IIf(IsNull(agent), "", Trim(agent)) + "')"
        M_RPTCONN.Execute cmdsql_insert
        
        'Update tarif ke tabel sentitems jika tarif di sentitems masih kosong
        If IsNull(M_objrs("tarif")) Then
            cmdsql_update = "update sentitems set tarif='"
            cmdsql_update = cmdsql_update + Trim(CStr(TxtTarifOpLain.Text)) + "' where "
            cmdsql_update = cmdsql_update + "destinationnumber='"
            cmdsql_update = cmdsql_update + Trim(CStr(M_objrs("destinationnumber"))) + "' and "
            cmdsql_update = cmdsql_update + "sendingdatetime='"
            cmdsql_update = cmdsql_update + Trim(CStr(Format(M_objrs("sendingdatetime"), "yyyy-mm-dd hh:mm:ss"))) + "'"
            M_OBJCONN1.Execute cmdsql_update
        End If
        
        
        M_objrs.MoveNext
    Wend
End Sub
