VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FormReport_distribusi 
   Caption         =   "Distribute Report"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtPath 
      Enabled         =   0   'False
      Height          =   480
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.CheckBox Check_all1 
      Caption         =   "Check All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4860
      TabIndex        =   12
      Top             =   2895
      Width           =   1455
   End
   Begin VB.TextBox txtlead 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   8325
      Width           =   915
   End
   Begin VB.ComboBox cmbregion 
      Height          =   315
      Left            =   1410
      TabIndex        =   5
      Top             =   1140
      Width           =   2220
   End
   Begin VB.CommandButton cmdProses 
      Caption         =   "Proses"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   345
      TabIndex        =   1
      Top             =   1950
      Width           =   975
   End
   Begin VB.CommandButton CmdExport 
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1350
      TabIndex        =   0
      Top             =   1935
      Width           =   975
   End
   Begin MSComctlLib.ListView List_ShowReport 
      Height          =   4830
      Left            =   225
      TabIndex        =   2
      Top             =   3390
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   8520
      View            =   3
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
   Begin MSComDlg.CommonDialog CD_Save 
      Left            =   3300
      Top             =   -30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TDBDate6Ctl.TDBDate TdTglCall1 
      Height          =   315
      Left            =   1425
      TabIndex        =   8
      Top             =   735
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   556
      Calendar        =   "FormReport_distribusi.frx":0000
      Caption         =   "FormReport_distribusi.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FormReport_distribusi.frx":0184
      Keys            =   "FormReport_distribusi.frx":01A2
      Spin            =   "FormReport_distribusi.frx":0200
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
      Format          =   "dd-mm-yyyy"
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
      Text            =   "__-__-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate TdTglCall2 
      Height          =   315
      Left            =   3210
      TabIndex        =   9
      Top             =   735
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   556
      Calendar        =   "FormReport_distribusi.frx":0228
      Caption         =   "FormReport_distribusi.frx":0340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FormReport_distribusi.frx":03AC
      Keys            =   "FormReport_distribusi.frx":03CA
      Spin            =   "FormReport_distribusi.frx":0428
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
      Format          =   "dd-mm-yyyy"
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
      Text            =   "__-__-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin MSComctlLib.ListView List_agent 
      Height          =   1965
      Left            =   4860
      TabIndex        =   13
      Top             =   915
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   3466
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Agent"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   4875
      TabIndex        =   14
      Top             =   660
      Width           =   1425
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9645
      TabIndex        =   11
      Top             =   8340
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   7
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   225
      TabIndex        =   6
      Top             =   1125
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   225
      TabIndex        =   4
      Top             =   675
      Width           =   780
   End
   Begin VB.Line Line1 
      X1              =   225
      X2              =   11580
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distribute Report"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   225
      TabIndex        =   3
      Top             =   75
      Width           =   2190
   End
End
Attribute VB_Name = "FormReport_distribusi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim sQuerySelectTemp As String
Dim sQuerySelectTempID As String
Dim sGetagent As String

Private Sub Check_all1_Click()
    Dim i As Integer
    i = 0
    If Check_all1.Value = 1 Then
        For i = 1 To List_agent.ListItems.Count
            List_agent.ListItems(i).Checked = True
        Next i
    ElseIf Check_all1.Value = 0 Then
        For i = 1 To List_agent.ListItems.Count
            List_agent.ListItems(i).Checked = False
        Next i
    End If
    Call GetAgents
End Sub

Private Sub cmbregion_DropDown()
    load_region
End Sub

Private Sub CmdProses_Click()
    If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
        If TdTglCall1.Value > TdTglCall2 Then
            MsgBox "Tanggal Tidak Sesuai", vbInformation, "Informasi"
            Exit Sub
        End If
    Else
        MsgBox "Tanggal Harus Diisi", vbInformation, "Informasi"
        Exit Sub
    End If
    
    If MDIForm1.txtlevel = "Supervisor" Or MDIForm1.txtlevel = "Admin" Then
        sWhere = " where sendby='" + MDIForm1.TxtUsername.text + "'"
    End If
    
    If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
        If Len(sWhere) = 0 Then
            sWhere = " where  date(tgl) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' "
            sWhere = sWhere + " and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
        Else
            sWhere = sWhere + "  and date(tgl) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' "
            sWhere = sWhere + " and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
        End If
    Else
        'MsgBox "Tanggal Call Harus Diisi", vbInformation, "Informasi"
    End If
    If cmbregion.text <> Empty Then
        If Len(sWhere) = 0 Then
            sWhere = sWhere + " where region ='" + cmbregion.text + "'"
        Else
            sWhere = sWhere + " and region ='" + cmbregion.text + "'"
        End If
    End If
    
    If MDIForm1.txtlevel.text = "Supervisor" Or MDIForm1.txtlevel.text = "Agent" Or MDIForm1.txtlevel.text = "Admin" Then
        If sGetagent <> Empty Then
            sWhere = sWhere + " and  userid in (select userid from usertbl where userid in (" + sGetagent + ")and aktif='1')"
        End If
    End If

    
   ' sQuerySelect = "SELECT nama_nasabah as ""NAMA"",c.email as ""EMAIL"",b.f_call_notelp as ""TELP"",tambahan_jangkawaktu as ""TENOR"",coalesce(tanggal_avalaible,'')   as ""TANGGAL AVALIABLE"", coalesce(time_avalaible,'') as ""TIME AVALIABLE"",jenis_kendaraan  as ""JENIS KENDARAAN"",tso_notes as ""TSO NOTE"" FROM tbl_submitweb a INNER JOIN mgm b ON a.id_cust=b.id INNER JOIN tbl_onaccount c ON b.id_onaccount=c.id_onaccount " & sWhere
    
    sQuerySelect = " SELECT tgl as ""Date"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,nama as ""Agent"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,custid as ""No Customer"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,name_ch as ""Nama Customer"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,region as ""Region"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,statuscall as ""Acct.Status"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,statuscall as ""Call Status"""
    sQuerySelect = sQuerySelect + vbCrLf + " ,tagihan as ""Jumlah Tagihan"""
    sQuerySelect = sQuerySelect + vbCrLf + " FROM tbllogdistribusi"
    sQuerySelect = sQuerySelect + vbCrLf + sWhere
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sQuerySelect, M_OBJCONN, adOpenKeyset, adLockOptimistic
    txtlead.text = rs.RecordCount
    sQuerySelectTemp = sQuerySelect
    sQuerySelectTempID = "SELECT custid FROM tbllogdistribusi " & sWhere
    If rs.EOF Then
        List_ShowReport.ListItems.CLEAR
        MsgBox "Data Kosong"
        Exit Sub
    End If
    
    Call ShowListView(rs, List_ShowReport)
End Sub
Public Sub ShowListView(ByRef rsS As ADODB.Recordset, list As ListView, Optional no As Boolean = True)
    Dim j As Double
    j = 1
    list.ColumnHeaders.CLEAR
    If no = True Then
        list.ColumnHeaders.ADD 1, , "NO"
    End If
    If no = True Then
        For i = 0 To rsS.fields.Count - 1
            list.ColumnHeaders.ADD i + 2, , rsS.fields(i).Name
        Next i
    Else
        For i = 0 To rsS.fields.Count - 1
            list.ColumnHeaders.ADD i + 1, , rsS.fields(i).Name
        Next i
    End If
    list.ListItems.CLEAR
    While Not rsS.EOF
        If no = True Then
            Set listv = list.ListItems.ADD(, , j)
            For i = 0 To rsS.fields.Count - 1
                listv.SubItems(i + 1) = cnull(rsS(i))
            Next i
        Else
            Set listv = list.ListItems.ADD(, , cnull(rsS(0)))
            For i = 1 To rsS.fields.Count - 1
                listv.SubItems(i) = cnull(rsS(i))
            Next i
        End If
        j = j + 1
        rsS.MoveNext
    Wend
End Sub
Public Sub load_region()
    sStrsql = " select distinct region  from tbllogdistribusi  "
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        cmbregion.CLEAR
        While Not M_objrs.EOF
                cmbregion.AddItem IIf(IsNull(M_objrs!region), "", M_objrs!region)
                M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub
Private Sub CmdExport_Click()
    If txtlead.text = 0 Then
        MsgBox "Data Kosong"
        Exit Sub
    Else
        Call isi_data_ex(sQuerySelectTemp)
    End If
End Sub
Private Sub isi_data_ex(strsql As String)
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim cmdsql_update As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i As Integer
    Dim m_msgbox As String
     i = 1
     'STRSQL = txtSintaks.Text
 Set M_objrs = New ADODB.Recordset
 M_objrs.CursorLocation = adUseClient
 M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
 
Form_Save:
    CD_save.ShowSave
    TxtPath.text = CD_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Download dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo Form_Save        '-> maka goto form_save
        End If
    End If
    'Set excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
    
    On Error GoTo Salah
    'Proses pengsisian nama field ke excel
    Dim X, Y    As Integer
        If M_objrs.State = 1 Then
            X = 0
            Y = M_objrs.fields().Count - 1
            Do Until X > Y
                DoEvents
                objSheet.Cells(1, i).Value = CStr(M_objrs.fields(X).Name)
                i = i + 1
                X = X + 1
            Loop
        End If
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs TxtPath.text, xlWorkbookNormal
    'objBook.SaveAs TxtPath.text, xlWorkbook
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
 
Salah:
    Exit Sub
End Sub
Private Sub GetAgents()
    Dim sWhere As String
    sWhere = ""
    sGetagent = ""
    sWhere = GETAgent
    If sWhere <> "" Then
        sGetagent = sWhere
        Exit Sub
    End If
End Sub
Public Function GETAgent() As Variant
    Dim row As Double
    row = 1
    strsql = ""
    For i = 1 To List_agent.ListItems.Count
       If List_agent.ListItems(i).Checked = True Then
            If row = 1 Then
                strsql = "'" + List_agent.ListItems(i).text + "'"
            Else
                strsql = strsql + ",'" + List_agent.ListItems(i).text + "'"
            End If
            row = row + 1
      End If
    Next i
    GETAgent = strsql
End Function

Private Sub Form_Load()
    List_agent.ColumnHeaders.ADD 1, , "Kode", 1400
    List_agent.ColumnHeaders.ADD 2, , "Nama Agent", 3000
    Call load_agent
    If MDIForm1.txtlevel.text = "Agent" Then
        Check_all1.Value = 1
        Check_all1.Enabled = False
        CmdExport.Enabled = False
    End If
End Sub

Private Sub List_agent_Click()
    Call GetAgents
End Sub
Private Sub load_agent()
    Dim listv As ListItem
    If MDIForm1.txtlevel.text = "Supervisor" Then
        sStrsql = "select userid , agent  from  usertbl  where  aktif ='1' and  level_name ='Agent' and spvcode='" + MDIForm1.TxtUsername.text + "'"
    ElseIf MDIForm1.txtlevel.text = "Agent" Then
        sStrsql = "select userid , agent  from  usertbl  where  aktif ='1' and  level_name ='Agent' and userid='" + MDIForm1.TxtUsername.text + "'"
    Else
        sStrsql = "select userid , agent  from  usertbl  where  aktif ='1' and  level_name ='Agent' "
    End If
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        List_agent.ListItems.CLEAR
        While Not M_objrs.EOF
                Set listv = List_agent.ListItems.ADD(, , IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID))
                listv.SubItems(1) = IIf(IsNull(M_objrs!AGENT), "", M_objrs!AGENT)
                M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub
