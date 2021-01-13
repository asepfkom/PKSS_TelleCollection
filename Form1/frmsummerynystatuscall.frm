VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frm_reportsummery 
   Caption         =   "Report Summery By Status Call"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16185
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   16185
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Criteria Report"
      Height          =   1875
      Left            =   30
      TabIndex        =   8
      Top             =   480
      Width           =   17715
      Begin VB.TextBox TxtPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5340
         TabIndex        =   12
         Top             =   1590
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.ComboBox CBOTEAMNAME 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   930
         Width           =   4815
      End
      Begin VB.ComboBox cboagentname 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   1290
         Width           =   4815
      End
      Begin VB.ComboBox cbocampaign 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   570
         Width           =   2580
      End
      Begin TDBDate6Ctl.TDBDate TdTglCall1 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   503
         Calendar        =   "frmsummerynystatuscall.frx":0000
         Caption         =   "frmsummerynystatuscall.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsummerynystatuscall.frx":0184
         Keys            =   "frmsummerynystatuscall.frx":01A2
         Spin            =   "frmsummerynystatuscall.frx":0200
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
         Height          =   285
         Left            =   4080
         TabIndex        =   14
         Top             =   240
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   503
         Calendar        =   "frmsummerynystatuscall.frx":0228
         Caption         =   "frmsummerynystatuscall.frx":0340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsummerynystatuscall.frx":03AC
         Keys            =   "frmsummerynystatuscall.frx":03CA
         Spin            =   "frmsummerynystatuscall.frx":0428
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
      Begin Threed.SSCommand cmdCari 
         Height          =   360
         Left            =   6480
         TabIndex        =   15
         Top             =   1230
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   635
         _Version        =   196610
         Font3D          =   5
         MousePointer    =   16
         BackColor       =   -2147483644
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmsummerynystatuscall.frx":0450
         Caption         =   "&Go"
         Alignment       =   6
         ButtonStyle     =   2
      End
      Begin MSComDlg.CommonDialog Cd_save 
         Left            =   4635
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Ms. Excel 97/2000/XP|*.xls"
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   5910
         TabIndex        =   16
         Top             =   210
         Visible         =   0   'False
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   360
         Left            =   7605
         TabIndex        =   23
         Top             =   1215
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         _Version        =   196610
         Font3D          =   5
         MousePointer    =   16
         BackColor       =   -2147483644
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmsummerynystatuscall.frx":0917
         Caption         =   "&Export"
         Alignment       =   6
         ButtonStyle     =   2
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "To :"
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
         Height          =   285
         Index           =   7
         Left            =   3240
         TabIndex        =   22
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telesales :"
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
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   1290
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Team Leader :"
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
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   930
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Campaign :"
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
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   570
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Periode :"
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
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label lblcampaign 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   9900
         TabIndex        =   17
         Top             =   2160
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5265
      Left            =   -90
      TabIndex        =   1
      Top             =   2250
      Width           =   16965
      Begin VB.Frame Frame4 
         Caption         =   "Agent"
         Height          =   4950
         Left            =   150
         TabIndex        =   2
         Top             =   135
         Width           =   16080
         Begin MSComctlLib.ListView ListView2 
            Height          =   4665
            Left            =   120
            TabIndex        =   3
            Top             =   180
            Width           =   15840
            _ExtentX        =   27940
            _ExtentY        =   8229
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   33023
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label LBLAGENT 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   900
            TabIndex        =   6
            Top             =   2400
            Width           =   795
         End
      End
   End
   Begin VB.Label Label20 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "CALL MANAGEMENT - ANTI ATTRITION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   12780
      TabIndex        =   7
      Top             =   5310
      Width           =   4485
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   5
      Left            =   12270
      Picture         =   "frmsummerynystatuscall.frx":0DDE
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   420
   End
   Begin VB.Label LBLREASON 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10590
      TabIndex        =   5
      Top             =   4980
      Width           =   4785
   End
   Begin VB.Label lblstatuscall 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   4980
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "Report Summary By Status Call"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   540
      TabIndex        =   0
      Top             =   0
      Width           =   7275
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   60
      Picture         =   "frmsummerynystatuscall.frx":18E8
      Stretch         =   -1  'True
      Top             =   30
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   465
      Left            =   30
      Picture         =   "frmsummerynystatuscall.frx":23F2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17700
   End
End
Attribute VB_Name = "frm_reportsummery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sQuerySelectTemp As String
Public Sub InitiateCbo()
Dim MOBJ As New ADODB.Recordset

'coba isi campaign

If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Then
    strsql = " SELECT * FROM DATASOURCETBL WHERE KODEDS IN (SELECT DISTINCT(recsource) AS CAMPAIGN_CODE FROM MGM WHERE AGENT IN (SELECT USERID FROM USERTBL WHERE spvcode='" + MDIForm1.TxtUsername.Text + "' and aktif='1'))"
Else
    strsql = "SELECT *  FROM DATASOURCETBL"
End If

Set MOBJ = New ADODB.Recordset
    MOBJ.CursorLocation = adUseClient
    MOBJ.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    cbocampaign.CLEAR
    While Not MOBJ.EOF
        cbocampaign.AddItem IIf(IsNull(MOBJ!KODEDS), "", MOBJ!KODEDS)
        MOBJ.MoveNext
    Wend
Set MOBJ = Nothing

'coba combo team
If UCase(MDIForm1.txtlevel) = "SUPERVISOR" Then
    strsql = "SELECT  USERID,AGENT FROM usertbl where USERID in (select USERID from usertbl where userid='" + MDIForm1.TxtUsername.Text + "' and aktif='1' ) "
Else
    strsql = "SELECT USERID,AGENT  FROM usertbl WHERE kdlevel ='2' and aktif='1' "
End If

Set MOBJ = New ADODB.Recordset
    MOBJ.CursorLocation = adUseClient
    MOBJ.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not MOBJ.EOF
        CBOTEAMNAME.AddItem IIf(IsNull(MOBJ!AGENT), "", MOBJ!AGENT)
        MOBJ.MoveNext
    Wend
Set MOBJ = Nothing

If UCase(MDIForm1.txtlevel) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel) = "AGENT" Then
    SSCommand2.Visible = False
    SSCommand4.Visible = False
End If


If UCase(MDIForm1.txtlevel) = "SUPERVISOR" Then
    strsql = "SELECT * FROM usertbl where spvcode in (select USERID from usertbl where userid='" + MDIForm1.TxtUsername.Text + "')  and  kdlevel=1 and aktif='1'  ORDER BY USERID "
Else
    strsql = "SELECT * FROM usertbl where kdlevel='1' and aktif='1' ORDER BY USERID "
End If

Set MOBJ = New ADODB.Recordset
    MOBJ.CursorLocation = adUseClient
    MOBJ.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not MOBJ.EOF
        cboagentname.AddItem IIf(IsNull(MOBJ!AGENT), "", MOBJ!AGENT)
        MOBJ.MoveNext
    Wend
Set MOBJ = Nothing
End Sub
Private Sub CmdCari_Click()
    cmdCari.Enabled = False
    'loadisentive
    cariisentive
    cmdCari.Enabled = True
End Sub
Private Sub Form_Load()
    InitiateCbo
    header
End Sub
Public Sub header()
    loadisentive
End Sub
Private Sub isi_data(strsql As String)
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
    Cd_save.ShowSave
    TxtPath.Text = Cd_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.Text = Empty Then
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
    objBook.SaveAs TxtPath.Text, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
 
Salah:
    Exit Sub
End Sub
Public Sub loadisentive()
On Error Resume Next
Dim MOBJ1 As New ADODB.Recordset
Dim MOBJ2 As New ADODB.Recordset
            'QUERYONE = " SELECT * from usertbl where usertype=1 order by agent "
            
            If MDIForm1.txtlevel.Text = "Supervisor" Then
                QUERYONE = " SELECT * from usertbl where kdlevel IN ('1') and spvcode ='" + MDIForm1.TxtUsername.Text + "' and aktif='1' order by agent "
            ElseIf MDIForm1.txtlevel.Text = "Agent" Then
                QUERYONE = " SELECT * from usertbl where kdlevel IN ('1') and aktif='1' and spvcode in (select spvcode from usertbl where userid ='" + MDIForm1.TxtUsername.Text + "' ) order by agent "
                'QUERYONE = " SELECT * from usertbl where usertype=1 and spvcode ='TALA000464' order by agent "
            
            Else
            
                QUERYONE = " SELECT * from usertbl where usertype IN ('1') and aktif='1' order by agent "
            End If
            
            
            Set MOBJ1 = New ADODB.Recordset
            MOBJ1.CursorLocation = adUseClient
            MOBJ1.Open QUERYONE, M_OBJCONN, adOpenDynamic, adLockOptimistic
            no = 1
            If MOBJ1.RecordCount = 0 Then Exit Sub
      
           M_OBJCONN.Execute "DROP TABLE TEMP"
           strsql = "create table temp ( agent character varying(100) "
            While Not MOBJ1.EOF
                strsql = strsql + "," + Chr(34) + MOBJ1!AGENT + Chr(34) + " character varying(100) "
                MOBJ1.MoveNext
          Wend
          strsql = strsql + ")"
          M_OBJCONN.Execute (strsql)
          strsql = "select tblstatuscall_keterangan from tblstatuscall where tblstatuscall_kdstatus='1' order by grp_call,tblstatuscall_keterangan "
          MOBJ2.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
            
            strsql = "insert into temp(agent) values ('All')"
            M_OBJCONN.Execute (strsql)
            While Not MOBJ2.EOF
          
                strsql = "insert into temp(agent) values ('" + cnull(MOBJ2!tblstatuscall_keterangan) + "')"
                M_OBJCONN.Execute (strsql)
                MOBJ2.MoveNext
        
            Wend
         
          
          strsql = "SELECT * FROM TEMP"
          Set MOBJ1 = New ADODB.Recordset
          MOBJ1.CursorLocation = adUseClient
          MOBJ1.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
          ' CREATE HEADER
          For i = 0 To MOBJ1.fields.Count - 1
                ListView2.ColumnHeaders.ADD , , MOBJ1.fields(i).Name, 15 * TXT
          Next i
          ListView2.ColumnHeaders.ADD , , "total"
          
         
ERRORA:
    'MsgBox Err.Description, vbCritical + vbOKOnly, "TINS"
End Sub
Public Sub cariisentive()
Dim MOBJ As New ADODB.Recordset
Dim MOBJ1 As New ADODB.Recordset
Dim MOBJ2 As New ADODB.Recordset
Dim list As ListItem
Dim strsql1 As String
Dim j As Double
Dim JMLAGENT As Double
Dim total As String
Set MOBJ = New ADODB.Recordset
MOBJ.CursorLocation = adUseClient
mwhere = ""
If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
  
        mwhere = mwhere + " and date(tglcall) between '"
        mwhere = mwhere + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '"
        mwhere = mwhere + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
  
End If


If scampaign <> Empty Then
        mwhere = mwhere + " and recsource like '%" + scampaign + "%'"
End If


If sSteam <> Empty Then
        mwhere = mwhere + " and  agent in (select userid from usertbl where spvcode='" + sSteam + "' and aktif='1' )"
    End If


If sAgent <> Empty Then
        mwhere = mwhere + " and  agent ='" + sAgent + "'"
    End If

strsql = " select mgm.agent,count(custid) as jml ,'All' as kdsts,usertbl.agent as namaagent from mgm left join usertbl on mgm.agent=usertbl.userid where custid <>''  " + mwhere + " AND usertbl.kdlevel IN('1','2') and usertbl.aktif='1' group by mgm.agent,usertbl.agent "
strsql = strsql + vbCrLf + " UNION ALL ("
strsql = strsql + vbCrLf + "  SELECT tbl.AGENT,COUNT(custid) AS jml,'Already Paid' as kdsts,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE STATUSCALL='Already Paid') AS TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.kdlevel IN('1','2') and usertbl.aktif='1'  " + mwhere + "GROUP BY TBL.AGENT,usertbl.agent"
strsql = strsql + vbCrLf + "  )"
strsql = strsql + vbCrLf + "  UNION ALL( "
strsql = strsql + vbCrLf + "  SELECT tbl.AGENT,COUNT(custid) AS jml,'BP' as kdsts,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + " SELECT * FROM MGM WHERE STATUSCALL='BP') AS TBL  left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.kdlevel IN('1','2') and usertbl.aktif='1' " + mwhere + "  GROUP BY tbl.AGENT,USERTBL.agent"
strsql = strsql + vbCrLf + "  )"
strsql = strsql + vbCrLf + "  UNION ALL ("
strsql = strsql + vbCrLf + "  SELECT tbl.AGENT,COUNT(custid) AS jml,'PTP' as kdsts,usertbl.agent as namaagent  FROM ("
strsql = strsql + vbCrLf + " SELECT * FROM MGM WHERE STATUSCALL ='PTP'  ) AS TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.kdlevel IN('1','2') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,USERTBL.agent  )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'Schedule Call' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  STATUSCALL IN ('Schedule Call') ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'Left Message' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  STATUSCALL IN ('Left Message') ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'Negosiasi' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  STATUSCALL IN ('Negosiasi') ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'Busy' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  STATUSCALL IN ('Busy') ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'Invalid' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  STATUSCALL IN ('Invalid') ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'Mailbox' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  STATUSCALL IN ('Mailbox') ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'Unknow' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  STATUSCALL IN ('Unknow') ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'Dead' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  STATUSCALL IN ('Dead') ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'Pindah Alamat' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  STATUSCALL IN ('Pindah Alamat') ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'Salah Sambung' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  STATUSCALL IN ('Salah Sambung') ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'Tidak Ada di Tempat' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  STATUSCALL IN ('Tidak Ada di Tempat') ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'Tidak Diangkat' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  STATUSCALL IN ('Tidak Diangkat') ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'Data Retur' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  STATUSCALL IN ('Data Retur') ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"
strsql = strsql + vbCrLf + "  UNION ALL("
strsql = strsql + vbCrLf + " SELECT tbl.AGENT,COUNT(custid) AS jml,'New Data' as kdsts ,usertbl.agent as namaagent FROM ("
strsql = strsql + vbCrLf + "  SELECT * FROM MGM  WHERE  COALESCE(STATUSCALL,'')='' ) AS"
strsql = strsql + vbCrLf + "  TBL left join usertbl on TBL.agent=usertbl.userid WHERE usertbl.USERTYPE IN('1','6') and usertbl.aktif='1' " + mwhere + " GROUP BY tbl.AGENT,usertbl.agent )"


Set MOBJ1 = New ADODB.Recordset
MOBJ1.CursorLocation = adUseClient
MOBJ1.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
strsql = ""
While Not MOBJ1.EOF
'On Error Resume Next
If IIf(IsNull(MOBJ1("NAMAAGENT")), "", MOBJ1("NAMAAGENT")) <> "" Then
strsql = "update temp set " + Chr(34) + MOBJ1("NAMAAGENT") + Chr(34) + " = " + CStr(MOBJ1!jml) + " where agent='" + MOBJ1!kdsts + "'"
        
        M_OBJCONN.Execute (strsql)
        End If
        MOBJ1.MoveNext
        
Wend
     Set MOBJ1 = Nothing

strsql1 = " select * from temp order by agent  "
Set MOBJ1 = New ADODB.Recordset
MOBJ1.CursorLocation = adUseClient
MOBJ1.Open strsql1, M_OBJCONN, adOpenDynamic, adLockOptimistic
M_OBJCONN.Execute "ALTER TABLE temp drop column IF EXISTS total"
strsql = "ALTER TABLE temp ADD COLUMN Total varchar(100)"
M_OBJCONN.Execute (strsql)

ListView2.ListItems.CLEAR
        While Not MOBJ1.EOF
            JMLAGENT = 0

            Set list = ListView2.ListItems.ADD(, , MOBJ1!AGENT)
            If MOBJ1!AGENT = "Left Message" Then
                rowfactor = MOBJ1.Bookmark
            ElseIf MOBJ1!AGENT = "All" Then
                rowfactor1 = MOBJ1.Bookmark
            ElseIf MOBJ1!AGENT = "Waiting CallBack" Then
                rowfactor2 = MOBJ1.Bookmark
            ElseIf MOBJ1!AGENT = "PTP" Then
                rowfactor6 = MOBJ1.Bookmark
            End If
                For i = 1 To MOBJ1.fields.Count - 1
                    list.SubItems(i) = IIf(IsNull(MOBJ1.fields(i).Value), 0, MOBJ1.fields(i).Value)
                    JMLAGENT = JMLAGENT + IIf(IsNull(MOBJ1.fields(i).Value), 0, MOBJ1.fields(i).Value)
                Next i
                countCol = i
                list.SubItems(ListView2.ColumnHeaders.Count - 1) = JMLAGENT
                total = JMLAGENT
                strsql = "update temp set total='" + total + "' WHERE agent='" + MOBJ1!AGENT + "'"
                M_OBJCONN.Execute (strsql)
            MOBJ1.MoveNext
          Wend
Set list = ListView2.ListItems.ADD(, , "Rate")
strsql = "insert into temp(agent) values ('Rate')"
M_OBJCONN.Execute (strsql)
j = 0
For i = 2 To ListView2.ColumnHeaders.Count
sName = ListView2.ColumnHeaders(i).Text
j = j + 1

nilaipembagi = Val(ListView2.ListItems(Val(rowfactor)).SubItems(j))
nilaipengurang = Val(ListView2.ListItems(Val(rowfactor1)).SubItems(j))
If rowfactor2 <> 0 Then
    factorpengurang = Val(ListView2.ListItems(Val(rowfactor2)).SubItems(j))
Else
    factorpengurang = 0
End If
factorfollow = Val(ListView2.ListItems(Val(rowfactor6)).SubItems(j))
pembagi = (nilaipengurang - factorpengurang)


If pembagi = 0 Then
    list.SubItems(j) = "0%"
Else
    If nilaipembagi = 0 And (nilaipengurang - factorpengurang - factorfollow) = 0 Then
        list.SubItems(j) = "0%"
        strsql = "update temp set " + Chr(34) + sName + Chr(34) + " = '" + list.SubItems(j) + "' where agent= 'Rate'"
        M_OBJCONN.Execute (strsql)
    Else
        list.SubItems(j) = CStr(Round((nilaipembagi / (nilaipengurang - factorpengurang - factorfollow)) * 100, 2)) + "%"
        strsql = "update temp set " + Chr(34) + sName + Chr(34) + " = '" + list.SubItems(j) + "' where agent= 'Rate'"
        M_OBJCONN.Execute (strsql)
    End If
'06122014 HERISCODE
   ' list.SubItems(j) = CStr(Round(nilaipembagi / (nilaipengurang - (factorpengurang + factorfollow)) * 100, 2)) + "%"

End If
Next i

strsql1 = " select * from temp "
Set MOBJ1 = New ADODB.Recordset
MOBJ1.CursorLocation = adUseClient
MOBJ1.Open strsql1, M_OBJCONN, adOpenDynamic, adLockOptimistic
sQuerySelectTemp = strsql1
End Sub

Private Sub SSCommand2_Click()
isi_data (sQuerySelectTemp)
End Sub
