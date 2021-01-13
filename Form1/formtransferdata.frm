VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form formtransferdata 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Data"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7770
   Begin VB.CommandButton cmdgarbage 
      Caption         =   "Garbage"
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   3960
      Width           =   1095
   End
   Begin VB.ComboBox cbsendapp 
      Height          =   315
      ItemData        =   "formtransferdata.frx":0000
      Left            =   3840
      List            =   "formtransferdata.frx":000D
      TabIndex        =   14
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton btnexit 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton btntransferdata 
      Caption         =   "Approval Transfer Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      TabIndex        =   11
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtdataupload 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   465
      Width           =   3495
   End
   Begin VB.CommandButton cmddata 
      Caption         =   "..."
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   435
      Width           =   615
   End
   Begin VB.ComboBox cbsheet 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   920
      Width           =   2415
   End
   Begin VB.TextBox txtcount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "0"
      Top             =   6510
      Width           =   465
   End
   Begin VB.CommandButton cbproses 
      Caption         =   "Proses"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   6480
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4725
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   8334
      _Version        =   393216
      BackColor       =   -2147483636
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000002&
      Caption         =   "Send Approval To"
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
      Left            =   2040
      TabIndex        =   13
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000002&
      Caption         =   "-Custid              -AgentLama           -AgentBaru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   10
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000002&
      Caption         =   "Penting : data (.xls) dengan header:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
      Caption         =   "Sheet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000002&
      Caption         =   "Jumlah Data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   6540
      Width           =   1095
   End
End
Attribute VB_Name = "formtransferdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_XLSCONN As New ADODB.Connection
Dim RSTEMP As ADODB.Recordset
Dim sbatch As String

Private Sub btnexit_Click()
    Unload Me
End Sub

Private Sub btntransferdata_Click()
    formapprovaltransferdata.Show vbModal
End Sub

Private Sub cbproses_Click()
    Dim S As String
    
    Dim rs As New ADODB.Recordset
    Dim temp_rs As ADODB.Recordset
    Dim str_sql As String
    Dim scustid As String
    Dim sagentlama As String
    Dim sagentbaru As String

        If CommonDialog1.FileName = "" Then
            MsgBox "Browse Data Excel Terlebih Dahulu", vbInformation + vbOKOnly, "Information"
            Exit Sub
        End If
        
        If cbsheet.Text = "" Then
           MsgBox "Pilih Sheet", vbInformation + vbOKOnly, "Information"
           cbsheet.SetFocus
           Exit Sub
        End If
        
        ssql = "SELECT * FROM [" & cbsheet.Text & "]   "
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        
        Set rsTemporary = New ADODB.Recordset
        rsTemporary.CursorLocation = adUseClient
        rsTemporary.CursorType = adOpenDynamic
        rsTemporary.ActiveConnection = M_OBJCONN
        rsTemporary.LockType = adLockOptimistic
        
        rs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        'M_OBJCONN.Execute "delete from tbl_uploadexcel"
        While Not rs.EOF
                scustid = IIf(IsNull(rs("custid")), "", rs("custid"))
                sagentlama = IIf(IsNull(rs("agentlama")), "", rs("agentlama"))
                sagentbaru = IIf(IsNull(rs("agentbaru")), "", rs("agentbaru"))
                str_sql = "INSERT INTO tampungtransferdata values ('" + scustid + "',"
                str_sql = str_sql + "'" + sagentlama + "', '" + sagentbaru + "', now(), '" & MDIForm1.TxtUsername.Text & "', '" & cbsendapp.Text & "')"
                M_OBJCONN.Execute str_sql
            rs.MoveNext
            'RS.MoveNext
        Wend
    
'    S = "insert into tbl_upload_data values (now(),'" & mdiform1.txtusername.text & "',"
'    S = S + "'" & txtdataupload.Text & "', '" & cbsheet.Text & "', '" & txtcount.Text & "')"
'    M_OBJCONN.Execute S
    
    Dim a As String
    a = MsgBox("Berhasil diUpload", vbOKOnly + vbInformation, "Konfirmasi")
    
    Call topup
    
End Sub

Private Sub cbsendapp_Click()
    If cbsendapp.Text = "" Then
        cbproses.Enabled = False
    Else
        cbproses.Enabled = True
    End If
End Sub

Private Sub cbsheet_Click()
    Dim OBJRECORD As New ADODB.Recordset
    
        Set RSTEMP = New ADODB.Recordset
        RSTEMP.CursorLocation = adUseClient
    
        ssql = "SELECT * FROM [" & cbsheet.Text & "] "
            RSTEMP.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
            Set RSTEMP = Nothing
            Set OBJRECORD = New ADODB.Recordset
            OBJRECORD.CursorLocation = adUseClient
            
        ssql = "SELECT * FROM [" & cbsheet.Text & "] "
            DoEvents
            OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
            Set DataGrid1.DATASOURCE = OBJRECORD
            txtcount.Text = OBJRECORD.RecordCount
End Sub

Private Sub topup()
Dim CMDSQL As String
Dim NAMA_MSG As String
Dim i As Integer
Dim NAMA_AM As String
Dim NAMA_SPVCODE As String
Dim M_OBJspv As ADODB.Recordset

For i = 1 To Len("JOKO      !JOKO;")
        Select Case Mid("JOKO      !JOKO;", i, 1)
        Case ";"
            CMDSQL = "INSERT INTO tblpermohonantransferdata"
            CMDSQL = CMDSQL + " ( penggaprove,"
            CMDSQL = CMDSQL + " tanggal,"
            CMDSQL = CMDSQL + " pemohon,"
            CMDSQL = CMDSQL + " ip)"
            CMDSQL = CMDSQL + " VALUES"
            CMDSQL = CMDSQL + " ( '" + cbsendapp.Text + "',"
            CMDSQL = CMDSQL + " '" + Format(Now(), "yyyy-mm-dd") + "',"
            CMDSQL = CMDSQL + " '" + Trim(MDIForm1.TxtUsername.Text) + "',"
            CMDSQL = CMDSQL + " '" + CStr(MDIForm1.Winsock1.LocalIP) + "')"
            M_OBJCONN.Execute CMDSQL
            NAMA_MSG = ""
        Case Else
            NAMA_MSG = NAMA_MSG + Mid("JOKO      !JOKO;", i, 1) 'add to txt
        End Select
        Next i
        Unload Me

End Sub

Private Sub cmddata_Click()
    With CommonDialog1
            .DialogTitle = "Import From File"
            .Filter = "Excel Files|*.xls"
            .ShowOpen
    End With
        
    txtdataupload.Text = CommonDialog1.FileName
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & txtdataupload.Text & ";Extended Properties=Excel 8.0;"
    Set RSTEMP = M_XLSCONN.OpenSchema(adSchemaTables)
    cbsheet.CLEAR
    If RSTEMP.EOF And RSTEMP.BOF Then Exit Sub
        
    While Not RSTEMP.EOF
        cbsheet.AddItem IIf(IsNull(RSTEMP!table_name), "", RSTEMP!table_name)
        RSTEMP.MoveNext
    Wend
    
    Set RSTEMP = Nothing
End Sub

Private Sub cmdgarbage_Click()
    frm_garbage.Show vbModal
End Sub

Private Sub Form_Load()
    If MDIForm1.txtlevel.Text = "Agent" Or MDIForm1.txtlevel.Text = "TeamLeader" Then
        btntransferdata.Visible = False
        cmdgarbage.Visible = False
    End If
    
    If cbsendapp.Text = "" Then
        cbproses.Enabled = False
    End If
End Sub
