VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_upload_fresh_wo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form Upload Fresh WO"
   ClientHeight    =   10350
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12420
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   12420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Upload Data"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   12135
      Begin VB.CommandButton Command3 
         Caption         =   "Execute"
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
         Left            =   10560
         TabIndex        =   13
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10560
         TabIndex        =   12
         Top             =   780
         Width           =   1365
      End
      Begin VB.ComboBox cboSheet 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Form_upload_fresh_wo.frx":0000
         Left            =   1815
         List            =   "Form_upload_fresh_wo.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   645
         Width           =   3525
      End
      Begin VB.TextBox txtListLocation 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1815
         TabIndex        =   6
         Top             =   270
         Width           =   3510
      End
      Begin VB.CommandButton cmdbrowse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5460
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Excel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   330
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Sheet"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label11 
         Caption         =   "* File Excel (.xls)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6060
         TabIndex        =   8
         Top             =   300
         Width           =   2175
      End
   End
   Begin VB.TextBox txtcount 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10710
      TabIndex        =   1
      Top             =   9870
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7245
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   12779
      _Version        =   393216
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
   Begin MSComDlg.CommonDialog CD_Brows 
      Left            =   7560
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   120
      Picture         =   "Form_upload_fresh_wo.frx":0004
      Stretch         =   -1  'True
      Top             =   60
      Width           =   540
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "Upload Fresh WO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   750
      TabIndex        =   11
      Top             =   45
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Jumlah Data"
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
      Left            =   9660
      TabIndex        =   3
      Top             =   9900
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "List Data Yang Akan di Upload :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   2160
      Width           =   2805
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "Form_upload_fresh_wo.frx":0B0E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20400
   End
End
Attribute VB_Name = "Form_upload_fresh_wo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSTEMP As New ADODB.Recordset
Dim M_XLSCONN As New ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub cbosheet_Click()
    Dim OBJRECORD As New ADODB.Recordset
    
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.CursorLocation = adUseClient

    ssql = "SELECT * FROM [" & cbosheet.Text & "] "
        RSTEMP.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Set RSTEMP = Nothing
        Set OBJRECORD = New ADODB.Recordset
        OBJRECORD.CursorLocation = adUseClient
        
    ssql = "SELECT * FROM [" & cbosheet.Text & "] "
        DoEvents
        OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Set DataGrid1.DATASOURCE = OBJRECORD
        txtcount.Text = OBJRECORD.RecordCount
End Sub

Private Sub CmdBrowse_Click()
    With CD_Brows
            .DialogTitle = "Import From File"
            .Filter = "Excel Files|*.xls"
            .ShowOpen
    End With
        
    txtListLocation.Text = CD_Brows.FileName
    
    If CD_Brows.FileName = "" Then Exit Sub
    
    If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & txtListLocation.Text & ";Extended Properties=Excel 8.0;"
    Set RSTEMP = M_XLSCONN.OpenSchema(adSchemaTables)
    cbosheet.CLEAR
    If RSTEMP.EOF And RSTEMP.BOF Then Exit Sub
        
    While Not RSTEMP.EOF
        cbosheet.AddItem IIf(IsNull(RSTEMP!table_name), "", RSTEMP!table_name)
        RSTEMP.MoveNext
    Wend
    
    Set RSTEMP = Nothing
End Sub

Private Sub Command3_Click()
    Call InsertData
End Sub

Private Sub InsertData()
    Dim rs As New ADODB.Recordset
    Dim temp_rs As ADODB.Recordset
    Dim str_sql As String
    Dim CustId As String
    Dim segment As String
    Dim action As String
    Dim asg_date As String
    Dim tag As String
    
    If CD_Brows.FileName = "" Then
        MsgBox "Browse Data Excel Terlebih Dahulu", vbInformation + vbOKOnly, "Information"
        Exit Sub
    End If
    
    If cbosheet.Text = "" Then
       MsgBox "Pilih Sheet", vbInformation + vbOKOnly, "Information"
       cbosheet.SetFocus
       Exit Sub
    End If
        
    ssql = "SELECT * FROM [" & cbosheet.Text & "]   "
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
        
    Set rsTemporary = New ADODB.Recordset
    rsTemporary.CursorLocation = adUseClient
    rsTemporary.CursorType = adOpenDynamic
    rsTemporary.ActiveConnection = M_OBJCONN
    rsTemporary.LockType = adLockOptimistic
        
    rs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic

    While Not rs.EOF
        CustId = IIf(IsNull(rs("ACCT")), "", rs("ACCT"))
        segment = IIf(IsNull(rs("segment")), "", rs("segment"))
        action = IIf(IsNull(rs("action")), "", rs("action"))
        asg_date = IIf(IsNull(rs("assignmt_date")), "", rs("assignmt_date"))
        tag = IIf(IsNull(rs("tag")), "", rs("tag"))
        
        If CustId <> "" Then
            If Len(asg_date) < 2 Then
                str_sql = "INSERT INTO mandiri.tbl_fresh_wo(custid,segment, action, asg_date,tag)"
                str_sql = str_sql + " VALUES ('" & CustId & "', '" & segment & "', '" & action & "',    "
                str_sql = str_sql + " null, '" & tag & "') "
            Else
                str_sql = "INSERT INTO mandiri.tbl_fresh_wo(custid,segment, action, asg_date,tag)"
                str_sql = str_sql + " VALUES ('" & CustId & "', '" & segment & "', '" & action & "',    "
                str_sql = str_sql + " '" & asg_date & "', '" & tag & "') "
            End If
        
            M_OBJCONN.Execute str_sql
        
        
            ' SIMPAN DI REMARKS ------------
            str_sql = "insert into mandiri.mgm_hst "
            str_sql = str_sql + " (custid,agent,hst,tgl) values ("
            str_sql = str_sql + "'" + CustId + "','fresh_wo',"
            str_sql = str_sql + "'" + action + "',now())"
            
            M_OBJCONN.Execute str_sql
            ' ------------------------------
        End If
        rs.MoveNext
    Wend
        
    Set rs = Nothing
    MsgBox "Data Berhasil Di - Upload!"
    DataGrid1.row = 0
    Unload Me
    Me.Show vbModal
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

