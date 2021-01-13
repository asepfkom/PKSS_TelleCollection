VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form form_add_history 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Special History"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnexit 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cbproses 
      Caption         =   "Proses"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txtcount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Text            =   "0"
      Top             =   6150
      Width           =   1425
   End
   Begin VB.ComboBox cbsheet 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   920
      Width           =   2415
   End
   Begin VB.CommandButton cmddata 
      Caption         =   "..."
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   435
      Width           =   615
   End
   Begin VB.TextBox txtdataupload 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   465
      Width           =   3495
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
      TabIndex        =   9
      Top             =   1320
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   8334
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
      TabIndex        =   7
      Top             =   6180
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000002&
      Caption         =   "* File Excel (.xls) Yang Berisi: -Cust_ID   -History  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   5400
      TabIndex        =   5
      Top             =   240
      Width           =   1215
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
      TabIndex        =   3
      Top             =   960
      Width           =   615
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
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "form_add_history"
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

Private Sub cbproses_Click()
    Dim S As String
    
    Dim rs As New ADODB.Recordset
    Dim temp_rs As ADODB.Recordset
    Dim str_sql As String
    Dim scustid As String
    Dim shst As String

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
            shst = IIf(IsNull(rs("hst")), "", rs("hst"))
                str_sql = "INSERT INTO mgm_hst(custid, hst, tgl_special_upload, f_special) values ('" + scustid + "',"
                str_sql = str_sql + "'" + shst + "', now(), '1')"
                M_OBJCONN.Execute str_sql
            rs.MoveNext
        Wend
    
    S = "insert into tbl_upload_data values (now(),'" & MDIForm1.TxtUsername.Text & "',"
    S = S + "'" & txtdataupload.Text & "', '" & cbsheet.Text & "', '" & txtcount.Text & "')"
    M_OBJCONN.Execute S
    
    Dim a As String
    a = MsgBox("Berhasil diUpload", vbOKOnly + vbInformation, "Konfirmasi")

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
