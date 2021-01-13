VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form form_upload_all 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   Caption         =   "Form Upload All"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11895
   LinkTopic       =   "Form2"
   ScaleHeight     =   2550
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1465
      TabIndex        =   8
      Top             =   550
      Width           =   5805
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose File"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton Command1 
         Caption         =   "UPLOAD"
         Height          =   495
         Left            =   10200
         TabIndex        =   9
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtlocation 
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   8445
      End
      Begin VB.CommandButton cmdbrowse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "...."
         Height          =   315
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   820
         Width           =   555
      End
      Begin VB.ComboBox cbosheet 
         Height          =   315
         Left            =   1350
         TabIndex        =   1
         Top             =   1230
         Width           =   2565
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         Caption         =   "Campaign"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   430
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Choose File"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Location"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "Sheet"
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   1260
         Width           =   795
      End
   End
End
Attribute VB_Name = "form_upload_all"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim get_header As String
Dim count_field As Integer
Public M_XLSCONN As New ADODB.Connection
Private Sub isicampaign()
    qs = "select * from information_schema.columns where table_name = 'tblalldata'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount = 0 Then
        qc = "create table tblalldata (id serial, tabel varchar);"
        M_OBJCONN.Execute qc
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    strsql = "select distinct recsource from mgm where recsource not in (select tabel from tblalldata) "
    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Combo1.CLEAR
        While Not M_objrs.EOF
          Combo1.AddItem IIf(IsNull(M_objrs!recsource), "", M_objrs!recsource)
            M_objrs.MoveNext
        Wend
     Set M_objrs = Nothing
End Sub

Private Sub cbosheet_Click()
    
    If txtlocation.text <> "" Then
        If M_XLSCONN.State = adStateOpen Then M_XLSCONN.Close
        M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 8.0;"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & cbosheet.text & "] "
        M_objrs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        get_header = ""
        count_field = 0
        If M_objrs.EOF And M_objrs.BOF Then Exit Sub
        For i = 0 To M_objrs.fields.Count - 1
            On Error Resume Next
            get_header = get_header & """" & M_objrs.fields(i).Name & """" & " varchar,"
            count_field = count_field + 1
            M_OBJCONN.Execute (strsql)
            lblstatus.Caption = "Field Terdefinisi"
        Next i
            get_header = Left(get_header, Len(get_header) - 1)
        Set M_objrs = Nothing
    End If
End Sub

Private Sub cmdbrowse_Click()
    With CommonDialog1
        .DialogTitle = "Import From File"
        '.Filter = "Excel Files|*.xls;*.xlsx"
        .Filter = "Excel Files|*.xls"
        .ShowOpen
    End With
    txtlocation.text = ""
    If CommonDialog1.FileName = "" Then Exit Sub
    txtlocation.text = CommonDialog1.FileName
    If M_XLSCONN.State = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 8.0;"
    Set M_objrs = M_XLSCONN.OpenSchema(adSchemaTables)
    cbosheet.CLEAR
    If M_objrs.EOF And M_objrs.BOF Then Exit Sub
    While Not M_objrs.EOF
        cbosheet.AddItem IIf(IsNull(M_objrs!table_name), "", M_objrs!table_name)
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    Set M_XLSCONN = Nothing

End Sub

Private Sub Command1_Click()
    Dim str_sql As String
    If Combo1.text = "" Then
        MsgBox "Harap Pilih Campaign"
        Exit Sub
    End If
    
    If cbosheet.text = "" Then
        MsgBox "Harap Pilih Sheet"
        Exit Sub
    End If
    
    qs = "select * from information_schema.columns where table_name = 'tblalldata'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount = 0 Then
        qc = "create table tblalldata (id serial, tabel varchar);"
        M_OBJCONN.Execute qc
    End If
    
    qi = "insert into tblalldata (tabel) values ('" & Combo1.text & "');"
    M_OBJCONN.Execute qi
    
    qc = "create table " & """" & Combo1.text & """" & " ( " & get_header & ") ;"
    M_OBJCONN.Execute qc
    
        ssql = "SELECT * FROM [" & cbosheet.text & "]   "
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
            
        Set rsTemporary = New ADODB.Recordset
        rsTemporary.CursorLocation = adUseClient
        rsTemporary.CursorType = adOpenDynamic
        rsTemporary.ActiveConnection = M_OBJCONN
        rsTemporary.LockType = adLockOptimistic
            
        rs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
    
        While Not rs.EOF
        
            str_sql = "INSERT INTO " & """" & Combo1.text & """" & " Values ("
            For i = 0 To count_field - 1
                str_sql = str_sql + " '" + cnull(rs(i)) + "' ,"
                
            Next i
            str_sql = Left(str_sql, Len(str_sql) - 1)
            str_sql = str_sql + ");"
            
            M_OBJCONN.Execute str_sql
            
            rs.MoveNext
        Wend
            
        Set rs = Nothing
        MsgBox "Data Berhasil Di - Upload!"
End Sub

Private Sub Form_Load()
    isicampaign
End Sub
