VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form_upload_absensi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Upload Absensi"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9030
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "&Browse"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   915
   End
   Begin VB.ComboBox cbosheet 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   4425
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   6300
      TabIndex        =   2
      Top             =   1320
      Width           =   1275
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   1320
      Width           =   1275
   End
   Begin VB.TextBox txtlocation 
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   6225
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama File Excel(.xls):"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1755
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pilih Sheet:"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1755
   End
   Begin VB.Label LblLokasiFile 
      Caption         =   "(Pilih lokasi file excel. Format .xls, .xlsx)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2940
      TabIndex        =   5
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Form_upload_absensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_XLSCONN As New ADODB.Connection
Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdBrowse_Click()
    With CommonDialog1
        .DialogTitle = "Pilih file excel"
        .Filter = "Excel Files|*.xls;*.xlsx"
        .ShowOpen
    End With
        
    txtlocation.Text = CommonDialog1.FileName
    LblLokasiFile.Caption = CommonDialog1.FileName
    
    If CommonDialog1.FileName = "" Then Exit Sub
    If M_XLSCONN.State = adStateOpen Then M_XLSCONN.Close
    'M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & txtlocation.Text & ";Extended Properties=Excel 8.0;"
    M_XLSCONN.Open "Provider = Microsoft.ACE.OLEDB.12.0;data source = " & txtlocation.Text & ";Extended Properties=Excel 12.0;"
        
    Set RSTEMP = M_XLSCONN.OpenSchema(adSchemaTables)
    cbosheet.CLEAR
    If RSTEMP.EOF And RSTEMP.BOF Then Exit Sub
    While Not RSTEMP.EOF
        cbosheet.AddItem IIf(IsNull(RSTEMP!table_name), "", RSTEMP!table_name)
        RSTEMP.MoveNext
    Wend
    Set RSTEMP = Nothing
End Sub

Private Sub cmdOK_Click()
    Dim OBJRECORD As ADODB.Recordset
    Dim ssql As String
    Dim CustId As String
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim M_objrs As ADODB.Recordset
    
    Dim jml_hours As Double
    
    On Error GoTo Salah
    
    Set OBJRECORD = New ADODB.Recordset
    OBJRECORD.CursorLocation = adUseClient
    ssql = "SELECT * FROM [" & cbosheet.Text & "] WHERE [NoPeg] IS NOT NULL"
    DoEvents
    OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        
    CustId = ""
    
    If OBJRECORD.RecordCount > 0 Then
        ProgressBar1.Max = OBJRECORD.RecordCount
        While Not OBJRECORD.EOF
            DoEvents
            ProgressBar1.Value = OBJRECORD.Bookmark
            jml_hours = DateDiff("n", Format(IIf(IsNull(OBJRECORD(5)) Or OBJRECORD(5) = "", "00:00", OBJRECORD(5)), "hh:mm"), Format(IIf(IsNull(OBJRECORD(6)) Or OBJRECORD(6) = "", "00:00", OBJRECORD(6)), "hh:mm")) / 60
            ' DELETE DATA EXISTING
            M_OBJCONN.Execute "DELETE FROM tblabsen WHERE nopeg='" & cnull(OBJRECORD(0)) & "' AND tanggal='" & cnull(OBJRECORD(4)) & "'"
            ' INSERT NEW
            M_OBJCONN.Execute "INSERT INTO tblabsen(nopeg,noakun,no,nama,tanggal,masuk,keluar,hours) VALUES " & _
                            "('" & cnull(OBJRECORD(0)) & "','" & cnull(OBJRECORD(1)) & "','" & cnull(OBJRECORD(2)) & "','" & cnull(OBJRECORD(3)) & "'," & _
                            "'" & cnull(OBJRECORD(4)) & "','" & IIf(IsNull(OBJRECORD(5)), "", OBJRECORD(5)) & "','" & IIf(IsNull(OBJRECORD(6)), "", OBJRECORD(6)) & "'," & jml_hours & ")"
            OBJRECORD.MoveNext
        Wend
        MsgBox OBJRECORD.RecordCount & " Data(s) berhasil di Upload!!", vbOKOnly + vbInformation, "INFO"
    Else
        MsgBox "Data di file excel anda kosong! Cek file excel anda atau mungkin anda salah memilih sheet!", vbOKOnly + vbExclamation, "Peringatan"
        Set OBJRECORD = Nothing
        Exit Sub
    End If
    Set OBJRECORD = Nothing
    
    Exit Sub
Salah:
    MsgBox "Maaf ada kesalahan! " & err.Description, vbOKOnly + vbExclamation, "Peringatan"
End Sub

