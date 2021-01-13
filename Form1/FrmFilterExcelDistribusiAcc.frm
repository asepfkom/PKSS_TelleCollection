VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmFilterExcelDistribusiAcc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filter Account Berdasarkan File Excel - Manage Distribusi Account"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9045
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check_decease 
      Caption         =   "Include Account Decease [ 835 ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtlocation 
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   6225
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   1320
      Width           =   1275
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   6300
      TabIndex        =   4
      Top             =   1320
      Width           =   1275
   End
   Begin VB.ComboBox cbosheet 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   4425
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "&Browse"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
   Begin VB.Label LblLokasiFile 
      Caption         =   "(Pilih lokasi file excel. Format .xls)"
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
      TabIndex        =   6
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pilih Sheet:"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1755
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama File Excel(.xls):"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "FrmFilterExcelDistribusiAcc"
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
            .Filter = "Excel Files|*.xls"
            .ShowOpen
        End With
        
        txtlocation.Text = CommonDialog1.FileName
        LblLokasiFile.Caption = CommonDialog1.FileName
        
        If CommonDialog1.FileName = "" Then Exit Sub
        If M_XLSCONN.state = adStateOpen Then M_XLSCONN.Close
                M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & txtlocation.Text & ";Extended Properties=Excel 8.0;"
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
    Dim OBJRECORD As New ADODB.Recordset
    Dim ssql As String
    Dim CustId As String
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim M_objrs As ADODB.Recordset
    
    On Error GoTo Salah
    
    Set OBJRECORD = New ADODB.Recordset
    OBJRECORD.CursorLocation = adUseClient
    ssql = "SELECT * FROM [" & cbosheet.Text & "] where [custid] is not null"
    DoEvents
    OBJRECORD.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        
    CustId = ""
    
    If OBJRECORD.RecordCount > 0 Then
        While Not OBJRECORD.EOF
            If CustId = "" Then
                CustId = "'" & IIf(IsNull(OBJRECORD(0)), "", OBJRECORD(0)) & "'"
            Else
                CustId = CustId & ",'" & IIf(IsNull(OBJRECORD(0)), "", OBJRECORD(0)) & "'"
            End If
            OBJRECORD.MoveNext
        Wend
    Else
        MsgBox "Data di file excel anda kosong! Cek file excel anda atau mungkin anda salah memilih sheet!", vbOKOnly + vbExclamation, "Peringatan"
        Set OBJRECORD = Nothing
        Exit Sub
    End If
    Set OBJRECORD = Nothing
    
    CMDSQL = "SELECT * FROM mandiri.mgm WHERE custid in (" & CustId & ") "
    CMDSQL = CMDSQL & " AND agent NOT IN ('LUNAS','COMPLAIN','CLAIM','AKSESALL','REVIEW','REVIEW1','REVIEW2','REVIEW3','REVIEW4','REVIEW5','REVIEW6','REVIEW7','REVIEW8','REVIEW9','REVIEW10') AND coalesce(agent,'')<>'' "
    CMDSQL = CMDSQL & " AND custid NOT IN (select distinct custid from mandiri.tblsendptp ) "
    ' TAMBAHAN AGAR CLASS 835 TIDAK KENA AKSES ALL
    ' DIGANTI 23 FEB 2015
    If Check_decease.Value = 0 Then
        CMDSQL = CMDSQL & " AND coalesce(cust_class,'')<>'835' "
    End If
    ' -------------------------------------------
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount > 0 Then
        With FrmDistribusiAcc
            .LvAcc.ListItems.CLEAR
            .PB1.Max = M_objrs.RecordCount
            While Not M_objrs.EOF
                .PB1.Value = M_objrs.Bookmark
                Set ListItem = .LvAcc.ListItems.ADD(, , M_objrs("custid"))
                ListItem.SubItems(1) = M_objrs("name")
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("f_cek_new")), "", M_objrs("f_cek_new"))
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
                ListItem.SubItems(4) = IIf(IsNull(M_objrs("agent_asli")), "", M_objrs("agent_asli"))
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("monitor_akses")), "", M_objrs("monitor_akses"))
                ListItem.SubItems(6) = IIf(IsNull(M_objrs("waktu_akses")), "", Format(M_objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
                If UCase(M_objrs("agent")) = "AKSESALL" Then
                    ListItem.ForeColor = vbRed
                    ListItem.ListSubItems(1).ForeColor = vbRed
                    ListItem.ListSubItems(2).ForeColor = vbRed
                    ListItem.ListSubItems(3).ForeColor = vbRed
                    ListItem.ListSubItems(4).ForeColor = vbRed
                    ListItem.ListSubItems(5).ForeColor = vbRed
                    ListItem.ListSubItems(6).ForeColor = vbRed
                End If
            
                If UCase(M_objrs("agent")) = "#KOSONG#" Then
                    ListItem.ForeColor = vbBlue
                    ListItem.ListSubItems(1).ForeColor = vbBlue
                    ListItem.ListSubItems(2).ForeColor = vbBlue
                    ListItem.ListSubItems(3).ForeColor = vbBlue
                    ListItem.ListSubItems(4).ForeColor = vbBlue
                    ListItem.ListSubItems(5).ForeColor = vbBlue
                    ListItem.ListSubItems(6).ForeColor = vbBlue
                End If
                M_objrs.MoveNext
            Wend
        End With
        MsgBox "Data berhasil di load!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        Unload Me
    Else
        MsgBox "Data tidak ditemukan!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    Exit Sub
Salah:
    MsgBox "Maaf ada kesalahan! " & err.Description, vbOKOnly + vbExclamation, "Peringatan"
    
End Sub

