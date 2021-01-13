VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmUpload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Upload Data Payment   ver1.1  deltagrandi.co.id"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   Icon            =   "frmUpload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmUpload.frx":15162
   ScaleHeight     =   5505
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand SSCommand1 
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   21
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   196610
      MousePointer    =   16
      PictureFrames   =   1
      Picture         =   "frmUpload.frx":24A43
      Caption         =   "         &UPLOAD"
      ButtonStyle     =   3
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   5310
      TabIndex        =   19
      Top             =   330
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Max             =   25
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1575
      TabIndex        =   15
      Top             =   285
      Width           =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1290
      Top             =   4875
   End
   Begin VB.ComboBox cboData 
      Height          =   315
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   6855
   End
   Begin VB.FileListBox FileImport 
      Height          =   1845
      Left            =   4200
      Pattern         =   "*.xls"
      TabIndex        =   5
      Top             =   720
      Width           =   3015
   End
   Begin VB.DirListBox DirImport 
      Height          =   1440
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   3615
   End
   Begin VB.DriveListBox driveImport 
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "&Upload"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&DataSource"
      Height          =   375
      Left            =   600
      Picture         =   "frmUpload.frx":24D74
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   22
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   196610
      MousePointer    =   16
      PictureFrames   =   1
      Picture         =   "frmUpload.frx":2C38D
      Caption         =   "    E&XIT"
      ButtonStyle     =   3
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proses :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   4155
      TabIndex        =   20
      Top             =   315
      Width           =   1125
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6165
      TabIndex        =   14
      Top             =   4950
      Width           =   1050
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   4815
      TabIndex        =   13
      Top             =   4950
      Width           =   1125
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   6165
      TabIndex        =   12
      Top             =   4440
      Width           =   1050
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Data Yang Diupload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   11
      Top             =   4485
      Width           =   3390
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Data Sebelumnya"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   10
      Top             =   3945
      Width           =   2115
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   6120
      TabIndex        =   9
      Top             =   3945
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data   :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   330
      TabIndex        =   8
      Top             =   255
      Width           =   900
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   5955
      TabIndex        =   16
      Top             =   3840
      Width           =   1305
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   2
      Left            =   5955
      TabIndex        =   18
      Top             =   4905
      Width           =   1305
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   5955
      TabIndex        =   17
      Top             =   4410
      Width           =   1305
   End
End
Attribute VB_Name = "frmUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstemp As ADODB.Recordset
Dim gFlag As Boolean
Dim sSQL As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Private Function HitungJamsostek(BscSalary As Currency, jNIK As String, _
'                                 JamsosID As String, Kondisi As String)
'    Dim JKK As Currency
'    Dim JKM As Currency
'    Dim JHT As Currency
'    Dim WajibComp As Currency
'    Dim WajibKaryawan As Currency
'
'    Set rsJamsos = New ADODB.Recordset
'    rsJamsos.CursorLocation = adUseClient
'
'    sSQL = "SELECT * FROM TblSettingPajak "
'    rsJamsos.Open sSQL, conn, adOpenDynamic, adLockOptimistic
'
'    If rsJamsos.EOF = True And rsJamsos.BOF = True Then Exit Function
'
'    JKK = rsJamsos("JKK") * BscSalary / 100
'    JKM = rsJamsos("JKM") * BscSalary / 100
'    JHT = rsJamsos("JHT") * BscSalary / 100
'
'    WajibComp = rsJamsos("WajibComp") * BscSalary / 100
'    WajibKaryawan = rsJamsos("WajibKaryawan") * BscSalary / 100
'
'    Set rsJamsos = Nothing
'
'    sSQL = "INSERT INTO TblJamsostek "
'    sSQL = sSQL & "(NIK, TglJamsostek, CompID, DivisiID, BscSalary, WajibComp, "
'    sSQL = sSQL & "WajibKaryawan, JKK, JKM, JHT, JamsosID, Kondisi) "
'    sSQL = sSQL & "VALUES "
'    sSQL = sSQL & "('" & jNIK & "', '" & Format(tdbUpload.Value, "yyyy/mm/01") & "', "
'    sSQL = sSQL & "'" & cboCompany(0).Text & "', '" & cboDivisi(0).Text & "', "
'    sSQL = sSQL & "'" & BscSalary & "',  "
'    sSQL = sSQL & "'" & ReplaceAngka(WajibComp) & "', '" & ReplaceAngka(WajibKaryawan) & "', "
'    sSQL = sSQL & "'" & ReplaceAngka(JKK) & "', '" & ReplaceAngka(JKM) & "', "
'    sSQL = sSQL & "'" & ReplaceAngka(JHT) & "', '" & JamsosID & "', '" & Kondisi & "') "
'    conn.Execute sSQL
'End Function

'Private Function CheckInsertedValue() As Boolean
'    CheckInsertedValue = True
'    CheckInsertedValue = CheckInsertedValue And cboCompany(0).Text <> vbNullString
'    CheckInsertedValue = CheckInsertedValue And cboDivisi(0).Text <> vbNullString
'    CheckInsertedValue = CheckInsertedValue And tdbUpload.ValueIsNull <> True
'    CheckInsertedValue = CheckInsertedValue And txtFile.Text <> vbNullString
'End Function

Private Sub DeleteUploadData()
'    conn.Execute "DELETE FROM TblGaji WHERE " & _
'                 "CompID = '" & cboCompany(0).Text & "' AND " & _
'                 "DivisiID = '" & cboDivisi(0).Text & "' AND " & _
'                 "TglGaji = '" & Format(tdbUpload, "yyyy/mm/01") & "'"
'
'    conn.Execute "DELETE FROM TblJamsostek WHERE " & _
'                 "CompID = '" & cboCompany(0).Text & "' AND " & _
'                 "DivisiID = '" & cboDivisi(0).Text & "' AND " & _
'                 "TglJamsostek = '" & Format(tdbUpload, "yyyy/mm/01") & "'"
'
'    conn.Execute "DELETE FROM TblGajiEx"
    
'    conn.Execute "DELETE FROM TblPajakBln WHERE " & _
'                 "CompID = '" & cboCompany(0).Text & "' AND " & _
'                 "DivisiID = '" & cboDivisi(0).Text & "' AND " & _
'                 "TglPajakBln = '" & Format(tdbUpload, "yyyy/mm/01") & "'"
End Sub

Private Sub ImportToSQL(ByVal mTableName As String)
   'import dari access ke sql itu -->> select dari access trus insert ke sql
    'Dim IuranJamsostek As Currency
    'Dim NIK As String
    
On Error GoTo errHandle
    M_OBJCONN.BeginTrans
    
    Set rstemp = New ADODB.Recordset
    rstemp.CursorLocation = adUseClient
    
    sSQL = "SELECT * FROM " & mTableName
    rstemp.Open sSQL, connAccess, adOpenDynamic, adLockOptimistic
    ProgressBar1.Max = rstemp.RecordCount + 3
        
        
    Label5.Caption = rstemp.RecordCount
    Label7.Caption = CInt(Label5.Caption) + CInt(Label2.Caption)
    'delete data yang uda masuk biar dinamis
    'Call DeleteUploadData
    
    rstemp.MoveFirst
    While Not rstemp.EOF
    ProgressBar1.Value = rstemp.Bookmark
            sSQL = "INSERT INTO tBLtEMP "
            sSQL = sSQL & "(CUSTID,Paydate,Payment ,Datafrom,Agent)"
            sSQL = sSQL & "VALUES "
            sSQL = sSQL & "('" & rstemp("custid") & "', "
            sSQL = sSQL & "'" & Format(rstemp("paydate"), "yyyy-mm-dd") & "', "
            sSQL = sSQL & "" & rstemp("payment") & ", "
            sSQL = sSQL & "'" & Text1.Text & "', "
            sSQL = sSQL & "'" & rstemp("agent") & "')"
            M_OBJCONN.Execute sSQL
            
        rstemp.MoveNext
    Wend
    Set rstemp = Nothing
    
    conn.CommitTrans
    MsgBox "Upload Done...", vbInformation, "Informasi"
    Label5.Caption = "-"
    ProgressBar1.Visible = False
    Exit Sub
    
errHandle:
    MsgBox Err.Description

    conn.RollbackTrans
End Sub

Private Sub Command1_Click()
FRM_DATASOURCE_LIST.Show
End Sub

Private Sub DirImport_Change()
    FileImport.Path = DirImport.Path
    txtFile.Text = DirImport.Path
End Sub

Private Sub driveImport_Change()
    DirImport.Path = driveImport.Drive
    txtFile.Text = driveImport.Drive
End Sub

Private Sub FileImport_Click()
    If FileImport.FileName = Null Then
        MsgBox "Pilih File Excelnya Dahulu", vbInformation + vbOKOnly, "Informasi"
    End If
    txtFile.Text = "" & DirImport.Path & "" + "\" + FileImport.FileName
End Sub

Private Sub DataComboCompany()
openConn
    cboData(1).Clear
    
    Set rstemp = New ADODB.Recordset
    rstemp.CursorLocation = adUseClient
    
    sSQL = "SELECT KODEDS FROM DATASOURCETBL ORDER BY KODEDS"
    

    
    
    rstemp.Open sSQL, conn, adOpenDynamic, adLockOptimistic
    
    If rstemp.BOF = True And rstemp.EOF = True Then Exit Sub
    
    rstemp.MoveFirst
    While Not rstemp.EOF
        
        cboData(1).AddItem rstemp("KODEDS")
        
        rstemp.MoveNext
    Wend
    
    Set rstemp = Nothing
End Sub

Private Sub cbodata_Click(Index As Integer)
    Set rstemp = New ADODB.Recordset
    rstemp.CursorLocation = adUseClient
    
    Select Case Index
'    Case 0
'        sSQL = "SELECT CompID, CompName FROM TblCompany " & _
'               "WHERE CompID = '" & cboCompany(0).Text & "' "
    Case 1
        sSQL = "SELECT KODEDS FROM DATASOURCETBL " & _
               "WHERE KODEDS  = '" & cboData(1).Text & "' "
    End Select
    
    rstemp.Open sSQL, conn, adOpenDynamic, adLockOptimistic
    If rstemp.EOF = True And rstemp.BOF = True Then Exit Sub
    
   ' cboCompany(0).Text = rstemp("CompID")
    cboData(1).Text = rstemp("KODEDS")
    
    Set rstemp = Nothing

    Call DataComboDivisi
End Sub

Private Sub cboDivisi_Click(Index As Integer)
'    Set rstemp = New ADODB.Recordset
'    rstemp.CursorLocation = adUseClient
'
'    Select Case Index
'    Case 0
'        sSQL = "SELECT DivisiID , DivisiName FROM TblDivisi " & _
'               "WHERE DivisiID = '" & cboDivisi(0).Text & "' "
'    Case 1
'        sSQL = "SELECT DivisiID , DivisiName FROM TblDivisi " & _
'               "WHERE DivisiName = '" & cboDivisi(1).Text & "' "
'    End Select
'
'    rstemp.Open sSQL, conn, adOpenDynamic, adLockOptimistic
'    If rstemp.EOF = True And rstemp.BOF = True Then Exit Sub
'
'    cboDivisi(0).Text = rstemp("DivisiID")
'    cboDivisi(1).Text = rstemp("DivisiName")
'
'    Set rstemp = Nothing
End Sub

Private Sub DataComboDivisi()
'    cboDivisi(0).Clear
'    cboDivisi(1).Clear
    
'    Set rstemp = New ADODB.Recordset
'    rstemp.CursorLocation = adUseClient
'
'    sSQL = "SELECT * FROM DATASOURCETBL " & _
'           "WHERE KODEDS = '" & cboCompany(0).Text & "'"
'    rstemp.Open sSQL, conn, adOpenDynamic, adLockOptimistic
'
'    If rstemp.RecordCount < 1 Then Exit Sub
    
'    rstemp.MoveFirst
'    While Not rstemp.EOF
'        cboDivisi(0).AddItem rstemp("DivisiID")
'        cboDivisi(1).AddItem rstemp("DivisiName")
'        rstemp.MoveNext
'    Wend
    
'    Set rstemp = Nothing
End Sub

Private Sub ShowDataCombo()
    Call DataComboCompany
End Sub

Private Sub Form_Load()
    'Skin1.ApplySkin Me.hWnd
    
    
    Call ShowDataCombo
ProgressBar1.Visible = False
    
End Sub

'Private Function SetPajak(ByRef BtsAwal1 As Currency, ByRef BtsAwal2 As Currency, _
'                          ByRef BtsAwal3 As Currency, ByRef BtsAwal4 As Currency, _
'                          ByRef BtsAwal5 As Currency, ByRef BtsAkhir1 As Currency, _
'                          ByRef BtsAkhir2 As Currency, ByRef BtsAkhir3 As Currency, _
'                          ByRef BtsAkhir4 As Currency, ByRef BtsAkhir5 As Currency, _
'                          ByRef PPH1 As Currency, ByRef PPH2 As Currency, _
'                          ByRef PPH3 As Currency, ByRef PPH4 As Currency, _
'                          ByRef PPH5 As Currency, ByRef PTKPpribadi As Currency, _
'                          ByRef PTKPkeluarga As Currency, ByRef PersenBiayaJb As Currency, _
'                          ByRef MaxJb As Currency)
'
'    Set rsPajak = New ADODB.Recordset
'    rsPajak.CursorLocation = adUseClient
'
'    sSQL = "SELECT * FROM TblSettingPajak "
'    rsPajak.Open sSQL, conn, adOpenDynamic, adLockOptimistic
'
'    If rsPajak.EOF = True And rsPajak.BOF = True Then Exit Function
'
'    BtsAwal1 = rsPajak("BatasAwal1")
'    BtsAwal2 = rsPajak("BatasAwal2")
'    BtsAwal3 = rsPajak("BatasAwal3")
'    BtsAwal4 = rsPajak("BatasAwal4")
'    BtsAwal5 = rsPajak("BatasAwal5")
'
'    BtsAkhir1 = rsPajak("BatasAkhir1")
'    BtsAkhir2 = rsPajak("BatasAkhir2")
'    BtsAkhir3 = rsPajak("BatasAkhir3")
'    BtsAkhir4 = rsPajak("BatasAkhir4")
'    BtsAkhir5 = rsPajak("BatasAkhir5")
'
'    PPH1 = rsPajak("PPH1")
'    PPH2 = rsPajak("PPH2")
'    PPH3 = rsPajak("PPH3")
'    PPH4 = rsPajak("PPH4")
'    PPH5 = rsPajak("PPH5")
'
'    PTKPpribadi = rsPajak("PTKPpribadi")
'    PTKPkeluarga = rsPajak("PTKPkeluarga")
'
'    PersenBiayaJb = rsPajak("PersenBiayaJb")
'    MaxJb = rsPajak("MaxJb")
'
'    Set rsPajak = Nothing
'End Function
'
'Private Sub HitungPajakBln(ByVal NIK As String)
'    Dim BtsAwal1 As Currency
'    Dim BtsAwal2 As Currency
'    Dim BtsAwal3 As Currency
'    Dim BtsAwal4 As Currency
'    Dim BtsAwal5 As Currency
'    Dim BtsAkhir1 As Currency
'    Dim BtsAkhir2 As Currency
'    Dim BtsAkhir3 As Currency
'    Dim BtsAkhir4 As Currency
'    Dim BtsAkhir5 As Currency
'    Dim PPH1 As Currency
'    Dim PPH2 As Currency
'    Dim PPH3 As Currency
'    Dim PPH4 As Currency
'    Dim PPH5 As Currency
'
'    Dim PTKPpribadi As Currency
'    Dim PTKPkeluarga As Currency
'    Dim PersenBiayaJb As Currency
'    Dim MaxJb As Currency
'    Dim TotalJamsostek As Currency
'    Dim MasaKerja As Integer
'    Dim BiayaJb As Currency
'
'    Dim PTKP As Currency
'    Dim PPH As Currency
'    Dim PPHsThn As Currency
'    Dim PdpGross As Currency
'    Dim PdpNetto As Currency
'    Dim Pkp1Thn As Currency
'    Dim PKP As Currency
'
'    Call SetPajak(BtsAwal1, BtsAwal2, BtsAwal3, BtsAwal4, BtsAwal5, _
'                  BtsAkhir1, BtsAkhir2, BtsAkhir3, BtsAkhir4, BtsAkhir5, _
'                  PPH1, PPH2, PPH3, PPH4, PPH5, PTKPpribadi, PTKPkeluarga, _
'                  PersenBiayaJb, MaxJb)
'
'    Set rsPajak = New ADODB.Recordset
'    rsPajak.CursorLocation = adUseClient
'
'    sSQL = "SELECT TblGaji.*, TblJamsostek.WajibKaryawan FROM TblGaji " & _
'           "LEFT JOIN TblJamsostek ON TblGaji.NIK = TblJamsostek.NIK " & _
'           "WHERE TblGaji.CompID = '" & cboCompany(0).Text & "' " & _
'           "AND TblGaji.DivisiID = '" & cboDivisi(0).Text & "' " & _
'           "AND TglGaji = '" & Format(tdbUpload.Value, "yyyy/mm/01") & "' " & _
'           "AND TblGaji.NIK = '" & NIK & "'"
'    rsPajak.Open sSQL, conn, adOpenDynamic, adLockOptimistic
'
'    If rsPajak.EOF = True And rsPajak.BOF = True Then Exit Sub
'
'    PdpGross = 0
'    TotalJamsostek = 0
'
''    --- Hitung Pendapatan Gross 1 Bulan ---
'    For i = 14 To 26
'        PdpGross = PdpGross + IIf(i > 24, -IIf(IsNull(rsPajak(i)), 0, rsPajak(i)), IIf(IsNull(rsPajak(i)), 0, rsPajak(i)))
'    Next i
'    TotalJamsostek = IIf(IsNull(rsPajak("WajibKaryawan")), 0, rsPajak("WajibKaryawan"))
'
''    --- Hitung Masa Kerja ---
'    If Year(rsPajak("JoinDate")) = tdbUpload.Year Then
'        MasaKerja = 12 - Month(rsPajak("joinDate")) + 1
'    Else
'        If tdbUpload.Year > Year(rsPajak("JoinDate")) Then
'            MasaKerja = 12
'        End If
'    End If
'
''    --- Hitung Biaya Jabatan ---
'    If rsPajak("JbID") <> "J01" Or rsPajak("JbID") <> "J02" Then
'        BiayaJb = (PersenBiayaJb / 100) * PdpGross
'        If BiayaJb > MaxJb / 12 Then BiayaJb = MaxJb / 12
'    Else
'        BiayaJb = 0
'    End If
'
'    PdpNetto = Round(PdpGross - BiayaJb - TotalJamsostek)
'
'    Select Case rsPajak("Status")
'    Case "TK"
'        PTKP = PTKPpribadi
'    Case "K0"
'        PTKP = PTKPpribadi + PTKPkeluarga
'    Case "K1"
'        PTKP = PTKPpribadi + (PTKPkeluarga * 2)
'    Case "K2"
'        PTKP = PTKPpribadi + (PTKPkeluarga * 3)
'    Case "K3"
'        PTKP = PTKPpribadi + (PTKPkeluarga * 4)
'    End Select
'
'    PKP = Round(PdpNetto - PTKP)
'    Pkp1Thn = Round(PKP * MasaKerja)
'
'    If Pkp1Thn < 0 Then
'        Pkp1Thn = 0
'    Else
'        If Pkp1Thn >= BtsAwal1 And Pkp1Thn <= BtsAkhir1 Then
'            PPHsThn = (PPH1 / 100) * Pkp1Thn
'        Else
'            If Pkp1Thn >= BtsAwal2 And Pkp1Thn <= BtsAkhir2 Then
'                PPHsThn = (PPH1 / 100) * BtsAkhir1
'                PPHsThn = ((PPH2 / 100) * (Pkp1Thn - BtsAkhir1)) + PPH
'            Else
'                If Pkp1Thn >= BtsAwal3 And Pkp1Thn <= BtsAkhir3 Then
'                    PPHsThn = (PPH1 / 100) * BtsAkhir1
'                    PPHsThn = ((PPH2 / 100) * (BtsAkhir2 - BtsAkhir1)) + PPH
'                    PPHsThn = ((PPH3 / 100) * (Pkp1Thn - BtsAkhir2)) + PPH
'                Else
'                    If Pkp1Thn >= BtsAwal4 And Pkp1Thn <= BtsAkhir4 Then
'                        PPHsThn = (PPH1 / 100) * BtsAkhir1
'                        PPHsThn = ((PPH2 / 100) * (BtsAkhir2 - BtsAkhir1)) + PPH
'                        PPHsThn = ((PPH3 / 100) * (BtsAkhir3 - BtsAkhir2)) + PPH
'                        PPHsThn = ((PPH4 / 100) * (Pkp1Thn - BtsAkhir3)) + PPH
'                    Else
'                        If Pkp1Thn >= BtsAwal5 Then
'                            PPHsThn = (PPH1 / 100) * BtsAkhir1
'                            PPHsThn = ((PPH2 / 100) * (BtsAkhir2 - BtsAkhir1)) + PPH
'                            PPHsThn = ((PPH3 / 100) * (BtsAkhir3 - BtsAkhir2)) + PPH
'                            PPHsThn = ((PPH4 / 100) * (BtsAkhir4 - BtsAkhir3)) + PPH
'                            PPHsThn = ((PPH5 / 100) * (Pkp1Thn - BtsAkhir4)) + PPH
'                        End If
'                    End If
'                End If
'            End If
'        End If
'    End If
'
'    PPHsThn = Round(PPHsThn)
'    PPH = Round(PPHsThn / MasaKerja)
'
'    Call InsertPajakBulanan(NIK, PPH, PPHsThn, PTKP, Pkp1Thn, PdpNetto, _
'                            BiayaJb, MasaKerja, TotalJamsostek)
'
'End Sub

Private Sub InsertPajakBulanan(ByVal NIK As String, ByVal PPH As Currency, _
                               ByVal PPHsThn As Currency, ByVal PTKP As Currency, _
                               ByVal PKP As Currency, ByVal pGrossSThn As Currency, _
                               ByVal BiayaJb As Currency, ByVal MasaKerja As Integer, _
                               ByVal TotalJamsostek As Currency)
        
'    sSQL = "INSERT INTO TblPajakBln "
'    sSQL = sSQL & "(NIK, TglPajakBln, CompID, DivisiID, PPH, PPHsThn, PTKP, "
'    sSQL = sSQL & "PKP, pGrossSThn, TotalJamsostek, BiayaJb, MasaKerja, UserID, "
'    sSQL = sSQL & "TglSource) "
'    sSQL = sSQL & "VALUES "
'    sSQL = sSQL & "('" & NIK & "', '" & Format(tdbUpload.Value, "yyyy/mm/01") & "', "
'    sSQL = sSQL & "'" & cboCompany(0).Text & "', '" & cboDivisi(0).Text & "', "
'    sSQL = sSQL & "'" & PPH & "', '" & PPHsThn & "', '" & PTKP & "',"
'    sSQL = sSQL & "'" & PKP & "', '" & pGrossSThn & "', '" & ReplaceAngka(TotalJamsostek) & "',"
'    sSQL = sSQL & "'" & ReplaceAngka(BiayaJb) & "', '" & MasaKerja & "', '" & NAMA & "', "
'    sSQL = sSQL & "'" & Format(Date, "yyyy/mm/dd") & "')"
'    conn.Execute sSQL
End Sub

Private Sub SSCommand1_Click(Index As Integer)
 Dim mExcelFile As String
    Dim mAccessFile As String
    Dim mWorkSheet As String
    Dim mTableName As String
    Dim mDataBase As Database
    Dim m_msgbox As VbMsgBoxResult
    
    
 Select Case Index
Case 0
On Error GoTo errHandle
        
    'If CheckInsertedValue = False Then
     '   Call msgUnComplete
      '  Exit Sub
    'End If
    ProgressBar1.Visible = True
    mExcelFile = txtFile.Text
    mAccessFile = App.Path & "\upload_new.mdb"
    mWorkSheet = "Sheet1"
    
    mTableName = "TblTemp"
    
    'Below you may use "Excel 7.0" or 8.0 depending on your installable ISAM.
    
    Set mDataBase = OpenDatabase(mExcelFile, True, False, "Excel 8.0")
    mDataBase.Execute "SELECT * INTO [;database=" & mAccessFile & "]." & mTableName & _
        " FROM [" & mWorkSheet & "$]"
    
    Call ImportToSQL(mTableName)
    
    Exit Sub
    
errHandle:
    If Err.Number = 3010 Then
        mDataBase.Execute "DROP TABLE [;database=" & mAccessFile & "]." & mTableName
        Resume
    ElseIf Err.Number = -2147217865 Then
        MsgBox Err.Description
        Resume
    
    ElseIf Right(txtFile.Text, 4) <> ".xls" Or txtFile.Text = Empty Then
            MsgBox "Pilih File Excel Yang akan di Upload..!"
            txtFile.SetFocus
     ElseIf Text1.Text = "" Then
            MsgBox "Field Data Harus Di isi..!"
            Text1.SetFocus
            
    Else
        co.RollbackTrans
        MsgBox Err.Number & "  " & Err.Description
        m_msgbox = MsgBox("Ulangi Lagi??", vbInformation + vbRetryCancel, "Konfirmasi")
        If m_msgbox = vbRetry Then
            conn.RollbackTrans
            Resume
        End If
    End If
    
 Case 1
  Unload Me
End Select

End Sub

Private Sub Timer1_Timer()
Dim rs As ADODB.Recordset
    Dim JUMLAH As Long
    Dim CMDSQL As String
    openConn
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    CMDSQL = "select * from tbllunas"
    rs.Open CMDSQL, conn, adOpenDynamic, adLockOptimistic
    Label2.Caption = rs.RecordCount
    Set rs = Nothing
    
If Label2.Visible = True Then
    Label2.Visible = False
    Else
    Label2.Visible = True
End If

End Sub
