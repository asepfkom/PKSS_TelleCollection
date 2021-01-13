VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRMMARKUP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Upload Untuk Focus Calling"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10020
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Upload"
      Height          =   6225
      Left            =   30
      TabIndex        =   1
      Top             =   450
      Width           =   9945
      Begin MSComDlg.CommonDialog Cdupdate 
         Left            =   7200
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid GridDetail 
         Height          =   4215
         Left            =   3840
         TabIndex        =   18
         Top             =   1440
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   7435
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
         Caption         =   "Detail Total Markup"
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
      Begin MSDataGridLib.DataGrid GridData 
         Height          =   4275
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   7541
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
         Caption         =   "GRID DATA"
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
      Begin VB.TextBox txttrigger 
         Height          =   285
         Left            =   5070
         TabIndex        =   11
         Top             =   1050
         Width           =   1965
      End
      Begin VB.TextBox TxtPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2190
         TabIndex        =   6
         Top             =   210
         Width           =   6015
      End
      Begin VB.ComboBox CmbSheet 
         Height          =   315
         Left            =   2190
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton CmdBrowse 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&Browse..."
         Height          =   345
         Left            =   8250
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   1065
      End
      Begin VB.TextBox TxtJmlData 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         Top             =   1050
         Width           =   1095
      End
      Begin VB.CommandButton CmdUpdateStatus 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&Upload..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   8250
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   60
         TabIndex        =   10
         Top             =   5880
         Visible         =   0   'False
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.ComboBox cbotype 
         Height          =   315
         Left            =   7980
         TabIndex        =   13
         Top             =   1140
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Field Diexcel"
         Height          =   285
         Left            =   7860
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   1365
      End
      Begin MSComctlLib.ListView LVCustid 
         Height          =   4260
         Left            =   7080
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   7514
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Campaign"
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
         Left            =   3360
         TabIndex        =   15
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pilih Sheet Excel :"
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
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "File excel:"
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
         Index           =   3
         Left            =   150
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jumlah data :"
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
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kriteria Manual List :"
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
         Index           =   0
         Left            =   7980
         TabIndex        =   12
         Top             =   1140
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Upload Untuk Focus Calling"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   540
      TabIndex        =   0
      Top             =   30
      Width           =   3405
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   5
      Left            =   60
      Picture         =   "FRMMARKUP.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   435
      Index           =   8
      Left            =   0
      Picture         =   "FRMMARKUP.frx":0B0A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20700
   End
End
Attribute VB_Name = "FRMMARKUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmbSheet_Click()
    Dim koneksi_excel As New ADODB.Connection
    Dim M_Obj As ADODB.Recordset
    Dim TempCustid As String
    Dim ListItem As ListItem
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    On Error GoTo Salah
    If CmbSheet.Text = "" Then
        MsgBox "Sheet tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    TempCustid = ""
    Set koneksi_excel = New ADODB.Connection
    koneksi_excel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & TxtPath.Text & _
                       ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
                       
    Set M_Obj = New ADODB.Recordset
    M_Obj.CursorLocation = adUseClient
    M_Obj.Open "Select * FROM [" & CmbSheet.Text & "] where [custid] is not null or [custid]='' ", _
                         koneksi_excel, adOpenStatic, adLockOptimistic, adCmdText
                         
    TxtJmlData.Text = M_Obj.RecordCount
    Set GridData.DATASOURCE = M_Obj
    
    'Jika data tidak ditemukan
    If M_Obj.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Set M_Obj = Nothing
        Set koneksi_excel = Nothing
        CmdUpdateStatus.Enabled = False
        Exit Sub
    End If
    
    
    If M_Obj.RecordCount > 0 Then
        ProgressBar1.Max = M_Obj.RecordCount
        While Not M_Obj.EOF
            ProgressBar1.Value = M_Obj.Bookmark
            If TempCustid = "" Then
                TempCustid = "'" & Trim(M_Obj(0)) & "'"
            Else
                TempCustid = TempCustid & ",'" & Trim(M_Obj(0)) & "'"
            End If
            
            Set ListItem = LVCustid.ListItems.ADD(, , Trim(M_Obj(0)))
            M_Obj.MoveNext
        Wend
        
        CMDSQL = "select agent,count(custid) as jumlah from mandiri.mgm where custid in ("
        CMDSQL = CMDSQL + TempCustid + ") group by agent order by agent asc"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        Set GridDetail.DATASOURCE = M_objrs
    End If
    CmdUpdateStatus.Enabled = True
    Set M_Obj = Nothing
    Set M_objrs = Nothing
    Set koneksi_excel = Nothing
    Exit Sub
Salah:
    MsgBox "Ada kesalahan: " & err.Description, vbOKOnly + vbInformation, "Informasi"
    Set koneksi_excel = Nothing
    CmdUpdateStatus.Enabled = False
End Sub

Private Sub CmdBrowse_Click()
Form_Save:
    With Cdupdate
    .CancelError = False
    .DialogTitle = "Cari data masukan Upload data"
    'On Error GoTo X
    .Filter = "Ms. Excel 9|*.xls"
    .ShowOpen
    TxtPath.Text = .FileName
    End With
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.Text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Update dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Update dibatalkan!", vbOKOnly + vbInformation, "Informasi"
              CmdUpdateStatus.Enabled = False
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo Form_Save        '-> maka goto form_save
        End If
    End If
 Call isi_sheet
 'CmdUpdateStatus.Enabled = True
End Sub

'Private Sub CmdUpdateStatus_Click()
''    Dim mobj As New ADODB.Recordset
''    Dim koneksi_excel As New ADODB.Connection
''    Set koneksi_excel = New ADODB.Connection
''    koneksi_excel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
''                       "Data Source=" & TxtPath.Text & _
''                       ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
''
''   Set mobj = New ADODB.Recordset
''   mobj.CursorLocation = adUseClient
''
''    '-> Membuka recordset Ms.Excel dengan status=gagal
''    mobj.Open "Select * FROM [" & CmbSheet.Text & "] where [custid] is not null or [custid]='' ", _
''                         koneksi_excel, adOpenStatic, adLockOptimistic, adCmdText
''    If mobj.RecordCount = 0 Then
''        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
''        Exit Sub
''    End If
''
''    TxtJmlData.Text = mobj.RecordCount
''    ProgressBar1.Max = mobj.RecordCount + 1
''    While Not mobj.EOF
''        ProgressBar1.Value = mobj.Bookmark
''        DoEvents
''        Strsql = ""
'''        If IIf(IsNull(mobj(0).Value), "", mobj(0).Value) <> "" Then
'''           'ddiskon = IIf(IsNull(mobj(1).Value), 0, mobj(1).Value)
'''            If Check1.Value = vbChecked Then
'''                ddiskon = IIf(IsNull(mobj(1).Value), 0, mobj(1).Value)
'''                Strsql = "update mgm set disapp=" & CStr(ddiskon) & ",typerumusdiscount ='Y', exclude='" + txttrigger.Text + "',typerumus='" + cbotype.Text + "' where custid='" + Trim(mobj(0).Value) + "'"
'''            Else
'''                Strsql = "update mgm set  exclude='" + txttrigger.Text + "',typerumus='" + cbotype.Text + "' where custid='" + Trim(mobj(0).Value) + "'"
'''            End If
'''            M_OBJCONN.Execute (Strsql)
'''        End If
''        '@@13092012
''        Strsql = "update mgm set  exclude='" + txttrigger.Text + "' where custid='" + Trim(mobj(0).Value) + "'"
''        mobj.MoveNext
''    Wend
''    MsgBox "Data telah di Markup", vbInformation + vbOKOnly, "Pesan"
''    CmdUpdateStatus.Enabled = False
'End Sub
Private Sub isi_sheet()
    Set koneksi_excel = CreateObject("ADODB.Connection")
    Set recordsetexcel = CreateObject("ADODB.Recordset")

    '-> Koneksi ke Ms.Excel
    koneksi_excel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & TxtPath.Text & _
                       ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
                       
    '-> Membuka recordset Ms.Excel dengan status=gagal
    Set recordsetexcel = koneksi_excel.OpenSchema(adSchemaTables)
       
       
                       
                         
    'Mengsisi sheet pada CmbSheet
    CmbSheet.CLEAR
    CmbSheet.AddItem ""
    
    While Not recordsetexcel.EOF
       If Left(recordsetexcel.fields("Table_Name").Value, 4) <> "MSys" And Left(recordsetexcel.fields("Table_Name").Value, 3) <> "Sys" Then
        CmbSheet.AddItem recordsetexcel.fields("Table_Name")
       End If
       recordsetexcel.MoveNext
    Wend
                       
End Sub
Public Sub AMBILRUMUS()
Dim RSNEW As New ADODB.Recordset
STRSQL = "SELECT DISTINCT(idkey)AS IDRMS FROM mandiri.tbloffering"
Set RSNEW = New ADODB.Recordset
RSNEW.CursorLocation = adUseClient
RSNEW.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not RSNEW.EOF
 cbotype.AddItem IIf(IsNull(RSNEW!IDRMS), "", RSNEW!IDRMS)
 RSNEW.MoveNext
Wend
Set RSNEW = Nothing

End Sub

Private Sub CmdUpdateStatus_Click()
    Dim K As Integer
    Dim CMDSQL As String
    
    On Error GoTo Salah
    If Trim(txttrigger.Text) = "" Then
        MsgBox "Campaign tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If Val(TxtJmlData.Text) > 0 Then
        ProgressBar1.Max = LVCustid.ListItems.Count
        For K = 1 To LVCustid.ListItems.Count
            ProgressBar1.Value = K
            CMDSQL = "update mandiri.mgm set  exclude='" + txttrigger.Text + "' where custid='" + CStr(Trim(LVCustid.ListItems(K).Text)) + "'"
            M_OBJCONN.Execute CMDSQL
        Next K
    End If
    MsgBox "Data berhasil di MarkUp!", vbOKOnly + vbInformation, "Informasi"
    Unload Me
    Exit Sub
Salah:
    MsgBox "Ada error: " & err.Description
End Sub

Private Sub Form_Load()
    AMBILRUMUS
    HeaderCustid
End Sub

Private Sub HeaderCustid()
    LVCustid.ColumnHeaders.ADD 1, , "Custid", 1000
End Sub
