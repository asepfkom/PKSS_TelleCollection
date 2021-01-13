VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frnubahstsaccount 
   Caption         =   "Change Status Account"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   6090
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Upload"
      Height          =   5565
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   11835
      Begin VB.ComboBox cbostsacc 
         Height          =   315
         Left            =   4770
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   1020
         Width           =   1785
      End
      Begin VB.Frame Frame2 
         Caption         =   "Transfer Account"
         Height          =   3675
         Left            =   180
         TabIndex        =   11
         Top             =   1830
         Width           =   11625
         Begin MSComDlg.CommonDialog Cdupdate 
            Left            =   10560
            Top             =   2520
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   10020
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1410
            Width           =   1545
         End
         Begin VB.TextBox txtjml 
            Height          =   285
            Left            =   1200
            TabIndex        =   29
            Top             =   3360
            Width           =   1515
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00F1E5DB&
            Caption         =   "Search"
            Height          =   285
            Left            =   8970
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   570
            Width           =   1035
         End
         Begin VB.ComboBox CBOACCOUNT 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5250
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   570
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   945
            TabIndex        =   22
            Top             =   570
            Width           =   2925
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   2
            Left            =   960
            MaxLength       =   20
            TabIndex        =   21
            Top             =   210
            Width           =   1965
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00F1E5DB&
            Caption         =   "Change Account"
            Height          =   495
            Left            =   9990
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   900
            Width           =   1515
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<<"
            Height          =   375
            Index           =   3
            Left            =   4740
            TabIndex        =   17
            Top             =   2100
            Width           =   1095
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">>"
            Height          =   375
            Index           =   2
            Left            =   4740
            TabIndex        =   16
            Top             =   1710
            Width           =   1095
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<"
            Height          =   375
            Index           =   1
            Left            =   4740
            TabIndex        =   15
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">"
            Height          =   375
            Index           =   0
            Left            =   4740
            TabIndex        =   14
            Top             =   930
            Width           =   1095
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2415
            Left            =   150
            TabIndex        =   12
            Top             =   900
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   4260
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   2445
            Left            =   5850
            TabIndex        =   13
            Top             =   870
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   4313
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label Label3 
            Caption         =   "Jml Lead"
            Height          =   225
            Left            =   390
            TabIndex        =   28
            Top             =   3390
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Status Telp."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   8
            Left            =   3720
            TabIndex        =   26
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   0
            Left            =   -660
            TabIndex        =   24
            Top             =   570
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "# Kartu :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   5
            Left            =   -330
            TabIndex        =   23
            Top             =   255
            Width           =   1125
         End
      End
      Begin VB.TextBox TxtJmlData 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   1050
         Width           =   1095
      End
      Begin VB.CommandButton CmdUpdateStatus 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&Upload..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   9420
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   1065
      End
      Begin VB.CommandButton CmdBrowse 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&Browse..."
         Height          =   345
         Left            =   9420
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1065
      End
      Begin VB.ComboBox CmbSheet 
         Height          =   315
         Left            =   2190
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox TxtPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2190
         TabIndex        =   1
         Top             =   210
         Width           =   6015
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   1380
         Visible         =   0   'False
         Width           =   9720
         _ExtentX        =   17145
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Status Account"
         Height          =   285
         Left            =   3510
         TabIndex        =   19
         Top             =   1050
         Width           =   1425
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
         TabIndex        =   8
         Top             =   1020
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
         TabIndex        =   7
         Top             =   240
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
         TabIndex        =   6
         Top             =   660
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   5
      Left            =   60
      Picture         =   "frnubahstsaccount.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   420
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Upload Untuk Ubah Status Account"
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
      TabIndex        =   9
      Top             =   30
      Width           =   5865
   End
   Begin VB.Image Image2 
      Height          =   435
      Index           =   8
      Left            =   0
      Picture         =   "frnubahstsaccount.frx":0B0A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20700
   End
End
Attribute VB_Name = "frnubahstsaccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click(Index As Integer)
Select Case Index
Case 1
If ListView2.ListItems.Count <> 0 Then
            Set lList = ListView1.ListItems.ADD(, , ListView2.SelectedItem.text)
            ListView2.ListItems.Remove ListView2.SelectedItem.Index
            End If
Case 3
 For i = 1 To ListView2.ListItems.Count
                Set lList = ListView1.ListItems.ADD(, , ListView2.SelectedItem.text)
                ListView2.ListItems.Remove ListView2.SelectedItem.Index
            Next
Case 0
If ListView1.ListItems.Count <> 0 Then
            Set lList = ListView2.ListItems.ADD(, , ListView1.SelectedItem.text)
           ListView1.ListItems.Remove ListView1.SelectedItem.Index
End If

Case 2
    For i = 1 To ListView1.ListItems.Count
            Set lList = ListView2.ListItems.ADD(, , ListView1.SelectedItem.text)
                   DoEvents
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
        Next
End Select

End Sub

Private Sub cmdbrowse_Click()
Form_Save:
    With Cdupdate
    .CancelError = False
    .DialogTitle = "Cari data masukan Upload data"
    'On Error GoTo X
    '.Filter = "Excel Files|*.xls;*.xlsx"
    .Filter = "Ms. Excel 9|*.xls"
    .ShowOpen
    TxtPath.text = .FileName
    End With
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.text = Empty Then
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
 CmdUpdateStatus.Enabled = True

End Sub

Private Sub CmdUpdateStatus_Click()
 Dim MOBJ As New ADODB.Recordset
 Dim koneksi_excel As New ADODB.Connection
 Set koneksi_excel = New ADODB.Connection
    koneksi_excel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & TxtPath.text & _
                       ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
   
   Set MOBJ = New ADODB.Recordset
   MOBJ.CursorLocation = adUseClient
   
    '-> Membuka recordset Ms.Excel dengan status=gagal
    MOBJ.Open "Select * FROM [" & CmbSheet.text & "]", _
                         koneksi_excel, adOpenStatic, adLockOptimistic, adCmdText
    TxtJmlData.text = MOBJ.RecordCount
    ProgressBar1.Max = MOBJ.RecordCount + 1
    While Not MOBJ.EOF
     ProgressBar1.Value = MOBJ.Bookmark
    DoEvents
    strsql = ""
      If Trim(Left(cbostsacc.text, 3)) = "PTP" Then
            strsql = " UPDATE  MGM SET F_CEK_NEW='" + Left(Combo1.text, 6) + "' ,KETHSLKERJA_NEW='" + Combo1.text + " ',KETHSLKERJADESC_NEW='" + Combo1.text + " ' WHERE CUSTID='" + MOBJ(0).Value + "'"
           ' M_OBJCONN.Execute (STRSQL)
        Else
            strsql = " UPDATE  MGM SET F_CEK_NEW='" + Left(Combo1.text, 3) + "' ,KETHSLKERJA_NEW='" + Combo1.text + " ',KETHSLKERJADESC_NEW='" + Combo1.text + " ' WHERE CUSTID='" + MOBJ(0).Value + "'"
           ' M_OBJCONN.Execute (STRSQL)
        End If
        
            M_OBJCONN.Execute (strsql)
       
        MOBJ.MoveNext
    Wend
    MsgBox "Data telah di Markup", vbInformation + vbOKOnly, "Pesan"
    CmdUpdateStatus.Enabled = False

End Sub

Private Sub Command1_Click()
Dim IJ As Double
If ListView2.ListItems.Count <> 0 Then
    For IJ = 1 To ListView2.ListItems.Count
        If Trim(Left(Combo1.text, 3)) = "PTP" Then
            strsql = " UPDATE  MGM SET F_CEK_NEW='" + Left(Combo1.text, 6) + "' ,KETHSLKERJA_NEW='" + Combo1.text + " ',KETHSLKERJADESC_NEW='" + Combo1.text + " ' WHERE CUSTID='" + ListView2.ListItems(IJ).text + "'"
           ' M_OBJCONN.Execute (STRSQL)
        Else
            strsql = " UPDATE  MGM SET F_CEK_NEW='" + Left(Combo1.text, 3) + "' ,KETHSLKERJA_NEW='" + Combo1.text + " ',KETHSLKERJADESC_NEW='" + Combo1.text + " ' WHERE CUSTID='" + ListView2.ListItems(IJ).text + "'"
           ' M_OBJCONN.Execute (STRSQL)
        End If
    Next IJ
    Debug.Print strsql
End If

End Sub

Private Sub Command2_Click()
Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim SqlCriteria As String
    Dim ListItem As ListItem
    Dim no As Integer
    
    SqlCriteria = ""
    
    If Text1(2).text <> "" Then
        If Len(SqlCriteria) = 0 Then
            SqlCriteria = " where CUSTID like '%"
            SqlCriteria = SqlCriteria + Trim(Text1(2).text) + "%'"
        Else
            SqlCriteria = SqlCriteria + " and CUSTID like '%"
            SqlCriteria = SqlCriteria + Trim(Text1(2).text) + "%'"
        End If
    End If
    
    If Text1(0).text <> "" Then
        If Len(SqlCriteria) = 0 Then
            SqlCriteria = " where NAME like '%"
            SqlCriteria = SqlCriteria + Trim(Text1(0).text) + "%'"
        Else
            SqlCriteria = SqlCriteria + " and NAME like '%"
            SqlCriteria = SqlCriteria + Trim(Text1(0).text) + "%'"
        End If
    End If
    
    
    If cboaccount.text <> "" Then
        If Len(SqlCriteria) = 0 Then
            SqlCriteria = " where KETHSLKERJA_NEW  like '%"
            SqlCriteria = SqlCriteria + Trim(cboaccount.text) + "%'"
        Else
            SqlCriteria = SqlCriteria + " and KETHSLKERJA_NEW like '%"
            SqlCriteria = SqlCriteria + Trim(cboaccount.text) + "%'"
        End If
    End If
    
    
    
    CMDSQL = "select * from MGM  "
    CMDSQL = CMDSQL + SqlCriteria
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
       
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data yang dicari tidak ditemukan!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        Exit Sub
    End If
    ListView1.ListItems.CLEAR
    no = 0
    txtjml.text = M_objrs.RecordCount
    While Not M_objrs.EOF
        no = no + 1
        Set ListItem = ListView1.ListItems.ADD(, , M_objrs("CUSTID"))
            ListItem.SubItems(1) = IIf(IsNull(M_objrs("name")), "", M_objrs("name"))
            
        M_objrs.MoveNext
    Wend
End Sub
Private Sub Form_Load()
   ListView1.ColumnHeaders.ADD 1, , "Custid", 10 * 200
   ListView1.ColumnHeaders.ADD 2, , "Nama", 10 * 200
   ListView2.ColumnHeaders.ADD 1, , "Custid", 20 * 200
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    strsql = "Select * from contacteddesc WHERE status=1"
    cbostsacc.CLEAR
    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_objrs.EOF
        cboaccount.AddItem M_objrs!KdNoProdPresented
        Combo1.AddItem M_objrs!KdNoProdPresented
        cbostsacc.AddItem M_objrs!KdNoProdPresented
        M_objrs.MoveNext
    Wend
    cboaccount.AddItem "PTP-POP"
    cboaccount.AddItem "PTP-NEW"
    Combo1.AddItem "PTP-POP"
    Combo1.AddItem "PTP-NEW"
    cbostsacc.AddItem "PTP-POP"
    cbostsacc.AddItem "PTP-NEW"
    Set M_objrs = Nothing
End Sub
Private Sub isi_sheet()
    Set koneksi_excel = CreateObject("ADODB.Connection")
    Set recordsetexcel = CreateObject("ADODB.Recordset")

    '-> Koneksi ke Ms.Excel
    koneksi_excel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & TxtPath.text & _
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


