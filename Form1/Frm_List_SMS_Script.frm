VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Frm_List_SMS_Script 
   Caption         =   "List SMS Script"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4920
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtJml 
      Alignment       =   2  'Center
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4500
      Width           =   1245
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   3
      Top             =   1320
      Width           =   1395
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   720
      Width           =   1395
   End
   Begin VB.CommandButton CmdTambah 
      Caption         =   "&Tambah"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9120
      TabIndex        =   1
      Top             =   150
      Width           =   1395
   End
   Begin MSComctlLib.ListView LvScriptSms 
      Height          =   4275
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   7541
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
   Begin VB.Label Label1 
      Caption         =   "Jumlah data:"
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
      Left            =   180
      TabIndex        =   4
      Top             =   4560
      Width           =   1035
   End
End
Attribute VB_Name = "Frm_List_SMS_Script"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub headerscriptsms()
    LvScriptSms.ColumnHeaders.ADD , , "Option", 3000
    LvScriptSms.ColumnHeaders.ADD , , "Sub option", 5000
    LvScriptSms.ColumnHeaders.ADD , , "Script SMS", 10000
    LvScriptSms.ColumnHeaders.ADD , , "ID", 500
End Sub

Private Sub LoadScriptSms()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    CMDSQL = "select * from tblscriptsms"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    txtjml.Text = M_objrs.RecordCount
    
    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    While Not M_objrs.EOF
        Set ListItem = LvScriptSms.ListItems.ADD(, , Trim(M_objrs("option")))
            ListItem.SubItems(1) = Trim(M_objrs("suboption"))
            ListItem.SubItems(2) = Trim(M_objrs("scriptsms"))
            ListItem.SubItems(3) = Trim(M_objrs("id"))
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub cmdEdit_Click()
    Dim CMDSQL As String
    
    If LvScriptSms.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    With Frm_Script_SMS
        .Caption = "Edit script sms"
        .CmbOption.Text = LvScriptSms.SelectedItem.Text
        .TxtSubOption.Text = LvScriptSms.SelectedItem.SubItems(1)
        .TxtSms.Text = LvScriptSms.SelectedItem.SubItems(2)
        .ok = False
        .Show vbModal
        If .ok Then
            CMDSQL = "update tblscriptsms set option='"
            CMDSQL = CMDSQL + Trim(.CmbOption.Text) + "',suboption='"
            CMDSQL = CMDSQL + Trim(.TxtSubOption.Text) + "',scriptsms='"
            CMDSQL = CMDSQL + Trim(.TxtSms.Text) + "' where id=" & LvScriptSms.SelectedItem.SubItems(3) & ""
            
            M_OBJCONN.Execute CMDSQL
            
            MsgBox "Data berhasil diupdate!", vbOKOnly + vbInformation, "Informasi"
            
            LvScriptSms.SelectedItem.Text = Trim(.CmbOption.Text)
            LvScriptSms.SelectedItem.SubItems(1) = Trim(.TxtSubOption.Text)
            LvScriptSms.SelectedItem.SubItems(2) = Trim(.TxtSms.Text)
        End If
    End With
End Sub

Private Sub cmdHapus_Click()
    Dim CMDSQL As String
    Dim m_msgbox As String
    
    If LvScriptSms.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    m_msgbox = MsgBox("Yakin data akan dihapus?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If m_msgbox = vbYes Then
        CMDSQL = "delete from tblscriptsms where id=" & LvScriptSms.SelectedItem.SubItems(3) & ""
        
        M_OBJCONN.Execute CMDSQL
        LvScriptSms.ListItems.Remove LvScriptSms.SelectedItem.Index
        txtjml.Text = Val(txtjml.Text) - 1
    End If
    
End Sub

Private Sub cmdkeluar_Click()
    Unload Me
End Sub

Private Sub cmdTambah_Click()
    Dim CMDSQL As String
    Dim ListItem As ListItem
    
    With Frm_Script_SMS
        .Caption = "Tambah script sms"
        .CmbOption.Text = ""
        .TxtSubOption.Text = ""
        .TxtSms.Text = ""
        .Show vbModal
        If .ok Then
            CMDSQL = "insert into tblscriptsms (option,suboption,scriptsms) values ('"
            CMDSQL = CMDSQL + Trim(.CmbOption.Text) + "','"
            CMDSQL = CMDSQL + Trim(.TxtSubOption.Text) + "','"
            CMDSQL = CMDSQL + Trim(.TxtSms.Text) + "')"
            
            M_OBJCONN.Execute CMDSQL
            MsgBox "Data berhasil disimpan!", vbOKOnly + vbInformation, "Informasi"
            
           Set ListItem = LvScriptSms.ListItems.ADD(, , .CmbOption.Text)
               ListItem.SubItems(1) = .TxtSubOption.Text
               ListItem.SubItems(2) = .TxtSms.Text
          txtjml.Text = Val(txtjml.Text) + 1
        End If
    End With
End Sub

Private Sub Form_Load()
    headerscriptsms
    LoadScriptSms
End Sub

Private Sub LvScriptSms_DblClick()
    cmdEdit_Click
End Sub

