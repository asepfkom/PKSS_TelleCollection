VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmListUnValidNumber 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form UnValid Number"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11520
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdUnCekAll 
      Caption         =   "&UnCek All"
      Height          =   315
      Left            =   7560
      TabIndex        =   11
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton CmdCekAll 
      Caption         =   "&Cek All"
      Height          =   315
      Left            =   6420
      TabIndex        =   10
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   315
      Left            =   5340
      TabIndex        =   9
      Top             =   180
      Width           =   675
   End
   Begin VB.TextBox TxtCustid 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   8
      Top             =   180
      Width           =   1515
   End
   Begin VB.TextBox TxtJmlh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Text            =   "0"
      Top             =   6300
      Width           =   1035
   End
   Begin VB.CommandButton CmdHapusUnValidNumber 
      Caption         =   "&Hapus Unvalid Number"
      Height          =   375
      Left            =   9480
      TabIndex        =   4
      Top             =   180
      Width           =   1935
   End
   Begin VB.CommandButton CmdCari 
      Caption         =   "&Cari"
      Height          =   315
      Left            =   4620
      TabIndex        =   3
      Top             =   180
      Width           =   675
   End
   Begin VB.TextBox TxtCariNoTelp 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   900
      TabIndex        =   2
      Top             =   180
      Width           =   1395
   End
   Begin MSComctlLib.ListView LVUnValidNumber 
      Height          =   5625
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   9922
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
      Caption         =   "Custid:"
      Height          =   195
      Left            =   2400
      TabIndex        =   7
      Top             =   240
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Jumlah Data:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   6300
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Cari Telp."
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmListUnValidNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
    With LVUnValidNumber.ColumnHeaders
        .ADD 1, , "ID", 1000
        .ADD 2, , "No.Telepon", 2000
        .ADD 3, , "Custid", 2000
        .ADD 4, , "Tgl.Input", 2000
        .ADD 5, , "Status", 1500
        .ADD 6, , "Keterangan", 2000
        .ADD 7, , "User Input", 2000
        .ADD 8, , "Telp Blok", 2000
    End With
End Sub

Private Sub IsiData()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    
    CMDSQL = "select * from tblunvalid_number where no_telp is not null  "
    
    If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        CMDSQL = CMDSQL + " and userid in (select userid from usertbl where team='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "' and usertype='1') "
    End If
    
    
    If TxtCariNoTelp.Text <> Empty Then
        CMDSQL = CMDSQL + " and no_telp like '%"
        CMDSQL = CMDSQL + CStr(Replace(TxtCariNoTelp.Text, " ", "")) + "%' "
    End If
    
    If txtcustid.Text <> Empty Then
        CMDSQL = CMDSQL + " and custid like '%"
        CMDSQL = CMDSQL + CStr(Replace(txtcustid.Text, " ", "")) + "%'"
    End If
    
    CMDSQL = CMDSQL + " order by tglinput,id desc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TxtJmlh.Text = M_objrs.RecordCount
    LVUnValidNumber.ListItems.CLEAR
    
    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    While Not M_objrs.EOF
        Set ListItem = LVUnValidNumber.ListItems.ADD(, , M_objrs("id"))
            ListItem.SubItems(1) = IIf(IsNull(M_objrs("no_telp")), "", M_objrs("no_telp"))
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("tglinput")), "", Format(M_objrs("tglinput"), "yyyy-mm-dd"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("status")), "", M_objrs("status"))
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("keterangan")), "", M_objrs("keterangan"))
            ListItem.SubItems(6) = IIf(IsNull(M_objrs("userinput")), "", M_objrs("userinput"))
            ListItem.SubItems(7) = IIf(IsNull(M_objrs("telpblok")), "", M_objrs("telpblok"))
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub CmdCari_Click()
    Call IsiData
End Sub

Private Sub HapusUnvalidNumber()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim W As Integer
    
    If LVUnValidNumber.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LVUnValidNumber.ListItems.Count
        If LVUnValidNumber.ListItems(W).Checked = True Then
            
            If LVUnValidNumber.ListItems(W).SubItems(7) <> Empty Then
                If LVUnValidNumber.ListItems(W).SubItems(7) = "Home 1" Then
                    CMDSQL = "update mgm set f_unvalid_home1=null,f_sts_unvalid_home1=null "
                    CMDSQL = CMDSQL + " where homeno='"
                    CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
                End If
                If LVUnValidNumber.ListItems(W).SubItems(7) = "Home 2" Then
                    CMDSQL = "update mgm set f_unvalid_home2=null,f_sts_unvalid_home2=null "
                    CMDSQL = CMDSQL + " where homeno2='"
                    CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
                End If
                If LVUnValidNumber.ListItems(W).SubItems(7) = "Office 1" Then
                    CMDSQL = "update mgm set f_unvalid_office1=null,f_sts_unvalid_office1=null "
                    CMDSQL = CMDSQL + " where officeno='"
                    CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
                End If
                If LVUnValidNumber.ListItems(W).SubItems(7) = "Office 2" Then
                    CMDSQL = "update mgm set f_unvalid_office2=null,f_sts_unvalid_office2=null "
                    CMDSQL = CMDSQL + " where officeno2='"
                    CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
                End If
                If LVUnValidNumber.ListItems(W).SubItems(7) = "Mobile 1" Then
                    CMDSQL = "update mgm set f_unvalid_mobile1=null,f_sts_unvalid_mobile1=null "
                    CMDSQL = CMDSQL + " where mobileno='"
                    CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
                End If
                If LVUnValidNumber.ListItems(W).SubItems(7) = "Mobile 2" Then
                    CMDSQL = "update mgm set f_unvalid_mobile2=null,f_sts_unvalid_mobile2=null "
                    CMDSQL = CMDSQL + " where mobileno2='"
                    CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
                End If
                If LVUnValidNumber.ListItems(W).SubItems(7) = "AddHome 1" Then
                    CMDSQL = "update mgm set f_unvalid_addhome1=null,f_sts_unvalid_addhome1=null "
                    CMDSQL = CMDSQL + " where homenoadd1='"
                    CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
                End If
                If LVUnValidNumber.ListItems(W).SubItems(7) = "AddHome 2" Then
                    CMDSQL = "update mgm set f_unvalid_addhome2=null,f_sts_unvalid_addhome2=null "
                    CMDSQL = CMDSQL + " where homenoadd2='"
                    CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
                End If
                If LVUnValidNumber.ListItems(W).SubItems(7) = "AddOffice 1" Then
                    CMDSQL = "update mgm set f_unvalid_addoffice1=null,f_sts_unvalid_addoffice1=null "
                    CMDSQL = CMDSQL + " where officenoadd1='"
                    CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
                End If
                If LVUnValidNumber.ListItems(W).SubItems(7) = "AddOffice 2" Then
                    CMDSQL = "update mgm set f_unvalid_addoffice2=null,f_sts_unvalid_addoffice2=null "
                    CMDSQL = CMDSQL + " where officenoadd2='"
                    CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
                End If
                If LVUnValidNumber.ListItems(W).SubItems(7) = "AddMobile 1" Then
                    CMDSQL = "update mgm set f_unvalid_addmobile1=null,f_sts_unvalid_addmobile1=null "
                    CMDSQL = CMDSQL + " where mobilenoadd1='"
                    CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
                End If
                If LVUnValidNumber.ListItems(W).SubItems(7) = "AddMobile 2" Then
                    CMDSQL = "update mgm set f_unvalid_addmobile2=null,f_sts_unvalid_addmobile2=null "
                    CMDSQL = CMDSQL + " where mobilenoadd2='"
                    CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
                End If
                If LVUnValidNumber.ListItems(W).SubItems(7) = "EC" Then
                    CMDSQL = "update mgm set f_unvalid_ec=null,f_sts_unvalid_ec=null "
                    CMDSQL = CMDSQL + " where ec_telp='"
                    CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
                End If
                If CMDSQL <> "" Then
                    M_OBJCONN.Execute CMDSQL
                End If
            End If
            
            CMDSQL = "DELETE FROM tblunvalid_number WHERE no_telp='"
            CMDSQL = CMDSQL + CStr(LVUnValidNumber.ListItems(W).SubItems(1)) + "'"
            M_OBJCONN.Execute CMDSQL
        End If
    Next W
    
    Call IsiData
    
    MsgBox "Proses selesai!", vbOKOnly + vbInformation, "Informasi"
    
End Sub

Private Sub CmdCekAll_Click()
    Dim W As Integer
    
    If LVUnValidNumber.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LVUnValidNumber.ListItems.Count
        LVUnValidNumber.ListItems(W).Checked = True
    Next W
End Sub

Private Sub CmdClear_Click()
    TxtCariNoTelp.Text = ""
    txtcustid.Text = ""
    TxtCariNoTelp.SetFocus
End Sub

Private Sub CmdHapusUnValidNumber_Click()
    Dim a As String
    
    a = MsgBox("Anda yakin akan menghapus data unvalid number?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbNo Then
        Exit Sub
    End If
    
    Call HapusUnvalidNumber
End Sub

Private Sub CmdUnCekAll_Click()
    Dim W As Integer
    
    If LVUnValidNumber.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LVUnValidNumber.ListItems.Count
        LVUnValidNumber.ListItems(W).Checked = False
    Next W
End Sub

Private Sub Form_Load()
    Call header
    Call IsiData
End Sub
