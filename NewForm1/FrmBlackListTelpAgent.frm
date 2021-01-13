VERSION 5.00
Begin VB.Form FrmBlackListTelpAgent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Valid/UnValid Number Telephone"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6150
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbStatusTelp 
      Height          =   315
      ItemData        =   "FrmBlackListTelpAgent.frx":0000
      Left            =   1440
      List            =   "FrmBlackListTelpAgent.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1020
      Width           =   1095
   End
   Begin VB.ComboBox CmbStatus 
      Height          =   315
      ItemData        =   "FrmBlackListTelpAgent.frx":0025
      Left            =   1440
      List            =   "FrmBlackListTelpAgent.frx":0032
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   300
      Width           =   2115
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   4620
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox TxtKeterangan 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1380
      Width           =   4635
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   3180
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtNotelp 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   660
      Width           =   2295
   End
   Begin VB.Label LblTelp 
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   300
      Width           =   2475
   End
   Begin VB.Label LblStatusTelp 
      Height          =   195
      Left            =   2700
      TabIndex        =   10
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Telp."
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1020
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Jenis"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   300
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Telepon:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1380
      Width           =   1575
   End
End
Attribute VB_Name = "FrmBlackListTelpAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ok As Boolean
Public STATUS As String

Private Sub CmbStatus_Click()
    If CmbStatus.Text <> "UnValid Number" Then
        CmbStatusTelp.Enabled = False
    Else
        CmbStatusTelp.Enabled = True
    End If
End Sub

Private Sub CmbStatusTelp_Click()
    If CmbStatusTelp.Text = "WN" Then
        LblStatusTelp.Caption = "Salah Sambung"
    ElseIf CmbStatusTelp.Text = "NK" Then
        LblStatusTelp.Caption = "CH tidak dikenal"
    ElseIf CmbStatusTelp.Text = "MV" Then
        LblStatusTelp.Caption = "CH pindah"
    ElseIf CmbStatusTelp.Text = "RSG" Then
        LblStatusTelp.Caption = "CH Resign"
    End If
End Sub

Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim VSAVE As Boolean
    
    VSAVE = True
    VSAVE = VSAVE And TxtNoTelp.Text <> Empty
    VSAVE = VSAVE And CmbStatus.Text <> Empty
    
    If VSAVE Then
         If Len(TxtNoTelp.Text) > 20 Then
            MsgBox "Maksimal jumlah digit no telp:20!", vbInformation + vbOKOnly, "Informasi"
            Exit Sub
        End If
        If CmbStatus.Text = "Black List Number" Then
            Call BlackListNumber
        ElseIf CmbStatus.Text = "Valid Number" Then
            STATUS = "Valid Number"
            Call InputValidNumber
             If ok = False Then
                Exit Sub
            End If
            Call UpdateStatusValidNumber
            Me.Hide
        '@@ 07-05-2012, Perubahan untuk Unvalid Number
        ElseIf CmbStatus.Text = "UnValid Number" Then
            If CmbStatusTelp.Text = "" Then
                MsgBox "Tentukan Status Telepon!", vbOKOnly + vbInformation, "Informasi"
                ok = False
                Exit Sub
            End If
            STATUS = "UNVALID NUMBER"
            Call UnValidNumber
            If ok = False Then
                Exit Sub
            End If
            Call UpdateStatusUnvalidNumber
            ok = True
            Me.Hide
        End If
    Else
      MsgBox "Data Yang Anda Masukan Tidak Lengkap", vbInformation, "Informasi"
      ok = False
    End If
    
End Sub



Private Sub Form_Load()
    ok = False
End Sub

Private Sub TxtNoTelp_KeyPress(KeyAscii As Integer)
 'Hanya numeric yang dapat diinput
 If KeyAscii < 47 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub InputValidNumber()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    '@@ 10052012 Cek Apakah data Masuk Dalam UnValid Number
    CMDSQL = "select * from tblunvalid_number where no_telp='"
    CMDSQL = CMDSQL + Trim(TxtNoTelp.Text) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount > 0 Then
        MsgBox "Nomor tersebut masuk dalam Unvalid Number! Tidak dapat dijadikan Valid Number sebelum data UnValid Number dihapus oleh TL/SPV!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        ok = False
        Exit Sub
    End If
    
    Set M_objrs = Nothing
    
    'Cek nomor telepon dulu, di blacklist apa ngga
    CMDSQL = "select * from tblblacklist where no_telp='"
    CMDSQL = CMDSQL + Trim(TxtNoTelp.Text) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        MsgBox "Mohon maaf, Valid Number tidak dapat dilakukan, karena nomor ini masuk dalam black list. Anda dapat meminta SPV/Admin untuk menghapus black list nomor ini!", vbOKOnly + vbInformation, "Informasi"
        ok = False
        Set M_objrs = Nothing
        Exit Sub
    End If
    Set M_objrs = Nothing
    
    CMDSQL = "insert into tblvalidnumber (no_telp,keterangan,tglinput,userinput,custid,agent) values ('"
    CMDSQL = CMDSQL + Trim(TxtNoTelp.Text) + "','"
    CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", Trim(TxtKeterangan.Text)) + "',"
    CMDSQL = CMDSQL + "'" & Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & "','" + MDIForm1.txtnama.Text + "',' "
    CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "','"
    CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "')"
    
    M_OBJCONN.Execute CMDSQL
    
    Remarks = "VALID NUMBER  "
    Remarks = Remarks & CStr(TxtNoTelp.Text)
    Remarks = Remarks & " ,Reason: "
    Remarks = Remarks & IIf(IsNull(TxtKeterangan.Text), "(Null)", TxtKeterangan.Text)
    
    CMDSQL = "insert into mgm_hst (custid,agent,hst) values ('"
    CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "','"
    CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblaoc.Caption) + "','"
    CMDSQL = CMDSQL + Remarks + "')"
    M_OBJCONN.Execute CMDSQL
    
    ok = True
    'Me.Hide
End Sub
Private Sub BlackListNumber()
    STATUS = "Black List Number"
    CMDSQL = "insert into tblblacklist (no_telp,keterangan,tglinput,userinput) values ('"
    CMDSQL = CMDSQL + Trim(TxtNoTelp.Text) + "','"
    CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", Trim(TxtKeterangan.Text)) + "',"
    CMDSQL = CMDSQL + "'" & Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & "','" + MDIForm1.txtnama.Text + "') "
    'Cek data no telepon yang sama
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
        Cmdsql_Cek = "select * from tblblacklist where no_telp='"
        Cmdsql_Cek = Cmdsql_Cek + CStr(Trim(TxtNoTelp.Text)) + "'"
    M_objrs.Open Cmdsql_Cek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount <> 0 Then
        m_msgbox = MsgBox("No Telepon sudah ada dalam blacklist. Data gagal disimpan!", vbOKOnly + vbExclamation, "Peringatan")
        ok = False
        Exit Sub
    End If
    Set M_objrs = Nothing
    
    M_OBJCONN.Execute CMDSQL
    ok = True
    Me.Hide
End Sub

Private Sub UnValidNumber()
    Dim CMDSQL As String
    Dim Remarks As String
    STATUS = "UNVALID NUMBER"
    
    
    CMDSQL = "insert into tblunvalid_number (no_telp,keterangan,tglinput,"
    CMDSQL = CMDSQL + "userinput,status,telpblok,custid,userid) values ('"
    CMDSQL = CMDSQL + Trim(TxtNoTelp.Text) + "','"
    CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", Trim(TxtKeterangan.Text)) + "',"
    CMDSQL = CMDSQL + "'" & Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & "','" + MDIForm1.txtnama.Text + "','"
    CMDSQL = CMDSQL + CmbStatusTelp.Text + "','"
    CMDSQL = CMDSQL + LblTelp.Caption + "','"
    CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "','"
    CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "')"
    
    'Cek data no telepon yang sama
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    Cmdsql_Cek = "select * from tblunvalid_number where no_telp='"
    Cmdsql_Cek = Cmdsql_Cek + CStr(Trim(TxtNoTelp.Text)) + "' and custid='"
    Cmdsql_Cek = Cmdsql_Cek + CStr(FrmCC_Colection.lblCustId.Text) + "'"
    M_objrs.Open Cmdsql_Cek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount <> 0 Then
        m_msgbox = MsgBox("No Telepon sudah ada dalam Unvalid Number. Data gagal disimpan!", vbOKOnly + vbExclamation, "Peringatan")
        ok = False
        Exit Sub
    End If
    Set M_objrs = Nothing
    
    M_OBJCONN.Execute CMDSQL
    
    '@@07-05-2012, Tulis Ke Remarks
    Remarks = "UNVALID NUMBER  "
    Remarks = Remarks & CStr(TxtNoTelp.Text)
    Remarks = Remarks & " ,Reason: "
    Remarks = Remarks & IIf(IsNull(TxtKeterangan.Text), "(Null)", TxtKeterangan.Text)
    
    CMDSQL = "insert into mgm_hst (custid,agent,hst) values ('"
    CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "','"
    CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblaoc.Caption) + "','"
    CMDSQL = CMDSQL + Remarks + "')"
    M_OBJCONN.Execute CMDSQL
    
    'Call UpdateStatusUnvalidNumber
    ok = True
    'Me.Hide
End Sub

Private Sub UpdateStatusUnvalidNumber()
    Dim CMDSQL As String
    
    Select Case LblTelp.Caption
        Case "Home 1"
            CMDSQL = "update mgm set f_unvalid_home1='1',f_valid_home1=null,f_sts_unvalid_home1='"
            CMDSQL = CMDSQL + CmbStatusTelp + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_valid_home1=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "Home 2"
            CMDSQL = "update mgm set f_unvalid_home2='1',f_valid_home2=null,f_sts_unvalid_home2='"
            CMDSQL = CMDSQL + CmbStatusTelp + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "', "
            CMDSQL = CMDSQL + " f_sts_valid_home2=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "Office 1"
            CMDSQL = "update mgm set f_unvalid_office1='1',f_valid_office1=null,f_sts_unvalid_office1='"
            CMDSQL = CMDSQL + CmbStatusTelp + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "', "
            CMDSQL = CMDSQL + " f_sts_valid_office1=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "Office 2"
            CMDSQL = "update mgm set f_unvalid_office2='1',f_valid_office2=null,f_sts_unvalid_office2='"
            CMDSQL = CMDSQL + CmbStatusTelp + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "', "
            CMDSQL = CMDSQL + " f_sts_valid_office2=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "Mobile 1"
            CMDSQL = "update mgm set f_unvalid_mobile1='1',f_valid_mobile1=null,f_sts_unvalid_mobile1='"
            CMDSQL = CMDSQL + CmbStatusTelp + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "', "
            CMDSQL = CMDSQL + " f_sts_valid_mobile1=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "Mobile 2"
            CMDSQL = "update mgm set f_unvalid_mobile2='1',f_valid_mobile2=null,f_sts_unvalid_mobile2='"
            CMDSQL = CMDSQL + CmbStatusTelp + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "', "
            CMDSQL = CMDSQL + " f_sts_valid_mobile2=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "AddHome 1"
            CMDSQL = "update mgm set f_unvalid_addhome1='1',f_valid_addhome1=null,f_sts_unvalid_addhome1='"
            CMDSQL = CMDSQL + CmbStatusTelp.Text + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "', "
            CMDSQL = CMDSQL + " f_sts_valid_addhome1=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "AddHome 2"
            CMDSQL = "update mgm set f_unvalid_addhome2='1',f_valid_addhome2=null,f_sts_unvalid_addhome2='"
            CMDSQL = CMDSQL + CmbStatusTelp + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "', "
            CMDSQL = CMDSQL + " f_sts_valid_addhome2=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "AddOffice 1"
            CMDSQL = "update mgm set f_unvalid_addoffice1='1',f_valid_addoffice1=null,f_sts_unvalid_addoffice1='"
            CMDSQL = CMDSQL + CmbStatusTelp + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "', "
            CMDSQL = CMDSQL + " f_sts_valid_addoffice1=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "AddOffice 2"
            CMDSQL = "update mgm set f_unvalid_addoffice2='1',f_valid_addoffice2=null,f_sts_unvalid_addoffice2='"
            CMDSQL = CMDSQL + CmbStatusTelp + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "', "
            CMDSQL = CMDSQL + " f_sts_valid_addoffice2=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "AddMobile 1"
            CMDSQL = "update mgm set f_unvalid_addmobile1='1',f_valid_addmobile1=null,f_sts_unvalid_addmobile1='"
            CMDSQL = CMDSQL + CmbStatusTelp + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "', "
            CMDSQL = CMDSQL + " f_sts_valid_addmobile1=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "AddMobile 2"
            CMDSQL = "update mgm set f_unvalid_addmobile2='1',f_valid_addmobile2=null,f_sts_unvalid_addmobile2='"
            CMDSQL = CMDSQL + CmbStatusTelp + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "', "
            CMDSQL = CMDSQL + " f_sts_valid_addmobile2=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
         Case "EC"
            CMDSQL = "update mgm set f_unvalid_ec='1',f_valid_ec=null,f_sts_unvalid_ec='"
            CMDSQL = CMDSQL + CmbStatusTelp + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "', "
            CMDSQL = CMDSQL + " f_sts_valid_ec=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
    End Select
    
    M_OBJCONN.Execute CMDSQL
    
End Sub


'@@11052012, UPdate Status Valid Number
Private Sub UpdateStatusValidNumber()
    Dim CMDSQL As String
    
    Select Case LblTelp.Caption
        Case "Home 1"
            CMDSQL = "update mgm set f_valid_home1='1', f_sts_valid_home1='"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_unvalid_home1=null,f_unvalid_home1=null "
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "Home 2"
            CMDSQL = "update mgm set f_valid_home2='1', f_sts_valid_home2='"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_unvalid_home2=null,f_unvalid_home2=null "
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "Office 1"
            CMDSQL = "update mgm set f_valid_office1='1', f_sts_valid_office1='"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_unvalid_office1=null,f_unvalid_office1=null "
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "Office 2"
            CMDSQL = "update mgm set f_valid_office2='1', f_sts_valid_office2='"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_unvalid_office2=null,f_unvalid_office2=null "
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "Mobile 1"
            CMDSQL = "update mgm set f_valid_mobile1='1', f_sts_valid_mobile1='"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_unvalid_mobile1=null,f_unvalid_mobile1=null "
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "Mobile 2"
            CMDSQL = "update mgm set f_valid_mobile2='1', f_sts_valid_mobile2='"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_unvalid_mobile2=null,f_unvalid_mobile2=null "
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "AddHome 1"
            CMDSQL = "update mgm set f_valid_addhome1='1', f_sts_valid_addhome1='"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_unvalid_addhome1=null,f_unvalid_addhome1=null "
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "AddHome 2"
            CMDSQL = "update mgm set f_valid_addhome2='1', f_sts_valid_addhome2='"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_unvalid_addhome2=null,f_unvalid_addhome2=null "
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "AddOffice 1"
            CMDSQL = "update mgm set f_valid_addoffice1='1', f_sts_valid_addoffice1='"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_unvalid_addoffice1=null,f_unvalid_addoffice1=null "
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "AddOffice 2"
            CMDSQL = "update mgm set f_valid_addoffice2='1', f_sts_valid_addoffice2='"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_unvalid_addoffice2=null,f_unvalid_addoffice2=null "
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "AddMobile 1"
            CMDSQL = "update mgm set f_valid_addmobile1='1', f_sts_valid_addmobile1='"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_unvalid_addmobile1=null,f_unvalid_addmobile1=null "
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
        Case "AddMobile 2"
            CMDSQL = "update mgm set f_valid_addmobile2='1', f_sts_valid_addmobile2='"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_unvalid_addmobile2=null,f_unvalid_addmobile2=null "
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
         Case "EC"
            CMDSQL = "update mgm set f_unvalid_ec='1',f_valid_ec=null,f_sts_unvalid_ec='"
            CMDSQL = CMDSQL + CmbStatusTelp + "-"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "', "
            CMDSQL = CMDSQL + " f_sts_valid_ec=null"
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
            
            CMDSQL = "update mgm set f_valid_ec='1', f_sts_valid_ec='"
            CMDSQL = CMDSQL + IIf(IsNull(TxtKeterangan.Text), "", TxtKeterangan.Text) + "',"
            CMDSQL = CMDSQL + " f_sts_unvalid_ec=null,f_unvalid_ec=null "
            CMDSQL = CMDSQL + " where custid='"
            CMDSQL = CMDSQL + CStr(FrmCC_Colection.lblCustId.Text) + "'"
    End Select
    
    M_OBJCONN.Execute CMDSQL
    
End Sub
