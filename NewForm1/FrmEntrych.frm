VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmEntryCH 
   Caption         =   "Entry Card Holder Data"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4890
   Icon            =   "FrmEntrych.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2865
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtIdReff 
      Height          =   375
      Left            =   1425
      TabIndex        =   14
      Top             =   120
      Width           =   2895
   End
   Begin VB.ComboBox cmbRecsourcech 
      Height          =   315
      Left            =   1425
      TabIndex        =   5
      Top             =   2370
      Width           =   1695
   End
   Begin Threed.SSCommand CmdSave 
      Height          =   390
      Left            =   3270
      TabIndex        =   6
      Top             =   2280
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   688
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Save"
   End
   Begin TDBDate6Ctl.TDBDate TdbDOBCH 
      Height          =   360
      Left            =   1425
      TabIndex        =   4
      Top             =   2010
      Width           =   1470
      _Version        =   65536
      _ExtentX        =   2593
      _ExtentY        =   635
      Calendar        =   "FrmEntrych.frx":000C
      Caption         =   "FrmEntrych.frx":0124
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmEntrych.frx":0190
      Keys            =   "FrmEntrych.frx":01AE
      Spin            =   "FrmEntrych.frx":020C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.15962147735399E-317
      CenturyMode     =   0
   End
   Begin VB.TextBox TxtHandPhone 
      Height          =   375
      Left            =   1425
      TabIndex        =   3
      Top             =   1620
      Width           =   2475
   End
   Begin VB.TextBox TxtTelpKantor 
      Height          =   375
      Left            =   1425
      TabIndex        =   2
      Top             =   1245
      Width           =   2475
   End
   Begin VB.TextBox TxtTelpRumah 
      Height          =   375
      Left            =   1425
      TabIndex        =   1
      Top             =   870
      Width           =   2475
   End
   Begin VB.TextBox TxtNama 
      Height          =   375
      Left            =   1425
      TabIndex        =   0
      Top             =   495
      Width           =   2895
   End
   Begin Threed.SSCommand CmdCancel 
      Height          =   390
      Left            =   4020
      TabIndex        =   7
      Top             =   2280
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   688
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Cancel"
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Id Ref :"
      Height          =   345
      Left            =   300
      TabIndex        =   15
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Batch : "
      Height          =   345
      Left            =   345
      TabIndex        =   13
      Top             =   2370
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "DOB : "
      Height          =   345
      Left            =   330
      TabIndex        =   12
      Top             =   2010
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Cellular : "
      Height          =   345
      Left            =   345
      TabIndex        =   11
      Top             =   1635
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Bussiness Phone : "
      Height          =   345
      Left            =   90
      TabIndex        =   10
      Top             =   1260
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Home Phone : "
      Height          =   345
      Left            =   345
      TabIndex        =   9
      Top             =   885
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name :"
      Height          =   345
      Left            =   300
      TabIndex        =   8
      Top             =   525
      Width           =   1095
   End
End
Attribute VB_Name = "FrmEntryCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public okReff As Boolean

Private Sub cmbRecsourcech_LostFocus()
Dim m_obj As New ADODB.Recordset
m_obj.CursorLocation = adUseClient
m_obj.Open "Select * from DATASOURCETBL WHERE KODEDS = '" + cmbRecsourcech.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_obj.RecordCount <> 0 Then
    cmbRecsourcech.Text = m_obj!KODEDS
Else
    cmbRecsourcech.Text = ""
End If
Set m_obj = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If Len(TxtNama.Text) < 2 Then
        MsgBox "Nama harus diisi", vbInformation + vbOKOnly, "Telegrandi"
        Exit Sub
    End If
    If Len(TxtTelpRumah.Text) < 2 And Len(TxtTelpKantor.Text) < 2 And Len(TxtHandPhone.Text) < 2 Then
        MsgBox "Minimal salah satu dari telp harus diisi", vbInformation + vbOKOnly, "Telegrandi"
        Exit Sub
    End If
    If Len(cmbRecsourcech.Text) < 2 Then
        MsgBox "Batch harus diisi", vbInformation + vbOKOnly, "Telegrandi"
        Exit Sub
    End If
    'CmdSave.Enabled = False
    Call cari_duplicate
End Sub

Private Sub cari_duplicate()
   Dim CMDSQL As String
Dim mrs_cek As ADODB.Recordset
    Dim kriteria1 As String
    Dim kriteria2 As String
    Dim CUSTID1 As String
    ' kriteria pertama
    'nama ama notelp
    If Len(TxtNama.Text) > 2 And Len(TxtTelpRumah.Text) > 2 Then
        kriteria2 = Left(TxtTelpRumah.Text, 5)
        CMDSQL = "Select * from MGM where name like '%" + TxtNama.Text + "%' "
        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
    
    Set mrs_cek = New ADODB.Recordset
        mrs_cek.CursorLocation = adUseClient
        
        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If mrs_cek.RecordCount <> 0 Then
            CUSTID1 = Empty
            While Not mrs_cek.EOF
                CUSTID1 = "MGMI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC_MGM"))
                CMDSQL = "Insert into TBL_DUPLIKASICH (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
                If TdbDOBCH.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "DOB,"
                End If
                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
                If TdbDOBCH.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "'" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "',"
                End If
                CMDSQL = CMDSQL + "'" + cmbRecsourcech.Text + "')"
                M_OBJCONN.Execute CMDSQL
                mrs_cek.MoveNext
            Wend
            
            ' tampilin yang duplicate deh...
                'Call show_Ch_Duplicate
                MsgBox " Nama dan Telp Rumah Ada yg sama", vbInformation + vbOKOnly, "Telegrandi"
            Set mrs_cek = Nothing
            Exit Sub
            
        End If
        Set mrs_cek = Nothing
    End If
    If Len(TxtNama.Text) > 2 And Len(TxtTelpKantor.Text) > 2 Then
        kriteria2 = Left(TxtTelpKantor.Text, 5)
        CMDSQL = "Select * from MGM where name like '%" + TxtNama.Text + "%' "
        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
    Set mrs_cek = New ADODB.Recordset
    mrs_cek.CursorLocation = adUseClient
        
        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If mrs_cek.RecordCount <> 0 Then
            CUSTID1 = Empty
            While Not mrs_cek.EOF
                CUSTID1 = "MGMI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC_MGM"))
                
                CMDSQL = "Insert into TBL_DUPLIKASICH (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
                If TdbDOBCH.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "DOB,"
                End If
                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
                If TdbDOBCH.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "'" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "',"
                End If
                CMDSQL = CMDSQL + "'" + cmbRecsourcech.Text + "')"
                M_OBJCONN.Execute CMDSQL
                mrs_cek.MoveNext
            Wend
            
            ' show data
            'Call show_Ch_Duplicate
            MsgBox " Nama dan Telp Kantor Ada yg sama", vbInformation + vbOKOnly, "Telegrandi"
            Set mrs_cek = Nothing
            Exit Sub
        End If
        Set mrs_cek = Nothing
        
    End If
    If Len(TxtNama.Text) > 2 And Len(TxtHandPhone.Text) > 2 Then
        kriteria2 = Left(TxtHandPhone.Text, 8)
        CMDSQL = "Select * from MGM where name like '%" + TxtNama.Text + "%' "
        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
    Set mrs_cek = New ADODB.Recordset
    mrs_cek.CursorLocation = adUseClient
        
        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If mrs_cek.RecordCount <> 0 Then
            
            
            CUSTID1 = Empty
            While Not mrs_cek.EOF
                CUSTID1 = "MGMI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC_MGM"))
                
                CMDSQL = "Insert into TBL_DUPLIKASICH (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
                If TdbDOBCH.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "DOB,"
                End If
                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
                If TdbDOBCH.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "'" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "',"
                End If
                CMDSQL = CMDSQL + "'" + cmbRecsourcech.Text + "')"
                
                M_OBJCONN.Execute CMDSQL
                mrs_cek.MoveNext
            Wend
            
            ' show data
            'Call show_Ch_Duplicate
            MsgBox "Nama dan Handphone Ada yg sama", vbInformation + vbOKOnly, "Telegrandi"
            Set mrs_cek = Nothing
            Exit Sub
        End If
        Set mrs_cek = Nothing
    
    End If
    If Len(TxtNama.Text) > 2 And TdbDOBCH.ValueIsNull = False Then
        kriteria2 = Format(TdbDOBCH.Value, "yyyy/mm/dd")
        CMDSQL = "Select * from MGM where name like '%" + TxtNama.Text + "%' "
        CMDSQL = CMDSQL + " and birthd = '" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "'"
        Set mrs_cek = New ADODB.Recordset
            mrs_cek.CursorLocation = adUseClient

        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If mrs_cek.RecordCount <> 0 Then
            
            
            CUSTID1 = Empty
            While Not mrs_cek.EOF
                CUSTID1 = "MGMI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC_MGM"))
                
                CMDSQL = "Insert into TBL_DUPLIKASICH (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
                If TdbDOBCH.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "DOB,"
                End If
                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
                If TdbDOBCH.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "'" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "',"
                End If
                CMDSQL = CMDSQL + "'" + cmbRecsourcech.Text + "')"
                
                M_OBJCONN.Execute CMDSQL
                mrs_cek.MoveNext
            Wend
            
            ' show data
            'Call show_Ch_Duplicate
            MsgBox "Nama dan DOB Ada yg sama", vbInformation + vbOKOnly, "Telegrandi"
            Set mrs_cek = Nothing
            Exit Sub
        End If
        Set mrs_cek = Nothing
    End If
        CUSTID1 = "MGMI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC_MGM"))

        CMDSQL = "Insert into MGM(CUSTID, NAME, HOMENO, MOBILENO, OFFICENO, AGENT, "
        If TdbDOBCH.ValueIsNull = False Then
            CMDSQL = CMDSQL + "BIRTHD,"
        End If
        CMDSQL = CMDSQL + " RECSOURCE) values"
        CMDSQL = CMDSQL + "('" + CUSTID1 + "',"
        CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
        CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
        CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
        CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
        CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
        If TdbDOBCH.ValueIsNull = False Then
            CMDSQL = CMDSQL + "'" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "',"
        End If
        CMDSQL = CMDSQL + "'" + cmbRecsourcech.Text + "')"
        M_OBJCONN.Execute CMDSQL
                    
        Dim listitem As listitem
        VIEW_MGMDATA.SSTab1.Tab = 0
        Set listitem = VIEW_MGMDATA.LstVwSearchMgm.ListItems.ADD(, , "99999")
            listitem.SubItems(1) = CUSTID1
            listitem.SubItems(2) = ""
            listitem.SubItems(3) = TxtNama.Text
            listitem.SubItems(4) = ""
            listitem.SubItems(5) = "Inbound CH"
            listitem.SubItems(6) = ""
            listitem.SubItems(7) = MDIForm1.Text1.Text
            listitem.SubItems(8) = MDIForm1.Text7.Text
            listitem.SubItems(9) = cmbRecsourcech.Text
            listitem.SubItems(10) = ""
            listitem.SubItems(11) = "1A"
            listitem.SubItems(12) = ""
            listitem.SubItems(13) = ""
            listitem.SubItems(14) = ""
        MsgBox "Data sudah tersimpan", vbInformation + vbOKOnly, "Telegrandi"
        okReff = True
        Me.Hide
    '    Unload Me
End Sub

Private Sub Form_Load()
Dim m_objrs As New ADODB.Recordset
    TdbDOBCH.Value = Empty
    TxtNama.Text = Empty
    TxtTelpRumah.Text = Empty
    TxtTelpKantor.Text = Empty
    TxtHandPhone.Text = Empty
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from DATASOURCETBL WHERE STATUS = 'I' ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not m_objrs.EOF
        cmbRecsourcech.AddItem m_objrs!KODEDS
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
End Sub

