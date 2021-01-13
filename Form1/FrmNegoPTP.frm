VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmNegoPTP 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1830
   ClientLeft      =   15
   ClientTop       =   -15
   ClientWidth     =   3510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3510
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   661
      _Version        =   196610
      Font3D          =   4
      ForeColor       =   12582912
      Caption         =   "Iregular to Paid Off"
      BevelWidth      =   2
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.TextBox TxtCustid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1125
      TabIndex        =   7
      Top             =   420
      Width           =   2055
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   345
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   1380
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   609
      _Version        =   196610
      Caption         =   "Add"
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   300
      Left            =   1125
      TabIndex        =   1
      Top             =   720
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   529
      Calendar        =   "FrmNegoPTP.frx":0000
      Caption         =   "FrmNegoPTP.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmNegoPTP.frx":0184
      Keys            =   "FrmNegoPTP.frx":01A2
      Spin            =   "FrmNegoPTP.frx":0200
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
      ValueVT         =   2010382337
      Value           =   2.12482692446619E-314
      CenturyMode     =   0
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
      Height          =   300
      Left            =   1125
      TabIndex        =   0
      Top             =   1035
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   529
      Calculator      =   "FrmNegoPTP.frx":0228
      Caption         =   "FrmNegoPTP.frx":0248
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmNegoPTP.frx":02B4
      Keys            =   "FrmNegoPTP.frx":02D2
      Spin            =   "FrmNegoPTP.frx":031C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   345
      Index           =   1
      Left            =   2070
      TabIndex        =   6
      Top             =   1365
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   609
      _Version        =   196610
      Caption         =   "Close"
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Tgl Janji :"
      Height          =   255
      Left            =   45
      TabIndex        =   4
      Top             =   750
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Nominal :"
      Height          =   255
      Left            =   285
      TabIndex        =   3
      Top             =   1065
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Cust Id :"
      Height          =   255
      Left            =   285
      TabIndex        =   2
      Top             =   435
      Width           =   735
   End
End
Attribute VB_Name = "FrmNegoPTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ok As Boolean

Private Sub ShowPTPnego()
Dim ShowNego As ADODB.Recordset
Dim CMDSQL As String
Set ShowNego = New ADODB.Recordset
ShowNego.CursorLocation = adUseClient
If FrmCC_Colection.LstPayment.ListItems.Count = 0 Then
Exit Sub
End If
CMDSQL = "SELECT * FROM tblnegoPTP where id =  '" + FrmCC_Colection.LstPayment.SelectedItem.SubItems(1) + "'"
ShowNego.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
If ShowNego.RecordCount <> 0 Then
    TDBDate1.Value = CStr(IIf(IsNull(ShowNego!PromiseDate), "", ShowNego!PromiseDate))
    TDBNumber1.Value = IIf(IsNull(ShowNego!PromisePay), "0", ShowNego!PromisePay)
End If
Set ShowNego = Nothing
End Sub

Private Sub UPDATE_NegoPTP()
Dim M_update As New ADODB.Recordset
Dim CMDSQL As String
 M_OBJCONN.Close
M_OBJCONN.Open
Set M_update = New ADODB.Recordset
M_update.CursorLocation = adUseClient
CMDSQL = "SELECT * FROM tblnegoPTP where ID =  '" + FrmCC_Colection.LstPayment.SelectedItem.SubItems(1) + "'"
M_update.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

M_update!PromiseDate = TDBDate1.Value
M_update!PromisePay = TDBNumber1.Value
M_update.update
MsgBox "Update Done...!"
'M_OBJCONN.CommitTrans

Set M_update = Nothing
FrmCC_Colection.LstPayment.SelectedItem.SubItems(2) = CStr(TDBDate1.Value)
FrmCC_Colection.LstPayment.SelectedItem.SubItems(3) = CStr(TDBNumber1.Value)

Unload Me
End Sub
Private Sub Insert_NegoPTP()
Dim M_update As New ADODB.Recordset
Dim CMDSQL As String
 M_OBJCONN.Close
M_OBJCONN.Open
Set M_update = New ADODB.Recordset
M_update.CursorLocation = adUseClient
CMDSQL = "SELECT * FROM tblnegoPTP where ID =  '" + FrmCC_Colection.LstPayment.SelectedItem.SubItems(1) + "'"
M_update.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

M_update!PromiseDate = TDBDate1.Value
M_update!PromisePay = TDBNumber1.Value
M_update.update
MsgBox "Update Done...!"
'M_OBJCONN.CommitTrans

Set M_update = Nothing
FrmCC_Colection.LstPayment.SelectedItem.SubItems(2) = CStr(TDBDate1.Value)
FrmCC_Colection.LstPayment.SelectedItem.SubItems(3) = CStr(TDBNumber1.Value)

Unload Me
End Sub


Private Sub Form_Load()
'If FrmNegoPTP.Caption = "Edit Data" Then
'SSCommand1(0).Caption = "Update"
'SSCommand1(1).Visible = True
'
'ElseIf FrmNegoPTP.Caption = "Insert Data" Then
'SSCommand1(0).Caption = "Save"
'SSCommand1(1).Visible = False
'End If
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
txtcustid.Text = FrmCC_Colection.lblCustId.Text
'Call ShowPTPnego
'frmCC_Colection.Check2.Value = 1
End Sub

Private Sub SSCommand1_Click(Index As Integer)
    Dim m_data As New ClsNegoPTP
    Dim VSAVE As Boolean
    Dim ListItem As ListItem
    Dim jenis As String
    Dim jatuhtempo As String
    
    
    jatuhtempo = Format(TDBDate1.Value, "yyyy-mm-dd")
    jenis = "IPO"
    VSAVE = True
Select Case Index
    Case 0
        VSAVE = VSAVE And txtcustid.Text <> Empty
        VSAVE = VSAVE And TDBDate1.ValueIsNull = False
        VSAVE = VSAVE And TDBNumber1.ValueIsNull = False

        If VSAVE Then
            ok = True
            '@@ 25 May 2012 Dinonaktifkan
'            If FrmCC_Colection.LstPayment.ListItems.Count > 5 Then
'                MsgBox "Maximal Insert PTP Sebanyak 6x !!!"
'                Exit Sub
'            End If
            With FrmNegoPTP
                .Caption = "Tambah Data"
                '@@08-02-2012, Cek jika tanggal PTP yang diinputkan lebih besar atau lebih kecil dari hari ini
'                If Format(TDBDate1.Value, "yyyy-mm-dd") < Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") Then
'                    MsgBox "Tanggal PTP tidak boleh lebih kecil dari tanggal sekarang!", vbOKOnly + vbExclamation, "Peringatan"
'                    ok = False
'                    Exit Sub
'                End If
                If TDBNumber1.Value = 0 Then
                    MsgBox "Nilai PTP tidak boleh=0!", vbOKOnly + vbExclamation, "Peringatan"
                    ok = False
                    Exit Sub
                End If
                FrmCC_Colection.TdbPTP.Value = Format(TDBDate1.Value, "yyyy/mm/dd")
                'FrmCC_Colection.TDBDate3.Value = Format(TDBDate1.Value, "dd/mm/yyyy")
                FrmCC_Colection.TDBDate3.Value = Format(TDBDate1.Value, "mm/dd/yyyy")
                Me.Hide
            End With
            
            
        Else
            MsgBox "Data Yang Anda Masukan Tidak Lengkap", vbInformation, "Informasi"
        End If
        
    Case 1
        ok = False
        Unload Me
        FrmCC_Colection.LstPayment.SetFocus
End Select
'Unload Me
End Sub
