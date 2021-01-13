VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Frmpaidoff 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2325
   ClientLeft      =   7290
   ClientTop       =   3165
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4050
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   2775
      TabIndex        =   10
      Top             =   1920
      Width           =   1020
   End
   Begin VB.TextBox txtbln 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1695
      TabIndex        =   1
      Top             =   810
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   1470
      TabIndex        =   0
      Top             =   1920
      Width           =   1020
   End
   Begin TDBDate6Ctl.TDBDate tgltempo 
      Height          =   315
      Left            =   1695
      TabIndex        =   2
      Top             =   1515
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "Frmpaidoff.frx":0000
      Caption         =   "Frmpaidoff.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Frmpaidoff.frx":0184
      Keys            =   "Frmpaidoff.frx":01A2
      Spin            =   "Frmpaidoff.frx":0200
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
      ForeColor       =   0
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
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBNumber6Ctl.TDBNumber txtjmlpay 
      Height          =   345
      Left            =   1695
      TabIndex        =   3
      Top             =   450
      Width           =   1725
      _Version        =   65536
      _ExtentX        =   3043
      _ExtentY        =   609
      Calculator      =   "Frmpaidoff.frx":0228
      Caption         =   "Frmpaidoff.frx":0248
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Frmpaidoff.frx":02B4
      Keys            =   "Frmpaidoff.frx":02D2
      Spin            =   "Frmpaidoff.frx":031C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
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
      MaxValue        =   99999999999999
      MinValue        =   -99999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   1701642245
      MinValueVT      =   3801093
   End
   Begin TDBNumber6Ctl.TDBNumber txtnominal 
      Height          =   300
      Left            =   1695
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   529
      Calculator      =   "Frmpaidoff.frx":0344
      Caption         =   "Frmpaidoff.frx":0364
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Frmpaidoff.frx":03D0
      Keys            =   "Frmpaidoff.frx":03EE
      Spin            =   "Frmpaidoff.frx":0438
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
      Enabled         =   0
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
      ReadOnly        =   -1
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   661
      _Version        =   196610
      Font3D          =   4
      ForeColor       =   12582912
      Caption         =   "Regular to Paid Off"
      BevelWidth      =   2
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Jumlah Pelunasan :"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   435
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Jatuh tempo :"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   1500
      Width           =   1665
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Jangka Waktu :"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   795
      Width           =   1665
   End
   Begin VB.Label Label4 
      Caption         =   "Bulan"
      Height          =   375
      Left            =   2535
      TabIndex        =   6
      Top             =   810
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Nominal :"
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   1185
      Width           =   1665
   End
End
Attribute VB_Name = "Frmpaidoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim M_DATA As New ClsNegoPTP
Dim listitem As listitem
Dim x As Integer
Dim jatuhtempo As String
Dim jenis As String
jenis = "RPO"
jatuhtempo = Format(tgltempo.Value, "yyyy-mm-dd")
For x = 1 To Val(txtbln.Text)
    With frmCC_Colection
    M_DATA.ADD_NegoPTP M_OBJCONN, .TxtCustid.Text, jatuhtempo, CStr(txtnominal.Value), MDIForm1.TDBDate1.Value, jenis
    On Error GoTo add_error
    If M_DATA.ADD_OK Then
        Set listitem = .LstPayment.ListItems.ADD(, , "")
        listitem.SubItems(1) = ""
        listitem.SubItems(2) = jatuhtempo
        listitem.SubItems(3) = txtnominal.Value
        listitem.SubItems(4) = jenis
        listitem.SubItems(5) = MDIForm1.TDBDate1.Value
        End If
    End With
    jatuhtempo = DateAdd("m", 1, Format(jatuhtempo, "yyyy-mm-dd"))
Next x
frmCC_Colection.TdbPTP.Value = tgltempo.Value
Exit Sub
add_error:
    MsgBox Err.Description
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
'txtjmlpay.Value = frmCC_Colection.lblPromPA.Value - (CCur(frmCC_Colection.txtDiscount.Text) * frmCC_Colection.lblPromPA.Value)
'txtnominal.Value = txtjmlpay.Value
txtjmlpay.Value = frmCC_Colection.txtPayment
End Sub

Private Sub txtbln_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'
'End If
End Sub

Private Sub txtbln_KeyUp(KeyCode As Integer, Shift As Integer)
If txtbln.Text = "" Then
Else
    If Val(txtbln.Text) > 6 Then
        MsgBox "Maximal PTP untuk 6 bulan !"
        txtbln.Text = 0
    Else
        If txtbln.Text = 0 Then
            MsgBox "Bulan tidak boleh Nol !"
        Else
            txtnominal = txtjmlpay.Value / Val(txtbln.Text)
        End If
    End If
End If
End Sub

Private Sub txtbln_LostFocus()
'If Val(txtbln.Text) > 6 Then
'    MsgBox "Maximal PTP untuk 6 bulan !"
'    txtbln.Text = 0
'Else
'    If txtbln.Text = 0 Then
'        MsgBox "Bulan tidak boleh Nol !"
'    Else
'        txtnominal = txtjmlpay.Value / Val(txtbln.Text)
'    End If
'End If
End Sub

