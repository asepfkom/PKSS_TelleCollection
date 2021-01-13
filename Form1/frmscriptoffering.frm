VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMSCRIPT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SCRIPT OFFERING"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Calculator discon"
      Height          =   3975
      Left            =   -60
      TabIndex        =   8
      Top             =   480
      Width           =   4635
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3480
         Width           =   885
      End
      Begin TDBNumber6Ctl.TDBNumber TdbMaxDisc 
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   1260
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   450
         Calculator      =   "frmscriptoffering.frx":0000
         Caption         =   "frmscriptoffering.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmscriptoffering.frx":008C
         Keys            =   "frmscriptoffering.frx":00AA
         Spin            =   "frmscriptoffering.frx":00F4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   1834745861
         MinValueVT      =   1970470917
      End
      Begin TDBNumber6Ctl.TDBNumber TdbBalance 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   900
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   556
         Calculator      =   "frmscriptoffering.frx":011C
         Caption         =   "frmscriptoffering.frx":013C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmscriptoffering.frx":01A8
         Keys            =   "frmscriptoffering.frx":01C6
         Spin            =   "frmscriptoffering.frx":0210
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
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
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TdbDiscon 
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   540
         Width           =   675
         _Version        =   65536
         _ExtentX        =   1191
         _ExtentY        =   450
         Calculator      =   "frmscriptoffering.frx":0238
         Caption         =   "frmscriptoffering.frx":0258
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmscriptoffering.frx":02C4
         Keys            =   "frmscriptoffering.frx":02E2
         Spin            =   "frmscriptoffering.frx":032C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   0
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
         MaxValueVT      =   1834745861
         MinValueVT      =   1970470917
      End
      Begin TDBNumber6Ctl.TDBNumber TDbBalanceAfter 
         Height          =   375
         Left            =   1860
         TabIndex        =   12
         Top             =   1560
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   661
         Calculator      =   "frmscriptoffering.frx":0354
         Caption         =   "frmscriptoffering.frx":0374
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmscriptoffering.frx":03E0
         Keys            =   "frmscriptoffering.frx":03FE
         Spin            =   "frmscriptoffering.frx":0448
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
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
         ForeColor       =   32768
         Format          =   "###,###,###,##0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label8 
         Caption         =   "Balance:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Discon:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Max.Disc:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "%"
         Height          =   255
         Left            =   1980
         TabIndex        =   16
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "%"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Balance after discon:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1620
         Width           =   1755
      End
      Begin VB.Label Label7 
         Caption         =   "(Balance after discon=balance - (balance*discon)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   2100
         Width           =   3495
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "S&end"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2910
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7230
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   30
      TabIndex        =   5
      Top             =   5730
      Width           =   4665
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1380
      TabIndex        =   4
      Top             =   5340
      Width           =   855
   End
   Begin VB.TextBox txtremarks 
      Appearance      =   0  'Flat
      BackColor       =   &H00B8E2D4&
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   5700
      Width           =   4725
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   405
      Left            =   -360
      TabIndex        =   0
      Top             =   6300
      Visible         =   0   'False
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   714
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
   Begin VB.Label LblTextGuide 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1275
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4665
   End
   Begin VB.Label Label1 
      Caption         =   "Ask Advice To"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   5370
      Width           =   1245
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Input Data Offer"
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
      TabIndex        =   1
      Top             =   6420
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.Image Image2 
      Height          =   435
      Index           =   8
      Left            =   60
      Picture         =   "frmscriptoffering.frx":0470
      Stretch         =   -1  'True
      Top             =   6240
      Visible         =   0   'False
      Width           =   20700
   End
End
Attribute VB_Name = "FRMSCRIPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Last Update #12042013 by Izuddin

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
   
            If Text1.Text = "" Then
                MsgBox "Pesan tidak boleh kosong!", vbOKOnly + vbExclamation, "Peringatan"
                Exit Sub
            End If
        
            CMDSQL = "INSERT INTO mandiri.msgtbl "
            CMDSQL = CMDSQL + " ( RECIPIENT,"
            CMDSQL = CMDSQL + " DATETIME,"
            CMDSQL = CMDSQL + " SENDER,"
            CMDSQL = CMDSQL + " SENTFROM,"
            CMDSQL = CMDSQL + " MSG)"
            CMDSQL = CMDSQL + " VALUES"
            CMDSQL = CMDSQL + " ( '" + Combo1.Text + "',"
            CMDSQL = CMDSQL + " '" + Format(Date, "yyyymmdd") + "',"
            CMDSQL = CMDSQL + " '" + Trim(MDIForm1.TxtUsername.Text) + "',"
            CMDSQL = CMDSQL + " '" + CStr(MDIForm1.Winsock1.LocalIP) + "',"
            CMDSQL = CMDSQL + " '" + Text1.Text + "')"
            M_OBJCONN.Execute CMDSQL
            
            MsgBox "Pesan telah dikirim ke TL anda!", vbOKOnly + vbInformation, "Informasi"
            Text1.Text = ""
           
        
Case 3
Unload Me
End Select

End Sub
Private Sub Form_Load()
Dim list As ListItem
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
header

STRSQL = " select * from mandiri.usertbl where userid='" + Trim(MDIForm1.TxtUsername.Text) + "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

If rs.BOF And rs.EOF Then
    Combo1.Text = ""
Else
    Combo1.Text = IIf(IsNull(rs!TEAM), "", rs!TEAM)
End If

While Not rs.EOF
    Combo1.AddItem IIf(IsNull(rs!TEAM), "", rs!TEAM)
    rs.MoveNext
Wend
Set rs = Nothing




Set rs1 = New ADODB.Recordset
rs1.CursorLocation = adUseClient
STRSQL = " select * from mandiri.mgm where custid='" + FrmCC_Colection.TxtCustid.Text + "'"
rs1.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic




STRSQL = " select * from mandiri.tbloffering WHERE IDKEY='" + FrmCC_Colection.TXTRUMUS.Text + "'"


Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    Set list = ListView1.ListItems.ADD(, , rs!keterangan)
    If rs1.BOF And rs1.EOF Then
        list.SubItems(1) = 0
    Else
    list.SubItems(1) = 0
    On Error Resume Next
    a = rs!fldrms
    If rs!OPERAND = "+" Then
    Value1 = IIf(IsNull(rs1.fields(a).Value), "0", rs1.fields(a).Value)
    If IIf(IsNull(rs!exispersentase), "", rs!exispersentase) = "Y" Then
            TEMP = Value1 + ((Value1 * FrmCC_Colection.Text6.Text) / 100)
        Else
            TEMP = Value1 + ((Value1 * rs!persentase) / 100)
        End If
        
    list.SubItems(1) = TEMP
    Else
    Value1 = IIf(IsNull(rs1.fields(a).Value), "0", rs1.fields(a).Value)
     If IIf(IsNull(rs!exispersentase), "", rs!exispersentase) = "Y" Then
            TEMP = Value1 - ((Value1 * FrmCC_Colection.Text6.Text) / 100)
        Else
            TEMP = Value1 - ((Value1 * rs!persentase) / 100)
        End If
        
    list.SubItems(1) = TEMP
 
    End If
 End If
 If IIf(IsNull(rs!Remarks), "", rs!Remarks) <> "" Then
        TxtRemarks.Text = rs!Remarks
    End If
    
    rs.MoveNext
Wend

End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "keterangan", 10 * 120
    ListView1.ColumnHeaders.ADD 2, , "amount", 20 * 120
   
    
End Sub


'@@ 09092011, Form Discon Offering Dipindahin ke sini
Private Sub TdbDiscon_Change()
    'jika diskon melebihi maksimal diskon maka keluar
    If TdbDiscon.Value > TdbMaxDisc.Value Then
        MsgBox "Maksimal diskon: " & TdbMaxDisc.Value & "%", vbOKOnly + vbInformation, "Informasi"
        TdbDiscon.Value = 0
        TDbBalanceAfter.Value = 0
        Exit Sub
    End If
    TDbBalanceAfter.Value = TdbBalance - (TdbBalance * TdbDiscon.Value / 100)
End Sub
