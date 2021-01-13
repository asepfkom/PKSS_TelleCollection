VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FRM_HASILUPLOAD 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   5235
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   7830
      Begin MSComctlLib.ListView ListView1 
         Height          =   4905
         Left            =   30
         TabIndex        =   1
         Top             =   135
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   8652
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
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
   End
   Begin Threed.SSCommand SSCommand2 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   6705
      TabIndex        =   2
      Top             =   5250
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   714
      _Version        =   196610
      Font3D          =   5
      MousePointer    =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
      ButtonStyle     =   2
   End
End
Attribute VB_Name = "FRM_HASILUPLOAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim M_OBJRS As ADODB.Recordset
Dim listitem As listitem
Call header
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS.Open "SELECT * FROM HSLUPLOAD ORDER BY DATASOURCE", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_OBJRS.EOF
    Set listitem = ListView1.ListItems.ADD(, , IIf(IsNull(M_OBJRS("TGL")), "", Format(M_OBJRS("TGL"), "dd-mmm-yyyy")))
    listitem.SubItems(1) = IIf(IsNull(M_OBJRS("DATASOURCE")), "", M_OBJRS("DATASOURCE"))
    listitem.SubItems(2) = IIf(IsNull(M_OBJRS("JMLALL")), 0, Format(M_OBJRS("JMLALL"), "###,###"))
    listitem.SubItems(3) = IIf(IsNull(M_OBJRS("JMLVALID")), 0, Format(M_OBJRS("JMLVALID"), "###,###"))
    listitem.SubItems(4) = IIf(IsNull(M_OBJRS("JMLINVALID")), 0, Format(M_OBJRS("JMLINVALID"), "###,###"))
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Tanggal", 12 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Data Source", 12 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Total Upload", 12 * TXT, 1
    ListView1.ColumnHeaders.ADD 4, , "Upload Sukses", 12 * TXT, 1
    ListView1.ColumnHeaders.ADD 5, , "Upload Gagal", 12 * TXT, 1
End Sub

Private Sub SSCommand2_Click()
    Unload Me
End Sub
