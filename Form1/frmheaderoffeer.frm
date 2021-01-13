VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form frmheaderoffeer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER OFFERING"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4245
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   7488
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Master Offering"
      TabPicture(0)   =   "frmheaderoffeer.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   3915
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   7860
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000004&
            Caption         =   "Pesentase Field ada di excel"
            Height          =   315
            Left            =   120
            TabIndex        =   20
            Top             =   1560
            Width           =   2355
         End
         Begin VB.ComboBox cbooperator 
            Height          =   315
            ItemData        =   "frmheaderoffeer.frx":001C
            Left            =   2580
            List            =   "frmheaderoffeer.frx":0026
            TabIndex        =   12
            Top             =   1230
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Ok"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3420
            UseMaskColor    =   -1  'True
            Width           =   825
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Batal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   5
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   3420
            UseMaskColor    =   -1  'True
            Width           =   810
         End
         Begin VB.TextBox txtremarks 
            Height          =   735
            Left            =   4530
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   840
            Width           =   3315
         End
         Begin VB.ComboBox cbopersentase 
            Height          =   315
            Left            =   1140
            TabIndex        =   8
            Top             =   1260
            Width           =   735
         End
         Begin VB.ComboBox cbomap 
            Height          =   315
            ItemData        =   "frmheaderoffeer.frx":0030
            Left            =   1140
            List            =   "frmheaderoffeer.frx":003D
            TabIndex        =   7
            Top             =   870
            Width           =   2295
         End
         Begin VB.TextBox txtketerangan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1140
            TabIndex        =   6
            Top             =   510
            Width           =   2265
         End
         Begin VB.ComboBox txtkey 
            Height          =   315
            Left            =   4590
            TabIndex        =   5
            Top             =   510
            Width           =   2175
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command2"
            Height          =   585
            Index           =   1
            Left            =   7290
            TabIndex        =   3
            Top             =   1200
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Cancel          =   -1  'True
            Caption         =   "&Remove"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   1950
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   3420
            UseMaskColor    =   -1  'True
            Width           =   810
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1515
            Left            =   60
            TabIndex        =   4
            Top             =   1890
            Width           =   7665
            _ExtentX        =   13520
            _ExtentY        =   2672
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
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "Input Data Offering"
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
            Left            =   570
            TabIndex        =   19
            Top             =   30
            Width           =   3405
         End
         Begin VB.Image Image1 
            Height          =   360
            Index           =   5
            Left            =   90
            Picture         =   "frmheaderoffeer.frx":005E
            Stretch         =   -1  'True
            Top             =   0
            Width           =   420
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tenor"
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Top             =   540
            Width           =   1155
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Field Map"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   930
            Width           =   885
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Persentase"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   1005
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Operator"
            Height          =   255
            Left            =   1890
            TabIndex        =   15
            Top             =   1290
            Width           =   705
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Key Manual List"
            Height          =   255
            Left            =   3420
            TabIndex        =   14
            Top             =   540
            Width           =   1695
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            Height          =   255
            Left            =   3420
            TabIndex        =   13
            Top             =   900
            Width           =   795
         End
         Begin VB.Image Image2 
            Height          =   435
            Index           =   8
            Left            =   0
            Picture         =   "frmheaderoffeer.frx":0B68
            Stretch         =   -1  'True
            Top             =   0
            Width           =   7860
         End
      End
   End
End
Attribute VB_Name = "frmheaderoffeer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ok As Boolean

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "ID", 0
    ListView1.ColumnHeaders.ADD 2, , "TENOR", 20 * 120
    ListView1.ColumnHeaders.ADD 3, , "AMOUNT", 20 * 120
    ListView1.ColumnHeaders.ADD 4, , "%", 6 * 120 '
    ListView1.ColumnHeaders.ADD 5, , "Remarks", 0
    ListView1.ColumnHeaders.ADD 6, , "Operator", 0
    ListView1.ColumnHeaders.ADD 7, , "Id  Rumus", 26 * 100
    ListView1.ColumnHeaders.ADD 8, , "Status", 26 * 100
    


End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim STRSQL As String
    Dim m_data As New CLSSPV_AGENT
    Dim ListItem As ListItem
    Dim cek As Integer
    Dim M_WHERE As String
    Dim i As Integer
    Call header

    STRSQL = "SELECT * FROM mandiri.TBLOFFERING"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    SSTab1.Tab = 0
    While Not M_objrs.EOF
         Set ListItem = ListView1.ListItems.ADD(, , M_objrs("id_offering"))
             ListItem.SubItems(1) = IIf(IsNull(M_objrs("keterangan")), "", M_objrs("keterangan"))
             ListItem.SubItems(2) = IIf(IsNull(M_objrs("fldrms")), "", M_objrs("fldrms"))
             ListItem.SubItems(3) = IIf(IsNull(M_objrs("persentase")), "", M_objrs("persentase"))
             ListItem.SubItems(4) = IIf(IsNull(M_objrs("remarks")), "", M_objrs("remarks"))
             ListItem.SubItems(5) = IIf(IsNull(M_objrs("operand")), "", M_objrs("operand"))
             ListItem.SubItems(6) = IIf(IsNull(M_objrs("idkey")), "", M_objrs("idkey"))
             ListItem.SubItems(7) = IIf(IsNull(M_objrs("exispersentase")), "", M_objrs("exispersentase"))
        M_objrs.MoveNext
    Wend
        M_objrs.Close
        Set M_objrs = Nothing
        
        
For i = 1 To 100
  cbopersentase.AddItem i
  
Next i


End Sub

Private Sub Command1_Click(Index As Integer)
Dim m_msgbox As Variant
Dim STATUS As String
Dim gaji As Currency
Dim gaji1 As String
Dim ListItem As ListItem
Dim stspersentase As String
Dim m_data As New CLSSPV_AGENT
Select Case Index
    Case 0
           SSTab1.Tab = 1
           Label45(0).Caption = "Tambah Data OFFERING"
    Case 1
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
                Label45(0).Caption = "Ubah Data offering"
                txtketerangan.Text = ListView1.SelectedItem.SubItems(1)
                cbomap.Text = ListView1.SelectedItem.SubItems(2)
                cbopersentase.Text = ListView1.SelectedItem.SubItems(3)
                TxtRemarks.Text = ListView1.SelectedItem.SubItems(4)
                cbooperator.Text = ListView1.SelectedItem.SubItems(5)
                txtkey.Text = ListView1.SelectedItem.SubItems(6)
                If ListView1.SelectedItem.SubItems(7) = "" Then
                    Check1.Value = vbUnchecked
                Else
                    Check1.Value = vbChecked
                End If
                
        Exit Sub
    Case 2
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
        m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
        If m_msgbox = 1 Then
            m_data.DELETE_ordering M_OBJCONN, ListView1.SelectedItem.Text
            If m_data.ADD_OK Then
                ListView1.ListItems.Remove ListView1.SelectedItem.Index
            End If
        End If
        Exit Sub
    Case 3
        Unload Me
        Exit Sub
    Case 4
        Dim VSAVE As Boolean
        VSAVE = True
            VSAVE = VSAVE And txtkey.Text <> Empty
            VSAVE = VSAVE And txtketerangan.Text <> Empty
            VSAVE = VSAVE And cbomap.Text <> Empty
            VSAVE = VSAVE And cbopersentase.Text <> Empty
            VSAVE = VSAVE And cbooperator.Text <> Empty
            If VSAVE Then
                ok = True
                   If ok Then
                   If Check1.Value = vbChecked Then
                            stspersentase = "Y"
                            Else
                            stspersentase = ""
                            End If
                        If Label45(0).Caption = "Tambah Data OFFERING" Or Label45(0).Caption = "Input Data Offering" Then
                            
                            
                            m_data.ADD_offering M_OBJCONN, txtketerangan, cbomap.Text, cbopersentase, TxtRemarks.Text, cbooperator.Text, txtkey.Text, stspersentase
                            On Error GoTo add_error
                            showup
                            clsClear
                            MsgBox "Successfuly Save", vbInformation + vbOKOnly, "Pesan"
                        ElseIf Label45(0).Caption = "Ubah Data offering" Then
                            On Error GoTo add_error
                            m_data.UPDATE_ordering M_OBJCONN, ListView1.SelectedItem.Text, txtketerangan, cbomap.Text, cbopersentase, TxtRemarks.Text, cbooperator.Text, txtkey.Text, stspersentase
                            showup
                            clsClear
                            
                            MsgBox "Successfuly Update", vbInformation + vbOKOnly, "Pesan"
                        End If
                        
                End If
                frmheaderoffeer.ListView1.SetFocus
            Else
                MsgBox "Data Yang Anda Masukan Tidak Lengkap", vbInformation, "Informasi"
            End If
    Case 5
        ok = False
        Unload Me
End Select


add_error:
End Sub

Private Sub ListView1_DblClick()
    Call Command1_Click(1)
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click(1)
    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub

Public Sub showup()
Dim STRSQL As String
Dim M_objrs As ADODB.Recordset
    Dim m_data As New CLSSPV_AGENT
    Dim ListItem As ListItem
    Dim cek As Integer
    Dim M_WHERE As String


    STRSQL = "SELECT * FROM TBLOFFERING"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ListView1.ListItems.CLEAR
    While Not M_objrs.EOF
 Set ListItem = ListView1.ListItems.ADD(, , M_objrs("id_offering"))
             ListItem.SubItems(1) = IIf(IsNull(M_objrs("keterangan")), "", M_objrs("keterangan"))
             ListItem.SubItems(2) = IIf(IsNull(M_objrs("fldrms")), "", M_objrs("fldrms"))
             ListItem.SubItems(3) = IIf(IsNull(M_objrs("persentase")), "", M_objrs("persentase"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("remarks")), "", M_objrs("remarks"))
             ListItem.SubItems(5) = IIf(IsNull(M_objrs("operand")), "", M_objrs("operand"))
             ListItem.SubItems(6) = IIf(IsNull(M_objrs("idkey")), "", M_objrs("idkey"))
             ListItem.SubItems(7) = IIf(IsNull(M_objrs("exispersentase")), "", M_objrs("exispersentase"))
        M_objrs.MoveNext
    Wend
        M_objrs.Close
End Sub
Private Sub txtkey_DropDown()
Dim OBJRS As New ADODB.Recordset
Dim STRSQL As String
txtkey.CLEAR
Set OBJRS = New ADODB.Recordset
OBJRS.CursorLocation = adUseClient
STRSQL = "SELECT DISTINCT(IDKEY) FROM TBLOFFERING"
OBJRS.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not OBJRS.EOF
   txtkey.AddItem IIf(IsNull(OBJRS(0).Value), "", OBJRS(0).Value)
    OBJRS.MoveNext
Wend

End Sub
Public Sub clsClear()
txtkey.Text = ""
txtketerangan.Text = ""
cbomap.Text = ""
cbooperator.Text = ""
cbopersentase.Text = ""
TxtRemarks.Text = ""
End Sub

