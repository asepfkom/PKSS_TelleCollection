VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmProductList 
   BorderStyle     =   0  'None
   Caption         =   "Product List"
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9660
   LinkTopic       =   "Form2"
   ScaleHeight     =   6150
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   5505
      Left            =   0
      TabIndex        =   0
      Top             =   645
      Width           =   9660
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Tambah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   0
         Left            =   8610
         Picture         =   "FrmProductList.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Hapus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   8610
         Picture         =   "FrmProductList.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1860
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Cancel          =   -1  'True
         Caption         =   "&Tutup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   8610
         Picture         =   "FrmProductList.frx":0BDC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2625
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Ubah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   1
         Left            =   8610
         Picture         =   "FrmProductList.frx":0D26
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1110
         Width           =   885
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5325
         Left            =   30
         TabIndex        =   5
         Top             =   135
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   9393
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         Picture         =   "FrmProductList.frx":0E70
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   1164
      _Version        =   196610
      Font3D          =   5
      ForeColor       =   0
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "List Informasi"
      BevelWidth      =   2
      BorderWidth     =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
End
Attribute VB_Name = "FrmProductList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub header()
    With ListView1.ColumnHeaders
        .ADD 1, , "Description", 25 * TXT
        .ADD 2, , "Direktori", 80 * TXT
        .ADD 3, , "Expiry Date", 20 * TXT
    End With
End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    Dim ssql As String
    Call header
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    ssql = "SELECT Description, Direktori, ExpiryDate FROM mandiri.TblInformationLokasi"
    M_objrs.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not M_objrs.EOF
         Set ListItem = ListView1.ListItems.ADD(, , M_objrs("Description"))
             ListItem.SubItems(1) = M_objrs("Direktori")
             ListItem.SubItems(2) = Format(M_objrs("ExpiryDate"), "yyyy/mm/dd")
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim m_msgbox As Variant
Dim ListItem As ListItem
Dim CMDSQL As String
Select Case Index
    Case 0
        With FrmProduct
            .Caption = "Tambah"
            .Show vbModal
            If .ok Then
                CMDSQL = "Insert into mandiri.TblInformationLokasi "
                CMDSQL = CMDSQL + "  (Description,Direktori,ExpiryDate)"
                CMDSQL = CMDSQL + " values"
                CMDSQL = CMDSQL + " ('" + .Text1.Text + "',"
                CMDSQL = CMDSQL + " '" + .Text2.Text + "',"
                CMDSQL = CMDSQL + " '" + Format(.TdbExpiryDate, "yyyy/mm/dd") + "'  )"
                    
                M_OBJCONN.Execute CMDSQL
                    
                On Error GoTo add_error
                Set ListItem = ListView1.ListItems.ADD(, , .Text1.Text)
                    ListItem.SubItems(1) = .Text2.Text
                    ListItem.SubItems(2) = .TdbExpiryDate.Value
                On Error GoTo 0
            End If
            Unload FrmProduct
        End With
    Exit Sub
         
    Case 1
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
        With FrmProduct
            .Caption = "Ubah"
            .Text1.Text = ListView1.SelectedItem.Text
            .Text2.Text = ListView1.SelectedItem.SubItems(1)
            .TdbExpiryDate.Value = ListView1.SelectedItem.SubItems(2)
            .Text1.Locked = True
            .Text1.TabStop = False
            .Text1.BackColor = &H8000000F
            .Text1.Appearance = 0
            .Show vbModal
            If .ok Then
                CMDSQL = "Update mandiri.TblInformationLokasi "
                CMDSQL = CMDSQL + " Set Direktori= '" + .Text2.Text + ","
                CMDSQL = CMDSQL + " ExpiryDate=  '" + Format(.TdbExpiryDate.Value, "dd/mm/yyyy") + "'"
                CMDSQL = CMDSQL + " Where Description= '" + .Text1.Text + "'"
                M_OBJCONN.Execute CMDSQL
                On Error GoTo add_error
                    ListView1.SelectedItem.SubItems(1) = .Text2.Text
                    ListView1.SelectedItem.SubItems(2) = .TdbExpiryDate.Value
                On Error GoTo 0
            End If
            Unload FrmProduct
        End With
    Exit Sub
        
    Case 2
        If ListView1.ListItems.Count = 0 Then
            Exit Sub
        End If
        
        'm_msgbox = MsgBox("Yakin Data Ini Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
        If MsgBox("Yakin Data Ini Akan Dihapus...??", vbYesNo + vbInformation, "Pemberitahuan") = vbYes Then
            M_OBJCONN.Execute "Delete from mandiri.TblInformationLokasi where Description ='" + ListView1.SelectedItem.Text + "'"
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
        Else
            MsgBox "Data Tidak Jadi Dihapus", vbOKOnly + vbInformation, "Informasi"
        End If
    Exit Sub
    Case 3
        Unload Me
    Exit Sub
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

