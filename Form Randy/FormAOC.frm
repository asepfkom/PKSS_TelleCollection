VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FormAOC 
   Caption         =   "Form AOC"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form3"
   ScaleHeight     =   6480
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnclear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   9120
      TabIndex        =   10
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton btndelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton btnupdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton btnsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtaoc 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtuserid 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4755
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   8387
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   4755
      Left            =   3840
      TabIndex        =   6
      Top             =   1560
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   8387
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Caption         =   "AOC"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblnama 
      Caption         =   "Nama"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "USERID"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "FormAOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnclear_Click()
Call clearfield
End Sub

Private Sub btndelete_Click()
    Dim M_DATA As New CLSSPV_AGENT
    M_DATA.DELETE_AOC M_OBJCONN
    'M_Objrs.Close
    Set M_Objrs = Nothing
    ListView2.ListItems.clear
    Call Form_Load
End Sub

Private Sub btnsave_Click()
    Dim M_DATA As New CLSSPV_AGENT
    CC = M_DATA.QUERY_AOC_CHECK(M_OBJCONN)
    'test.Caption = M_DATA.QUERY_AOC_CHECK(M_OBJCONN)
    If CC = True Then
     MsgBox "USERID INI SUDAH ADA", vbOKOnly + vbInformation, "Informasi"
    Else
    M_DATA.ADD_AOC M_OBJCONN
    End If
    'M_Objrs.Close
    Set M_Objrs = Nothing
    ListView2.ListItems.clear
    Call Form_Load
End Sub

Private Sub btnupdate_Click()
    Dim M_DATA As New CLSSPV_AGENT
    M_DATA.UPDATE_AOC M_OBJCONN
    'M_Objrs.Close
    Set M_Objrs = Nothing
    ListView2.ListItems.clear
    Call Form_Load
End Sub

Private Sub Form_Load()
    Call LS1
    Call LS2
    Call clearfield
End Sub

Private Sub header1()
    ListView1.ColumnHeaders.ADD 1, , "User Id", 10 * 120
    ListView1.ColumnHeaders.ADD 2, , "Nama Agent", 20 * 120
End Sub

Private Sub header2()
    ListView2.ColumnHeaders.ADD 1, , "User Id", 10 * 120
    ListView2.ColumnHeaders.ADD 2, , "Nama Agent", 20 * 120
    ListView2.ColumnHeaders.ADD 3, , "AOC", 30 * 120
End Sub

Private Sub clearfield()
txtaoc.Text = ""
txtuserid.Text = ""
lblnama.Caption = ""
End Sub



Private Sub ListView1_DblClick()
    If ListView1.SelectedItem.Text <> "" Then
        lblnama.Caption = ListView1.SelectedItem.Text
        txtuserid.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
    End If
End Sub

Private Sub LS1()
    Dim M_Objrs As ADODB.Recordset
    Dim M_DATA As New CLSSPV_AGENT
    Dim LS As listItem
    Call header1
    Set M_Objrs = M_DATA.QUERY_AOC(M_OBJCONN)
    While Not M_Objrs.EOF
         Set LS = ListView1.ListItems.ADD(, , Trim(M_Objrs("USERID")))
             LS.SubItems(1) = M_Objrs("AGENT")
         M_Objrs.MoveNext
    Wend
        M_Objrs.Close
        Set M_Objrs = Nothing
End Sub

Private Sub LS2()
    Dim M_Objrs As ADODB.Recordset
    Dim M_DATA As New CLSSPV_AGENT
    Dim LS As listItem
    Call header2
    Set M_Objrs = M_DATA.SHOW_AOC(M_OBJCONN)
    While Not M_Objrs.EOF
         Set LS = ListView2.ListItems.ADD(, , Trim(M_Objrs("USERID")))
             LS.SubItems(1) = M_Objrs("AGENT")
             LS.SubItems(2) = M_Objrs("AOC")
         M_Objrs.MoveNext
    Wend
        M_Objrs.Close
        Set M_Objrs = Nothing
End Sub

Private Sub ListView2_DblClick()
    If ListView2.SelectedItem.Text <> "" Then
        lblnama.Caption = ListView2.SelectedItem.Text
        txtuserid.Text = ListView2.ListItems(ListView2.SelectedItem.Index).SubItems(1)
        txtaoc.Text = ListView2.ListItems(ListView2.SelectedItem.Index).SubItems(2)
    End If
End Sub
