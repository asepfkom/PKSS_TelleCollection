VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form4"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12570
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   12570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "BTN"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5355
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   9446
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
   Begin VB.Label Label1 
      Caption         =   "Table"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim LS As listItem
    
    ListView1.ListItems.CLEAR
    query = "select column_name from information_schema.columns where table_name = '" + Combo1.Text + "'"
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
        While Not RS_Lv.EOF
            Set LS = ListView1.ListItems.ADD(, , Trim(RS_Lv("column_name")) + ",")
            RS_Lv.MoveNext
        Wend
    
        RS_Lv.Close
        Set RS_Lv = Nothing
End Sub

Private Sub Form_Load()
    Call header1
    Call Isi_Table
End Sub

Private Sub Isi_Table()
    sQuery = "select distinct(table_name) as table from information_schema.columns order by 1"
    Set RS_Lv = New ADODB.Recordset
    RS_Lv.CursorLocation = adUseClient
    RS_Lv.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If RS_Lv.RecordCount > 0 Then
        While Not RS_Lv.EOF
            Combo1.AddItem RS_Lv!Table
            RS_Lv.MoveNext
        Wend
    End If
End Sub

Private Sub header1()
    ListView1.ColumnHeaders.ADD 1, , "Query", 10 * 1000
End Sub
