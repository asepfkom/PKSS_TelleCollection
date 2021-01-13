VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form formlogdistribute 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Distribute"
   ClientHeight    =   4905
   ClientLeft      =   6240
   ClientTop       =   3045
   ClientWidth     =   8730
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8730
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.ComboBox cbosheet 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   300
         Width           =   3165
      End
      Begin MSComctlLib.ListView LvPTP 
         Height          =   3900
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   6879
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "formlogdistribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub gettanggal()
    qs = "select replace(tanggal,'i-',':')::timestamp without time zone tanggal, tabel from (" & vbCrLf
    qs = qs + "select replace(tanggal,'i--',' ') tanggal, tabel from (" & vbCrLf
    qs = qs + "select replace(tanggal, '_','-') as tanggal, tabel from (" & vbCrLf
    qs = qs + "select split_part(table_name,'___', 2) as tanggal, tabel from (" & vbCrLf
    qs = qs + "select distinct table_name, table_name as tabel from information_schema.columns where table_name ilike '%backupdistribute__%' group by 2" & vbCrLf
    qs = qs + ") a" & vbCrLf
    qs = qs + ") b" & vbCrLf
    qs = qs + ") c" & vbCrLf
    qs = qs + ") d"
    
    'qs = "select * from (select distinct table_name from information_schema.columns where table_name ilike '%backupdistribute__%');"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    cbosheet.CLEAR
    
    While Not M_objrs.EOF
        cbosheet.AddItem Format(M_objrs!Tanggal, "yyyy-mm-dd hh:m:ss")
        M_objrs.MoveNext
    Wend
End Sub

Private Sub cbosheet_Click()
    Call called
End Sub

Private Sub Form_Load()
    Call header
    Call gettanggal
End Sub

Private Sub called()
    qs = "select * from ("
    qs = qs + "select replace(tanggal,'i-',':')::timestamp without time zone tanggal, tabel from (" & vbCrLf
    qs = qs + "select replace(tanggal,'i--',' ') tanggal, tabel from (" & vbCrLf
    qs = qs + "select replace(tanggal, '_','-') as tanggal, tabel from (" & vbCrLf
    qs = qs + "select split_part(table_name,'___', 2) as tanggal, tabel from (" & vbCrLf
    qs = qs + "select distinct table_name, table_name as tabel from information_schema.columns where table_name ilike '%backupdistribute__%' group by 2" & vbCrLf
    qs = qs + ") a" & vbCrLf
    qs = qs + ") b" & vbCrLf
    qs = qs + ") c" & vbCrLf
    qs = qs + ") d" & vbCrLf
    qs = qs + ") b where tanggal = '" & cbosheet.text & "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    qs = "select * from " & M_objrs!tabel
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    LvPTP.ListItems.CLEAR
    
    While Not M_objrs.EOF
        Set ListItem = LvPTP.ListItems.ADD(, , cnull(M_objrs("custid")))
            ListItem.SubItems(1) = cnull(M_objrs("agent_lama"))
            ListItem.SubItems(2) = cnull(M_objrs("agent_baru"))
            ListItem.SubItems(3) = cnull(M_objrs("distributeby"))
            ListItem.SubItems(4) = cnull(M_objrs("Tanggal"))
        M_objrs.MoveNext
    Wend
End Sub

Private Sub header()
    LvPTP.ColumnHeaders.CLEAR
    With LvPTP.ColumnHeaders
        .ADD 1, , "CUSTID"
        .ADD 2, , "AGENT LAMA"
        .ADD 3, , "AGENT BARU"
        .ADD 4, , "DISTRIBUTED"
        .ADD 5, , "TANGGAL"
    End With

End Sub
