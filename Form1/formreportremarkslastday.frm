VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form formreportremarkslastday 
   Caption         =   "Report Remarks Last Day"
   ClientHeight    =   8115
   ClientLeft      =   1995
   ClientTop       =   1515
   ClientWidth     =   16935
   LinkTopic       =   "Form2"
   ScaleHeight     =   8115
   ScaleWidth      =   16935
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   13440
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Export"
      Height          =   375
      Left            =   15240
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16935
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "formreportremarkslastday.frx":0000
         Left            =   1680
         List            =   "formreportremarkslastday.frx":001F
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
      Begin MSComctlLib.ListView lv1 
         Height          =   7380
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   16740
         _ExtentX        =   29528
         _ExtentY        =   13018
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
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   0
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         Caption         =   "Harap Pilih Bank/Fintech"
         Height          =   255
         Left            =   5280
         TabIndex        =   6
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "BANK/FinTech"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   315
         Width           =   1215
      End
   End
End
Attribute VB_Name = "formreportremarkslastday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub header()
    lv1.ColumnHeaders.CLEAR
    With lv1.ColumnHeaders
        .ADD 1, , "CUSTID", 3000
        .ADD 2, , "AGENT", 3000
        .ADD 3, , "HISTORY", 3000
        .ADD 4, , "STATUS DATA", 3000
        .ADD 5, , "DATE", 3000
        .ADD 6, , "PHONE NUMBER", 3000
        .ADD 7, , "STATUSCALL", 3000
    End With
End Sub

Private Sub search()
    
    strsql = " select a.custid, a.agent, a.hst, a.f_cek, a.tgl, a.phoneno, a.statuscall from mgm_hst a inner join ( " & vbCrLf
    strsql = strsql & " select a.* from ( " & vbCrLf
    strsql = strsql & " select custid, tgl " & vbCrLf
    strsql = strsql & "  from mgm_hst where custid in (select custid from mgm where recsource  ilike '%" & Combo1.text & "%')) a " & vbCrLf
    strsql = strsql & " inner join ( " & vbCrLf
    strsql = strsql & " select custid, tgl::date as tgl from ( " & vbCrLf
    strsql = strsql & " select max(tgl) as tgl, custid from mgm_hst where custid in ( " & vbCrLf
    strsql = strsql & " select custid from mgm where recsource  ilike '%" & Combo1.text & "%') group by 2 " & vbCrLf
    strsql = strsql & " ) a " & vbCrLf
    strsql = strsql & " ) b on a.custid = b.custid and date(a.tgl) = b.tgl order by 1,2 " & vbCrLf
    strsql = strsql & " ) b on a.custid = b.custid and a.tgl = b.tgl "

    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    lv1.ListItems.CLEAR

    While Not M_objrs.EOF
        Set ListItem = lv1.ListItems.ADD(, , cnull(M_objrs("custid")))
            ListItem.SubItems(1) = cnull(M_objrs("agent"))
            ListItem.SubItems(2) = cnull(M_objrs("hst"))
            ListItem.SubItems(3) = cnull(M_objrs("f_cek"))
            ListItem.SubItems(4) = cnull(M_objrs("tgl"))
            ListItem.SubItems(5) = cnull(M_objrs("phoneno"))
            ListItem.SubItems(6) = cnull(M_objrs("statuscall"))
        M_objrs.MoveNext
    Wend
     
    Set M_objrs = Nothing

End Sub

Private Sub Command1_Click()
    If Combo1.text = "" Then
        MsgBox "Harap Pilih Bank/Fintect"
        Exit Sub
    End If
    
    Call search
End Sub

Private Sub Command2_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, row As Integer
    Dim a As String
    If lv1.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To lv1.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = lv1.ColumnHeaders(col)
        Next
     
        For row = 2 To lv1.ListItems.Count + 1
            For col = 1 To lv1.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(row, col).Value = lv1.ListItems(row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    If col <> 12 And col <> 14 Then
                        hasil1 = "'" + lv1.ListItems(row - 1).SubItems(col - 1)
                        objExcelSheet.Cells(row, col).Value = hasil1
                    Else
                        hasil1 = lv1.ListItems(row - 1).SubItems(col - 1)
                        objExcelSheet.Cells(row, col).Value = hasil1
                    End If
                End If
            Next
        Next
     
        objExcelSheet.Columns.AutoFit
        CD_save.ShowOpen
        a = CD_save.FileName
     
        'objExcelSheet.SaveAs a & ".xlsx"
        objExcelSheet.SaveAs a & ".xls"
        MsgBox "Export Completed", vbInformation, Me.Caption
     
        'objExcel.Workbooks.Open a & ".xlsx"
        objExcel.Workbooks.Open a & ".xls"
        objExcel.Visible = True
    Else
        MsgBox "No data to export", vbInformation, Me.Caption
    End If

End Sub

Private Sub Form_Load()
    Call header
    Call list_client(Combo1)
End Sub
