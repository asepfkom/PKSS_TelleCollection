VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form formmonitoringperpindahandata 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitoring Perpindahan Data"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13635
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "List"
      ForeColor       =   &H8000000E&
      Height          =   6015
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   13335
      Begin MSComctlLib.ListView list1 
         Height          =   5565
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   13050
         _ExtentX        =   23019
         _ExtentY        =   9816
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Filter"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      Begin VB.CommandButton Command2 
         Caption         =   "Export"
         Height          =   615
         Left            =   12120
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   615
         Left            =   12120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   1080
         Width           =   4695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "formmonitoringperpindahandata.frx":0000
         Left            =   1560
         List            =   "formmonitoringperpindahandata.frx":0022
         TabIndex        =   5
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   195
         Width           =   4695
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Campaign"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank/Fintech"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Custid :"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "formmonitoringperpindahandata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header()
    list1.ColumnHeaders.ADD , , "CUSTID", 2000
    list1.ColumnHeaders.ADD , , "NAMA", 2000
    list1.ColumnHeaders.ADD , , "CAMPAIGN", 2000
    list1.ColumnHeaders.ADD , , "AGENT NOW", 2000
    list1.ColumnHeaders.ADD , , "AGENT HST", 2000
    list1.ColumnHeaders.ADD , , "TANGGAL PERTAMA TOUCH", 2000
    list1.ColumnHeaders.ADD , , "STATUS TERAKHIR", 2000
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
    Call isilist
End Sub

Private Sub Command2_Click()
Dim objExcel As New Excel.Application
Dim objExcelSheet As Excel.Worksheet
Dim col, row As Integer
Dim a As String
If list1.ListItems.Count > 0 Then
    objExcel.Workbooks.ADD
    Set objExcelSheet = objExcel.Worksheets.ADD
 

    For col = 1 To list1.ColumnHeaders.Count
        objExcelSheet.Cells(1, col).Value = list1.ColumnHeaders(col)
    Next
 
    For row = 2 To list1.ListItems.Count + 1
        For col = 1 To list1.ColumnHeaders.Count
        If col = 1 Then
                objExcelSheet.Cells(row, col).Value = list1.ListItems(row - 1).text
        Else
            '" 'cararandy 29032016 "
            Dim hasil1 As String
                If col <> 12 And col <> 14 Then
                    hasil1 = "'" + list1.ListItems(row - 1).SubItems(col - 1)
                    objExcelSheet.Cells(row, col).Value = hasil1
                Else
                    hasil1 = list1.ListItems(row - 1).SubItems(col - 1)
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
    Call campaign
    Call list_client(Combo1)
End Sub

Private Sub campaign()
    sStrsql = "select * from datasourcetbl where   status ='1' order by  tglentry,  kodeds "
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        Combo2.CLEAR
        While Not M_objrs.EOF
            Combo2.AddItem IIf(IsNull(M_objrs!KODEDS), "", M_objrs!KODEDS)
            M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub

Private Sub isilist()
    Dim qsel As String
    
    qsel = "select * from (" & vbCrLf
    qsel = qsel & "select c.*, coalesce(d.f_cek_new,'') status_terakhir from (" & vbCrLf
    qsel = qsel & "select a.custid, name, recsource, a.agent agent_now, b.agent agent_hst, tanggal_touch_pertama from (" & vbCrLf
    qsel = qsel & "select custid, name, recsource, agent from mgm ) a left join" & vbCrLf
    qsel = qsel & "(" & vbCrLf
    qsel = qsel & "select custid, agent, min(tgl) tanggal_touch_pertama from (" & vbCrLf
    qsel = qsel & "select distinct custid, agent, date(tgl) as tgl from (" & vbCrLf
    qsel = qsel & "select custid, agent, tgl from mgm_hst where custid in (select custid from mgm)) abc order by 1,3" & vbCrLf
    qsel = qsel & ") abc group by 1,2 order by 1,3" & vbCrLf
    qsel = qsel & ") b on a.custid = b.custid order by 1,5" & vbCrLf
    qsel = qsel & ") c" & vbCrLf
    qsel = qsel & "left join (" & vbCrLf
    qsel = qsel & "select custid, f_cek_new from mgm_hst where id in (" & vbCrLf
    qsel = qsel & "select max(id) id from mgm_hst where custid in (select custid from mgm) group by custid)) d" & vbCrLf
    qsel = qsel & "on c.custid = d.custid" & vbCrLf
    qsel = qsel & ") e where 1 = 1 " & vbCrLf
    
    If Text1.text <> "" Then
        qsel = qsel & " and custid = '" & Text1.text & "' " & vbCrLf
    End If
    If Combo1.text <> "" Then
        qsel = qsel & " and recsource ilike '%" & Combo1.text & "%' " & vbCrLf
    End If
    If Combo2.text <> "" Then
        qsel = qsel & " and recsource = '" & Combo2.text & "' "
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open qsel & " order by 1,6", M_OBJCONN, adOpenKeyset, adLockOptimistic
    
    list1.ListItems.CLEAR
    While Not rs.EOF
         Set ListItem = list1.ListItems.ADD(, , cnull(rs(0)))
            For i = 1 To 6
                ListItem.SubItems(i) = cnull(rs(i))
            Next i
         rs.MoveNext
    Wend
    Set rs = Nothing
    
End Sub

