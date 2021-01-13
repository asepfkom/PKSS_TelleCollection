VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form formattemptpercustidperagent 
   Caption         =   "Form Attempt Per Custid Per Agent"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6120
   LinkTopic       =   "Form2"
   ScaleHeight     =   6645
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Export"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   6060
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   10689
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
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "formattemptpercustidperagent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    Call load
End Sub

Private Sub header()
    lv1.ColumnHeaders.CLEAR
    With lv1.ColumnHeaders
        .ADD 1, , "AGENT", 3000
        .ADD 2, , "CUSTID", 1500
        .ADD 3, , "TOUCH", 1500
    End With
End Sub

Private Sub load()
    strsql = " select b.*,a.touch from ( " & vbCrLf
    strsql = strsql & " select agent, count(agent) touch from mgm_hst where tgl between date(now()) and date(now() + interval '1 day') group by 1 " & vbCrLf
    strsql = strsql & " ) a,( " & vbCrLf
    strsql = strsql & " select agent, count(custid) custid from ( " & vbCrLf
    strsql = strsql & " select distinct(custid),agent from mgm_hst where tgl between date(now()) and date(now() + interval '1 day') group by 1,2) a group by 1 " & vbCrLf
    strsql = strsql & " ) b where a.agent = b.agent "
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    lv1.ListItems.CLEAR

    While Not M_objrs.EOF
        Set ListItem = lv1.ListItems.ADD(, , cnull(M_objrs("agent")))
            ListItem.SubItems(1) = cnull(M_objrs("custid"))
            ListItem.SubItems(2) = cnull(M_objrs("touch"))
        M_objrs.MoveNext
    Wend
     
    Set M_objrs = Nothing

End Sub
