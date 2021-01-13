VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmtarikdata 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tarik Data"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9870
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Cara Pakai"
      Height          =   1575
      Left            =   6240
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Label Label6 
         BackColor       =   &H80000007&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   $"frmtarikdata.frx":0000
         Height          =   1095
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Searching"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.Timer Timer1 
         Interval        =   600
         Left            =   0
         Top             =   0
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   285
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmtarikdata.frx":00DC
         Left            =   1920
         List            =   "frmtarikdata.frx":00FB
         TabIndex        =   5
         Top             =   765
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Export"
         Height          =   375
         Left            =   5880
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   0
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "Cara Pakai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Found :"
         Height          =   255
         Left            =   8280
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Campaign"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "BANK/FinTech"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView LvPTP 
      Height          =   4620
      Left            =   0
      TabIndex        =   7
      Top             =   1440
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   8149
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
End
Attribute VB_Name = "frmtarikdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_objrs As ADODB.Recordset

Private Sub Command1_Click()
    If Combo2.text = "" Then
        MsgBox "Campaign tidak boleh kosong"
    'If Combo1.text = "" And Combo2.text = "" Then
        'MsgBox "BANK/Fintect atau Campaign tidak boleh kosong"
        Exit Sub
    End If
    Call search
End Sub

Private Sub Command2_Click()
Dim objExcel As New Excel.Application
Dim objExcelSheet As Excel.Worksheet
Dim col, row As Integer
Dim a As String
Dim sql As String

If LvPTP.ListItems.Count > 0 Then
    objExcel.Workbooks.ADD
    Set objExcelSheet = objExcel.Worksheets.ADD
 

    For col = 1 To LvPTP.ColumnHeaders.Count
        objExcelSheet.Cells(1, col).Value = LvPTP.ColumnHeaders(col)
    Next
 
    For row = 2 To LvPTP.ListItems.Count + 1
        For col = 1 To LvPTP.ColumnHeaders.Count
        If col = 1 Then
                objExcelSheet.Cells(row, col).Value = LvPTP.ListItems(row - 1).text
        Else
            '" 'cararandy 29032016 "
            Dim hasil1 As String
                If col <> 12 And col <> 14 Then
                    hasil1 = "'" + LvPTP.ListItems(row - 1).SubItems(col - 1)
                    objExcelSheet.Cells(row, col).Value = hasil1
                Else
                    hasil1 = LvPTP.ListItems(row - 1).SubItems(col - 1)
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
    '====asep===='
    
    M_OBJCONN.Execute "create Table " & Combo2.text & "  as select * from mgm where recsource = '" & Combo2.text & "'  "
    M_OBJCONN.Execute "delete from mgm where recsource = '" & Combo2.text & "' "
    'recsource = 'PKSS-'||to_char(now(),'yyyymmddhhmiss')
    '============'
Else
    MsgBox "No data to export", vbInformation, Me.Caption
End If


End Sub

Private Sub Form_Load()
    Call campaign
    Call header
    Call supervisorole
    Call list_client(Combo1)
End Sub

Private Sub supervisorole()
    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        q = "select distinct recsource from mgm where agent in (select userid from usertbl where team = '" & MDIForm1.TxtUsername.text & "' or userid = '" & MDIForm1.TxtUsername.text & "' )  "
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        Dim aa
        Dim zz As String
        Dim sss As String
        aa = Array("BCA", "BRI", "HCI", "MANDIRI", "MAYBANK", "PANIN", "PLUS", "EXPRES", "GLOBAL", "COURT")
        
        
        sss = ""
        zz = ""
        Combo1.CLEAR
        While Not M_objrs.EOF
        'If M_objrs.RecordCount > 0 Then
                For i = 1 To 10
                    a = aa(i - 1)
                    If M_objrs!recsource Like "*" & a & "*" Then
                        If aa(i - 1) = "PLUS" Then
                            If sss Like "*PLUS*" Then
                            Else
                                Combo1.AddItem "RUPIAH PLUS"
                                sss = sss & " PLUS "
                            End If
                        ElseIf aa(i - 1) = "EXPRES" Then
                            If sss Like "*EXPRES*" Then
                            Else
                                Combo1.AddItem "UANGEXPRESS"
                                sss = sss & " EXPRES "
                            End If
                        ElseIf aa(i - 1) = "GLOBAL" Then
                            If sss Like "*GLOBAL*" Then
                            Else
                                Combo1.AddItem "GLOBALINDO"
                                sss = sss & " GLOBAL "
                            End If
                        Else
                            'If zz Like "*" & aa(i - 1) & "*" Then
                            'Else
                            If sss Like "*" & aa(i - 1) & "*" Then
                            Else
                                Combo1.AddItem aa(i - 1)
                                zz = zz & " " & aa(i - 1)
                                sss = sss & " " & aa(i - 1) & " "
                            End If
                        End If
                    End If
                Next i
            M_objrs.MoveNext
        'End If
        Wend
        
    End If

End Sub


Private Sub header()
    LvPTP.ColumnHeaders.CLEAR
    With LvPTP.ColumnHeaders
        .ADD 1, , "CR_NAME_1"
        .ADD 2, , "ADDRESSNOW"
        .ADD 3, , "HOMEPHONE"
        .ADD 4, , "MOBILEPHONE"
        .ADD 5, , "ADDRESSOFFICE"
        .ADD 6, , "OFFICEPHONE"
        .ADD 7, , "CARDNO"
        .ADD 8, , "REGION"
        .ADD 9, , "RECSOURCE"
        .ADD 10, , "CUSTID"
        .ADD 11, , "CM_TOT_BALANCE"
        .ADD 12, , "PAYDATE"
        .ADD 13, , "LASTPAY"
        .ADD 14, , "ECPHONE"
        .ADD 15, , "RO"
        .ADD 16, , "REMARKSOLD"
        '.ADD 17, , "ECPHONE2"
        .ADD 17, , "DPD"
        '========asep==========='
        .ADD 18, , "CR_ADDR"
        .ADD 19, , "CO_DATE"
        .ADD 20, , "CM_STATUS"
        .ADD 21, , "CR_ZIP_CODE"
        .ADD 22, , "CR_EU_SEX"
        .ADD 23, , "JENIS KELAMIN"
        .ADD 24, , "ECDESC"
        .ADD 25, , "CM_SHORT_NAME"
        .ADD 26, , "CM_BLOCK_CODE"
        '.ADD 28, , "ECPHONE"
        .ADD 27, , "AGENT"
        '======================='
    End With

End Sub


Private Sub Label4_Click()
    Frame2.Visible = True
End Sub

Private Sub Label6_Click()
    Frame2.Visible = False
End Sub

Private Sub Timer1_Timer()
    If Label4.BackColor = &H8000000D Then
        Label4.BackColor = &H8000000F
    Else
        Label4.BackColor = &H8000000D
    End If
End Sub

Private Sub campaign()
    'sStrsql = "select * from datasourcetbl where   status ='1' "
    sStrsql = "select distinct recsource from mgm "
    
    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        'sStrsql = sStrsql & " and KODEDS in (select distinct recsource from mgm where agent in (select userid from usertbl where spvcode = '" & MDIForm1.txtusername.text & "' or userid = '" & MDIForm1.txtusername.text & "'))"
        '=====asep====='
        sStrsql = "select distinct recsource from mgm where agent in (select userid from usertbl where spvcode = '" & MDIForm1.TxtUsername.text & "' or userid = '" & MDIForm1.TxtUsername.text & "')"
        '============='
    End If
    
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        'M_objrs.Open sStrsql & " order by  kodeds ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        '====asep====='
        M_objrs.Open sStrsql & " order by  recsource ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        '============='
        Combo2.CLEAR
        While Not M_objrs.EOF
            'Combo2.AddItem IIf(IsNull(M_objrs!KODEDS), "", M_objrs!KODEDS)
            '====asep===='
            Combo2.AddItem IIf(IsNull(M_objrs!recsource), "", M_objrs!recsource)
            '============'
            M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub

Private Sub search()
    Field = "custid,name,addrnow,homeno,mobileno,addrpt,officeno,nocard,region,recsource,tglincoming,cycle,stskathomeadd1,stskathomeadd2,stskatofficeadd2,f_sts_valid_home2,f_sts_valid_office1,curbal,pay_dt,lastpay,product_desc,batchdiskon,remarks_old,afaxno,delq_history,co_office_phone_1,block_code_1,zipnow,'" & limit & "',agent"
    'Field = "name,addrnow,homeno,mobileno,officeno,nocard,region,dob,recsource,custid,curbal,pay_dt,lastpay,batchdiskon,remarks_old,delq_history"

    If Combo1.text = "RUPIAHPLUS" Then
        sStrsql = "select " & Field & " from mgm where recsource ilike '%PLUS%' or recsource = '" & Combo2.text & "' "
    ElseIf Combo1.text = "UANGEXPRESS" Then
        sStrsql = "select " & Field & " from mgm where recsource ilike '%EXPRES%' or recsource = '" & Combo2.text & "' "
    ElseIf Combo1.text = "GLOBALINDO" Then
        sStrsql = "select " & Field & " from mgm where recsource ilike '%GLOBAL%' or recsource = '" & Combo2.text & "' "
    Else
        If Combo2.text <> "" Then
            sStrsql = "select " & Field & " from mgm where recsource = '" & Combo2.text & "' "
        Else
            sStrsql = "select " & Field & " from mgm where recsource ilike '%" & Combo1.text & "%' or recsource = '" & Combo2.text & "' "
        End If
    End If
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    LvPTP.ListItems.CLEAR
    While Not M_objrs.EOF
        Set ListItem = LvPTP.ListItems.ADD(, , cnull(M_objrs("name")))
            ListItem.SubItems(1) = cnull(M_objrs("AddrNow"))
            ListItem.SubItems(2) = cnull(M_objrs("homeno"))
            ListItem.SubItems(3) = cnull(M_objrs("mobileno"))
            ListItem.SubItems(4) = cnull(M_objrs("addrpt"))
            ListItem.SubItems(5) = cnull(M_objrs("officeno"))
            ListItem.SubItems(6) = cnull(M_objrs("nocard"))
            ListItem.SubItems(7) = cnull(M_objrs("region"))
            ListItem.SubItems(8) = cnull(M_objrs("recsource"))
            ListItem.SubItems(9) = cnull(M_objrs("CustId"))
            ListItem.SubItems(10) = cnull(M_objrs("curbal"))
            ListItem.SubItems(11) = cnull(M_objrs("pay_dt"))
            ListItem.SubItems(12) = cnull(M_objrs("lastpay"))
            ListItem.SubItems(13) = cnull(M_objrs("product_desc"))
            ListItem.SubItems(14) = cnull(M_objrs("batchdiskon"))
            ListItem.SubItems(15) = cnull(M_objrs("remarks_old"))
            'ListItem.SubItems(16) = cnull(M_objrs("afaxno"))
            ListItem.SubItems(16) = cnull(M_objrs("delq_history"))
            '===asep===='
            ListItem.SubItems(17) = cnull(M_objrs("addrpt"))
            ListItem.SubItems(18) = cnull(M_objrs("tglincoming"))
            ListItem.SubItems(19) = cnull(M_objrs("stskathomeadd1"))
            ListItem.SubItems(20) = cnull(M_objrs("zipnow"))
            ListItem.SubItems(21) = cnull(M_objrs("f_sts_valid_home2"))
            ListItem.SubItems(22) = cnull(M_objrs("f_sts_valid_office1"))
            ListItem.SubItems(23) = cnull(M_objrs("co_office_phone_1"))
            ListItem.SubItems(24) = cnull(M_objrs("stskathomeadd2"))
            ListItem.SubItems(25) = cnull(M_objrs("block_code_1"))
            'ListItem.SubItems(27) = cnull(M_objrs("stskatofficeadd2"))
            ListItem.SubItems(26) = cnull(M_objrs("agent"))
            'ListItem.SubItems(29) = cnull(M_objrs("delq_history"))
            '==========='
        M_objrs.MoveNext
    Wend
    
    Label3.Caption = "Found : " & M_objrs.RecordCount
    
    Set M_objrs = Nothing
End Sub
