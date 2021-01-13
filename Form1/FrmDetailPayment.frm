VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmDetailPayment 
   Appearance      =   0  'Flat
   BackColor       =   &H00ABE18E&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Detail Payment"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   10815
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstPayment 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4683
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
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmDetailPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim ListItem As ListItem

TXT_X = 70
LstPayment.ColumnHeaders.ADD 1, , "No", 10 * TXT_X
LstPayment.ColumnHeaders.ADD 2, , "Tahun", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 3, , "Jan", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 4, , "Feb", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 5, , "Mar", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 6, , "Apr", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 7, , "Mei", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 8, , "Jun", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 9, , "Jul", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 10, , "Aug", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 11, , "Sep", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 12, , "Okt", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 13, , "Nop", 15 * TXT_X
LstPayment.ColumnHeaders.ADD 14, , "Des", 15 * TXT_X

'Cmdsql = " SELECT m.custid, m.tahun, COALESCE(m.""1"", 0) AS ""Jan"", COALESCE(m.""2"", 0) AS ""Feb"", COALESCE(m.""3"", 0) AS ""Mar"", COALESCE(m.""4"", 0) AS ""Apr"", COALESCE(m.""5"", 0) AS ""Mei"", COALESCE(m.""6"", 0) AS ""Jun"", COALESCE(m.""7"", 0) AS ""Jul"", COALESCE(m.""8"", 0) AS ""Aug"", COALESCE(m.""9"", 0) AS ""Sep"", COALESCE(m.""10"", 0) AS ""Okt"", COALESCE(m.""11"", 0) AS ""Nop"", COALESCE(m.""12"", 0) AS ""Des"""
'Cmdsql = Cmdsql + "  FROM crosstab('select custid, date_part(''year'',paydate)as tahun, date_part(''month'',paydate) as bulan, sum(payment) as payment from tbllunas"
'Cmdsql = Cmdsql + "   where custid=''" + FrmCC_Colection.lblCustId.text + "'' group by custid, tahun,bulan order by  tahun,custid'::text, 'select m from generate_series(1,12) m'::text) m(custid text, ""tahun"" text,  ""1"" numeric, ""2"" numeric, ""3"" numeric, ""4"" numeric, ""5"" numeric, ""6"" numeric, ""7"" numeric, ""8"" numeric, ""9"" numeric, ""10"" numeric, ""11"" numeric, ""12"" numeric);"

CMDSQL = "SELECT m.custid, m.tahun, COALESCE(m.""1"", 0) AS ""Jan"","
CMDSQL = CMDSQL + " COALESCE(m.""2"", 0) AS ""Feb"", COALESCE(m.""3"", 0) AS ""Mar"", COALESCE(m.""4"", 0) AS ""Apr"", COALESCE(m.""5"", 0) AS ""Mei"","
CMDSQL = CMDSQL + " COALESCE(m.""6"", 0) AS ""Jun"", COALESCE(m.""7"", 0) AS ""Jul"", COALESCE(m.""8"", 0) AS ""Aug"", COALESCE(m.""9"", 0) AS ""Sep"","
CMDSQL = CMDSQL + " COALESCE(m.""10"", 0) AS ""Okt"", COALESCE(m.""11"", 0) AS ""Nop"", COALESCE(m.""12"", 0) AS ""Des""  "
CMDSQL = CMDSQL + " FROM crosstab('select date_part(''year'',paydate)as tahun,custid,date_part(''month'',paydate) as bulan, "
CMDSQL = CMDSQL + " sum(payment) as payment from tbllunas where custid=''" + FrmCC_Colection.lblCustId.Text + "''"
CMDSQL = CMDSQL + " group by tahun,custid,bulan order by  tahun,custid'::text, 'select m from generate_series(1,12) m'::text) m(custid text, ""tahun"" text,  ""1"" numeric, ""2"" numeric, ""3"" numeric, ""4"" numeric, ""5"" numeric, ""6"" numeric, ""7"" numeric, ""8"" numeric, ""9"" numeric, ""10"" numeric, ""11"" numeric, ""12"" numeric);"

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not rs.EOF
 Set ListItem = LstPayment.ListItems.ADD(, , rs.Bookmark)
     ListItem.SubItems(1) = IIf(IsNull(rs!Tahun), "", rs!Tahun)
     ListItem.SubItems(2) = Format(IIf(IsNull(rs!Jan), "", rs!Jan), "#,###,###")
     ListItem.SubItems(3) = Format(IIf(IsNull(rs!feb), "", rs!feb), "#,###,###")
     ListItem.SubItems(4) = Format(IIf(IsNull(rs!Mar), "", rs!Mar), "#,###,###")
     ListItem.SubItems(5) = Format(IIf(IsNull(rs!Apr), "", rs!Apr), "#,###,###")
     ListItem.SubItems(6) = Format(IIf(IsNull(rs!Mei), "", rs!Mei), "#,###,###")
     ListItem.SubItems(7) = Format(IIf(IsNull(rs!Jun), "", rs!Jun), "#,###,###")
     ListItem.SubItems(8) = Format(IIf(IsNull(rs!Jul), "", rs!Jul), "#,###,###")
     ListItem.SubItems(9) = Format(IIf(IsNull(rs!Aug), "", rs!Aug), "#,###,###")
     ListItem.SubItems(10) = Format(IIf(IsNull(rs!Sep), "", rs!Sep), "#,###,###")
     ListItem.SubItems(11) = Format(IIf(IsNull(rs!Okt), "", rs!Okt), "#,###,###")
     ListItem.SubItems(12) = Format(IIf(IsNull(rs!Nop), "", rs!Nop), "#,###,###")
     ListItem.SubItems(13) = Format(IIf(IsNull(rs!des), "", rs!des), "#,###,###")
     rs.MoveNext
Wend

Set rs = Nothing
End Sub
