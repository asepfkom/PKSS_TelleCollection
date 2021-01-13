VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form_List_CPA 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form List CPA"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   18915
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   18915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Download Data To Coint"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last Payment"
      Height          =   6495
      Left            =   9480
      TabIndex        =   17
      Top             =   2760
      Width           =   9255
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cek All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.ListView LvLastPayment 
         Height          =   5190
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   9155
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
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
   Begin VB.ComboBox cmb_status 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Form_List_CPA.frx":0000
      Left            =   1200
      List            =   "Form_List_CPA.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Monthly CPA"
      Height          =   6495
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   9255
      Begin VB.CheckBox cek_all_ptp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cek All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin MSComctlLib.ListView LvCPA 
         Height          =   5190
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   9155
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
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
      Begin VB.Label lbldata 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   6000
         Width           =   2655
      End
      Begin VB.Label lbltotal 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL : IDR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   6000
         Width           =   8055
      End
   End
   Begin VB.ComboBox cmb_agent 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Form_List_CPA.frx":0014
      Left            =   1200
      List            =   "Form_List_CPA.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmd_showcpa 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Show CPA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9360
      Width           =   1215
   End
   Begin VB.ComboBox cmb_team 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Form_List_CPA.frx":0018
      Left            =   1200
      List            =   "Form_List_CPA.frx":001A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CheckBox chk_team 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3720
      TabIndex        =   0
      Top             =   1320
      Width           =   195
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Index           =   1
      Left            =   9720
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM-yyyy"
      Format          =   45940739
      CurrentDate     =   41610
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   21
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   101908483
      CurrentDate     =   41444
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Index           =   2
      Left            =   6060
      TabIndex        =   22
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   101580803
      CurrentDate     =   41444
   End
   Begin VB.Label Label3 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   23
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "Status  :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   4
      Left            =   120
      Picture         =   "Form_List_CPA.frx":001C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "List CPA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   630
      TabIndex        =   14
      Top             =   60
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "Agent  :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lbl_team 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "Team   :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "Form_List_CPA.frx":0B26
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20040
   End
End
Attribute VB_Name = "Form_List_CPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs_list As ADODB.Recordset
Dim f_team As Boolean

Private Sub koneksi()
    Set Rs_list = New ADODB.Recordset
    Rs_list.CursorLocation = adUseClient
    Rs_list.ActiveConnection = M_OBJCONN
    Rs_list.CursorType = adOpenDynamic
    Rs_list.LockType = adLockOptimistic
End Sub

Private Sub cek_all_ptp_Click()
    Dim W As Integer
    
    If LvCPA.ListItems.Count = 0 Then
        MsgBox "Maaf data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If cek_all_ptp.Value = 1 Then
        For W = 1 To LvCPA.ListItems.Count
            LvCPA.ListItems(W).Checked = True
        Next W
    Else
        For W = 1 To LvCPA.ListItems.Count
            LvCPA.ListItems(W).Checked = False
        Next W
    End If
End Sub

Private Sub chk_team_Click()
    If chk_team.Value = vbChecked Then
        Call Isi_TL
        cmb_agent.ListIndex = 0
        cmb_agent.Enabled = False
        cmb_team.Enabled = True
        f_team = True
    Else
        cmb_agent.Enabled = True
        cmb_team.Enabled = False
        cmb_team.ListIndex = 0
        f_team = False
    End If
End Sub

Private Sub Isi_TL()
    If Rs_list.state = 1 Then Rs_list.Close
    

    Rs_list.Open "SELECT DISTINCT team FROM mandiri.usertbl where team ilike  'TL%' "
    
    cmb_team.AddItem " "
    
    While Not Rs_list.EOF
        cmb_team.AddItem Rs_list("team")
        Rs_list.MoveNext
    Wend
End Sub

Private Sub cmd_showcpa_Click()
    If cmb_status.ListIndex = 0 Then
        Call IsiListApprove
    Else
        Call IsiListReject
    End If
End Sub

Private Sub IsiListApprove()
    Dim ListItem As ListItem
    Dim AGENT As String
    Dim total_ptp As Double
    Dim bulan_sekarang As String
    Dim tahun_sekarang As String
    Dim id_payment As String
    Dim TotalKeseluruhan As Double
    
'    tanggal_sekarang = Format(DTPicker1.Value, "yyyy-mm-dd")
'
'    bulan_sekarang = Format(tanggal_sekarang, "MM")
'    tahun_sekarang = Format(tanggal_sekarang, "YYYY")
    
    
    AGENT = cmb_agent.Text
    
    If cmb_agent = " " Then
        AGENT = ""
    End If
    
    LvCPA.ListItems.CLEAR
    If Rs_list.state = 1 Then Rs_list.Close
    
    If cmb_agent = "ALL" Then
'        Rs_list.Open " SELECT coalesce(r.custid,'') as custid_bayar, agent, s.custid as custid_cpa, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
'                     " FROM ( " & _
'                     " (SELECT Distinct custid FROM tbllunas " & _
'                     " WHERE payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
'                     " AND date_part('year',paydate) = '" & tahun_sekarang & "' order by custid) AS r " & _
'                     " RIGHT JOIN " & _
'                     " (SELECT  agent, a.custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
'                     " FROM ( " & _
'                     " (SELECT agent, custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance " & _
'                     " FROM tblsendptp_log_approve " & _
'                     " WHERE date_part('month',tgldata) = '" & bulan_sekarang & "' " & _
'                     " AND date_part('year',tgldata) = '" & tahun_sekarang & "' ) As a " & _
'                     " LEFT JOIN (SELECT custid,acc_type FROM mgm) As b ON a.custid = b.custid )) As s" & _
'                     " ON r.custid = s.custid) " & _
'                     " ORDER BY custid_bayar DESC "
        Rs_list.Open " SELECT coalesce(r.custid,'') as custid_bayar, agent, s.custid as custid_cpa, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
                     " FROM ( " & _
                     " (SELECT Distinct custid FROM mandiri.tbllunas " & _
                     " WHERE payment > 100 AND date(paydate) BETWEEN '" & Format(DTPicker1(0).Value, "yyyy-mm-dd") & "' " & _
                     " AND '" & Format(DTPicker1(2).Value, "yyyy-mm-dd") & "' order by custid) AS r " & _
                     " RIGHT JOIN " & _
                     " (SELECT  agent, a.custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
                     " FROM ( " & _
                     " (SELECT agent, custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance " & _
                     " FROM mandiri.tblsendptp_log_approve " & _
                     " WHERE date(tgldata) BETWEEN '" & Format(DTPicker1(0).Value, "yyyy-mm-dd") & "' " & _
                     " AND '" & Format(DTPicker1(2).Value, "yyyy-mm-dd") & "' ) As a " & _
                     " LEFT JOIN (SELECT custid,acc_type FROM mandiri.mgm) As b ON a.custid = b.custid )) As s" & _
                     " ON r.custid = s.custid) " & _
                     " ORDER BY custid_bayar DESC "
    Else
        If f_team = False Then
'            Rs_list.Open "SELECT agent, a.custid as custid_cpa, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type FROM (" & _
'                         "(SELECT agent, custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance FROM tblsendptp_log_approve " & _
'                         "WHERE date_part('month',tgldata) = '" & bulan_sekarang & "' " & _
'                         "AND date_part('year',tgldata) = '" & tahun_sekarang & "' AND agent = '" & agent & "' ) As a " & _
'                         "LEFT JOIN " & _
'                         "(SELECT custid, acc_type FROM mgm) As b " & _
'                         "ON a.custid = b.custid )"
            Rs_list.Open " SELECT coalesce(r.custid,'') as custid_bayar, agent, s.custid as custid_cpa, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
                         " FROM ( " & _
                         " (SELECT Distinct custid FROM mandiri.tbllunas " & _
                         " WHERE payment > 100 AND date(paydate) BETWEEN '" & Format(DTPicker1(0).Value, "yyyy-mm-dd") & "' " & _
                         " AND '" & Format(DTPicker1(2).Value, "yyyy-mm-dd") & "' AND agent = '" & AGENT & "' order by custid) AS r " & _
                         " RIGHT JOIN " & _
                         " (SELECT  agent, a.custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
                         " FROM ( " & _
                         " (SELECT agent, custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance " & _
                         " FROM mandiri.tblsendptp_log_approve " & _
                         " WHERE date(paydate) BETWEEN '" & Format(DTPicker1(0).Value, "yyyy-mm-dd") & "' " & _
                         " AND '" & Format(DTPicker1(2).Value, "yyyy-mm-dd") & "' AND agent = '" & AGENT & "' ) As a " & _
                         " LEFT JOIN (SELECT custid,acc_type FROM mandiri.mgm) As b ON a.custid = b.custid )) As s" & _
                         " ON r.custid = s.custid) " & _
                         " ORDER BY custid_bayar DESC "
        Else
'            Rs_list.Open "SELECT agent, a.custid as custid_cpa, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type FROM (" & _
'                         "(SELECT agent, custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance FROM tblsendptp_log_approve " & _
'                         "WHERE date_part('month',tgldata) = '" & bulan_sekarang & "' " & _
'                         "AND date_part('year',tgldata) = '" & tahun_sekarang & "' AND agent in (select userid from usertbl where team = '" & cmb_team.Text & "' AND userid ilike  'D%') ) As a " & _
'                         "LEFT JOIN " & _
'                         "(SELECT custid, acc_type FROM mgm) As b " & _
'                         "ON a.custid = b.custid ) ORDER BY agent"
            Rs_list.Open " SELECT coalesce(r.custid,'') as custid_bayar, agent, s.custid as custid_cpa, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
                         " FROM ( " & _
                         " (SELECT Distinct custid FROM mandiri.tbllunas " & _
                         " WHERE payment > 100 AND date(paydate) BETWEEN '" & Format(DTPicker1(0).Value, "yyyy-mm-dd") & "' " & _
                         " AND '" & Format(DTPicker1(2).Value, "yyyy-mm-dd") & "' AND agent in (select userid from mandiri.usertbl where team = '" & cmb_team.Text & "' AND userid ilike  'D%') order by custid) AS r " & _
                         " RIGHT JOIN " & _
                         " (SELECT  agent, a.custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
                         " FROM ( " & _
                         " (SELECT agent, custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance " & _
                         " FROM mandiri.tblsendptp_log_approve " & _
                         " WHERE date(paydate) BETWEEN '" & Format(DTPicker1(0).Value, "yyyy-mm-dd") & "' " & _
                         " AND '" & Format(DTPicker1(2).Value, "yyyy-mm-dd") & "'  AND agent in (select userid from mandiri.usertbl where team = '" & cmb_team.Text & "' AND userid ilike  'D%') ) As a " & _
                         " LEFT JOIN (SELECT custid,acc_type FROM mandiri.mgm) As b ON a.custid = b.custid )) As s" & _
                         " ON r.custid = s.custid) " & _
                         " ORDER BY custid_bayar DESC "
        End If
    End If
        
'    Rs_list.Open "SELECT tblnegoptp.id, agent, tblnegoptp.custid, promisedate, promisepay FROM mgm, tblnegoptp  " & _
'                 "WHERE mgm.custid = tblnegoptp.custid " & _
'                 "AND agent = '" & agent & "' AND promisepay > 100 AND date_part('month',promisedate) = '" & bulan_sekarang & "' " & _
'                 "AND date_part('year',promisedate) = '" & tahun_sekarang & "' order by promisedate desc "
    
    If Rs_list.RecordCount > 0 Then
        nomor = 0
          Do Until Rs_list.EOF
              nomor = nomor + 1
              Set ListItem = LvCPA.ListItems.ADD(, , nomor)
                              ListItem.SubItems(1) = IIf(IsNull(Rs_list!AGENT), "", Rs_list!AGENT)
                              ListItem.SubItems(2) = IIf(IsNull(Rs_list!custid_cpa), "", Rs_list!custid_cpa)
                              ListItem.SubItems(3) = IIf(IsNull(Rs_list!vcustname), "", Rs_list!vcustname)
                              ListItem.SubItems(4) = IIf(IsNull(Rs_list!jenis_ptp), "", Rs_list!jenis_ptp)
                              ListItem.SubItems(5) = Format(IIf(IsNull(Rs_list!tgl_tagih), "", Rs_list!tgl_tagih), "DD-MM-YYYY")
                              ListItem.SubItems(6) = Format(IIf(IsNull(Rs_list!total_amount_deal), "", Rs_list!total_amount_deal), "##,###")
                              ListItem.SubItems(7) = Format(IIf(IsNull(Rs_list!Pembayaran_awal), "", Rs_list!Pembayaran_awal), "##,###")
                              ListItem.SubItems(8) = cnull(IIf(IsNull(Rs_list!Tenor), "", Rs_list!Tenor))
                              ListItem.SubItems(9) = Format(IIf(IsNull(Rs_list!balance), "", Rs_list!balance), "##,###")
                              ListItem.SubItems(10) = cnull(IIf(IsNull(Rs_list!acc_type), "", Rs_list!acc_type))
                              
                              Total = Total + IIf(IsNull(Rs_list!total_amount_deal), "", Rs_list!total_amount_deal)
                              
                              id_payment = IIf(IsNull(Rs_list!custid_bayar), "", Rs_list!custid_bayar)

                              If id_payment <> "" Then
                                    For K = 1 To 10
                                          'LvPTP.ListItems(Rs_list.Bookmark).ListSubItems(K).ForeColor = vbBlue
                                          ListItem.ListSubItems(K).ForeColor = vbBlue
                                          ListItem.ForeColor = vbBlue
                                    Next K
                              End If
              Rs_list.MoveNext
          Loop
          lbldata.Caption = "Jumlah Data  : " & Rs_list.RecordCount & " Rows"
          lbltotal.Caption = "Amount Deal : IDR " & Format(Total, "##,###") & " "
          'txt_total_ptp.Text = Total
      Else
          MsgBox "Data Tidak Tersedia !", vbOKOnly + vbInformation, "Info"
          
          LvLastPayment.ListItems.CLEAR
          lbldata.Caption = "Rows : 0"
          lbltotal.Caption = "Amount Deal : IDR 0 "
          'txt_total_ptp.Text = 0
      End If
End Sub

Private Sub IsiListReject()
    Dim ListItem As ListItem
    Dim AGENT As String
    Dim total_ptp As Double
    Dim bulan_sekarang As String
    Dim tahun_sekarang As String
    Dim id_payment As String
    Dim TotalKeseluruhan As Double
    
'    tanggal_sekarang = Format(DTPicker1.Value, "yyyy-mm-dd")
'
'    bulan_sekarang = Format(tanggal_sekarang, "MM")
'    tahun_sekarang = Format(tanggal_sekarang, "YYYY")
    
    
    AGENT = cmb_agent.Text
    
    If cmb_agent = " " Then
        AGENT = ""
    End If
    
    LvCPA.ListItems.CLEAR
    If Rs_list.state = 1 Then Rs_list.Close
    
    If cmb_agent = "ALL" Then
        Rs_list.Open " SELECT coalesce(r.custid,'') as custid_bayar, agent, s.custid as custid_cpa, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
                     " FROM ( " & _
                     " (SELECT Distinct custid FROM mandiri.tbllunas " & _
                     " WHERE payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                     " AND date_part('year',paydate) = '" & tahun_sekarang & "' order by custid) AS r " & _
                     " RIGHT JOIN " & _
                     " (SELECT  agent, a.custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
                     " FROM ( " & _
                     " (SELECT agent, custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance " & _
                     " FROM mandiri.tblsendptp_log_reject " & _
                     " WHERE date_part('month',tgldata) = '" & bulan_sekarang & "' " & _
                     " AND date_part('year',tgldata) = '" & tahun_sekarang & "' ) As a " & _
                     " LEFT JOIN (SELECT custid,acc_type FROM mandiri.mgm) As b ON a.custid = b.custid )) As s" & _
                     " ON r.custid = s.custid) " & _
                     " ORDER BY custid_bayar DESC "
    Else
        If f_team = False Then
            Rs_list.Open " SELECT coalesce(r.custid,'') as custid_bayar, agent, s.custid as custid_cpa, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
                         " FROM ( " & _
                         " (SELECT Distinct custid FROM mandiri.tbllunas " & _
                         " WHERE payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                         " AND date_part('year',paydate) = '" & tahun_sekarang & "' AND agent = '" & AGENT & "' order by custid) AS r " & _
                         " RIGHT JOIN " & _
                         " (SELECT  agent, a.custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
                         " FROM ( " & _
                         " (SELECT agent, custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance " & _
                         " FROM mandiri.tblsendptp_log_reject " & _
                         " WHERE date_part('month',tgldata) = '" & bulan_sekarang & "' " & _
                         " AND date_part('year',tgldata) = '" & tahun_sekarang & "' AND agent = '" & AGENT & "' ) As a " & _
                         " LEFT JOIN (SELECT custid,acc_type FROM mandiri.mgm) As b ON a.custid = b.custid )) As s" & _
                         " ON r.custid = s.custid) " & _
                         " ORDER BY custid_bayar DESC "
        Else
            Rs_list.Open " SELECT coalesce(r.custid,'') as custid_bayar, agent, s.custid as custid_cpa, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
                         " FROM ( " & _
                         " (SELECT Distinct custid FROM mandiri.tbllunas " & _
                         " WHERE payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                         " AND date_part('year',paydate) = '" & tahun_sekarang & "' AND agent in (select userid from mandiri.usertbl where team = '" & cmb_team.Text & "' AND userid ilike  'D%') order by custid) AS r " & _
                         " RIGHT JOIN " & _
                         " (SELECT  agent, a.custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance, acc_type " & _
                         " FROM ( " & _
                         " (SELECT agent, custid, vcustname, jenis_ptp, tgl_tagih, total_amount_deal, pembayaran_awal, tenor, balance " & _
                         " FROM mandiri.tblsendptp_log_reject " & _
                         " WHERE date_part('month',tgldata) = '" & bulan_sekarang & "' " & _
                         " AND date_part('year',tgldata) = '" & tahun_sekarang & "' AND agent in (select userid from mandiri.usertbl where team = '" & cmb_team.Text & "' AND userid ilike  'D%') ) As a " & _
                         " LEFT JOIN (SELECT custid,acc_type FROM mandiri.mgm) As b ON a.custid = b.custid )) As s" & _
                         " ON r.custid = s.custid) " & _
                         " ORDER BY custid_bayar DESC "
        End If
    End If
            
    If Rs_list.RecordCount > 0 Then
        nomor = 0
          Do Until Rs_list.EOF
              nomor = nomor + 1
              Set ListItem = LvCPA.ListItems.ADD(, , nomor)
                              ListItem.SubItems(1) = IIf(IsNull(Rs_list!AGENT), "", Rs_list!AGENT)
                              ListItem.SubItems(2) = IIf(IsNull(Rs_list!custid_cpa), "", Rs_list!custid_cpa)
                              ListItem.SubItems(3) = IIf(IsNull(Rs_list!vcustname), "", Rs_list!vcustname)
                              ListItem.SubItems(4) = IIf(IsNull(Rs_list!jenis_ptp), "", Rs_list!jenis_ptp)
                              ListItem.SubItems(5) = Format(IIf(IsNull(Rs_list!tgl_tagih), "", Rs_list!tgl_tagih), "DD-MM-YYYY")
                              ListItem.SubItems(6) = Format(IIf(IsNull(Rs_list!total_amount_deal), "", Rs_list!total_amount_deal), "##,###")
                              ListItem.SubItems(7) = Format(IIf(IsNull(Rs_list!Pembayaran_awal), "", Rs_list!Pembayaran_awal), "##,###")
                              ListItem.SubItems(8) = cnull(IIf(IsNull(Rs_list!Tenor), "", Rs_list!Tenor))
                              ListItem.SubItems(9) = Format(IIf(IsNull(Rs_list!balance), "", Rs_list!balance), "##,###")
                              ListItem.SubItems(10) = cnull(IIf(IsNull(Rs_list!acc_type), "", Rs_list!acc_type))
                              
                              Total = Total + IIf(IsNull(Rs_list!total_amount_deal), "", Rs_list!total_amount_deal)
                              
                              id_payment = IIf(IsNull(Rs_list!custid_bayar), "", Rs_list!custid_bayar)

                              If id_payment <> "" Then
                                    For K = 1 To 10
                                          'LvPTP.ListItems(Rs_list.Bookmark).ListSubItems(K).ForeColor = vbBlue
                                          ListItem.ListSubItems(K).ForeColor = vbBlue
                                          ListItem.ForeColor = vbBlue
                                    Next K
                              End If
              Rs_list.MoveNext
          Loop
          lbldata.Caption = "Jumlah Data  : " & Rs_list.RecordCount & " Rows"
          lbltotal.Caption = "Amount Deal : IDR " & Format(Total, "##,###") & " "
          'txt_total_ptp.Text = Total
      Else
          MsgBox "Data Tidak Tersedia !", vbOKOnly + vbInformation, "Info"
          
          LvLastPayment.ListItems.CLEAR
          lbldata.Caption = "Rows : 0"
          lbltotal.Caption = "Amount Deal : IDR 0 "
          'txt_total_ptp.Text = 0
      End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Call ExportCoint
End Sub

Private Sub ExportCoint()
    Dim sQuery As String
    Dim RsCoint As ADODB.Recordset
    Dim W, K, S As Integer
    Dim CustId As String
    Dim ExlObj As Excel.Application
    
    CustId = ""
    
    If LvCPA.ListItems.Count = 0 Then
        MsgBox "Maaf, data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    S = 0
    For K = 1 To LvCPA.ListItems.Count
        If LvCPA.ListItems(K).Checked = True Then
            S = S + 1
            Exit For
        End If
    Next K
    
    If S = 0 Then
        MsgBox "Maaf anda belum memilih data yang akan di lihat!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
        

    'Ambil nilai custid
    For W = 1 To LvCPA.ListItems.Count
        If LvCPA.ListItems(W).Checked = True Then
            If CustId = "" Then
                CustId = "'" & CStr(LvCPA.ListItems(W).ListSubItems(2)) & "'"
            Else
                CustId = CustId & ",'" & CStr(LvCPA.ListItems(W).ListSubItems(2)) & "'"
            End If
        End If
    Next W
    
    sQuery = "SELECT region, now() as proposal_date, Ref_Number, ndiscountamt, produk,  customer_id,"
    sQuery = sQuery + " name, opendate, 'RIT' as area, nbalance, jml_payment, nttlpayment,"
    sQuery = sQuery + " nprincipal, vosprincipal, vosbalance , 'FINANCIAL PROBLEM' AS justification,"
    sQuery = sQuery + " dpropsal,aoc, kode_deskcol, nama_deskcol, nperiod FROM ("
    sQuery = sQuery + " SELECT region, now() as proposal_date, Ref_Number, ndiscountamt, produk,  customer_id, name, opendate, "
    sQuery = sQuery + " 'RIT' as area, nbalance, jml_payment, nttlpayment,  nprincipal, vosprincipal, vosbalance ,"
    sQuery = sQuery + " 'FINANCIAL PROBLEM' AS justification, dpropsal, descol, nperiod FROM ("
    sQuery = sQuery + " SELECT region, now() as proposal_date, Ref_Number, ndiscountamt, produk, "
    sQuery = sQuery + " customer_id, name, opendate, 'RIT' as area, nbalance, jml_payment, nttlpayment, "
    sQuery = sQuery + " nprincipal, vosprincipal, vosbalance , 'FINANCIAL PROBLEM' AS justification , dpropsal, descol, nperiod"
    sQuery = sQuery + " FROM ("
    sQuery = sQuery + " SELECT region, now() as proposal_date, mgm.acc_type as produk, "
    sQuery = sQuery + " mgm.custid as customer_id, name, opendate, 'RIT' as area, tblcpa.nbalance, "
    sQuery = sQuery + " tblcpa.nttlpayment, tblcpa.nprincipal, tblcpa.ndiscountamt,"
    sQuery = sQuery + " CASE WHEN  tblcpa.ndiscountamt = 0 THEN 'X' "
    sQuery = sQuery + " WHEN  tblcpa.ndiscountamt > 0 THEN 'D'"
    sQuery = sQuery + " END AS Ref_Number, vosprincipal, vosbalance, dpropsal, mgm.agent as descol, nperiod  "
    sQuery = sQuery + " FROM mandiri.mgm inner join mandiri.tblcpa on tblcpa.vcustid = mgm.custid ) AS a1"
    sQuery = sQuery + " LEFT JOIN ( "
    sQuery = sQuery + " SELECT * FROM mandiri.temp_proses_lunas ) AS a2"
    sQuery = sQuery + " ON a1.customer_id=a2.custid WHERE customer_id in(" & CustId & ") ) as query1"
    sQuery = sQuery + " INNER JOIN ( "
    sQuery = sQuery + " SELECT vcustid, max(dpropsal)AS maxdrpopsal FROM mandiri.tblcpa WHERE vcustid in(" & CustId & ")"
    sQuery = sQuery + " group by vcustid ) as query2 ON query1.dpropsal = query2.maxdrpopsal) as queryjoin1"
    sQuery = sQuery + " LEFT JOIN"
    sQuery = sQuery + " (SELECT kode_deskcol, nama_deskcol, aoc FROM mandiri.tbl_data_karyawan) as queryjoin2"
    sQuery = sQuery + " ON queryjoin1.descol = queryjoin2.kode_deskcol"
    Set RsCoint = New ADODB.Recordset
    RsCoint.CursorLocation = adUseClient
    RsCoint.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Set ExlObj = CreateObject("excel.application")
    ExlObj.Workbooks.ADD
    ExlObj.Visible = True
    
    With ExlObj.ActiveSheet
        .Cells(1, 1).Value = "NO"
        .Cells(1, 2).Value = "AOC"
        .Cells(1, 3).Value = "NAMA DESK COLLECTION"
        .Cells(1, 4).Value = "REGION"
        .Cells(1, 5).Value = "PROPOSAL DATE"
        .Cells(1, 6).Value = "REF. No."
        .Cells(1, 7).Value = "PRODUCT"
        .Cells(1, 8).Value = "CARD / Acc. No."
        .Cells(1, 9).Value = "CUSTOMER NAME"
        .Cells(1, 10).Value = "CARD OPEN DATE"
        .Cells(1, 11).Value = "AREA"
        .Cells(1, 12).Value = "OUTSTANDING BALANCE"
        .Cells(1, 13).Value = "TOTAL PAYMENT"
        .Cells(1, 14).Value = "PRINCIPAL"
        .Cells(1, 15).Value = "DISCOUNT AMOUNT"
        .Cells(1, 16).Value = "TENOR"
        .Cells(1, 17).Value = "+/- FROM O/S %"
        .Cells(1, 18).Value = "+/- FROM PRINCIPAL %"
        .Cells(1, 19).Value = "JUSTIFICATION"
        
        iRow = 2
        If RsCoint.RecordCount > 0 Then
            i = 0
            Do Until RsCoint.EOF
                i = i + 1
                iRow = iRow + 1
                
                .Cells(iRow, 1).Value = i
                .Cells(iRow, 2).Value = IIf(IsNull(RsCoint!aoc), "", RsCoint!aoc)
                .Cells(iRow, 3).Value = IIf(IsNull(RsCoint!nama_deskcol), "", RsCoint!nama_deskcol)
                .Cells(iRow, 4).Value = IIf(IsNull(RsCoint!region), "", RsCoint!region)
                .Cells(iRow, 5).Value = Format(IIf(IsNull(RsCoint!proposal_date), "", RsCoint!proposal_date), "DD-MMM-YYYY")
                .Cells(iRow, 6).Value = IIf(IsNull(RsCoint!Ref_Number), "", RsCoint!Ref_Number)
                .Cells(iRow, 7).Value = IIf(IsNull(RsCoint!produk), "", RsCoint!produk)
                .Cells(iRow, 8).Value = IIf(IsNull(RsCoint!customer_id), "", RsCoint!customer_id)
                .Cells(iRow, 9).Value = IIf(IsNull(RsCoint!Name), "", RsCoint!Name)
                .Cells(iRow, 10).Value = Format(IIf(IsNull(RsCoint!opendate), "", RsCoint!opendate), "DD-MMM-YYYY")
                .Cells(iRow, 11).Value = IIf(IsNull(RsCoint!area), "", RsCoint!area)
                .Cells(iRow, 12).Value = IIf(IsNull(RsCoint!nbalance), "", RsCoint!nbalance)
                .Cells(iRow, 13).Value = IIf(IsNull(RsCoint!jml_payment), "", RsCoint!jml_payment)
                .Cells(iRow, 14).Value = IIf(IsNull(RsCoint!nprincipal), "", RsCoint!nprincipal)
                .Cells(iRow, 15).Value = IIf(IsNull(RsCoint!ndiscountamt), "", RsCoint!ndiscountamt)
                .Cells(iRow, 16).Value = IIf(IsNull(RsCoint!nperiod), "1", RsCoint!nperiod)
                .Cells(iRow, 17).Value = IIf(IsNull(RsCoint!vosbalance), "", RsCoint!vosbalance)
                .Cells(iRow, 18).Value = IIf(IsNull(RsCoint!vosprincipal), "", RsCoint!vosprincipal)
                .Cells(iRow, 19).Value = IIf(IsNull(RsCoint!justification), "", RsCoint!justification)
       
                RsCoint.MoveNext
            Loop
        End If
    
        'OTOMATISASI CELL
        For iColom = 1 To 16
            ExlObj.Cells(2, iColom).EntireColumn.AutoFit
        Next
        
        MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"

        Set ExlObj = Nothing
        Set RsCoint = Nothing
    End With
End Sub

Private Sub Form_Load()
    Call koneksi
    Call IsiAgent
    
    f_team = False
    
    cmb_status.ListIndex = 0
    
    DTPicker1(0).Value = Now
    DTPicker1(2).Value = Now

    
    
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        lbl_team.Visible = False
        cmb_team.Visible = False
        chk_team.Visible = False
        'cmd_download.Visible = False
        'cmd_download_payment.Visible = False
    Else
        If Rs_list.state = 1 Then Rs_list.Close

        If Left(MDIForm1.TxtUsername.Text, 2) = "TL" Then
            Rs_list.Open "select userid from mandiri.usertbl where usertype = '1' AND userid ilike 'D%' AND  team = '" & MDIForm1.TxtUsername.Text & "' Order by userid"
        Else
            Rs_list.Open "SELECT DISTINCT team from mandiri.usertbl WHERE team ilike 'TL%'"
        End If

    End If
    Call HeaderLvCPA
    Call HeaderLastPayment
End Sub

Private Sub HeaderLastPayment()
    LvLastPayment.ColumnHeaders.ADD , , "Custid", 2100
    LvLastPayment.ColumnHeaders.ADD , , "Agent", 1000
    LvLastPayment.ColumnHeaders.ADD , , "LPA", 1300
    LvLastPayment.ColumnHeaders.ADD , , "LPD", 1300
    LvLastPayment.ColumnHeaders.ADD , , "CH Name", 2350
    LvLastPayment.ColumnHeaders.ADD , , "Product", 1500
End Sub


Private Sub IsiAgent()
    If Rs_list.state = 1 Then Rs_list.Close
    
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        Rs_list.Open "select userid from mandiri.usertbl where usertype = '1' and userid = '" & MDIForm1.TxtUsername.Text & "' Order by userid"
    ElseIf UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        Rs_list.Open "select userid from mandiri.usertbl where usertype = '1' AND userid ilike 'D%' AND  team = '" & MDIForm1.TxtUsername.Text & "' Order by userid"
    Else
        Rs_list.Open "select userid from mandiri.usertbl where usertype = '1' and userid ilike 'D%' Order by userid"
    End If
    
    cmb_agent.AddItem ""
    
    If UCase(MDIForm1.txtlevel.Text) <> "AGENT" And UCase(MDIForm1.txtlevel.Text) <> "TEAMLEADER" Then
        cmb_agent.AddItem "ALL"
    End If
    
    While Not Rs_list.EOF
        cmb_agent.AddItem Rs_list("USERID")
        Rs_list.MoveNext
    Wend
    cmb_agent.ListIndex = 1
End Sub

Private Sub HeaderLvCPA()
 LvCPA.ColumnHeaders.CLEAR
    With LvCPA.ColumnHeaders
        .ADD 1, , "No", 550
        .ADD 2, , "Agent", 1500
        .ADD 3, , "Custid", 2700
        .ADD 4, , "CH Name", 3000
        .ADD 5, , "Jenis PTP", 1500
        .ADD 6, , "Tanggal Tagih", 2000
        .ADD 7, , "Amount Deal", 3000
        .ADD 8, , "Down Payment", 3000
        .ADD 9, , "Tenor", 700
        .ADD 10, , "Total Amount", 2000
        .ADD 11, , "Product", 1500
    End With
End Sub
   
'Private Sub LvCPA_Click()
'    If LvCPA.ListItems.Count <= 0 Then
'        MsgBox "Tampilkan Data Terlebih Dahulu !", vbOKOnly + vbInformation, "Perhatian"
'    Exit Sub
'    End If
'
'    Call Isi_payment
'End Sub

Private Sub Isi_payment()
    Dim ListItem As ListItem
    Dim AGENT As String
    Dim total_ptp As Double
    Dim bulan_sekarang As String
    Dim tahun_sekarang As String
        
'    tanggal_sekarang = Format(DTPicker1.Value, "yyyy-mm-dd")
'
'    bulan_sekarang = Format(tanggal_sekarang, "MM")
'    tahun_sekarang = Format(tanggal_sekarang, "YYYY")
    
'    bulan_sekarang = "06"
'    tahun_sekarang = "2015"
    
    LvLastPayment.ListItems.CLEAR
    If Rs_list.state = 1 Then Rs_list.Close

    Rs_list.Open "SELECT * FROM ( " & _
                 "(SELECT custid, paydate, payment, agent, id FROM mandiri.tbllunas  " & _
                 "WHERE custid = '" & LvCPA.SelectedItem.SubItems(2) & "' " & _
                 "AND Payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                 "AND date_part('year',paydate) = '" & tahun_sekarang & "' " & _
                 "AND paydate = (SELECT MAX(paydate) FROM mandiri.tbllunas WHERE custid = '" & LvCPA.SelectedItem.SubItems(2) & "' " & _
                 "AND Payment > 100 AND date_part('month',paydate) = '" & bulan_sekarang & "' " & _
                 "AND date_part('year',paydate) = '" & tahun_sekarang & "')) AS a " & _
                 "LEFT JOIN (SELECT custid, acc_type, name FROM mandiri.mgm where custid =  '" & LvCPA.SelectedItem.SubItems(2) & "') AS b on a.custid = b.custid)"

    If Rs_list.RecordCount > 0 Then
          Do Until Rs_list.EOF
              Set ListItem = LvLastPayment.ListItems.ADD(, , IIf(IsNull(Rs_list!CustId), "", CStr(Rs_list!CustId)))
                              ListItem.SubItems(1) = IIf(IsNull(Rs_list!AGENT), "", Rs_list!AGENT)
                              ListItem.SubItems(2) = Format(IIf(IsNull(Rs_list!Payment), "", Rs_list!Payment), "##,###")
                              ListItem.SubItems(3) = cnull(IIf(IsNull(Rs_list!paydate), "", Rs_list!paydate))
                              ListItem.SubItems(4) = cnull(IIf(IsNull(Rs_list!Name), "", Rs_list!Name))
                              ListItem.SubItems(5) = cnull(IIf(IsNull(Rs_list!acc_type), "", Rs_list!acc_type))
              Rs_list.MoveNext
          Loop
    End If
End Sub

Private Sub LvCPA_DblClick()
    If LvCPA.ListItems.Count > 0 Then
        VIEW_MGMDATA.txtnocard.Text = LvCPA.SelectedItem.SubItems(2)
        Form_List_CPA.Hide
        VIEW_MGMDATA.Show
    Else
        MsgBox "Data tidak ada!!", vbOKOnly + vbInformation, "INFO"
    End If
End Sub
