VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_deskcoll_performance_reguler 
   BackColor       =   &H00F0FFF0&
   Caption         =   "DeskColl Performance Insentif Regular"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13470
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   13470
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid msfx 
      Height          =   6015
      Left            =   240
      TabIndex        =   24
      Top             =   1440
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   10610
      _Version        =   393216
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F0FFF0&
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   7440
      Width           =   12975
      Begin MSComDlg.CommonDialog CD 
         Left            =   2880
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   10980
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "0"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   10980
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "0"
         Top             =   600
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0"
         Top             =   600
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0"
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   600
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0"
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   600
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Keluar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1155
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "Export to Excel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1150
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "::: REGULAR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   360
         TabIndex        =   35
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Kept Amount"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   9720
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   12960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Data"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   9720
         TabIndex        =   29
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Kept + Broken"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   7080
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Kept"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   7080
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "RPC 2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "RPC 1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5040
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PTP"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dialer Hours"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Paid Hours"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0FFF0&
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5520
         TabIndex        =   32
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4440
         TabIndex        =   30
         Top             =   720
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   8040
         TabIndex        =   27
         Top             =   720
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.ComboBox cb_sort 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4440
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Proses"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker tgl_laporan 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM-yyyy"
         Format          =   96141315
         CurrentDate     =   41610
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   31
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sort Column By"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan dan Tahun"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   6240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
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
Attribute VB_Name = "Form_deskcoll_performance_reguler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs_calc As ADODB.Recordset
Private rs_temp As ADODB.Recordset

Private tgl_lap As Date
Private sql_str As String

Private sPaid_hours As Long
Private sDialer_hours As Long
Private sPTP As Long
Private sRPC1 As Long
Private sRPC2 As Long
Private sKept As Long
Private sKept_old As Long
Private sKeptBroken As Long
Private sKeptBrokenOld As Long
Private sKeptAmount As Double
Private sPTP_Reg As Double
Private sPymnGlobal As Double

Private tPaid_hours As Long
Private tDialer_hours As Long
Private tPTP As Long
Private tRPC1 As Long
Private tRPC2 As Long
Private tKept As Long
Private tKept_old As Long
Private tKeptBroken As Long
Private tKeptBrokenOld As Long
Private tKeptAmount As Double
Private tPTP_Reg As Double
Private tPymnGlobal As Double

Private sqlfilter As String
Private m_SortColumn As Integer
Private m_SortOrder As Integer

Private Sub cb_sort_Click()
    SortByColumn cb_sort.ListIndex
End Sub

Private Sub Combo1_Click()
    ' ---- OPSI AGENT ----
    If Combo1.Text <> "" And Combo1.Text <> "ALL" Then
        If rs_temp.state = 1 Then rs_temp.Close
        rs_temp.Open "SELECT userid,agent FROM usertbl WHERE userid like 'D%' AND team='" & Combo1.Text & "' ORDER BY userid"
        Combo2.CLEAR
        Combo3.CLEAR
        Do Until rs_temp.EOF
            Combo2.AddItem IIf(IsNull(rs_temp!Userid), "", rs_temp!Userid)
            Combo3.AddItem IIf(IsNull(rs_temp!agent), "", rs_temp!agent)
            rs_temp.MoveNext
        Loop
        ' -------------------
    Else
        If rs_temp.state = 1 Then rs_temp.Close
        rs_temp.Open "SELECT userid,agent FROM usertbl WHERE userid like 'D%' ORDER BY userid"
        Combo2.CLEAR
        Combo3.CLEAR
        Do Until rs_temp.EOF
            Combo2.AddItem IIf(IsNull(rs_temp!Userid), "", rs_temp!Userid)
            Combo3.AddItem IIf(IsNull(rs_temp!agent), "", rs_temp!agent)
            rs_temp.MoveNext
        Loop
        ' -------------------
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo2_Click()
    Combo3.ListIndex = Combo2.ListIndex
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo3_Click()
    Combo2.ListIndex = Combo3.ListIndex
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
    Dim lstItem         As listItem
    Dim xx              As Integer
    Dim warna_belang    As Integer
    
    Dim sql_tahun       As String
    Dim sql_bulan       As String
    
    Dim temp_payment    As Double
    Dim temp_paydate    As String
    
    Dim f_kept          As String
    
    Dim dial_paidh As Double
    Dim rpc_dialh As Double
    Dim rpc_paidh As Double
    Dim keptbroken_rpc As Double
    Dim ptp_rpc As Double
    Dim kept_keptbroken As Double
    Dim avg_paymentsize As Double
    Dim cev As Double
    Dim evph As Double

    sqlfilter = ""
    Command1.Enabled = False
    
    tgl_lap = Format(tgl_laporan.Value, "yyyy-mm-dd")
    sql_str = ""

    sql_tahun = Format(tgl_lap, "yyyy")
    sql_bulan = Format(tgl_lap, "mm")
    
    M_OBJCONN.Execute "DELETE FROM tblreport_komisi_reg WHERE date_part('month',tgl_report)=" & sql_bulan & " AND date_part('year',tgl_report)=" & sql_tahun & ""

    M_OBJCONN.Execute " INSERT INTO tblreport_komisi_reg(custid,promisedate,agent,tglinput,promisepay,tgl_report) " & _
                    " SELECT custid,promisedate,agent,tglinput,promisepay,to_date('" & Format(tgl_lap, "yyyy-mm-01") & "','YYYY-MM-DD') " & _
                    " FROM tblnegoptp_log WHERE date_part('month',promisedate)=" & sql_bulan & " AND date_part('year',promisedate)=" & sql_tahun & ";"

'    M_OBJCONN.Execute "INSERT INTO tblreport_komisi_reg(custid,promisedate,agent,tglinput,promisepay,tgl_report) " & _
'                    "SELECT b.* FROM (SELECT custid,max(promisedate) as tgl_max,agent FROM tblnegoptp_log  " & _
'                    " WHERE (date_part('month',promisedate)=" & sql_bulan & " AND date_part('year',promisedate)=" & sql_tahun & ") " & _
'                    " GROUP BY custid,agent) a, (SELECT custid,promisedate,agent,tglinput,promisepay,to_date('" & Format(tgl_lap, "yyyy-mm-01") & "','YYYY-MM-DD') " & _
'                    " FROM tblnegoptp_log  WHERE date_part('month',promisedate)=" & sql_bulan & " AND date_part('year',promisedate)=" & sql_tahun & " ) b WHERE a.custid=b.custid AND a.agent=b.agent AND a.tgl_max=b.promisedate;"

    If rs_temp.state = 1 Then rs_temp.Close
    rs_temp.Open "SELECT id,custid,promisedate,agent,promisepay FROM tblreport_komisi_reg"
    If rs_temp.RecordCount > 0 Then
        ProgressBar1.Max = rs_temp.RecordCount
        Do Until rs_temp.EOF
            DoEvents
            ProgressBar1.Value = rs_temp.Bookmark
         
            sql_str = "SELECT sum(a.payment) as total_bayar,max(a.paydate) as tgl_akhir FROM (SELECT * FROM tbllunas WHERE date_part('month',paydate)=" & sql_bulan & " AND date_part('year',paydate)=" & sql_tahun & " ) a " & _
                        " WHERE a.custid='" & cnull(rs_temp!CustId) & "'"
            
            
            temp_paydate = "null"
            temp_payment = 0
            
            If rs_calc.state = 1 Then rs_calc.Close
            rs_calc.Open sql_str
            If rs_calc.RecordCount > 0 Then
               temp_paydate = IIf(IsNull(rs_calc!tgl_akhir), "Null", "'" & Format(rs_calc!tgl_akhir, "yyyy-mm-dd") & "'")
               temp_payment = IIf(IsNull(rs_calc!total_bayar), 0, rs_calc!total_bayar)
            End If
            
            If temp_payment >= cnull(rs_temp!PromisePay) Then
                f_kept = "KEPT"
            Else
                f_kept = "BP"
            End If
            
            ' UPDATE untuk agent yang terakhir ---------------
            M_OBJCONN.Execute "UPDATE tblreport_komisi_reg SET payment=" & temp_payment & ",paydate=" & temp_paydate & ",f_kept='" & f_kept & "' FROM (SELECT custid,max(promisedate) as tgl_akhirPTP FROM tblreport_komisi_reg WHERE date(tgl_report)='" & Format(tgl_lap, "yyyy-mm-01") & "' AND custid='" & cnull(rs_temp!CustId) & "' GROUP BY custid) a WHERE tblreport_komisi_reg.custid=a.custid AND promisedate=a.tgl_akhirPTP "
            
            rs_temp.MoveNext
        Loop
    End If
    
    M_OBJCONN.Execute "DELETE FROM result_komisi_reg ;"
    
    sql_str = "INSERT INTO result_komisi_reg(userid,name,tl,paid_hours,dialer_hours,ptp,rpc1,rpc2,kept,keptbroken,keptamount) SELECT userid,agent as nama,team as TL, floor(jml_jam) as Paid_hours,floor(jml_dialer) as Dialer_hours,jml_ptp as PTP, sts1 as RPC1, sts2 as RPC2 ,coalesce(jml_kept,0) as Kept, coalesce(jml_kept,0) + coalesce(broken,0) as ""Kept+Broken"", coalesce(keptamount,0) as kept_amount FROM " & _
                "(SELECT z.*,aa.sts2 FROM (SELECT x.*,y.sts1 FROM (SELECT v.*,w.broken FROM (SELECT t.*,u.jml_kept,u.keptamount FROM (SELECT r.*,s.jml_ptp FROM (SELECT o.*,p.jml_dialer FROM (SELECT a.userid,a.agent,a.team,b.jml_jam FROM (SELECT * FROM usertbl WHERE usertype='1' AND userid like 'D%') a LEFT JOIN "
    ' PAID HOURS
    sql_str = sql_str + " (SELECT y.userid,sum(x.hours) as jml_jam FROM tblabsen x, usertbl y WHERE x.nopeg=y.nik_absensi AND date_part('month',tanggal)=" & sql_bulan & " AND date_part('year',tanggal)=" & sql_tahun & " GROUP BY userid,nopeg) b ON a.userid=b.userid) o LEFT JOIN  "
    ' DIALER HOURS
    sql_str = sql_str + " (SELECT userid,sum(hours) as jml_dialer FROM tblabsen_aplikasi WHERE date_part('month',tanggal)=" & sql_bulan & " AND date_part('year',tanggal)=" & sql_tahun & " GROUP BY userid ) p ON o.userid=p.userid) r LEFT JOIN "
    ' JUMLAH PTP
    sql_str = sql_str + " (Select a.agent,count(a.custid) as jml_PTP from (SELECT custid,agent,promisedate FROM reportPTP WHERE date_part('month',promisedate)=" & sql_bulan & " AND date_part('year',promisedate)=" & sql_tahun & ") x,(SELECT custid,agent,max(promisedate) as Tgl_akhir FROM reportPTP WHERE date_part('month',promisedate)=" & sql_bulan & " AND date_part('year',promisedate)=" & sql_tahun & " GROUP BY custid,agent) a,mgm b where a.custid=b.custid AND x.custid=a.custid AND x.promisedate=a.Tgl_akhir GROUP BY a.agent) s ON r.userid=s.agent ) t LEFT JOIN "
    ' KEPT
    sql_str = sql_str + " (SELECT agent,count(custid) as jml_kept,sum(payment) as keptamount FROM tblreport_komisi_reg WHERE date_part('month',tgl_report)=" & sql_bulan & " AND date_part('year',tgl_report)=" & sql_tahun & " AND f_kept='KEPT' GROUP BY agent) u ON t.userid=u.agent ) v LEFT JOIN "
    ' BP
    sql_str = sql_str + " (SELECT agent,count(custid) as broken FROM tblreport_komisi_reg WHERE date_part('month',tgl_report)=" & sql_bulan & " AND date_part('year',tgl_report)=" & sql_tahun & " AND f_kept='BP' GROUP BY agent) w ON v.userid=w.agent ) x LEFT JOIN "
    ' RPC 1
'    sql_str = sql_str + " (SELECT agent,count(id) as sts1 FROM mgm WHERE substring(f_cek_new,1,3) in('PTP','KP-','BP-','ON-','PR-') AND agent like 'D%' AND date_part('month',tglcall)=" & sql_bulan & " AND date_part('year',tglcall)=" & sql_tahun & " GROUP by agent) y ON x.userid=y.agent) z LEFT JOIN "
'    ' RPC 2
'    sql_str = sql_str + " (SELECT agent,count(id) as sts2 FROM mgm WHERE trim(lower(statuscall)) in ('ch','spouse') AND agent like 'D%' AND date_part('month',tglcall)=" & sql_bulan & " AND date_part('year',tglcall)=" & sql_tahun & " GROUP by agent) aa ON z.userid=aa.agent) ORDER BY userid "
    sql_str = sql_str + " (SELECT agent,count(a.custid) as sts1 FROM (SELECT custid,agent,tgl,ststelpwith FROM mgm_hst WHERE date_part('month',tgl)=" & sql_bulan & " AND date_part('year',tgl)=" & sql_tahun & " AND substring(f_cek_new,1,3) in('PTP','KP-','BP-','ON-','PR-')) a,(SELECT custid, max(tgl) as tgl_akhir FROM mgm_hst WHERE date_part('month',tgl)=" & sql_bulan & " AND date_part('year',tgl)=" & sql_tahun & " AND substring(f_cek_new,1,3) in('PTP','KP-','BP-','ON-','PR-') GROUP BY custid ) b WHERE a.custid=b.custid AND a.tgl=b.tgl_akhir GROUP BY agent) y ON x.userid=y.agent) z LEFT JOIN "
    ' RPC 2
    sql_str = sql_str + " (SELECT agent,count(a.custid) as sts2 FROM (SELECT custid,agent,tgl,ststelpwith FROM mgm_hst WHERE date_part('month',tgl)=" & sql_bulan & " AND date_part('year',tgl)=" & sql_tahun & " AND ststelpwith in ('CH','SPOUSE')) a,(SELECT custid, max(tgl) as tgl_akhir FROM mgm_hst WHERE date_part('month',tgl)=" & sql_bulan & " AND date_part('year',tgl)=" & sql_tahun & " AND ststelpwith in ('CH','SPOUSE') GROUP BY custid ) b WHERE a.custid=b.custid AND a.tgl=b.tgl_akhir GROUP BY agent) aa ON z.userid=aa.agent) ORDER BY userid "


    M_OBJCONN.Execute sql_str
    
    If Combo1.Text <> "" And Combo1.Text <> "ALL" Then
        sqlfilter = " AND tl='" & Combo1.Text & "'"
    End If

    If Combo2.Text <> "" Then
        sqlfilter = sqlfilter & " AND userid='" & Combo2.Text & "'"
    End If
    
    sPaid_hours = 0
    sDialer_hours = 0
    sPTP = 0
    sRPC1 = 0
    sRPC2 = 0
    sKept = 0
    sKeptBroken = 0
    sKeptAmount = 0
    
    If rs_temp.state = 1 Then rs_temp.Close
    rs_temp.Open "SELECT * FROM result_komisi_reg WHERE userid is not null " & sqlfilter & " ORDER BY userid"
    If rs_temp.RecordCount > 0 Then
        ProgressBar1.Max = rs_temp.RecordCount
        warna_belang = 0
        msfx.Rows = 1
        msfx.Rows = rs_temp.RecordCount + 1
        
        Do Until rs_temp.EOF
            DoEvents
            ProgressBar1.Value = rs_temp.Bookmark
            tPaid_hours = Val(cnull(rs_temp!paid_hours))
            tDialer_hours = Val(cnull(rs_temp!Dialer_hours))
            tPTP = Val(cnull(rs_temp!ptp))
            tRPC1 = Val(cnull(rs_temp!rpc1))
            tRPC2 = Val(cnull(rs_temp!rpc2))
            tKept = Val(cnull(rs_temp!kept))
            tKeptBroken = Val(cnull(rs_temp("Keptbroken")))
            tKeptAmount = Val(cnull(rs_temp("Keptamount")))

            With msfx
                xx = rs_temp.Bookmark
                .TextMatrix(xx, 1) = xx
                .TextMatrix(xx, 2) = cnull(rs_temp!Userid)
                .TextMatrix(xx, 3) = cnull(rs_temp!Name)
                .TextMatrix(xx, 4) = cnull(rs_temp!TL)
                .TextMatrix(xx, 5) = tPaid_hours
                .TextMatrix(xx, 6) = tDialer_hours
                .TextMatrix(xx, 7) = tPTP
                .TextMatrix(xx, 8) = tRPC1
                .TextMatrix(xx, 9) = tRPC2
                .TextMatrix(xx, 10) = tKept
                .TextMatrix(xx, 11) = tKeptBroken
                .TextMatrix(xx, 12) = Format(tKeptAmount, "#,###,###")
                
                ' Dialer Op / Paid Hours
                If tDialer_hours > 0 And tPaid_hours > 0 Then
                    dial_paidh = Round(tDialer_hours / tPaid_hours, 2)
                    .TextMatrix(xx, 13) = dial_paidh
                End If

                ' RPC / Dialer
                If tRPC2 > 0 And tDialer_hours > 0 Then
                    rpc_dialh = Round(tRPC2 / tDialer_hours, 2)
                    .TextMatrix(xx, 14) = rpc_dialh
                End If

                ' RPC / Paid Hours
                If tRPC2 > 0 And tPaid_hours > 0 Then
                    rpc_paidh = Round(tRPC2 / tPaid_hours, 2)
                    .TextMatrix(xx, 15) = rpc_paidh
                End If

                ' KeptBroken / RPC
                If tKeptBroken > 0 And tRPC2 > 0 Then
                    keptbroken_rpc = Round(tKeptBroken / tRPC2, 2)
                    .TextMatrix(xx, 16) = keptbroken_rpc
                End If

                ' PTP / RPC
                If tPTP > 0 And tRPC2 > 0 Then
                    ptp_rpc = Round(tPTP / tRPC2, 2)
                    .TextMatrix(xx, 17) = ptp_rpc
                End If

                ' KEPT / KEPT BROKEN
                If tKept > 0 And tKeptBroken > 0 Then
                    kept_keptbroken = Round(tKept / tKeptBroken, 2)
                    .TextMatrix(xx, 18) = kept_keptbroken
                End If

                ' Average Payment Size
                If tKeptAmount > 0 And tKept > 0 Then
                    avg_paymentsize = Format(Round(tKeptAmount / tKept, 0), "#,###,###")
                    .TextMatrix(xx, 19) = avg_paymentsize
                End If

                ' CEV
                cev = Round(Val(.TextMatrix(xx, 16)) * Val(.TextMatrix(xx, 18)) * Val(Format(.TextMatrix(xx, 19), "#")), 0)
                .TextMatrix(xx, 20) = Format(cev, "#,###,###")

                ' EVPH
                evph = Round(Val(.TextMatrix(xx, 15)) * Val(Format(.TextMatrix(xx, 20), "#")), 0)
                .TextMatrix(xx, 21) = Format(evph, "#,###,###")

                sPaid_hours = sPaid_hours + Val(cnull(rs_temp!paid_hours))
                sDialer_hours = sDialer_hours + Val(cnull(rs_temp!Dialer_hours))
                sPTP = sPTP + Val(cnull(rs_temp!ptp))
                sRPC1 = sRPC1 + Val(cnull(rs_temp!rpc1))
                sRPC2 = sRPC2 + Val(cnull(rs_temp!rpc2))
                sKept = sKept + Val(cnull(rs_temp!kept))
                sKeptBroken = sKeptBroken + Val(cnull(rs_temp("KeptBroken")))
                sKeptAmount = sKeptAmount + Val(cnull(rs_temp("Keptamount")))
                
                If warna_belang = 1 Then
                    For i = 1 To msfx.Cols - 1
                        .Col = i
                        .Row = xx
                        .CellBackColor = &HF0FFF0
                    Next i
                    warna_belang = 0
                Else
                    warna_belang = 1
                End If
                
                M_OBJCONN.Execute "UPDATE result_komisi_reg SET dial_paidh=" & dial_paidh & ",rpc_dialh=" & rpc_dialh & ",rpc_paidh=" & rpc_paidh & "," & _
                                "keptbroken_rpc=" & keptbroken_rpc & ",ptp_rpc=" & ptp_rpc & ",kept_keptbroken=" & kept_keptbroken & ",avg_paymentsize=" & avg_paymentsize & ",cev=" & cev & ",epvh=" & evph & " WHERE " & _
                                "userid='" & rs_temp!Userid & "'"
                
                rs_temp.MoveNext
            End With
        Loop
    Else
        MsgBox "Data tidak ditemukan!!", vbOKOnly + vbInformation, "INFO"
    End If
    
    ' ======= TOTAL DATA ======
    Text1(0).Text = Format(sPaid_hours, "#,###,###")
    Text1(1).Text = Format(sDialer_hours, "#,###,###")
    Text1(2).Text = Format(sPTP, "#,###,###")
    Text1(3).Text = Format(sRPC1, "#,###,###")
    Text1(4).Text = Format(sRPC2, "#,###,###")
    Text1(5).Text = Format(sKept, "#,###,###")
    Text1(6).Text = Format(sKeptBroken, "#,###,###")
    Text1(7).Text = Format(sKeptAmount, "#,###,###")
    Text1(10).Text = Format(rs_temp.RecordCount, "#,###,###")
    ' =========================
    Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Form_upload_absensi.Show 1
End Sub

Private Sub Command3_Click()
    CD.Filter = "Excel Files (*.xls)|*.xls"
    CD.ShowSave
    If CD.FileName <> "" Then
        If rs_calc.state = 1 Then rs_calc.Close
        rs_calc.Open "SELECT * FROM result_komisi_reg ORDER BY agent;"
        
        If rs_calc.RecordCount > 0 Then
            ConvertToExcel rs_calc, CD.FileName
        Else
            MsgBox "Tidak ada data yang didownload!!", vbOKOnly + vbInformation, "INFO"
        End If
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Call koneksi
    
    ' ---- OPSI TL ----
    If rs_temp.state = 1 Then rs_temp.Close
    rs_temp.Open "SELECT distinct team as team_TL FROM usertbl WHERE team is not null AND lower(team) not in ('reserved','septian','wulan','admin')"
    Combo1.CLEAR
    Combo1.AddItem "ALL"
    Do Until rs_temp.EOF
        Combo1.AddItem IIf(IsNull(rs_temp!team_TL), "", rs_temp!team_TL)
        rs_temp.MoveNext
    Loop
    ' ------------------
    
    ' ---- OPSI AGENT ----
    If rs_temp.state = 1 Then rs_temp.Close
    rs_temp.Open "SELECT userid,agent FROM usertbl WHERE userid like 'D%' ORDER BY userid"
    Combo2.CLEAR
    Combo3.CLEAR
    Do Until rs_temp.EOF
        Combo2.AddItem IIf(IsNull(rs_temp!Userid), "", rs_temp!Userid)
        Combo3.AddItem IIf(IsNull(rs_temp!agent), "", rs_temp!agent)
        rs_temp.MoveNext
    Loop
    ' -------------------
    
    msfx.Cols = 22
    With msfx
        .TextMatrix(0, 1) = "No"
        .TextMatrix(0, 2) = "User ID"
        .TextMatrix(0, 3) = "Name"
        .TextMatrix(0, 4) = "TL"
        .TextMatrix(0, 5) = "Paid Hours"
        .TextMatrix(0, 6) = "Dialer Hours"
        .TextMatrix(0, 7) = "PTP"
        .TextMatrix(0, 8) = "RPC 1"
        .TextMatrix(0, 9) = "RPC 2"
        .TextMatrix(0, 10) = "Kept"
        .TextMatrix(0, 11) = "Kept+Broken"
        .TextMatrix(0, 12) = "Kept Amount"
        .TextMatrix(0, 13) = "Dialer Op/Paid Hrs"
        .TextMatrix(0, 14) = "RPC/Dialer Op Hrs"
        .TextMatrix(0, 15) = "RPC/Paid Hrs"
        .TextMatrix(0, 16) = "(Kept+Broken)/RPC"
        .TextMatrix(0, 17) = "PTP/RPC"
        .TextMatrix(0, 18) = "Kept#/(Kept+Broken)"
        .TextMatrix(0, 19) = "Average Payment Size"
        .TextMatrix(0, 20) = "CEV"
        .TextMatrix(0, 21) = "EVPH"
    End With

    For i = 0 To msfx.Cols - 1
        cb_sort.AddItem msfx.TextMatrix(0, i)
    Next i
End Sub

Private Sub koneksi()
    Set rs_calc = New ADODB.Recordset
    rs_calc.CursorLocation = adUseClient
    rs_calc.CursorType = adOpenDynamic
    rs_calc.LockType = adLockOptimistic
    rs_calc.ActiveConnection = M_OBJCONN
    
    Set rs_temp = New ADODB.Recordset
    rs_temp.CursorLocation = adUseClient
    rs_temp.CursorType = adOpenDynamic
    rs_temp.LockType = adLockOptimistic
    rs_temp.ActiveConnection = M_OBJCONN
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_temp = Nothing
    Set rs_calc = Nothing
End Sub


Private Sub msfx_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
 ' If this is not row 0, do nothing.
    If msfx.MouseRow <> 0 Then Exit Sub

    ' Sort by the clicked column.
    SortByColumn msfx.MouseCol
End Sub

Private Sub SortByColumn(ByVal sort_column As Integer)
    ' Hide the FlexGrid.
    msfx.Visible = False
    msfx.Refresh

    ' Sort using the clicked column.
    msfx.Col = sort_column
    msfx.ColSel = sort_column
    msfx.Row = 0
    msfx.RowSel = 0

    ' If this is a new sort column, sort ascending.
    ' Otherwise switch which sort order we use.
    If m_SortColumn <> sort_column Then
        m_SortOrder = flexSortGenericAscending
    ElseIf m_SortOrder = flexSortGenericAscending Then
        m_SortOrder = flexSortGenericDescending
    Else
        m_SortOrder = flexSortGenericAscending
    End If
    msfx.Sort = m_SortOrder

    ' Restore the previous sort column's name.
    If m_SortColumn >= 0 Then
        msfx.TextMatrix(0, m_SortColumn) = Mid$(msfx.TextMatrix(0, m_SortColumn), 3)
    End If

    ' Display the new sort column's name.
    m_SortColumn = sort_column
    If m_SortOrder = flexSortGenericAscending Then
        msfx.TextMatrix(0, m_SortColumn) = "> " & msfx.TextMatrix(0, m_SortColumn)
    Else
        msfx.TextMatrix(0, m_SortColumn) = "< " & msfx.TextMatrix(0, m_SortColumn)
    End If

    ' Display the FlexGrid.
    msfx.Visible = True
End Sub

