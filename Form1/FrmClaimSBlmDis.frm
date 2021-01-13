VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmClaimSBlmDis 
   Caption         =   "Data Belum Di Distribusi"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12285
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   12285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Verifikasi Ok"
      Height          =   585
      Left            =   10875
      TabIndex        =   1
      Top             =   450
      Width           =   915
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6900
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   12171
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmClaimSBlmDis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Id", 10 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Nama", 10 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Telp Rumah", 20 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Telp Rumah 2", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "HandPhone", 10 * TXT
    ListView1.ColumnHeaders.ADD 6, , "HandPhone 2", 15 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Telp Kantor", 15 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Telp Kantor 2", 15 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Nama Pt", 15 * TXT
    ListView1.ColumnHeaders.ADD 10, , "RecSource", 15 * TXT
End Sub


Private Sub Command1_Click()
Dim cmdsql As String
If ListView1.ListItems.Count = 0 Then
    Exit Sub
End If
    cmdsql = "Insert into MGM "
    cmdsql = cmdsql + "(Custid,Name,HOMENO,HOMENO2,MOBILENO,MOBILENO2,OFFICENO,OFFICENO2,NAMAPT,RECSOURCE,agent)"
    cmdsql = cmdsql + " Values"
    cmdsql = cmdsql + " ('" + ListView1.SelectedItem.Text + "',"
    cmdsql = cmdsql + " '" + ListView1.SelectedItem.SubItems(1) + "',"
    cmdsql = cmdsql + " '" + ListView1.SelectedItem.SubItems(2) + "',"
    cmdsql = cmdsql + " '" + ListView1.SelectedItem.SubItems(3) + "',"
    cmdsql = cmdsql + " '" + ListView1.SelectedItem.SubItems(4) + "',"
    cmdsql = cmdsql + " '" + ListView1.SelectedItem.SubItems(5) + "',"
    cmdsql = cmdsql + " '" + ListView1.SelectedItem.SubItems(6) + "',"
    cmdsql = cmdsql + " '" + ListView1.SelectedItem.SubItems(7) + "',"
    cmdsql = cmdsql + " '" + ListView1.SelectedItem.SubItems(8) + "',"
    cmdsql = cmdsql + " '" + ListView1.SelectedItem.SubItems(9) + "',"
    cmdsql = cmdsql + " '" + FrmClaimVerifikasi.TxtAgentClaim.Text + "')"
M_OBJCONN.Execute cmdsql

M_OBJCONN.Execute "delete from tempcc_custtbl where custid ='" + ListView1.SelectedItem.Text + "'"
    cmdsql = "UPDATE ClaimSheet SET "
    cmdsql = cmdsql + " AgentLama ='Data Belum Didistribusi', "
    cmdsql = cmdsql + " KodeStatus ='1', "
    cmdsql = cmdsql + " Keterangan ='Telah Diverifikasi oleh " + MDIForm1.Text1.Text + " ' "
    cmdsql = cmdsql + " where id = " + FrmClaimVerifikasi.TxtId.Text + ""
M_OBJCONN.Execute cmdsql
FrmClaimList.ListView1.SelectedItem.SubItems(2) = "Data Belum Didistribusi"
FrmClaimList.ListView1.SelectedItem.SubItems(5) = "1"
FrmClaimList.ListView1.SelectedItem.SubItems(6) = "Telah Diverifikasi oleh " & MDIForm1.Text1.Text
MsgBox "Proses Selesai", vbInformation + vbOKOnly, "Telegrandi"
Unload Me
Unload FrmClaimVerifikasi
End Sub

Private Sub form_load()
Dim listitem As listitem
Dim m_cari As New ADODB.Recordset
Call header
m_cari.CursorLocation = adUseClient
m_cari.Open "Select * from tempcc_custtbl where name like '%" + FrmClaimVerifikasi.TxtNamaDiKartu.Text + "%'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_cari.EOF
     Set listitem = ListView1.ListItems.ADD(, , m_cari("CUSTID"))
        listitem.SubItems(1) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
        listitem.SubItems(2) = IIf(IsNull(m_cari("HOMENO")), "", m_cari("HOMENO"))
        listitem.SubItems(3) = IIf(IsNull(m_cari("HOMENO2")), "", m_cari("HOMENO2"))
        listitem.SubItems(4) = IIf(IsNull(m_cari("MOBILENO")), "", m_cari("MOBILENO"))
        listitem.SubItems(5) = IIf(IsNull(m_cari("MOBILENO2")), "", m_cari("MOBILENO2"))
        listitem.SubItems(6) = IIf(IsNull(m_cari("OFFICENO")), "", m_cari("OFFICENO"))
        listitem.SubItems(7) = IIf(IsNull(m_cari("OFFICENO2")), "", m_cari("OFFICENO2"))
        listitem.SubItems(8) = IIf(IsNull(m_cari("NAMAPT")), "", m_cari("NAMAPT"))
        listitem.SubItems(9) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
        m_cari.MoveNext
Wend
m_cari.Close
Set m_cari = Nothing
End Sub
