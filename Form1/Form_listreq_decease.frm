VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_listreq_decease 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Request Account Decease"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7650
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "REJECT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "APPROVE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CheckBox chk_all 
      Caption         =   "Check All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   1815
   End
   Begin MSComctlLib.ListView LvRequest 
      Height          =   5100
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   8996
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   7440
      Y1              =   5280
      Y2              =   5280
   End
End
Attribute VB_Name = "Form_listreq_decease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs_list As ADODB.Recordset

Private Sub Command1_Click()
    Dim xxx As Integer
    Dim bceklist As Boolean
    
    bceklist = False
    konfirmasi_pesan = MsgBox("Data yang diceklist akan diset menjadi Account Decease ???", vbYesNo + vbQuestion, "Konfirmasi")
    If konfirmasi_pesan = vbYes Then
        For xxx = 1 To LvRequest.ListItems.Count
            If LvRequest.ListItems(xxx).Checked = True Then
                bceklist = True
                Exit For
            End If
        Next xxx
        
        If bceklist = True Then
            For xxx = 1 To LvRequest.ListItems.Count
                If LvRequest.ListItems(xxx).Checked = True Then
                    M_OBJCONN.Execute "UPDATE mgm SET f_decease=1 WHERE " & _
                                "custid='" & LvRequest.ListItems(xxx).Text & "'"
                    M_OBJCONN.Execute "DELETE FROM tblapprove_decease WHERE custid='" & Trim(LvRequest.ListItems(xxx).Text) & "'"
                    M_OBJCONN.Execute "INSERT INTO tblapprove_decease(custid,agent,approve_by) VALUES('" & LvRequest.ListItems(xxx).Text & "','" & LvRequest.ListItems(xxx).SubItems(2) & "','" & MDIForm1.TxtUsername.Text & "') "
                    ' Hapus log 5x Call diblock - Update 2013-04-25 By Izuddin
                    M_OBJCONN.Execute "DELETE FROM tblreq_decease WHERE custid='" & Trim(LvRequest.ListItems(xxx).Text) & "'"
                End If
            Next xxx
            MsgBox "Account(s) yg diceklist Telah diset Decease!! ", vbOKOnly + vbInformation, "INFO"
            Call IsiCustidOtomatis
        Else
            MsgBox "Anda belum mencentang ceklist data yang dipilih!!", vbOKOnly
            Exit Sub
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim xxx As Integer
    Dim bceklist As Boolean
    
    bceklist = False
    konfirmasi_pesan = MsgBox("Data yang diceklist akan di reject dari Account Decease ???", vbYesNo + vbQuestion, "Konfirmasi")
    If konfirmasi_pesan = vbYes Then
        For xxx = 1 To LvRequest.ListItems.Count
            If LvRequest.ListItems(xxx).Checked = True Then
                bceklist = True
                Exit For
            End If
        Next xxx
        
        If bceklist = True Then
            For xxx = 1 To LvRequest.ListItems.Count
                If LvRequest.ListItems(xxx).Checked = True Then
                    M_OBJCONN.Execute "DELETE FROM tblreq_decease WHERE custid='" & Trim(LvRequest.ListItems(xxx).Text) & "'"
                End If
            Next xxx
            MsgBox "Account(s) yg diceklist telah di reject !! ", vbOKOnly + vbInformation, "INFO"
            Call IsiCustidOtomatis
        Else
            MsgBox "Anda belum mencentang ceklist data yang dipilih!!", vbOKOnly
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    Call koneksi
    Call HeaderListTransfer
    Call IsiCustidOtomatis
End Sub

Private Sub chk_all_Click()
    Dim W As Integer
    If LvRequest.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvRequest.ListItems.Count
        If chk_all.Value = 1 Then
            LvRequest.ListItems(W).Checked = True
        Else
            LvRequest.ListItems(W).Checked = False
        End If
    Next W
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Rs_list = Nothing
    VIEW_MGMDATA.WindowState = 2
End Sub

Private Sub HeaderListTransfer()
    LvRequest.ColumnHeaders.ADD , , "Custid", 1500
    LvRequest.ColumnHeaders.ADD , , "Nama Cust", 2000
    LvRequest.ColumnHeaders.ADD , , "Agent", 1000
    LvRequest.ColumnHeaders.ADD , , "Tanggal", 1500
End Sub

Private Sub IsiCustidOtomatis()
    Dim ListItem As ListItem
    
    LvRequest.ListItems.CLEAR
    If Rs_list.State = 1 Then Rs_list.Close
    If LCase(MDIForm1.txtlevel.Text) = "supervisor" Then
        Rs_list.Open "SELECT a.*,b.name as nm_cust FROM tblreq_decease a,mgm b WHERE a.custid=b.custid "
    Else
        Rs_list.Open "SELECT a.*,b.name as nm_cust FROM tblreq_decease a,mgm b WHERE a.custid=b.custid AND a.agent in (SELECT userid FROM usertbl WHERE team='" & MDIForm1.TxtUsername.Text & "')"
    End If
    If Rs_list.RecordCount > 0 Then
        Do Until Rs_list.EOF
            Set ListItem = LvRequest.ListItems.ADD(, , IIf(IsNull(Rs_list!CustId), "", Rs_list!CustId))
                            ListItem.SubItems(1) = IIf(IsNull(Rs_list!nm_cust), "", Rs_list!nm_cust)
                            ListItem.SubItems(2) = IIf(IsNull(Rs_list!AGENT), "", Rs_list!AGENT)
                            ListItem.SubItems(3) = IIf(IsNull(Rs_list!tgl_request), "", Rs_list!tgl_request)

            Rs_list.MoveNext
        Loop
        Label1.Caption = "Rows : " & Rs_list.RecordCount
    Else
        MsgBox "List Request Account Decease tidak ada / kosong", vbOKOnly + vbInformation, "Info"
        Label1.Caption = "Rows : 0"
    End If
End Sub

Private Sub LvRequest_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvRequest.SortKey = ColumnHeader.Index - 1
    LvRequest.Sorted = True
End Sub

Private Sub koneksi()
    Set Rs_list = New ADODB.Recordset
    Rs_list.CursorLocation = adUseClient
    Rs_list.ActiveConnection = M_OBJCONN
    Rs_list.CursorType = adOpenDynamic
    Rs_list.LockType = adLockOptimistic
End Sub

Private Sub LvRequest_DblClick()
    If LvRequest.ListItems.Count > 0 Then
        sReminder_CUST_ID = LvRequest.SelectedItem.Text
        If bAktif_form_customer Then
            Unload FrmCC_Colection
        End If
        bAktif_Cust_Review = True
        FrmCC_Colection.Show vbModal
    End If
End Sub


