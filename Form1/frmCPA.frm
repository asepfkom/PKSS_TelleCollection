VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmCPA 
   BackColor       =   &H009AD6C2&
   Caption         =   "Create CPA"
   ClientHeight    =   5820
   ClientLeft      =   8985
   ClientTop       =   1755
   ClientWidth     =   10755
   LinkTopic       =   "Form2"
   ScaleHeight     =   5820
   ScaleWidth      =   10755
   Begin MSComctlLib.ListView lstCpa 
      Height          =   3435
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   6059
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin Threed.SSCommand cmdcpa 
      Height          =   840
      Index           =   0
      Left            =   9180
      TabIndex        =   1
      Top             =   90
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1482
      _Version        =   196610
      PictureFrames   =   1
      Picture         =   "frmCPA.frx":0000
      AutoSize        =   1
      Alignment       =   8
   End
   Begin Threed.SSCommand cmdcpa 
      Height          =   840
      Index           =   1
      Left            =   9270
      TabIndex        =   3
      Top             =   1305
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1482
      _Version        =   196610
      PictureFrames   =   1
      Picture         =   "frmCPA.frx":0589
      AutoSize        =   1
      Alignment       =   8
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   45
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   688
      _Version        =   196610
      Font3D          =   4
      ForeColor       =   12582912
      BackColor       =   10147522
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "List Create CPA"
      BevelWidth      =   2
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand cmdcpa 
      Cancel          =   -1  'True
      Height          =   615
      Index           =   3
      Left            =   9360
      TabIndex        =   6
      Top             =   2520
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   12582912
      PictureMaskColor=   -2147483644
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPA.frx":0B12
      AutoSize        =   1
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      Height          =   240
      Left            =   9405
      TabIndex        =   7
      Top             =   3240
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remove"
      Height          =   240
      Left            =   9360
      TabIndex        =   4
      Top             =   2160
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADD"
      Height          =   240
      Left            =   9315
      TabIndex        =   2
      Top             =   990
      Width           =   1140
   End
End
Attribute VB_Name = "frmCPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    createHeader
    showlist
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
    cmdcpa(1).Enabled = False
    cmdcpa(0).Enabled = False
    End If
    
End Sub

Public Sub createHeader()
    With LstCpa
        .ColumnHeaders.ADD 1, , "ID", 5
        .ColumnHeaders.ADD 2, , "custid", 1000
        .ColumnHeaders.ADD 3, , "cust name", 2000
        .ColumnHeaders.ADD 4, , "Proposal Date", 1200
        .ColumnHeaders.ADD 5, , "reff no", 1200
        .ColumnHeaders.ADD 6, , "Product", 1300
        .ColumnHeaders.ADD 7, , "Arrangement", 1500
        .ColumnHeaders.ADD 8, , "card status", 1000
        .ColumnHeaders.ADD 9, , "Total Payment", 1500
        .ColumnHeaders.ADD 10, , "Down Payment", 1500
        .ColumnHeaders.ADD 11, , "future Pay", 1500
        .ColumnHeaders.ADD 12, , "Charges", 1500
        .ColumnHeaders.ADD 13, , "discount aount", 1500
        .ColumnHeaders.ADD 14, , " O/S balance (%)", 1000
        .ColumnHeaders.ADD 15, , " Principal (%)", 1000
        .ColumnHeaders.ADD 16, , " verify", 1000
        .ColumnHeaders.ADD 17, , " Approvel ", 1000
        .ColumnHeaders.ADD 18, , " Tanggal Pelunasan ", 1200
        .ColumnHeaders.ADD 19, , "Justification ", 1500
        .ColumnHeaders.ADD 20, , "Balance ", 1500
        .ColumnHeaders.ADD 21, , "Principal", 1500
        .ColumnHeaders.ADD 22, , "Tanggal lunas", 1500
        .ColumnHeaders.ADD 23, , "Tanggal Update", 1500
        .ColumnHeaders.ADD 24, , "Occupation", 1500
        .ColumnHeaders.ADD 25, , "Reason", 1500
        .ColumnHeaders.ADD 26, , "DLQ", 1500
        .ColumnHeaders.ADD 27, , "Payment Handle", 1500
        .ColumnHeaders.ADD 28, , "Justification", 1500
        .ColumnHeaders.ADD 29, , "Verify", 1500
        .ColumnHeaders.ADD 30, , "Approvel", 1500
    End With

End Sub
Public Sub showlist()
strsql = "SELECT * from tblcpa WHERE vcustid='" + FrmCC_Colection.lblCustId.text + "'"
'       strsql = "SELECT * from tblcpa WHERE nid IN ( SELECT max(tblcpa.nid) "
'       strsql = strsql + " FROM tblcpa where vcustid='" + FrmCC_Colection.lblCustId.text + "')"
       Set rsTemporary = New ADODB.Recordset
       rsTemporary.CursorLocation = adUseClient
       rsTemporary.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
       LstCpa.ListItems.CLEAR
       While Not rsTemporary.EOF
            Set iListitem = LstCpa.ListItems.ADD(, , rsTemporary("nid"))
                iListitem.SubItems(1) = IIf(IsNull(rsTemporary("vcustid")), "", rsTemporary("vcustid"))
                iListitem.SubItems(2) = IIf(IsNull(rsTemporary("vcustname")), "", rsTemporary("vcustname"))
                iListitem.SubItems(3) = IIf(IsNull(rsTemporary("dpropsal")), "", rsTemporary("dpropsal"))
                iListitem.SubItems(4) = IIf(IsNull(rsTemporary("vreffno")), "", rsTemporary("vreffno"))
                iListitem.SubItems(5) = IIf(IsNull(rsTemporary("vproduct")), "", rsTemporary("vproduct"))
                iListitem.SubItems(6) = IIf(IsNull(rsTemporary("varragement")), "", rsTemporary("varragement"))
                iListitem.SubItems(7) = IIf(IsNull(rsTemporary("vcardsts")), "", rsTemporary("vcardsts"))
                iListitem.SubItems(8) = Format(IIf(IsNull(rsTemporary("nttlpayment")), 0, rsTemporary("nttlpayment")), "##,###")
                iListitem.SubItems(9) = Format(IIf(IsNull(rsTemporary("ndownpay")), 0, rsTemporary("ndownpay")), "##,###")
                iListitem.SubItems(10) = Format(IIf(IsNull(rsTemporary("nfuturepay")), 0, rsTemporary("nfuturepay")), "##,###")
                iListitem.SubItems(11) = Format(IIf(IsNull(rsTemporary("ncharge")), 0, rsTemporary("ncharge")), "##,###")
                iListitem.SubItems(12) = Format(IIf(IsNull(rsTemporary("ndiscountamt")), 0, rsTemporary("ndiscountamt")), "##,###")
                iListitem.SubItems(13) = IIf(IsNull(rsTemporary("vosbalance")), "", rsTemporary("vosbalance"))
                iListitem.SubItems(14) = IIf(IsNull(rsTemporary("vosprincipal")), "", rsTemporary("vosprincipal"))
                iListitem.SubItems(15) = IIf(IsNull(rsTemporary("vverify")), "", rsTemporary("vverify"))
                iListitem.SubItems(16) = IIf(IsNull(rsTemporary("votority")), "", rsTemporary("votority"))
                iListitem.SubItems(17) = IIf(IsNull(rsTemporary("dtglpelunasan")), "", Format(rsTemporary("dtglpelunasan"), "dd/mm/yyyy"))
               iListitem.SubItems(18) = IIf(IsNull(rsTemporary("vjust")), "", rsTemporary("vjust"))
               iListitem.SubItems(19) = Format(IIf(IsNull(rsTemporary("nbalance")), 0, rsTemporary("nbalance")), "##,###")
               iListitem.SubItems(20) = Format(IIf(IsNull(rsTemporary("nprincipal")), 0, rsTemporary("nprincipal")), "##,###")
               iListitem.SubItems(21) = IIf(IsNull(rsTemporary("dtglpelunasan")), "", Format(rsTemporary("dtglpelunasan"), "dd/mm/yyyy"))
                iListitem.SubItems(22) = IIf(IsNull(rsTemporary("dtgllastupdate")), "", Format(rsTemporary("dtgllastupdate"), "dd/mm/yyyy"))
                iListitem.SubItems(23) = IIf(IsNull(rsTemporary("voccupation")), "", rsTemporary("voccupation"))
                iListitem.SubItems(24) = IIf(IsNull(rsTemporary("vreason")), "", rsTemporary("vreason"))
                iListitem.SubItems(25) = IIf(IsNull(rsTemporary("vnodlq")), "", rsTemporary("vnodlq"))
                iListitem.SubItems(26) = IIf(IsNull(rsTemporary("vpaymenthandle")), "", rsTemporary("vpaymenthandle"))
                 iListitem.SubItems(27) = IIf(IsNull(rsTemporary("vjust")), "", rsTemporary("vjust"))
                iListitem.SubItems(28) = IIf(IsNull(rsTemporary("intverify")), "0", rsTemporary("intverify"))
                iListitem.SubItems(29) = IIf(IsNull(rsTemporary("intapprovel")), "0", rsTemporary("intapprovel"))
            rsTemporary.MoveNext
       Wend
       Set rsTemporary = Nothing
       Set iListitem = Nothing
       
End Sub
Private Sub cmdcpa_Click(Index As Integer)
    Select Case Index
    Case 0
        frmCPA.WindowState = 1
        With frmisicpa
        
            .Caption = "Add"
            .SSCommand1(0).Tag = 1
            .Label5.text = IIf(FrmCC_Colection.lblAmount.ValueIsNull, "0", FrmCC_Colection.lblAmount)
            .Label8.text = IIf(FrmCC_Colection.LblPrompA.ValueIsNull, "0", FrmCC_Colection.LblPrompA)
            .txtregion.text = FrmCC_Colection.lblregion
            .txtcardno.text = FrmCC_Colection.lblCustId
            .TxtName.text = FrmCC_Colection.lblNama.text
            .txtproduct.text = "CARD"
            .dtcardopen.Value = FrmCC_Colection.lblOpenDate.Value
            .txtplace.text = "CardHolder"
            .txtcollect.text = VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11)
            .Show 1
        End With
        
    Case 1
         If LstCpa.ListItems.Count <> 0 Then
            If MsgBox("Yakin Akan dihapus...!!!!", vbQuestion + vbYesNo, "Peringatan") = vbYes Then
                strsql = "select custid ,intverify , intapprovel from mgm where custid='" + LstCpa.SelectedItem.SubItems(1) + "'"
                Set rsTemporary = New ADODB.Recordset
                    rsTemporary.CursorLocation = adUseClient
                    rsTemporary.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
                    If Not rsTemporary.EOF Then
                           If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
                                If (rsTemporary("intverify") = 0 Or rsTemporary("intapprovel") = 0) And LstCpa.SelectedItem.SubItems(22) = Format(MDIForm1.TDBDate1.Value, "dd/mm/yyyy") Then
                                        strsql = "delete from tblcpa where nid='" + LstCpa.SelectedItem.text + "'"
                                        M_OBJCONN.Execute (strsql)
                                        strsql = "update mgm set stscpa=0 where custid ='" + FrmCC_Colection.lblCustId.text + "'"
                                         M_OBJCONN.Execute (strsql)
                                        LstCpa.ListItems.Remove LstCpa.SelectedItem.Index
                                        Exit Sub
                                End If
                            End If
                            
                            If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
                                    If rsTemporary("intapprovel") = 0 And LstCpa.SelectedItem.SubItems(22) = Format(MDIForm1.TDBDate1.Value, "dd/mm/yyyy") Then
                                        strsql = "delete from tblcpa where nid='" + LstCpa.SelectedItem.text + "'"
                                        M_OBJCONN.Execute (strsql)
                                        strsql = "update mgm set intverify=0,vnameverify='',stscpa=0,intapprovel=0 where custid ='" + FrmCC_Colection.lblCustId.text + "'"
                                        M_OBJCONN.Execute (strsql)
                                        LstCpa.ListItems.Remove LstCpa.SelectedItem.Index
                                        Exit Sub
                                End If
                            Else
                            strsql = "delete from tblcpa where nid='" + LstCpa.SelectedItem.text + "'"
                            M_OBJCONN.Execute (strsql)
                            LstCpa.ListItems.Remove LstCpa.SelectedItem.Index
                            End If
                            
                    End If
            End If
         End If
     Case 3
     Unload Me
     
     
    End Select
End Sub
Private Sub lstCpa_DblClick()
    If LstCpa.ListItems.Count <> 0 Then
        With frmisicpa
         
            .Caption = "Edit"
            .SSCommand1(0).Tag = 2
            .txtregion.text = FrmCC_Colection.lblregion
            .txtcardno.text = FrmCC_Colection.lblCustId.text
            .TxtName.text = Replace(FrmCC_Colection.lblNama.text, "'", "")
            .txtproduct.text = "CARD"
            .dtcardopen.Value = FrmCC_Colection.lblOpenDate.Value
            .lblLastPay.Value = IIf(LstCpa.SelectedItem.SubItems(8) = "", "0", LstCpa.SelectedItem.SubItems(8))
            .txtdownpayment.Value = IIf(LstCpa.SelectedItem.SubItems(9) = "", "0", LstCpa.SelectedItem.SubItems(9))
            .txtplace.text = "CardHolder"
            .Label5.text = IIf(FrmCC_Colection.lblAmount.ValueIsNull, "0", FrmCC_Colection.lblAmount)
            .Label8.text = IIf(FrmCC_Colection.LblPrompA.ValueIsNull, "0", FrmCC_Colection.LblPrompA)
            .txtreff = LstCpa.SelectedItem.SubItems(4)
            .txtcharge = IIf(LstCpa.SelectedItem.SubItems(10) = "", "0", LstCpa.SelectedItem.SubItems(10))
            .txtprincipal.Value = IIf(LstCpa.SelectedItem.SubItems(20) = "", "0", LstCpa.SelectedItem.SubItems(20))
            .txtcollect.text = VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11)
            .cbosts.text = IIf(LstCpa.SelectedItem.SubItems(7) = "", "WO", LstCpa.SelectedItem.SubItems(7))
            .txtbalance.Value = IIf(LstCpa.SelectedItem.SubItems(19) = "", "0", LstCpa.SelectedItem.SubItems(19))
            .txtarrangement.text = LstCpa.SelectedItem.SubItems(6)
            .txtfrombalancepersen.text = LstCpa.SelectedItem.SubItems(13)
            .txtpersenprincipal.text = LstCpa.SelectedItem.SubItems(14)
            .dtpropsal.Value = Format(LstCpa.SelectedItem.SubItems(3), "dd/mm/yyyy")
            .dtpelunasan = Format(LstCpa.SelectedItem.SubItems(21), "dd/mm/yyyy")
            .txtoccupation.text = LstCpa.SelectedItem.SubItems(23)
            .txtreason.text = LstCpa.SelectedItem.SubItems(24)
            .txtnodlq.text = LstCpa.SelectedItem.SubItems(25)
            .txtpaymenthandle.text = LstCpa.SelectedItem.SubItems(26)
            .txtjust.text = LstCpa.SelectedItem.SubItems(27)
            If strStatusCpa = "GAGAL" Then
            .chkcek(2).Value = 1
            Else
            
            End If
            
            If LstCpa.SelectedItem.SubItems(28) = "0" Then
            .chkcek(0).Value = 0
            Else
            .chkcek(0).Value = 1
            End If
            
            If LstCpa.SelectedItem.SubItems(29) = "0" Then
            .chkcek(1).Value = 0
            Else
            .chkcek(1).Value = 1
            End If
            .Show 1
        End With
    End If
 
End Sub


