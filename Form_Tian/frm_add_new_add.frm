VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frm_add_new_add 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Address"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8790
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Text1 
      Height          =   315
      ItemData        =   "frm_add_new_add.frx":0000
      Left            =   1560
      List            =   "frm_add_new_add.frx":0028
      TabIndex        =   32
      Top             =   1270
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   1920
      Width           =   6735
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   5400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      _Version        =   196610
      Font3D          =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Add"
      PictureAlignment=   4
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   28
      Top             =   4920
      Width           =   5055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   7
      Left            =   5280
      TabIndex        =   26
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   6
      Left            =   3600
      TabIndex        =   24
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   23
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   3
      Left            =   4200
      TabIndex        =   18
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   17
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   12
      Top             =   3450
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   3450
      Width           =   1575
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1931
      _Version        =   196610
      BackColor       =   16579836
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   7080
         TabIndex        =   2
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FCFCFC&
         Caption         =   "Labelappid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FCFCFC&
         Caption         =   "AppID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblcustid 
         BackColor       =   &H00FCFCFC&
         Caption         =   "Labelcustid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lbl1 
         BackColor       =   &H00FCFCFC&
         Caption         =   "Card No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FCFCFC&
         Caption         =   "Choose Address Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   615
      Index           =   1
      Left            =   1680
      TabIndex        =   30
      Top             =   5400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      _Version        =   196610
      Font3D          =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
      PictureAlignment=   4
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   27
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   25
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mobile Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   21
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Office Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   20
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Home Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Kelurahan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Kecamatan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   15
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "RT/RW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Zip Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   11
      Top             =   3195
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   3200
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Address Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   0
      X2              =   8760
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "frm_add_new_add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub getadrtype()
    query = "SELECT adr_type from bca.tbl_address where custid = '" + lblcustid.Caption + "' and f_app = 1"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
        While Not M_objrs.EOF
            Combo1.AddItem M_objrs!adr_type
            M_objrs.MoveNext
        Wend
        Set M_objrs = Nothing
End Sub

Private Sub Combo1_Click()
    query = "SELECT * FROM bca.tbl_address where custid = '" + lblcustid.Caption + "' and adr_type = '" + Combo1.Text + "' "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
        While Not M_objrs.EOF
            Text1.Text = IIf(IsNull(M_objrs("Adr_type")), "", M_objrs("Adr_type"))
            Text3(0).Text = IIf(IsNull(M_objrs("address1")), "", M_objrs("address1"))
            Text3(1).Text = IIf(IsNull(M_objrs("address2")), "", M_objrs("address2"))
            Text3(2).Text = IIf(IsNull(M_objrs("address3")), "", M_objrs("address3"))
            Text3(3).Text = IIf(IsNull(M_objrs("address4")), "", M_objrs("address4"))
            Text2(0).Text = IIf(IsNull(M_objrs("city")), "", M_objrs("city"))
            Text2(1).Text = IIf(IsNull(M_objrs("zipcode")), "", M_objrs("zipcode"))
            Text3(4).Text = IIf(IsNull(M_objrs("contact1")), "", M_objrs("contact1"))
            Text3(5).Text = IIf(IsNull(M_objrs("contact2")), "", M_objrs("contact2"))
            Text3(6).Text = IIf(IsNull(M_objrs("mobileno")), "", M_objrs("mobileno"))
            Text3(7).Text = IIf(IsNull(M_objrs("fax")), "", M_objrs("fax"))
            Text3(8).Text = IIf(IsNull(M_objrs("email")), "", M_objrs("email"))
            M_objrs.MoveNext
        Wend
        Set M_objrs = Nothing
        
    If Combo1.Text <> "" Then
        SSCommand1(0).Caption = "UPDATE"
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    lblcustid.Caption = FrmCC_Colection.lblnocard.Caption
    Label3.Caption = FrmCC_Colection.lvaddress.ListItems(1).SubItems(1)
    Call getadrtype
End Sub

Private Sub SSCommand1_Click(Index As Integer)
    Select Case Index
        Case 1
            Unload Me
        Case 0
            
            If Text1.Text = "" Then
                MsgBox "Address Type Can't be Empty"
                Exit Sub
            End If
            
            query = "SELECT * FROM bca.tbl_address where custid = '" + lblcustid.Caption + "' and adr_type = '" + Text1.Text + "' "
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic
            
            If Combo1.Text <> "" Then
                query = " UPDATE bca.tbl_address set address1 = '" + Text3(0).Text + "',"
                query = query + " address2 = '" + Text3(1).Text + "', "
                query = query + " address3 = '" + Text3(2).Text + "', "
                query = query + " address4 = '" + Text3(3).Text + "', "
                query = query + " city = '" + Text2(0).Text + "', "
                query = query + " zipcode = '" + Text2(1).Text + "', "
                query = query + " contact1 = '" + Text3(4).Text + "', "
                query = query + " contact2 = '" + Text3(5).Text + "', "
                query = query + " mobileno = '" + Text3(6).Text + "', "
                query = query + " fax = '" + Text3(7).Text + "', "
                query = query + " email = '" + Text3(8).Text + "', "
                query = query + " agent =  '" + MDIForm1.TxtUsername.Text + "' "
                query = query + " WHERE custid = '" + lblcustid.Caption + "' and adr_type = '" + Text1.Text + "'"
                M_OBJCONN.Execute query
                
                MsgBox "Data Updated"
            Else
                If M_objrs.RecordCount > 0 Then
                    MsgBox "Tipe Address Tidak Boleh Sama"
                    Exit Sub
                End If
                
                query = "INSERT INTO bca.tbl_address(custid,appid,adr_type,address1,address2,address3,address4,city,zipcode,contact1,contact2,mobileno,fax,email,f_app,app_spv,tglreq,agent) values "
                query = query + " ( '" + lblcustid.Caption + "', '" + Label3.Caption + "', '" + Text1.Text + "', '" + Text3(0).Text + "' "
                query = query + " , '" + Text3(1).Text + "', '" + Text3(2).Text + "', '" + Text3(3).Text + "', '" + Text2(0).Text + "', '" + Text2(1).Text + "'"
                query = query + " , '" + Text3(4).Text + "', '" + Text3(5).Text + "', '" + Text3(6).Text + "', '" + Text3(7).Text + "', '" + Text3(8).Text + "', 1, 1, now(), '" + MDIForm1.TxtUsername.Text + "'"
                query = query + " ) "
                M_OBJCONN.Execute query
                
                CMDSQL = "INSERT INTO bca.tblnotif_info "
                CMDSQL = CMDSQL & "( type_notif,notif_from) values ('"
                CMDSQL = CMDSQL & "address','" & Trim$(MDIForm1.TxtUsername.Text) & "')"
                M_OBJCONN.Execute CMDSQL
                
                MsgBox "Address Has Been Inserted"
            
                Unload Me
            End If
    End Select
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Text2_Change(Index As Integer)
    Select Case Index
        Case 1
            textval = Text2(1).Text
            If IsNumeric(textval) Then
              numval = textval
            Else
              Text2(1).Text = CStr(numval)
            End If
    End Select
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 4
            If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
                KeyAscii = 0
                KeyAscii = vbKeyBack
            End If
        Case 5
            If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
                KeyAscii = 0
                KeyAscii = vbKeyBack
            End If
        Case 6
            If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
                KeyAscii = 0
                KeyAscii = vbKeyBack
            End If
        Case 7
            If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
                KeyAscii = 0
                KeyAscii = vbKeyBack
            End If
    End Select
End Sub
