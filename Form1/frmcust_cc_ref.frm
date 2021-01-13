VERSION 5.00
Begin VB.Form frmcust_cc_ref 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Isi Referensi"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   ControlBox      =   0   'False
   Icon            =   "frmcust_cc_ref.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1230
      TabIndex        =   7
      Top             =   1860
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3660
      TabIndex        =   9
      Top             =   2235
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   390
      Left            =   2655
      TabIndex        =   8
      Top             =   2235
      Width           =   930
   End
   Begin VB.TextBox Text7 
      Height          =   330
      Left            =   1230
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1515
      Width           =   1665
   End
   Begin VB.TextBox Text6 
      Height          =   330
      Left            =   1230
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1170
      Width           =   1665
   End
   Begin VB.TextBox Text5 
      Height          =   330
      Left            =   3360
      TabIndex        =   4
      Top             =   825
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   330
      Left            =   1245
      MaxLength       =   20
      TabIndex        =   3
      Top             =   825
      Width           =   1665
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   1245
      MaxLength       =   20
      TabIndex        =   1
      Top             =   480
      Width           =   1665
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1245
      MaxLength       =   50
      TabIndex        =   0
      Top             =   135
      Width           =   2805
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Hubungan :"
      Height          =   315
      Left            =   45
      TabIndex        =   17
      Top             =   1890
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Telp Rumah 2 :"
      Height          =   315
      Left            =   45
      TabIndex        =   16
      Top             =   1530
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Telp Rumah :"
      Height          =   315
      Left            =   165
      TabIndex        =   15
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Ext :"
      Height          =   315
      Left            =   2910
      TabIndex        =   14
      Top             =   840
      Width           =   435
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Ext :"
      Height          =   315
      Left            =   2925
      TabIndex        =   13
      Top             =   495
      Width           =   435
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Telp Kantor 2 :"
      Height          =   315
      Left            =   75
      TabIndex        =   12
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Telp Kantor :"
      Height          =   315
      Left            =   180
      TabIndex        =   11
      Top             =   510
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nama :"
      Height          =   315
      Left            =   180
      TabIndex        =   10
      Top             =   165
      Width           =   975
   End
End
Attribute VB_Name = "frmcust_cc_ref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ok As Boolean
Public ADD As Boolean

Private Sub Command1_Click()
    If Text1.Text = Empty Then
        MsgBox "Nama Harus Diisi", vbOKOnly + vbCritical, "Telegrandi"
        Text1.SetFocus
        Exit Sub
    End If
    If Text2.Text = Empty And Text4.Text = Empty And Text6.Text = Empty And Text7.Text = Empty Then
        MsgBox "Salah Satu Telpon Harus DiIsi", vbOKOnly + vbCritical, "Informasi"
        Exit Sub
    End If
    ok = True
    Me.Hide
    FRMCUST_CC_mgm.ListView1(0).SetFocus
End Sub

Private Sub Command2_Click()
    ok = False
    Unload Me
End Sub

Private Sub Form_Load()
Dim M_OBJRS As New ADODB.Recordset
    Me.MousePointer = vbHourglass
    Combo1.AddItem "Teman"
    Combo1.AddItem "Keluarga"
    
    Me.Top = 5000
    Me.Left = 2500
    If ADD = False Then
        M_OBJRS.CursorLocation = adUseClient
        M_OBJRS.Open "Select * from MGM_REF where mgm_id ='" + FRMCUST_CC_mgm.ListView1(0).SelectedItem.SubItems(3) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If Not M_OBJRS.EOF Then
            Text1.Text = IIf(IsNull(M_OBJRS!NAMA), "", M_OBJRS!NAMA)
            Text2.Text = IIf(IsNull(M_OBJRS!OFFICENO), "", M_OBJRS!OFFICENO)
            Text3.Text = IIf(IsNull(M_OBJRS!EXTOFFICENO), "", M_OBJRS!EXTOFFICENO)
            Text4.Text = IIf(IsNull(M_OBJRS!OFFICENO2), "", M_OBJRS!OFFICENO2)
            Text5.Text = IIf(IsNull(M_OBJRS!EXTOFFICENO2), "", M_OBJRS!EXTOFFICENO2)
            Text6.Text = IIf(IsNull(M_OBJRS!HOMENO), "", M_OBJRS!HOMENO)
            Text7.Text = IIf(IsNull(M_OBJRS!HOMENO2), "", M_OBJRS!HOMENO2)
            Combo1.Text = IIf(IsNull(M_OBJRS!HUBUNGAN), "", M_OBJRS!HUBUNGAN)
        End If
    End If
    Set M_OBJRS = Nothing
    Me.MousePointer = vbNormal
End Sub


