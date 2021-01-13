VERSION 5.00
Begin VB.Form frmextensioncc 
   BorderStyle     =   0  'None
   Caption         =   "&H00ABE18E&"
   ClientHeight    =   4800
   ClientLeft      =   3840
   ClientTop       =   510
   ClientWidth     =   4350
   LinkTopic       =   "Form2"
   ScaleHeight     =   4800
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FCFCFC&
      Caption         =   "OTHER INFO"
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   11
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   4320
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   9
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3600
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   8
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3240
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   7
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2880
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   10
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3960
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   0
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   5
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2160
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   4
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1800
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   3
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1440
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   2
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   1
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   6
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2520
         Width           =   2355
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3960
         TabIndex        =   25
         Top             =   120
         Width           =   375
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Roll Over Amount"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   23
         Top             =   4320
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Deft Amount"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   11
         Top             =   3960
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Int Amount"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Principal Amount"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Over Due"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Contacted Mobile 2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   1680
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Relationship 2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Name 2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Roll Pay Code BCA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Code BCA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Roll Pay Code"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Code"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmextensioncc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    Call isi
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub isi()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    CMDSQL = "SELECT * FROM mgm where custid = '" & FrmCC_Colection.lblCustId.text & "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    txta(0).text = cnull(M_objrs!stskathomeadd1)
    txta(1).text = cnull(M_objrs!stskathomeadd2)
    txta(2).text = cnull(M_objrs!stskatofficeadd1)
    txta(3).text = cnull(M_objrs!stskatofficeadd2)
    txta(4).text = cnull(M_objrs!stskathpadd1)
    txta(5).text = cnull(M_objrs!stskathpadd2)
    txta(6).text = cnull(M_objrs!f_sts_valid_home1)
    txta(7).text = cnull(M_objrs!f_sts_valid_home2)
    txta(8).text = cnull(M_objrs!f_sts_valid_office1)
    txta(9).text = cnull(M_objrs!f_sts_valid_office2)
    txta(10).text = cnull(M_objrs!f_sts_valid_mobile1)
    txta(11).text = cnull(M_objrs!f_sts_valid_mobile2)
    
    Set M_objrs = Nothing
End Sub

Private Sub txta_click(Index As Integer)
    
    
    If Index = 6 Then
        FrmCC_Colection.txtgetnomor.text = txta(6).text
        FrmCC_Colection.Txtperiod.Caption = txta(6).text
        FrmCC_Colection.CmbPhone.text = "CON MOBILE 2"
        Unload Me
    End If
    
End Sub
