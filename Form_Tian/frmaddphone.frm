VERSION 5.00
Begin VB.Form frmaddphone 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4845
   ClientLeft      =   12045
   ClientTop       =   645
   ClientWidth     =   4350
   LinkTopic       =   "Form2"
   ScaleHeight     =   4845
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
         TabIndex        =   25
         Top             =   4320
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   10
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3960
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   9
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   3600
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   6
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2520
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   1
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   2
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   3
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   4
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1800
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   5
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2160
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   0
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   7
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2880
         Width           =   2355
      End
      Begin VB.TextBox txta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         Height          =   285
         Index           =   8
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   3240
         Width           =   2355
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "EC PHONE 2"
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
         Index           =   12
         Left            =   120
         TabIndex        =   26
         Top             =   4320
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "EC DESC 2"
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
         TabIndex        =   24
         Top             =   3960
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "EC NAME 2"
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
         TabIndex        =   21
         Top             =   3600
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "CO HOME 1"
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
         TabIndex        =   19
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "CO OFFICE 1"
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
         TabIndex        =   18
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "CO HOME 2"
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
         TabIndex        =   17
         Top             =   1080
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "CO OFFICE 2"
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
         TabIndex        =   16
         Top             =   1440
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "HM PHONE 2"
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
         TabIndex        =   15
         Top             =   1800
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "OFF PHONE 2"
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
         TabIndex        =   14
         Top             =   2160
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "ALT PHONE 1"
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
         TabIndex        =   13
         Top             =   2520
         Width           =   1680
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "ALT PHONE 2"
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
         TabIndex        =   12
         Top             =   2880
         Width           =   1560
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00ABE18E&
         BackStyle       =   0  'Transparent
         Caption         =   "ALT PHONE 3"
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
         TabIndex        =   11
         Top             =   3240
         Width           =   1560
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
         TabIndex        =   10
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00ABE18E&
      BackStyle       =   0  'Transparent
      Caption         =   "ALT PHONE 3"
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
      Left            =   0
      TabIndex        =   22
      Top             =   4080
      Width           =   1560
   End
End
Attribute VB_Name = "frmaddphone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    addphone = True
    Call isi
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    addphone = False
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
    
    txta(0).text = cnull(M_objrs!co_home_phone_1)
    txta(1).text = cnull(M_objrs!co_office_phone_1)
    txta(2).text = cnull(M_objrs!co_home_phone_2)
    txta(3).text = cnull(M_objrs!co_office_phone_2)
    txta(4).text = cnull(M_objrs!home_phone_2)
    txta(5).text = cnull(M_objrs!office_phone_2)
    txta(6).text = cnull(M_objrs!alt_phone_1)
    txta(7).text = cnull(M_objrs!alt_phone_2)
    txta(8).text = cnull(M_objrs!alt_phone_3)
    txta(9).text = cnull(M_objrs!stskathpadd1)
    txta(10).text = cnull(M_objrs!stskathpadd2)
    txta(11).text = cnull(M_objrs!f_sts_valid_home1)
    
    
    Set M_objrs = Nothing
End Sub

Private Sub txta_Click(Index As Integer)
    Dim i As Integer
        
    If lg_call = False Then
        i = Index
        
        If i <> 9 And i <> 10 Then
            If Len(txta(i).text) > 3 Then
                If txta(Index).text <> Empty Then
                    FrmCC_Colection.txtPhone.text = txta(Index).text
                    FrmCC_Colection.txtPhone.text = Replace(FrmCC_Colection.txtPhone.text, " ", "")
                    FrmCC_Colection.txtPhone.text = Replace(FrmCC_Colection.txtPhone.text, "'", "")
    
                    FrmCC_Colection.txtgetnomor.text = txta(Index).text
                    FrmCC_Colection.Txtperiod.Caption = txta(Index).text
                End If
                If Index = 0 Then
                    FrmCC_Colection.CmbPhone.text = "ICO HOME 1"
                ElseIf Index = 1 Then
                    FrmCC_Colection.CmbPhone.text = "ICO OFFICE 1"
                ElseIf Index = 2 Then
                    FrmCC_Colection.CmbPhone.text = "ICO HOME 2"
                ElseIf Index = 3 Then
                    FrmCC_Colection.CmbPhone.text = "ICO OFFICE 2"
                ElseIf Index = 4 Then
                    FrmCC_Colection.CmbPhone.text = "ICO HM PHONE 2"
                ElseIf Index = 5 Then
                    FrmCC_Colection.CmbPhone.text = "ICO OFF PHONE 2"
                ElseIf Index = 6 Then
                    FrmCC_Colection.CmbPhone.text = "ICO ALT PHONE 1"
                ElseIf Index = 7 Then
                    FrmCC_Colection.CmbPhone.text = "ICO ALT PHONE 2"
                ElseIf Index = 8 Then
                    FrmCC_Colection.CmbPhone.text = "ICO ALT PHONE 3"
                ElseIf Index = 11 Then
                    FrmCC_Colection.CmbPhone.text = "EC 2"
                End If
            Else
                FrmCC_Colection.CmbPhone.text = ""
            End If
        End If
        FrmCC_Colection.Frame3.Caption = "1"
    End If
    Unload Me
End Sub
