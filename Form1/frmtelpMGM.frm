VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmtelp_mgm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   2280
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmtelpMGM.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Height          =   4020
      Left            =   30
      TabIndex        =   11
      Top             =   15
      Width           =   2250
      Begin MSCommLib.MSComm MSComm1 
         Left            =   2745
         Top             =   135
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   540
         TabIndex        =   9
         Top             =   3435
         Width           =   1125
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000004&
         Height          =   3210
         Left            =   60
         TabIndex        =   10
         Top             =   150
         Width           =   2100
         Begin VB.Frame Frame3 
            BackColor       =   &H80000004&
            Height          =   660
            Left            =   45
            TabIndex        =   12
            Top             =   3435
            Visible         =   0   'False
            Width           =   3345
            Begin VB.CommandButton Command3 
               BackColor       =   &H00C0C0C0&
               Caption         =   "&Call"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2655
               TabIndex        =   7
               Top             =   180
               Width           =   615
            End
            Begin VB.TextBox Text16 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   60
               MaxLength       =   30
               TabIndex        =   6
               Top             =   210
               Width           =   2550
            End
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OfficePhone 2"
            Height          =   480
            Index           =   1
            Left            =   90
            TabIndex        =   1
            Top             =   660
            UseMaskColor    =   -1  'True
            Width           =   1935
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OfficePhone 1"
            Height          =   480
            Index           =   0
            Left            =   90
            TabIndex        =   0
            Top             =   165
            Width           =   1935
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "HandPhone 2"
            Height          =   480
            Index           =   3
            Left            =   75
            TabIndex        =   3
            Top             =   1650
            Width           =   1935
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "HandPhone 1"
            Height          =   480
            Index           =   2
            Left            =   75
            TabIndex        =   2
            Top             =   1155
            Width           =   1935
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "HomePhone  2"
            Height          =   480
            Index           =   5
            Left            =   75
            TabIndex        =   5
            Top             =   2655
            Width           =   1935
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "HomePhone 1"
            Height          =   480
            Index           =   4
            Left            =   75
            TabIndex        =   4
            Top             =   2160
            Width           =   1935
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Keluar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   420
         TabIndex        =   8
         Top             =   4395
         Visible         =   0   'False
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmtelp_mgm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cancelflag As Boolean
Public SUKSESPHONE As Boolean
Public TELPRUMAH1 As String
Public TELPRUMAH2 As String
Public TELPKANTOR1 As String
Public TELPKANTOR2 As String
Public HANDPHONE1 As String
Public HANDPHONE2 As String

Private Function ILANGIN_AREA(TELP As String) As String
    ILANGIN_AREA = Replace(TELP, "021", "")
End Function

Private Sub Command1_Click(Index As Integer)
Dim NOTELP As String
cancelflag = False
NOTELP = Empty

Dim m_objcek As New ADODB.Recordset
m_objcek.CursorLocation = adUseClient
m_objcek.Open "Select * from vwcallcfg1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If Format(m_objcek!tglsystem, "hh:nn:ss") < MDIForm1.TxtJamMulaiTelp.Text Or Format(m_objcek!tglsystem, "hh:nn:ss") > MDIForm1.TxtJamSelesaiTelp.Text Then
    MsgBox "Jam Operasional Telephone dari 08:00:00 sampai dengan 20:00:00", vbInformation + vbOKOnly, "Telegrandi"
    Set m_objcek = Nothing
    Exit Sub
End If
Set m_objcek = Nothing

If MSComm1.CommPort = False Then
       MSComm1.CommPort = MDIForm1.TxtCommPort.Text
       MSComm1.Settings = "9600,N,8,1"
       MSComm1.PortOpen = True
End If

Select Case Index
Case 0
    If Len(TELPKANTOR1) < 3 Then
        Me.MousePointer = 0
        MsgBox "No Telepon Tidak Valid", vbCritical + vbOKOnly, "TeleGrandi"
    Else
        If FRMCUST_CC_MGM.TDBMaskAOffice(0).Value = "021" Or FRMCUST_CC_MGM.TDBMaskAOffice(0).Value = Empty Then
            NOTELP = MDIForm1.TxtAuthPrefix.Text & MDIForm1.TxtAuth.Text & MDIForm1.TxtModemAcod.Text & TELPKANTOR1
            Call Dial(GetNumber(NOTELP))
            MDIForm1.m_TelpNoTelp = TELPKANTOR1
            MDIForm1.m_TelpUserId = FRMCUST_CC_MGM.Text1(1).Text
        Else
            NOTELP = MDIForm1.TxtAuthPrefix.Text & MDIForm1.TxtAuth.Text & MDIForm1.TxtModemAcod.Text & FRMCUST_CC_MGM.TDBMaskAOffice(0).Value & TELPKANTOR1
            Call Dial(GetNumber(NOTELP))
            MDIForm1.m_TelpNoTelp = FRMCUST_CC_MGM.TDBMaskAOffice(0).Value & TELPKANTOR1
            MDIForm1.m_TelpUserId = FRMCUST_CC_MGM.Text1(1).Text
        End If
    End If
Case 1
    If Len(TELPKANTOR2) < 3 Then
        Me.MousePointer = 0
        MsgBox "No Telepon Tidak Valid", vbCritical + vbOKOnly, "TeleGrandi"
    Else
        If FRMCUST_CC_MGM.TDBMaskAOffice(1).Value = "021" Or FRMCUST_CC_MGM.TDBMaskAOffice(1).Value = Empty Then
            NOTELP = MDIForm1.TxtAuthPrefix.Text & MDIForm1.TxtAuth.Text & MDIForm1.TxtModemAcod.Text & TELPKANTOR2
            Call Dial(GetNumber(NOTELP))
            MDIForm1.m_TelpNoTelp = TELPKANTOR2
            MDIForm1.m_TelpUserId = FRMCUST_CC_MGM.Text1(1).Text
        Else
            NOTELP = MDIForm1.TxtAuthPrefix.Text & MDIForm1.TxtAuth.Text & MDIForm1.TxtModemAcod.Text & FRMCUST_CC_MGM.TDBMaskAOffice(1).Value & TELPKANTOR2
            Call Dial(GetNumber(NOTELP))
            MDIForm1.m_TelpNoTelp = FRMCUST_CC_MGM.TDBMaskAOffice(1).Value & TELPKANTOR2
            MDIForm1.m_TelpUserId = FRMCUST_CC_MGM.Text1(1).Text
        End If
    End If
Case 2
    If Len(HANDPHONE1) < 2 Then
        Me.MousePointer = 0
        MsgBox "No Telepon Tidak Valid", vbCritical + vbOKOnly, "TeleGrandi"
    Else
        NOTELP = MDIForm1.TxtAuthPrefix.Text & MDIForm1.TxtAuth.Text & MDIForm1.TxtModemAcod.Text & HANDPHONE1
        Call Dial(GetNumber(NOTELP))
            MDIForm1.m_TelpNoTelp = HANDPHONE1
            MDIForm1.m_TelpUserId = FRMCUST_CC_MGM.Text1(1).Text
    End If
Case 3
    If Len(HANDPHONE2) < 3 Then
        Me.MousePointer = 0
        MsgBox "No Telepon Tidak Valid", vbCritical + vbOKOnly, "TeleGrandi"
    Else
        NOTELP = MDIForm1.TxtAuthPrefix.Text & MDIForm1.TxtAuth.Text & MDIForm1.TxtModemAcod.Text & HANDPHONE2
        Call Dial(GetNumber(NOTELP))
            MDIForm1.m_TelpNoTelp = HANDPHONE2
            MDIForm1.m_TelpUserId = FRMCUST_CC_MGM.Text1(1).Text
    End If
Case 4
    If Len(TELPRUMAH1) < 3 Then
        Me.MousePointer = 0
        MsgBox "No Telepon Tidak Valid", vbCritical + vbOKOnly, "TeleGrandi"
    Else
        If FRMCUST_CC_MGM.TDBMaskAHome(0).Value = "021" Or FRMCUST_CC_MGM.TDBMaskAHome(0).Value = Empty Then
            NOTELP = MDIForm1.TxtAuthPrefix.Text & MDIForm1.TxtAuth.Text & MDIForm1.TxtModemAcod.Text & TELPRUMAH1
            Call Dial(GetNumber(NOTELP))
            MDIForm1.m_TelpNoTelp = TELPRUMAH1
            MDIForm1.m_TelpUserId = FRMCUST_CC_MGM.Text1(1).Text
        Else
            NOTELP = MDIForm1.TxtAuthPrefix.Text & MDIForm1.TxtAuth.Text & MDIForm1.TxtModemAcod.Text & FRMCUST_CC_MGM.TDBMaskAHome(0).Value & TELPRUMAH1
            Call Dial(GetNumber(NOTELP))
            MDIForm1.m_TelpNoTelp = FRMCUST_CC_MGM.TDBMaskAHome(0).Value & TELPRUMAH1
            MDIForm1.m_TelpUserId = FRMCUST_CC_MGM.Text1(1).Text
        End If
    End If
Case 5
    If Len(TELPRUMAH2) < 3 Then
        Me.MousePointer = 0
        MsgBox "No Telepon Tidak Valid", vbCritical + vbOKOnly, "TeleGrandi"
    Else
        If FRMCUST_CC_MGM.TDBMaskAHome(1).Value = "021" Or FRMCUST_CC_MGM.TDBMaskAHome(1).Value = Empty Then
            NOTELP = MDIForm1.TxtAuthPrefix.Text & MDIForm1.TxtAuth.Text & MDIForm1.TxtModemAcod.Text & TELPRUMAH2
            Call Dial(GetNumber(NOTELP))
            MDIForm1.m_TelpNoTelp = TELPRUMAH2
            MDIForm1.m_TelpUserId = FRMCUST_CC_MGM.Text1(1).Text
        Else
            NOTELP = MDIForm1.TxtAuthPrefix.Text & MDIForm1.TxtAuth.Text & MDIForm1.TxtModemAcod.Text & FRMCUST_CC_MGM.TDBMaskAHome(1).Value & TELPRUMAH2
            Call Dial(GetNumber(NOTELP))
            MDIForm1.m_TelpNoTelp = FRMCUST_CC_MGM.TDBMaskAHome(1).Value & TELPRUMAH2
            MDIForm1.m_TelpUserId = FRMCUST_CC_MGM.Text1(1).Text
        End If
    End If
End Select
End Sub

Private Sub Dial(Number$)
Dim M_TELP As ADODB.Recordset
Dim cmdsql As String
Dim DialString$, FromModem$, dummy
    DialString$ = "ATDT" + Number$ + ";" + vbCr
    On Error Resume Next
    If MSComm1.PortOpen Then
    Else
        If MDIForm1.TxtCommPort.Text = Empty Then
            MsgBox "Tidak Ada Variable buat Comport", vbInformation + vbOKOnly
            Exit Sub
        End If
        MSComm1.CommPort = MDIForm1.TxtCommPort.Text
        MSComm1.Settings = "9600,N,8,1"
        MSComm1.PortOpen = True
    End If
Me.MousePointer = 11
If MDIForm1.Text6.Text <> Empty Then
    WaitSecs (CCur(MDIForm1.Text6.Text))
Else
    WaitSecs (0)
End If
    If Err Then
        MsgBox Err.Description, vbCritical + vbOKOnly, "TeleGrandi"
        MSComm1.PortOpen = False
        cancelflag = True
        Me.MousePointer = 0
        Exit Sub
    End If
    MSComm1.InBufferCount = 0
    MSComm1.Output = DialString$
    Me.MousePointer = 0
    Do
        dummy = DoEvents()
        If MSComm1.InBufferCount Then
            FromModem$ = FromModem$ + MSComm1.Input
            If InStr(FromModem$, "OK") Then
                Beep
                MsgBox "Angkat Telephone Kemudian Click OK", vbInformation + vbOKOnly, "TeleGrandi"
                Me.Hide
                cmdsql = "Insert Into TblPhoneMonitorHst(UserId, CustId, NamaCh,StartDate, EndDate, TelpNo, Recsource) Values ('" + MDIForm1.Text1.Text + "' , '" + FRMCUST_CC_MGM.Text1(1).Text + "','" + FRMCUST_CC_MGM.Text1(0).Text + "', '" + Format(CStr(MDIForm1.TDBDate1.Value), "mm/dd/yyyy") & " " & Format(Now, "hh:nn") + "' , ' ', '" + MDIForm1.m_TelpNoTelp + "','" + FRMCUST_CC_MGM.Combo1(0).Text + "')"
                M_OBJCONN.Execute cmdsql
                
                cmdsql = " INSERT INTO PHONENO_CALL"
                cmdsql = cmdsql + " (TGL,"
                cmdsql = cmdsql + " AGENT,"
                cmdsql = cmdsql + " CUSTID,"
                cmdsql = cmdsql + " RECSOURCE,"
                cmdsql = cmdsql + " NOTELP)"
                cmdsql = cmdsql + " VALUES"
                cmdsql = cmdsql + " ('" + Format(MDIForm1.TDBDate1.Text, "MM/DD/YY") & " " & Format(Now, "HH:MM") + "',"
                cmdsql = cmdsql + " '" + MDIForm1.Text1.Text + "',"
                cmdsql = cmdsql + " '" + FRMCUST_CC_MGM.Text1(0).Text + "',"
                cmdsql = cmdsql + " '" + FRMCUST_CC_MGM.Combo1(0).Text + "',"
                cmdsql = cmdsql + " '" + CStr(Number) + "')"
                WaitSecs (0.05)
                'M_OBJCONN.Execute cmdsql
        '        frmtelp_mgm.SUKSESPHONE = True
         '       FRMCUST_CC_MGM.SSCommand1(2).Enabled = True
                cancelflag = True
                Exit Do
            End If
            If InStr(FromModem$, "NO DIALTONE") Then
                Beep
                Beep
                MsgBox Err.Description, vbInformation + vbOKOnly, "TeleGrandi"
                cancelflag = True
                Exit Do
            End If
        End If
        If cancelflag Then
            cancelflag = False
            Me.MousePointer = 0
            Exit Do
        End If
    Loop
    If MSComm1.PortOpen = True And cancelflag = True Then
        MSComm1.Output = "ATH" + vbCr
        MSComm1.PortOpen = False
    End If
    Me.MousePointer = 0
    Unload Me
End Sub
   
Private Sub Command2_Click(Index As Integer)
Select Case Index
    Case 1
        cancelflag = True
    Case 2
        cancelflag = True
        If MSComm1.PortOpen Then
            MSComm1.PortOpen = False
        End If
        Unload Me
End Select
End Sub

