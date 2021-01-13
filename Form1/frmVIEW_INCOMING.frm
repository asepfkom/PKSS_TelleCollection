VERSION 5.00
Begin VB.Form frmVIEW_INCOMING 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000004&
      Caption         =   "Tampilkan Berdasarkan.."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   165
      TabIndex        =   2
      Top             =   1035
      Width           =   2565
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
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
      Height          =   420
      Left            =   3000
      TabIndex        =   7
      Top             =   3420
      UseMaskColor    =   -1  'True
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000004&
      Height          =   2190
      Left            =   60
      TabIndex        =   12
      Top             =   1095
      Width           =   4230
      Begin VB.ComboBox Combo1 
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
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   375
         Width           =   2985
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   375
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
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
         Index           =   3
         Left            =   1200
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   705
         Width           =   2985
      End
      Begin VB.ComboBox Combo1 
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
         Index           =   5
         Left            =   1200
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   1020
         Width           =   2985
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tampilkan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   570
         TabIndex        =   6
         Top             =   1530
         Width           =   3180
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   2
         Left            =   1200
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   705
         Visible         =   0   'False
         Width           =   2910
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   4
         Left            =   1200
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   1020
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Sumber Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   420
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Sales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   14
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Caption         =   "Supervisor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   165
         TabIndex        =   13
         Top             =   1080
         Width           =   900
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000004&
      Caption         =   "Tampilkan Semua..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   240
      Width           =   2145
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   60
      TabIndex        =   11
      Top             =   285
      Width           =   4230
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tampilkan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   585
         TabIndex        =   1
         Top             =   225
         Width           =   3180
      End
   End
End
Attribute VB_Name = "frmVIEW_INCOMING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public HEADER_JUDUL As String

Private Sub Combo1_Click(Index As Integer)
Dim M_DATA As New VIEW
Dim m_objrs As ADODB.Recordset

Select Case Index
    Case 1
        Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("KODEDS")
            Combo1(1).Text = m_objrs("KETERANGAN")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 3
        Set m_objrs = M_DATA.QUERY_AGENT(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(2).Text = m_objrs("USERID")
            Combo1(3).Text = m_objrs("AGENT")
        Else
            Combo1(2).Text = Empty
            Combo1(3).Text = Empty
        End If
    Case 5
        Set m_objrs = M_DATA.QUERY_SPV(M_OBJCONN, "SPVNAME = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(4).Text = m_objrs("SPVCODE")
            Combo1(5).Text = m_objrs("SPVNAME")
        Else
            Combo1(4).Text = Empty
            Combo1(5).Text = Empty
        End If
End Select
Set M_DATA = Nothing
Set m_objrs = Nothing
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim M_DATA As New VIEW
Dim m_objrs As ADODB.Recordset

Select Case Index
    Case 1
        Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("KODEDS")
            Combo1(1).Text = m_objrs("KETERANGAN")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 3
        Set m_objrs = M_DATA.QUERY_AGENT(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(2).Text = m_objrs("USERID")
            Combo1(3).Text = m_objrs("AGENT")
        Else
            Combo1(2).Text = Empty
            Combo1(3).Text = Empty
        End If
    Case 5
        Set m_objrs = M_DATA.QUERY_SPV(M_OBJCONN, "SPVNAME = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(4).Text = m_objrs("SPVCODE")
            Combo1(5).Text = m_objrs("SPVNAME")
        Else
            Combo1(4).Text = Empty
            Combo1(5).Text = Empty
        End If
End Select
Set M_DATA = Nothing
Set m_objrs = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New VIEW
Dim M_KONDISI As String

Select Case Index
    Case 3
        Me.Hide
        With VIEWCUSTINCOMING
            HEADER_JUDUL = Command1(Index).Caption
            .Show vbModal
        End With
    Case 4
        If Combo1(0).Text = Empty And Combo1(2).Text = Empty And Combo1(4).Text = Empty Then
            MsgBox "Masukan Kriteria Yang Akan Di cari...", vbCritical + vbOKOnly, "TeleGrandi"
        Else
        Me.Hide
        VIEWCUSTINCOMING.Show vbModal
        End If
    
End Select
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New VIEW
Dim M_OBJ As Object
    

    Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "")
    While Not m_objrs.EOF
        Combo1(0).AddItem m_objrs("KODEDS")
        Combo1(0).DataField = m_objrs("KODEDS")
        Combo1(1).AddItem m_objrs("KETERANGAN")
        Combo1(1).DataField = m_objrs("KETERANGAN")
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    
    Set m_objrs = M_DATA.QUERY_AGENT(M_OBJCONN, "")
    While Not m_objrs.EOF
        Combo1(2).AddItem m_objrs("USERID")
        Combo1(2).DataField = m_objrs("USERID")
        Combo1(3).AddItem m_objrs("AGENT")
        Combo1(3).DataField = m_objrs("AGENT")
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    
    Set m_objrs = M_DATA.QUERY_SPV(M_OBJCONN, "")
    While Not m_objrs.EOF
        Combo1(4).AddItem m_objrs("SPVCODE")
        Combo1(4).DataField = m_objrs("SPVCODE")
        Combo1(5).AddItem m_objrs("SPVNAME")
        Combo1(5).DataField = m_objrs("SPVNAME")
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Option1(0).Value = True
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
    Case 0
        Option1(1).Value = False
        Frame2.Enabled = True
        Frame3.Enabled = False
        Command1(3).Enabled = True
        Command1(4).Enabled = False
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
        Combo1(4).Text = Empty
        Combo1(5).Text = Empty
    Case 1
        Option1(0).Value = False
        Frame2.Enabled = False
        Frame3.Enabled = True
        Command1(3).Enabled = False
        Command1(4).Enabled = True
End Select
End Sub
