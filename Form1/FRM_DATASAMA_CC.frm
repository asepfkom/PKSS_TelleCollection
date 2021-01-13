VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRM_DATASAMA_CC 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   Icon            =   "FRM_DATASAMA_CC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   11670
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Tutup"
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
      Left            =   10830
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   11640
      Begin MSComctlLib.ListView ListView1 
         Height          =   6315
         Left            =   30
         TabIndex        =   1
         Top             =   135
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   11139
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   12582912
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
End
Attribute VB_Name = "FRM_DATASAMA_CC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim M_DATA As New CLS_FRMCUST_CC
Dim m_objrs As ADODB.Recordset
Dim NAMA1 As String
Dim NAMA2 As String
Dim LISTITEM As LISTITEM
Dim HP As String
Dim HP2 As String
Dim CMDSQL As String
Call header
'cari buat add nama ada hampir sama
    If NAMA_SAMA Then
        NAMA1 = GetNamaNoSpace(FRMCUST_CC.Text1(0).Text)
        Set m_objrs = M_DATA.QUERY_CUST(M_OBJCONN, "")
            While Not m_objrs.EOF
                NAMA2 = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
                
                If Len(NAMA2) <> 0 Then
                    NAMA2 = GetNamaNoSpace(NAMA2)
                End If
                If Len(NAMA2) <> 0 Then
                    If NAMA1 = NAMA2 Then
                        Set LISTITEM = ListView1.ListItems.ADD(, , m_objrs("CUSTID"))
                                LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
                                LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("ADDRNOW")), "", m_objrs("ADDRNOW"))
                                If IsNull(m_objrs("BIRTHD")) Then
                                    LISTITEM.SubItems(3) = " "
                                Else
                                    LISTITEM.SubItems(3) = Right(m_objrs("BIRTHD"), 2) + "/" + Mid(m_objrs("BIRTHD"), 5, 2) + "/" + Left(m_objrs("BIRTHD"), 4)
                                End If
                                LISTITEM.SubItems(4) = IIf(IsNull(m_objrs("HOMENO")), "", m_objrs("HOMENO"))
                                LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("HOMENO2")), "", m_objrs("HOMENO2"))
                                LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("OFFICENO")), "", m_objrs("OFFICENO"))
                                LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("OFFICENO2")), "", m_objrs("OFFICENO2"))
                                LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("MOBILENO")), "", m_objrs("MOBILENO"))
                                LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("MOBILENO2")), "", m_objrs("MOBILENO2"))
                                LISTITEM.SubItems(10) = IIf(IsNull(m_objrs("FAXNO")), "", m_objrs("FAXNO"))
                                LISTITEM.SubItems(11) = IIf(IsNull(m_objrs("FAXNO2")), "", m_objrs("FAXNO2"))
                                LISTITEM.SubItems(12) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
                                LISTITEM.SubItems(13) = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
                    End If
                End If
                m_objrs.MoveNext
            Wend
        Set m_objrs = Nothing
        Set M_DATA = Nothing
    Exit Sub
    End If
'cari update telp ada yang sama
            If update_TELP_SAMA Then
                With FRMCUST_CC
                    If .TDBMask1(0).ReadOnly = False Then
                        If Len(.TDBMask1(0).Value) > 4 Then
                            CMDSQL = "(HOMENO = '" + .TDBMask1(0).Value + "'"
                            CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(0).Value + "'"
                            CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(0).Value + "'"
                            CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(0).Value + "'"
                            CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(0).Value + "'"
                            CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(0).Value + "'"
                            CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(0).Value + "'"
                            CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(0).Value + "'"
                        End If
                    End If
                    If .TDBMask1(1).ReadOnly = False Then
                        If Len(.TDBMask1(1).Value) > 4 Then
                            If CMDSQL = Empty Then
                                CMDSQL = CMDSQL + " (HOMENO = '" + .TDBMask1(1).Value + "'"
                            Else
                                CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(1).Value + "'"
                            End If
                                CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(1).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(1).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(1).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(1).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(1).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(1).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(1).Value + "'"
                        End If
                    End If
                    If .TDBMask1(2).ReadOnly = False Then
                        If Len(.TDBMask1(2).Value) > 4 Then
                            If CMDSQL = Empty Then
                                CMDSQL = CMDSQL + " (HOMENO = '" + .TDBMask1(2).Value + "'"
                            Else
                                CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(2).Value + "'"
                            End If
                                CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(2).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(2).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(2).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(2).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(2).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(2).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(2).Value + "'"
                        End If
                    End If
                    If .TDBMask1(3).ReadOnly = False Then
                        If Len(.TDBMask1(3).Value) > 4 Then
                            If CMDSQL = Empty Then
                                CMDSQL = CMDSQL + " (HOMENO = '" + .TDBMask1(3).Value + "'"
                            Else
                                CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(3).Value + "'"
                            End If
                                CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(3).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(3).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(3).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(3).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(3).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(3).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(3).Value + "'"
                        End If
                    End If
                    If .TDBMask1(4).ReadOnly = False Then
                        If Len(.TDBMask1(4).Value) > 4 Then
                            If CMDSQL = Empty Then
                                CMDSQL = CMDSQL + " (HOMENO = '" + .TDBMask1(4).Value + "'"
                            Else
                                CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(4).Value + "'"
                            End If
                                CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(4).Value + "'"
                                CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(4).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(4).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(4).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(4).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(4).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(4).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(4).Value + "'"
                        End If
                    End If
                    If .TDBMask1(5).ReadOnly = False Then
                        If Len(.TDBMask1(5).Value) > 4 Then
                            If CMDSQL = Empty Then
                                CMDSQL = CMDSQL + " (HOMENO = '" + .TDBMask1(5).Value + "'"
                            Else
                                CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(5).Value + "'"
                            End If
                                CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(5).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(5).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(5).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(5).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(5).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(5).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(5).Value + "'"
                        End If
                    End If
                    If .TDBMask1(6).ReadOnly = False Then
                        If Len(.TDBMask1(6).Value) > 4 Then
                            If CMDSQL = Empty Then
                                CMDSQL = CMDSQL + " (HOMENO = '" + .TDBMask1(6).Value + "'"
                            Else
                                CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(6).Value + "'"
                            End If
                                CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(6).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(6).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(6).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(6).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(6).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(6).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(6).Value + "'"
                        End If
                    End If
                    If .TDBMask1(7).ReadOnly = False Then
                        If Len(.TDBMask1(7).Value) > 4 Then
                            If CMDSQL = Empty Then
                                CMDSQL = CMDSQL + " (HOMENO = '" + .TDBMask1(7).Value + "'"
                            Else
                                CMDSQL = CMDSQL + " or HOMENO = '" + .TDBMask1(7).Value + "'"
                            End If
                                CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(7).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(7).Value + "'"
                                CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(7).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(7).Value + "'"
                                CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(7).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(7).Value + "'"
                                CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(7).Value + "'"
                        End If
                    End If
                    If Len(CMDSQL) <> 0 Then
                        CMDSQL = CMDSQL + ") AND CUSTID <> '" + .Text1(1).Text + "'"
                    End If
                End With
            Set m_objrs = M_DATA.QUERY_CEK_ADDCUST(M_OBJCONN, CMDSQL)
            Me.Caption = "No Telp Yang Ditambah Ada Yang Sama"
                
            While Not m_objrs.EOF
                Set LISTITEM = ListView1.ListItems.ADD(, , m_objrs("CUSTID"))
                        LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
                        LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("ADDRNOW")), "", m_objrs("ADDRNOW"))
                        If IsNull(m_objrs("BIRTHD")) Then
                            LISTITEM.SubItems(3) = " "
                        Else
                            LISTITEM.SubItems(3) = Format(m_objrs("BIRTHD"), "dd-mmm-yyyy")
                        End If
                        LISTITEM.SubItems(4) = IIf(IsNull(m_objrs("HOMENO")), "", m_objrs("HOMENO"))
                        LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("HOMENO2")), "", m_objrs("HOMENO2"))
                        LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("OFFICENO")), "", m_objrs("OFFICENO"))
                        LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("OFFICENO2")), "", m_objrs("OFFICENO2"))
                        LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("MOBILENO")), "", m_objrs("MOBILENO"))
                        LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("MOBILENO2")), "", m_objrs("MOBILENO2"))
                        LISTITEM.SubItems(10) = IIf(IsNull(m_objrs("FAXNO")), "", m_objrs("FAXNO"))
                        LISTITEM.SubItems(11) = IIf(IsNull(m_objrs("FAXNO2")), "", m_objrs("FAXNO2"))
                        LISTITEM.SubItems(12) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
                        LISTITEM.SubItems(13) = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
                    m_objrs.MoveNext
            Wend
                Exit Sub
            End If
'cari buat add telpon sama
                        If TELP_SAMA Then
                            With FRMCUST_CC
                                    If Len(.TDBMask1(0).Value) < 5 Then
                                Else
                                    CMDSQL = "HOMENO = '" + .TDBMask1(0).Value + "'"
                                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(0).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(0).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(0).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(0).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(0).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(0).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(0).Value + "'"
                                End If
                                If Len(.TDBMask1(1).Value) < 5 Then
                                Else
                                    If CMDSQL = Empty Then
                                        CMDSQL = CMDSQL + " HOMENO = '" + .TDBMask1(1).Value + "'"
                                    Else
                                        CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(1).Value + "'"
                                    End If
                                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(1).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(1).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(1).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(1).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(1).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(1).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(1).Value + "'"
                                End If
                                If Len(.TDBMask1(2).Value) < 5 Then
                                Else
                                    If CMDSQL = Empty Then
                                        CMDSQL = CMDSQL + " HOMENO = '" + .TDBMask1(2).Value + "'"
                                    Else
                                        CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(2).Value + "'"
                                    End If
                                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(2).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(2).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(2).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(2).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(2).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(2).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(2).Value + "'"
                                End If
                                If Len(.TDBMask1(3).Value) = Empty Then
                                Else
                                    If CMDSQL = Empty Then
                                        CMDSQL = CMDSQL + " HOMENO = '" + .TDBMask1(3).Value + "'"
                                    Else
                                        CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(3).Value + "'"
                                    End If
                                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(3).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(3).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(3).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(3).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(3).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(3).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(3).Value + "'"
                                End If
                                If Len(.TDBMask1(4).Value) < 5 Then
                                Else
                                    If CMDSQL = Empty Then
                                        CMDSQL = CMDSQL + " HOMENO = '" + .TDBMask1(4).Value + "'"
                                    Else
                                        CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(4).Value + "'"
                                    End If
                                    CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(4).Value + "'"
                                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(4).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(4).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(4).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(4).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(4).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(4).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(4).Value + "'"
                                End If
                                
                                If Len(.TDBMask1(5).Value) < 5 Then
                                Else
                                    If CMDSQL = Empty Then
                                        CMDSQL = CMDSQL + " HOMENO = '" + .TDBMask1(5).Value + "'"
                                    Else
                                        CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(5).Value + "'"
                                    End If
                                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(5).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(5).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(5).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(5).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(5).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(5).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(5).Value + "'"
                                End If
                                
                                If Len(.TDBMask1(6).Value) < 5 Then
                                Else
                                    If CMDSQL = Empty Then
                                        CMDSQL = CMDSQL + " HOMENO = '" + .TDBMask1(6).Value + "'"
                                    Else
                                        CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(6).Value + "'"
                                    End If
                                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(6).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(6).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(6).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(6).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(6).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(6).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(6).Value + "'"
                                End If
                                If Len(.TDBMask1(7).Value) < 5 Then
                                Else
                                    If CMDSQL = Empty Then
                                        CMDSQL = CMDSQL + " HOMENO = '" + .TDBMask1(7).Value + "'"
                                    Else
                                        CMDSQL = CMDSQL + " OR HOMENO = '" + .TDBMask1(7).Value + "'"
                                    End If
                                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + .TDBMask1(7).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO = '" + .TDBMask1(7).Value + "'"
                                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + .TDBMask1(7).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO = '" + .TDBMask1(7).Value + "'"
                                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + .TDBMask1(7).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO = '" + .TDBMask1(7).Value + "'"
                                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + .TDBMask1(7).Value + "'"
                                End If
                                If Len(CMDSQL) > 0 Then
                                    CMDSQL = CMDSQL + " AND (LEFT(RECSTATUS,1)<>'0')"
                                End If
                            End With
                            Set m_objrs = M_DATA.QUERY_CEK_ADDCUST(M_OBJCONN, CMDSQL)
                            Me.Caption = "Nomor Telepone Ada Yang Sama"
                        Else
' add nama ada yang sama
                        With FRMCUST_CC
                            Set m_objrs = M_DATA.QUERY_CEK_ADDCUST(M_OBJCONN, "NAME = '" + .Text1(0).Text + "'AND LEFT(RECSTATUS,1)<>'0'")
                            Me.Caption = "Nama Ada Yang Sama"
                        End With
                        End If
' add tgllahir ada yang sama
    If TGLLHR_SAMA Then
        With FRMCUST_CC
            Set m_objrs = M_DATA.QUERY_CEK_ADDCUST(M_OBJCONN, "BIRTHD = '" + Format(.TDBDate1(0).Value, "mm/dd/yy") + "'AND LEFT(RECSTATUS,1)<>'0'")
            Me.Caption = "Tanggal Lahir Ada Yang Sama"
        End With
    End If
    While Not m_objrs.EOF
        Set LISTITEM = ListView1.ListItems.ADD(, , m_objrs("CUSTID"))
                LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
                LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("ADDRNOW")), "", m_objrs("ADDRNOW"))
                If IsNull(m_objrs("BIRTHD")) Then
                    LISTITEM.SubItems(3) = " "
                Else
                    LISTITEM.SubItems(3) = Format(m_objrs("BIRTHD"), "dd-mmm-yyyy")
                End If
                LISTITEM.SubItems(4) = IIf(IsNull(m_objrs("HOMENO")), "", m_objrs("HOMENO"))
                LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("HOMENO2")), "", m_objrs("HOMENO2"))
                LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("OFFICENO")), "", m_objrs("OFFICENO"))
                LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("OFFICENO2")), "", m_objrs("OFFICENO2"))
                LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("MOBILENO")), "", m_objrs("MOBILENO"))
                LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("MOBILENO2")), "", m_objrs("MOBILENO2"))
                LISTITEM.SubItems(10) = IIf(IsNull(m_objrs("FAXNO")), "", m_objrs("FAXNO"))
                LISTITEM.SubItems(11) = IIf(IsNull(m_objrs("FAXNO2")), "", m_objrs("FAXNO2"))
                LISTITEM.SubItems(12) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
                LISTITEM.SubItems(13) = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
            m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Set M_DATA = Nothing
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Customers Id", 15 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Customers Name", 40 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Alamat", 50 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Tanggal Lahir", 15 * TXT
    ListView1.ColumnHeaders.ADD 5, , "No. Telephone", 15 * TXT
    ListView1.ColumnHeaders.ADD 6, , "No. Telephone2", 15 * TXT
    ListView1.ColumnHeaders.ADD 7, , "No. Telp. Kantor", 15 * TXT
    ListView1.ColumnHeaders.ADD 8, , "No. Telp. Kantor2", 15 * TXT
    ListView1.ColumnHeaders.ADD 9, , "No. Mobile", 18 * TXT
    ListView1.ColumnHeaders.ADD 10, , "No. Mobile2", 18 * TXT
    ListView1.ColumnHeaders.ADD 11, , "No. Fax", 18 * TXT
    ListView1.ColumnHeaders.ADD 12, , "No. Fax2", 18 * TXT
    ListView1.ColumnHeaders.ADD 13, , "Data Source", 15 * TXT
     ListView1.ColumnHeaders.ADD 14, , "Agent Name", 50 * TXT
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub

