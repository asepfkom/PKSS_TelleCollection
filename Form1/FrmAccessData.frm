VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmAccessData 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7365
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   6135
   ControlBox      =   0   'False
   Icon            =   "FrmAccessData.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "WO DATE"
      Height          =   3015
      Left            =   210
      TabIndex        =   19
      Top             =   3570
      Width           =   5655
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         Height          =   2595
         ItemData        =   "FrmAccessData.frx":1CFA
         Left            =   1440
         List            =   "FrmAccessData.frx":1D2B
         MultiSelect     =   2  'Extended
         TabIndex        =   21
         Top             =   270
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "WO DATE"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1931
      _Version        =   196610
      BackColor       =   14737632
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   3120
         TabIndex        =   12
         Top             =   600
         Width           =   2655
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "All TeleCollection"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   11
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "Pilih TeleCollection"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "Pilih SPV Name"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "TeleCollection Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2295
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4048
      _Version        =   196610
      BackColor       =   14737632
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SK - SKIP"
         Height          =   255
         Index           =   12
         Left            =   3030
         TabIndex        =   25
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON - On Nego"
         Height          =   255
         Index           =   11
         Left            =   90
         TabIndex        =   24
         Top             =   1590
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OS - On Process"
         Height          =   255
         Index           =   10
         Left            =   3030
         TabIndex        =   23
         Top             =   1590
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data Blank"
         Height          =   255
         Index           =   8
         Left            =   3030
         TabIndex        =   22
         Top             =   1290
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "S P - Settled Payment"
         Height          =   255
         Index           =   5
         Left            =   3030
         TabIndex        =   18
         Top             =   990
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "RP - Refuse Payment"
         Height          =   255
         Index           =   7
         Left            =   135
         TabIndex        =   17
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "VL - VALID"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   150
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PR- PROSPECT"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "P T P - Promise To Pay"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1230
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B P - Broken Promise"
         Height          =   255
         Index           =   3
         Left            =   3030
         TabIndex        =   5
         Top             =   120
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "P O P - Progress Of Payment"
         Height          =   255
         Index           =   4
         Left            =   3030
         TabIndex        =   4
         Top             =   390
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Uncontacted"
         Height          =   255
         Index           =   6
         Left            =   3030
         TabIndex        =   3
         Top             =   690
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   6660
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Execute"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   6660
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Release"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   15
      Top             =   6660
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "E&xit"
      ButtonStyle     =   3
   End
End
Attribute VB_Name = "FrmAccessData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StsVl As String
Dim StsPR As String
Dim StsON As String
Dim StsSK As String
Dim StsOS As String
Dim StsPTP As String
Dim StsBP As String
Dim StsPOP As String
Dim StsSP As String
Dim StsUC As String
Dim StsRP As String
Dim StsFresh As String
Dim StsWO_Date As String
Dim StsWO_2009 As String
Dim StsWO_2008 As String
Dim StsWO_2007 As String
Dim StsWO_2006 As String
Dim StsWO_2005 As String
Dim StsWO_2004 As String
Dim StsWO_2003 As String
Dim StsWO_2002 As String
Dim StsWO_2001 As String
Dim StsWO_2000 As String
Dim StsWO_1999 As String
Dim StsWO_2010 As String
Dim Stsblank As String
Dim cmdsql As String
Dim spv As Boolean



Private Sub Check2_Click()
If Check2.Value = vbChecked Then
    List1.Enabled = True
Else
    List1.Enabled = False
End If
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim M_DATA As New CLS_FRMSEARCH
Dim M_objrs As ADODB.Recordset
Select Case Index
Case 0
    If spv = False Then
        Set M_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("USERID")
            Combo1(1).Text = M_objrs("AGENT")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Else
        Set M_objrs = New ADODB.Recordset
        M_objrs.Open "select * from SPVTBL where SPVCODE='" + Combo1(0) + "'", M_OBJCONN, adOpenDynamic, adLockBatchOptimistic
            While Not M_objrs.EOF
                Combo1(0).Text = M_objrs("SPVCODE")
                Combo1(1).Text = M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
    End If
Case 1
    If spv = False Then
        Set M_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("USERID")
            Combo1(1).Text = M_objrs("AGENT")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Else
        Set M_objrs = New ADODB.Recordset
        M_objrs.Open "select * from SPVTBL where SPVNAME='" + Combo1(1) + "'", M_OBJCONN, adOpenDynamic, adLockBatchOptimistic
            While Not M_objrs.EOF
                Combo1(0).Text = M_objrs("SPVCODE")
                Combo1(1).Text = M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
    End If
    
 End Select
 
 Set M_DATA = Nothing
Set M_objrs = Nothing
End Sub

Private Sub Command1_Click()
Dim i As Integer
If List1.ListIndex = -1 Then Exit Sub
Text1.Text = ""
For i = List1.ListCount - 1 To 0 Step -1
     If List1.Selected(i) = True Then
        Select Case List1.list(i)
         Case "2009"
            StsWO_2009 = "2009"
            Text1.Text = List1.list(i)
         Case "2008"
            StsWO_2008 = "2008"
            Text1.Text = List1.list(i)
         Case "2007"
            StsWO_2007 = "2007"
         Case "2006"
            StsWO_2006 = "2006"
         Case "2005"
            StsWO_2005 = "2005"
         Case "2004"
            StsWO_2004 = "2004"
         Case "2003"
            StsWO_2003 = "2003"
         Case "2002"
            StsWO_2002 = "2002"
         Case "2001"
            StsWO_2001 = "2001"
         Case "2000"
            StsWO_2000 = "2000"
         Case "1999"
            StsWO_1999 = "1999"
        End Select
        
    ' Text1.Text = Text1.Text + "'" + List1.LIST(i) + "',"
     Else
End If
Next i
 'Text1.Text = Right(Text1.Text, 1)
End Sub

Private Sub Form_Load()
Dim M_objrs As ADODB.Recordset
Dim M_DATA As New CLS_FRMSEARCH

Set M_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "")
    While Not M_objrs.EOF
        Combo1(0).AddItem M_objrs("USERID")
        Combo1(1).AddItem M_objrs("AGENT")
        M_objrs.MoveNext
    Wend
Set M_objrs = Nothing
SSOption1(0).Value = True
spv = False


End Sub


Private Sub SSCommand1_Click(Index As Integer)
Dim M_objrs As New ADODB.Recordset
'Dim CMDSQL As String

Select Case Index
Case 0
        If SSOption1(0).Value = False And SSOption1(1).Value = False And SSOption1(2).Value = False Then
            MsgBox "Select DCR Name To Proccess OR All"
         Else
         
         
                If SSOption1(0).Value Then
                   Set M_objrs = New ADODB.Recordset
                   M_objrs.CursorLocation = adUseClient
                   M_objrs.Open "select * from usertbl where lockdarispv  is not null limit 1", M_OBJCONN, adOpenDynamic, adLockOptimistic
                   
                   If Not M_objrs.EOF Then
                       If Check1(0).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intNa = InStr(1, vrsumber, "VL-", vbTextCompare)
                            If intNa > 0 Then StrNa = "[VL]"
                       End If
                       
                       If Check1(7).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intRP = InStr(1, vrsumber, "RP-", vbTextCompare)
                            If intRP > 0 Then strRp = "[Rp]"
                       End If
                       
                       If Check1(1).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intOp = InStr(1, vrsumber, "PR-", vbTextCompare)
                            If intOp > 0 Then strOp = "[PR]"
                       End If
                       
                       
                       If Check1(2).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intPTP = InStr(1, vrsumber, "PTP", vbTextCompare)
                            If intPTP > 0 Then strPtp = "[PTP]"
                       End If
                       
                       If Check1(4).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intpop = InStr(1, vrsumber, "POP", vbTextCompare)
                            If intpop > 0 Then strpop = "[POP]"
                       End If
                       
                       If Check1(3).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intBP = InStr(1, vrsumber, "BP-", vbTextCompare)
                            If intpop > 0 Then strbp = "[BP]"
                       End If
                       
                       If Check1(5).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intSp = InStr(1, vrsumber, "SP-", vbTextCompare)
                            If intSp > 0 Then strSp = "[SP]"
                       End If
                       
                       If Check1(10).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intSp = InStr(1, vrsumber, "OS-", vbTextCompare)
                            If intSp > 0 Then strSp = "[OS]"
                       End If
                       
                     
                       If Check1(11).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intSp = InStr(1, vrsumber, "ON-", vbTextCompare)
                            If intSp > 0 Then strSp = "[ON]"
                       End If
                       
                        If Check1(12).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intSp = InStr(1, vrsumber, "SK-", vbTextCompare)
                            If intSp > 0 Then strSp = "[SK]"
                       End If
                       
                       
                   End If
                
                   Set M_objrs = Nothing
                    Call CLEA
                    cmdsql = cmdsql + " Where usertype='1'"
                    M_OBJCONN.Execute cmdsql
                    Call ceksts
                    
                
                    
                    cmdsql = "UPDATE usertbl SET f_flagrender=1,F_VL='" + StsVl + "', F_PR='" + StsPR + "',F_SK='" + StsSK + "', F_ON='" + StsON + "',F_OS='" + StsOS + "',F_PTP='" + StsPTP + "' "
                    cmdsql = cmdsql + ", F_BP='" + StsBP + "', F_POP='" + StsPOP + "', F_SP='" + StsSP + "', F_UC='" + StsUC + "', F_RP='" + StsRP + "' , f_blank='" + Stsblank + "'"
                    cmdsql = cmdsql + ", F_WO_DATE='" + StsWO_Date + "',F_WO_2009='" + StsWO_2009 + "', F_WO_2008='" + StsWO_2008 + "',F_WO_2007='" + StsWO_2007 + "' "
                    cmdsql = cmdsql + ", F_WO_2006='" + StsWO_2006 + "', F_WO_2005='" + StsWO_2005 + "', F_WO_2004='" + StsWO_2004 + "', F_WO_2003='" + StsWO_2003 + "' "
                    cmdsql = cmdsql + ", F_WO_2002='" + StsWO_2002 + "', F_WO_2001='" + StsWO_2001 + "', F_WO_2000='" + StsWO_2000 + "', F_WO_1999='" + StsWO_1999 + "',F_WO_2010='" + StsWO_2010 + "' "
                    cmdsql = cmdsql + " Where usertype='1'"
                    M_OBJCONN.Execute cmdsql
                    MsgBox "Proccess to All DCR Name Done.....!"
                   strall = "" + StrNa + strRp + strOp + strPtp + strpop + strbp + strSp
                    If strall <> Empty Then MsgBox " data terblok oleh Supervisor "
                  End If
        
                If SSOption1(1).Value Then
                    If Combo1(0).Text = "" Then
                        MsgBox "Select DCR Name To Proccess..!"
                        Combo1(0).SetFocus
                    Else
                   Set M_objrs = New ADODB.Recordset
                        M_objrs.CursorLocation = adUseClient
                        M_objrs.Open "select * from usertbl where  userid='" + Combo1(0).Text + "' and lockdarispv  is not null limit 1", M_OBJCONN, adOpenDynamic, adLockOptimistic
                   
                     If Not M_objrs.EOF Then
                       If Check1(0).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intNa = InStr(1, vrsumber, "VL-", vbTextCompare)
                            If intNa > 0 Then StrNa = "[VL]"
                       End If
                       
                       If Check1(7).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intRP = InStr(1, vrsumber, "RP-", vbTextCompare)
                            If intRP > 0 Then strRp = "[Rp]"
                       End If
                       
                       If Check1(1).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intOp = InStr(1, vrsumber, "PR-", vbTextCompare)
                            If intOp > 0 Then strOp = "[PR]"
                       End If
                       
                       
                       If Check1(2).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intPTP = InStr(1, vrsumber, "PTP", vbTextCompare)
                            If intPTP > 0 Then strPtp = "[PTP]"
                       End If
                       
                       If Check1(4).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intpop = InStr(1, vrsumber, "POP", vbTextCompare)
                            If intpop > 0 Then strpop = "[POP]"
                       End If
                       
                       If Check1(3).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intBP = InStr(1, vrsumber, "BP-", vbTextCompare)
                            If intpop > 0 Then strbp = "[BP]"
                       End If
                       
                       If Check1(5).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intSp = InStr(1, vrsumber, "SP-", vbTextCompare)
                            If intSp > 0 Then strSp = "[SP]"
                       End If
                       
                          If Check1(10).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intSp = InStr(1, vrsumber, "OS-", vbTextCompare)
                            If intSp > 0 Then strSp = "[OS]"
                       End If
                       
                     
                       If Check1(11).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intSp = InStr(1, vrsumber, "ON-", vbTextCompare)
                            If intSp > 0 Then strSp = "[ON]"
                       End If
                       
                        If Check1(12).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intSp = InStr(1, vrsumber, "SK-", vbTextCompare)
                            If intSp > 0 Then strSp = "[SK]"
                       End If
                       
                       
                   End If
                
                   Set M_objrs = Nothing
                   
                        Call CLEA
                        cmdsql = cmdsql + " Where userid='" + Combo1(0).Text + "'"
                        M_OBJCONN.Execute cmdsql
                        Call ceksts
                        cmdsql = "UPDATE usertbl SET f_flagrender=1,F_VL='" + StsVl + "', F_PR='" + StsPR + "',F_SK='" + StsSK + "', F_ON='" + StsON + "',F_OS='" + StsOS + "', F_PTP='" + StsPTP + "' "
                        cmdsql = cmdsql + ", F_BP='" + StsBP + "', F_POP='" + StsPOP + "', F_SP='" + StsSP + "', F_UC='" + StsUC + "',  F_RP='" + StsRP + "',f_blank='" + Stsblank + "'"
                        cmdsql = cmdsql + ", F_WO_DATE='" + StsWO_Date + "',F_WO_2009='" + StsWO_2009 + "', F_WO_2008='" + StsWO_2008 + "',F_WO_2007='" + StsWO_2007 + "' "
                        cmdsql = cmdsql + ", F_WO_2006='" + StsWO_2006 + "', F_WO_2005='" + StsWO_2005 + "', F_WO_2004='" + StsWO_2004 + "', F_WO_2003='" + StsWO_2003 + "' "
                        cmdsql = cmdsql + ", F_WO_2002='" + StsWO_2002 + "', F_WO_2001='" + StsWO_2001 + "', F_WO_2000='" + StsWO_2000 + "', F_WO_1999='" + StsWO_1999 + "', F_WO_2010='" + StsWO_2010 + "' "
                        cmdsql = cmdsql + " Where userid='" + Combo1(0).Text + "'"
                        M_OBJCONN.Execute cmdsql
                        MsgBox "Proccess To  " + Combo1(0).Text + "  " + Combo1(1).Text + " Done.....!"
                        strall = "" + StrNa + strRp + strOp + strPtp + strpop + strbp + strSp
                        If strall <> Empty Then MsgBox " data terblok oleh Supervisor dengan account " + strall
                    
                    End If
                Else
                    If SSOption1(2).Value And spv = True Then
                        If Combo1(0).Text = "" Then
                            MsgBox "Select SPV Name To Proccess..!"
                        Else
                            Set M_objrs = New ADODB.Recordset
                                M_objrs.CursorLocation = adUseClient
                                M_objrs.Open "select * from usertbl where spvcode='" + Combo1(0).Text + "' AND lockdarispv  is not null limit 1", M_OBJCONN, adOpenDynamic, adLockOptimistic
                            If Not M_objrs.EOF Then
                                If Check1(0).Value = 1 Then
                                    vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                                    intNa = InStr(1, vrsumber, "VL-", vbTextCompare)
                                    If intNa > 0 Then StrNa = "[VL]"
                                End If
                       
                                If Check1(7).Value = 1 Then
                                        vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                                        intRP = InStr(1, vrsumber, "RP-", vbTextCompare)
                                        If intRP > 0 Then strRp = "[Rp]"
                                End If
                       
                                If Check1(1).Value = 1 Then
                                    vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                                    intOp = InStr(1, vrsumber, "PR-", vbTextCompare)
                                    If intOp > 0 Then strOp = "[PR]"
                                End If
                       
                       
                                If Check1(2).Value = 1 Then
                                    vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                                    intPTP = InStr(1, vrsumber, "PTP", vbTextCompare)
                                    If intPTP > 0 Then strPtp = "[PTP]"
                                End If
                       
                                If Check1(4).Value = 1 Then
                                    vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                                    intpop = InStr(1, vrsumber, "POP", vbTextCompare)
                                    If intpop > 0 Then strpop = "[POP]"
                                End If
                       
                                If Check1(3).Value = 1 Then
                                     vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                                     intBP = InStr(1, vrsumber, "BP-", vbTextCompare)
                                     If intpop > 0 Then strbp = "[BP]"
                                End If
                       
                                If Check1(5).Value = 1 Then
                                     vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                                     intSp = InStr(1, vrsumber, "SP-", vbTextCompare)
                                     If intSp > 0 Then strSp = "[SP]"
                                End If
                                
                                   If Check1(10).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intSp = InStr(1, vrsumber, "OS-", vbTextCompare)
                            If intSp > 0 Then strSp = "[OS]"
                       End If
                       
                     
                       If Check1(11).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intSp = InStr(1, vrsumber, "ON-", vbTextCompare)
                            If intSp > 0 Then strSp = "[ON]"
                       End If
                       
                        If Check1(12).Value = 1 Then
                            vrsumber = IIf(IsNull(M_objrs("lockdarispv")), "", M_objrs("lockdarispv"))
                            intSp = InStr(1, vrsumber, "SK-", vbTextCompare)
                            If intSp > 0 Then strSp = "[SK]"
                       End If
                       
                   End If
                
                   Set M_objrs = Nothing
                            Call CLEA
                            cmdsql = cmdsql + " Where spvcode='" + Combo1(0).Text + "'"
                            M_OBJCONN.Execute cmdsql
                            Call ceksts
                            cmdsql = "UPDATE usertbl SET f_flagrender=1,F_VL='" + StsVl + "', F_PR='" + StsPR + "',F_SK='" + StsSK + "', F_ON='" + StsON + "',F_OS='" + StsOS + "', F_PTP='" + StsPTP + "' "
                            cmdsql = cmdsql + ", F_BP='" + StsBP + "', F_POP='" + StsPOP + "', F_SP='" + StsSP + "', F_UC='" + StsUC + "',  F_RP='" + StsRP + "',f_blank='" + Stsblank + "'"
                            cmdsql = cmdsql + ", F_WO_DATE='" + StsWO_Date + "',F_WO_2009='" + StsWO_2009 + "', F_WO_2008='" + StsWO_2008 + "',F_WO_2007='" + StsWO_2007 + "' "
                            cmdsql = cmdsql + ", F_WO_2006='" + StsWO_2006 + "', F_WO_2005='" + StsWO_2005 + "', F_WO_2004='" + StsWO_2004 + "', F_WO_2003='" + StsWO_2003 + "' "
                            cmdsql = cmdsql + ", F_WO_2002='" + StsWO_2002 + "', F_WO_2001='" + StsWO_2001 + "', F_WO_2000='" + StsWO_2000 + "', F_WO_1999='" + StsWO_1999 + "' ,F_WO_2010='" + StsWO_2010 + "' "
                            cmdsql = cmdsql + " Where spvcode='" + Combo1(0).Text + "'"
                            M_OBJCONN.Execute cmdsql
                            MsgBox "Proccess To  " + Combo1(0).Text + "  " + Combo1(1).Text + " Done.....!"
                        strall = "" + StrNa + strRp + strOp + strPtp + strpop + strbp + strSp
                        If strall <> Empty Then MsgBox " data terblok oleh Supervisor dengan account " + strall
                        End If
                    End If
             End If
        End If
        StsVl = ""
        StsPR = ""
        StsOS = ""
        StsON = ""
        StsSK = ""
        
       StsPTP = ""
       StsBP = ""
       StsPOP = ""
       StsSP = ""
       StsUC = ""
       StsRP = ""
       StsWO_Date = ""
       StsWO_2009 = ""
       StsWO_2008 = ""
       StsWO_2007 = ""
       StsWO_2006 = ""
       StsWO_2005 = ""
       StsWO_2004 = ""
       StsWO_2003 = ""
       StsWO_2002 = ""
       StsWO_2001 = ""
       StsWO_2000 = ""
       StsWO_1999 = ""
       
Case 1
        'CMDSQL = "UPDATE usertbl SET F_NA='NA-', F_OP='OP-', F_PTP='POP', F_BP='BP-', F_POP='POP', F_SP='SP-', F_UC='UC'"
        
        cmdsql = "UPDATE usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL, F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL ,F_blank=NULL"
        cmdsql = cmdsql + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
        cmdsql = cmdsql + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,F_WO_2010=NULL"
        
        Select Case UCase(MDIForm1.txtlevel.Text)
            Case "SUPERVISOR"
                    cmdsql = cmdsql + " Where spvcode='" + Combo1(0).Text + "' "
            Case "TEAMLEADER"
                   cmdsql = cmdsql + " WHERE TEAM='" + MDIForm1.txtusername.Text + "'"
        End Select
        M_OBJCONN.Execute cmdsql
        MsgBox "Reset Done.....!"
Case 2
        Unload Me


End Select

End Sub

Private Sub SSOption1_Click(Index As Integer, Value As Integer)
Dim M_objrs As ADODB.Recordset
Select Case Index

Case 0
        Combo1(0).Enabled = False
        Combo1(1).Enabled = False
Case 1
        Combo1(0).Enabled = True
        Combo1(1).Enabled = True
        Combo1(0).CLEAR
        Combo1(1).CLEAR
        Dim M_DATA As New CLS_FRMSEARCH
        Set M_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "")
            While Not M_objrs.EOF
                Combo1(0).AddItem M_objrs("USERID")
                Combo1(1).AddItem M_objrs("AGENT")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        'SSOption1(0).Value = True
        spv = False
Case 2
        Combo1(0).Enabled = True
        Combo1(1).Enabled = True
        Combo1(0).CLEAR
        Combo1(1).CLEAR
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Then
            M_objrs.Open "select * from SPVTBL ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        ElseIf UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Or UCase(MDIForm1.txtlevel.Text) = "ADMIN" Then
            M_objrs.Open "select * from SPVTBL", M_OBJCONN, adOpenDynamic, adLockOptimistic
        ElseIf UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        M_objrs.Open "select * from SPVTBL where team='" + MDIForm1.txtusername.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
        End If
            While Not M_objrs.EOF
                Combo1(0).AddItem M_objrs("SPVCODE")
                Combo1(1).AddItem M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
'        SSOption1(0).Value = True
'        SSOption1(1).Value = True
        
End Select

End Sub

Sub CLEA()
cmdsql = "UPDATE usertbl SET F_NA=NULL,F_blank=NULL, F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL, "
cmdsql = cmdsql & " F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL, F_WO_2004=NULL, "
cmdsql = cmdsql & " F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,F_WO_2010=NULL "
End Sub

Sub ceksts()
If Check1(0).Value Then
    StsVl = "VL-"
End If
If Check1(1).Value Then
          
                StsPR = "PR-"
            
End If

If Check1(2).Value Then
            StsPTP = "PTP"
End If
 If Check1(3).Value Then
            StsBP = "BP-"
End If
       If Check1(4).Value Then
            StsPOP = "POP"
 End If
If Check1(5).Value Then
            StsSP = "SP-"
End If
If Check1(6).Value Then
            StsUC = "UC"
End If
If Check1(7).Value Then
            StsRP = "RP-"
End If

If Check1(8).Value Then
            Stsblank = ""
End If


If Check1(11).Value Then
            StsON = "ON-"
End If

If Check1(10).Value Then
            StsOS = "OS-"
End If

If Check1(12).Value Then
            StsSK = "SK-"
End If



If Check2.Value Then
        StsWO_Date = "1"
    Dim i As Integer
    If List1.ListIndex = -1 Then Exit Sub
    
    For i = List1.ListCount - 1 To 0 Step -1
         If List1.Selected(i) = True Then
            Select Case List1.list(i)
             Case "2010"
                StsWO_2010 = "2010"
             Case "2009"
                StsWO_2009 = "2009"
                'Text1.Text = List1.LIST(i)
             Case "2008"
                StsWO_2008 = "2008"
                Text1.Text = List1.list(i)
             Case "2007"
                StsWO_2007 = "2007"
             Case "2006"
                StsWO_2006 = "2006"
             Case "2005"
                StsWO_2005 = "2005"
             Case "2004"
                StsWO_2004 = "2004"
             Case "2003"
                StsWO_2003 = "2003"
             Case "2002"
                StsWO_2002 = "2002"
             Case "2001"
                StsWO_2001 = "2001"
             Case "2000"
                StsWO_2000 = "2000"
             Case "1999"
                StsWO_1999 = "1999"
            End Select
            
        ' Text1.Text = Text1.Text + "'" + List1.LIST(i) + "',"
         Else
    End If
    Next i
End If

End Sub
