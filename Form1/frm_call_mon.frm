VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_call_mon 
   Caption         =   "Call Monitoring To Excel"
   ClientHeight    =   10575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15090
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   15090
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   11880
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BROWSE"
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
      Left            =   8040
      TabIndex        =   2
      Top             =   180
      Width           =   1215
   End
   Begin VB.TextBox Txtpath 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1680
      TabIndex        =   3
      Top             =   180
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2295
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   14295
      Begin VB.TextBox txtch 
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   1320
         TabIndex        =   13
         Top             =   360
         Width           =   3015
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "SIMPAN"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtch 
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   1320
         TabIndex        =   11
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtch 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   1320
         TabIndex        =   10
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtch 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   1320
         TabIndex        =   9
         Top             =   1800
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   1
         Left            =   9120
         TabIndex        =   7
         Top             =   600
         Width           =   2790
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   0
         Left            =   7920
         TabIndex        =   6
         Top             =   600
         Width           =   1200
      End
      Begin VB.TextBox txtlen_service 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   9960
         TabIndex        =   5
         Text            =   "0"
         Top             =   1380
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   0
         Left            =   5880
         TabIndex        =   8
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy HH:mm"
         Format          =   132251651
         CurrentDate     =   41813
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   1
         Left            =   5880
         TabIndex        =   14
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy HH:mm"
         Format          =   132251651
         CurrentDate     =   41813
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   2
         Left            =   5880
         TabIndex        =   15
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy HH:mm"
         Format          =   132251651
         CurrentDate     =   41813
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   3
         Left            =   5880
         TabIndex        =   16
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy HH:mm"
         Format          =   132251651
         CurrentDate     =   41813
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Index           =   4
         Left            =   7920
         TabIndex        =   17
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy HH:mm"
         Format          =   132251651
         CurrentDate     =   41813
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Third Party :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CH 1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CH 2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "CH 3"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Time"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Time"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   24
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Time"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Time"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   22
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tele Collection :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Length Of Service :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   20
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10680
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   7920
         TabIndex        =   18
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.CheckBox text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
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
      Left            =   10560
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7575
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   13361
      _Version        =   393216
      Enabled         =   0   'False
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Excel File :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   30
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "Form_call_mon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sCurrent_Row    As Integer
Private sCurrent_Col    As Integer

Private xcol()          As String

Private objExcel        As Excel.Application
Private objBook         As Excel.Workbook
Private objSheet        As Excel.Worksheet
Private xx              As Integer

Public Sub ResizeGrid(pGrid As MSFlexGrid, pForm As Form)
    Dim intRow As Integer
    Dim intCol As Integer
    
    With pGrid
        For intCol = 0 To .Cols - 1
            For intRow = 0 To .Rows - 1
                If .ColWidth(intCol) < pForm.TextWidth(.TextMatrix(intRow, intCol)) + 100 Then
                   .ColWidth(intCol) = pForm.TextWidth(.TextMatrix(intRow, intCol)) + 100
                End If
            Next
        Next
    End With
End Sub

Private Sub load_master_callmon()
    Dim rs      As New ADODB.Recordset
    Dim i       As Integer
    Dim z       As Integer
    Dim X       As Integer
    Dim yy      As Integer
    
    CMDSQL = "SELECT * FROM mandiri.tblresult_callmon ORDER BY row_sheet"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    i = 2: X = 1
    MSFlexGrid1.CLEAR
    
    MSFlexGrid1.Cols = 8
    MSFlexGrid1.Rows = 2
    
    MSFlexGrid1.TextMatrix(1, 1) = "DESC. PENILAIAN"
    MSFlexGrid1.TextMatrix(1, 2) = "Third Party"
    MSFlexGrid1.TextMatrix(1, 3) = "CH 1"
    MSFlexGrid1.TextMatrix(1, 4) = "CH 2"
    MSFlexGrid1.TextMatrix(1, 5) = "CH 3"
    
    If rs.RecordCount > 0 Then
        For yy = 1 To MSFlexGrid1.Cols - 1
            MSFlexGrid1.col = yy
            MSFlexGrid1.row = 1
            MSFlexGrid1.CellBackColor = &H8000000F
            MSFlexGrid1.CellFontBold = True
        Next yy

        While Not rs.EOF
            xcol() = Split(IIf(IsNull(rs!column_sheet), "", rs!column_sheet), "|")
            
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(i, 1) = IIf(IsNull(rs!keterangan), "", rs!keterangan)
            MSFlexGrid1.TextMatrix(i, 6) = cnull(rs!row_sheet)
            MSFlexGrid1.TextMatrix(i, 7) = cnull(rs!row_disabled)
            If cnull(rs!row_disabled) < 1 Then
                MSFlexGrid1.TextMatrix(i, 0) = X
                
                MSFlexGrid1.TextMatrix(i, 2) = 0
                MSFlexGrid1.TextMatrix(i, 3) = 0
                MSFlexGrid1.TextMatrix(i, 4) = 0
                MSFlexGrid1.TextMatrix(i, 5) = 0
                X = X + 1
            Else
                For yy = 1 To MSFlexGrid1.Cols - 1
                    MSFlexGrid1.col = yy
                    MSFlexGrid1.row = i
                    MSFlexGrid1.CellBackColor = &HC0FFC0
                    MSFlexGrid1.CellFontBold = True
                Next yy
            End If
            i = i + 1
            
            rs.MoveNext
        Wend
    End If
    
    Set rs = Nothing
End Sub

Private Sub Combo1_Click(Index As Integer)
    If Index = 0 Then
        Combo1(1).ListIndex = Combo1(0).ListIndex
    Else
        Combo1(0).ListIndex = Combo1(1).ListIndex
    End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
    CD.ShowSave
    TxtPath.Text = CD.FileName
    
    If CD.FileName <> "" Then
        Set objExcel = CreateObject("Excel.Application")
        Set objBook = objExcel.Workbooks.Open(TxtPath.Text)
        Set objSheet = objBook.Sheets(1)
        
        Frame1.Enabled = False: MSFlexGrid1.Enabled = False
        
        If Trim(LCase(objBook.Sheets(1).Name)) = "collmon" Then
        
            If IsDate(objSheet.Cells(4, "E")) Then
                DTPicker1(4).Value = CDate(objSheet.Cells(4, "E"))
            End If
            Combo1(0).Text = objSheet.Cells(1, "A")
            Combo1(1).Text = objSheet.Cells(5, "E")
            txtlen_service.Text = objSheet.Cells(6, "E")
        
            For xx = 0 To UBound(xcol)
                txtch(xx).Text = objSheet.Cells(4, xcol(xx))
                DTPicker1(xx).Value = CDate(objSheet.Cells(5, xcol(xx)))
            Next xx
            
            With MSFlexGrid1
                For i = 2 To .Rows - 1
                    If .TextMatrix(i, 7) < 1 Then
                        For xx = 0 To UBound(xcol)
                            ' Start from 2
                            colFlex = xx + 2
                            .TextMatrix(i, colFlex) = IIf(objSheet.Cells(.TextMatrix(i, 6), xcol(xx)) = "v", 1, 0)
                        Next xx
                    End If
                Next i
            End With
            
            Frame1.Enabled = True: MSFlexGrid1.Enabled = True
            
        Else
            MsgBox "Format excel salah!!", vbCritical + vbOKOnly, "INFO"
        End If
        
        objExcel.Quit
        Set objExcel = Nothing
        Set objBook = Nothing
        Set objSheet = Nothing
    End If
    Exit Sub
err:
    MsgBox err.Description, vbCritical + vbInformation, "INFO"
    objExcel.Quit
    Set objExcel = Nothing
    Set objBook = Nothing
    Set objSheet = Nothing
End Sub

Private Sub Command3_Click()
    Dim rs              As New ADODB.Recordset
    Dim i               As Double
    Dim m_msgbox        As String
    Dim iCell           As Integer
    Dim iLastColumn     As Integer
    Dim arrAlpha
    Dim ilastrow        As Integer
    Dim colFlex         As Integer

    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath = Empty Then
        MsgBox "Nama file tidak boleh kosong, download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.Open(TxtPath.Text)
    Set objSheet = objBook.Sheets(1)

    objSheet.Cells(4, "E") = DTPicker1(4).Value
    objSheet.Cells(1, "A") = Combo1(0).Text
    objSheet.Cells(5, "E") = Combo1(1).Text
    objSheet.Cells(6, "E") = txtlen_service.Text
    
    ' SET HEADER
    For xx = 0 To UBound(xcol)
        objSheet.Cells(4, xcol(xx)) = "'" & txtch(xx).Text
        objSheet.Cells(5, xcol(xx)) = Format(DTPicker1(xx).Value, "dd-mm-yyyy hh:mm")
    Next xx
    
    With MSFlexGrid1
        For i = 2 To .Rows - 1
            If .TextMatrix(i, 7) < 1 Then
                For xx = 0 To UBound(xcol)
                    ' Start from 2
                    colFlex = xx + 2
                    objSheet.Cells(.TextMatrix(i, 6), xcol(xx)) = IIf(.TextMatrix(i, colFlex) = "1", "v", "")
                Next xx
            End If
        Next i
    End With
    
    objBook.Save
    MsgBox "Simpan data Berhasil !!", vbOKOnly + vbInformation, "INFO"
    objExcel.Quit

    Set objExcel = Nothing
    Set objBook = Nothing
    Set objSheet = Nothing
End Sub

Private Sub Form_Load()
    Dim rs              As New ADODB.Recordset
    Dim sqlstr          As String
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    sqlstr = "SELECT userid,agent FROM mandiri.usertbl WHERE userid IS NOT NULL "
    
    If MDIForm1.txtlevel.Text = "TeamLeader" Then
        sqlstr = sqlstr & " AND team='" & MDIForm1.TxtUsername.Text & "' "
    End If
    
    rs.Open sqlstr, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        Do Until rs.EOF
            Combo1(0).AddItem cnull(rs!USERID)
            Combo1(1).AddItem cnull(rs!AGENT)
            rs.MoveNext
        Loop
    End If

    Set rs = Nothing
    
    Call load_master_callmon
    ResizeGrid MSFlexGrid1, Me
    MSFlexGrid1.ColWidth(6) = 0
    MSFlexGrid1.ColWidth(7) = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objExcel = Nothing
    Set objBook = Nothing
    Set objSheet = Nothing
End Sub

Private Sub MSFlexGrid1_Click()
    sCurrent_Col = MSFlexGrid1.col
    sCurrent_Row = MSFlexGrid1.row
    
    If IIf(MSFlexGrid1.TextMatrix(sCurrent_Row, 7) = "", 0, MSFlexGrid1.TextMatrix(sCurrent_Row, 7)) < 1 Then
        If sCurrent_Row > 1 And sCurrent_Col > 1 Then
            Text1.Value = MSFlexGrid1.TextMatrix(sCurrent_Row, sCurrent_Col)
            Call GridEdit(0)
        End If
    End If
End Sub

Sub GridEdit(KeyAscii As Integer)
   'use correct font
   Text1.FontName = MSFlexGrid1.FontName
   Text1.FontSize = MSFlexGrid1.FontSize
   'text1 = MSFlexGrid1
   Select Case KeyAscii
      Case 0 To Asc(" ")
         'text1 = MSFlexGrid1
         'text1.SelStart = 1000
      Case Else
         'text1 = Chr(KeyAscii)
         'text1.SelStart = 1
   End Select

   'position the edit box
   Text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
   Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
   Text1.Width = MSFlexGrid1.CellWidth
   'text1.Height = MSFlexGrid1.CellHeight
   Text1.Visible = True
   Text1.SetFocus
End Sub

Private Sub MSFlexGrid1_LeaveCell()
   If Text1.Visible Then
      If Len(MSFlexGrid1.Text) > 6 Then
      Else
        MSFlexGrid1 = Text1
      End If
      Text1.Visible = False
   End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyEscape
         Text1.Visible = False
         MSFlexGrid1.SetFocus
      Case vbKeyReturn
         MSFlexGrid1.SetFocus
      Case vbKeyDown
         MSFlexGrid1.SetFocus
         DoEvents
         If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then
            MSFlexGrid1.row = MSFlexGrid1.row + 1
         End If
      Case vbKeyUp
         MSFlexGrid1.SetFocus
         DoEvents
         If MSFlexGrid1.row > MSFlexGrid1.FixedRows Then
            MSFlexGrid1.row = MSFlexGrid1.row - 1
         End If
   End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   'noise suppression
   If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub txtlen_service_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 58, vbKeyBack
        Exit Sub
    Case Else
        KeyAscii = 0
    End Select
End Sub
