VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_CopyFIleCPA 
   Caption         =   "Copy FIle  CPA-Dokumen Pendukung CH"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16500
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   16500
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   11160
      TabIndex        =   23
      Top             =   360
      Width           =   5175
   End
   Begin VB.FileListBox File3 
      Height          =   4575
      Left            =   11160
      TabIndex        =   22
      Top             =   3480
      Width           =   5175
   End
   Begin VB.FileListBox File2 
      Height          =   4575
      Left            =   5640
      TabIndex        =   18
      Top             =   3600
      Width           =   5415
   End
   Begin VB.FileListBox File1 
      Height          =   4575
      Left            =   240
      TabIndex        =   17
      Top             =   3600
      Width           =   5295
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8640
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      Begin VB.CommandButton Command3 
         Caption         =   "REFRESH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   25
         Top             =   1800
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   9360
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*dir"
      End
      Begin VB.CommandButton Command2 
         Caption         =   "START COPY FILE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   2475
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SIMPAN "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton Browse 
         Caption         =   "...."
         Height          =   255
         Index           =   2
         Left            =   8880
         TabIndex        =   10
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton Browse 
         Caption         =   "...."
         Height          =   255
         Index           =   1
         Left            =   8880
         TabIndex        =   9
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Browse 
         Caption         =   "...."
         Height          =   255
         Index           =   0
         Left            =   8880
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TxtDestinasiFolder 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   5055
      End
      Begin VB.TextBox TxtDirketoridocPendukung 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox Txtdirektori1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   5055
      End
      Begin MSComDlg.CommonDialog CD2 
         Left            =   9360
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CD3 
         Left            =   9360
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   2520
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.TextBox Txtfolder1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Txtfolder1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   5160
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Txtfolder1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   6480
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   12
         Top             =   2400
         Width           =   10935
      End
      Begin VB.Label Label2 
         Caption         =   "*(scan ktp, surat permohonan,  bukti pembayaran,etc"
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
         Left            =   600
         TabIndex        =   4
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "Destinasi Directory"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Directory Folder Dokumen Pendukung. "
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   2
         Top             =   840
         Width           =   3135
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Directory Folder File CPA"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Contoh : xxxxxxxxxxxxxxxx-1.pdf"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   26
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Direktori Tujuan (hasil Copy File)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   11160
      TabIndex        =   24
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "Source Dokumen Pendukung, Scan FC KTP, bukti Pembayaran, etc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   5760
      TabIndex        =   20
      Top             =   3120
      Width           =   5415
   End
   Begin VB.Label Label4 
      Caption         =   "Source FIle CPA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   3240
      Width           =   1695
   End
End
Attribute VB_Name = "Form_CopyFIleCPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As Scripting.FileSystemObject
Dim logfile As Integer
Dim tarih As String

Private Sub Browse_Click(Index As Integer)
    Dim spli1() As String
    
    Select Case Index
    Case 0
        Txtdirektori1.Text = BrowseForFolder(hwnd, "Direktori File CPA")
'        Txtdirektori1.Text = CD1.FileName
'        split1 = Split(Txtdirektori1.Text, "\")
'        For I = 0 To UBound(split1)
'            If I = UBound(split1) Then
'            'MsgBox ""
'            Else
'                Txtfolder1(0).Text = Txtfolder1(0).Text & split1(I) & "\"
'            End If
'        Next I
'        Txtdirektori1.Text = Txtfolder1(0).Text
    Case 1
        TxtDirketoridocPendukung.Text = BrowseForFolder(hwnd, "Direktori Dokumen Pendukung")
'        CD2.ShowOpen
'        TxtDirketoridocPendukung.Text = CD2.FileName
'        split1 = Split(TxtDirketoridocPendukung.Text, "\")
'        For I = 0 To UBound(split1)
'            If I = UBound(split1) Then
'            'MsgBox ""
'            Else
'            Txtfolder1(1).Text = Txtfolder1(1).Text & split1(I) & "\"
'            End If
'        Next I
'        TxtDirketoridocPendukung.Text = Txtfolder1(1).Text
    Case 2
        TxtDestinasiFolder.Text = BrowseForFolder(hwnd, "Direktori tujuan file")
'        CD3.ShowOpen
'        TxtDestinasiFolder.Text = CD3.FileName
'        split1 = Split(TxtDestinasiFolder.Text, "\")
'        For I = 0 To UBound(split1)
'            If I = UBound(split1) Then
'            'MsgBox ""
'            Else
'            Txtfolder1(2).Text = Txtfolder1(2).Text & split1(I) & "\"
'            End If
'        Next I
'        TxtDestinasiFolder.Text = Txtfolder1(2).Text
    End Select
End Sub

Private Sub Command1_Click()
    Dim CMDSQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "select * from mandiri.tblsetingcopyfolder", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        CMDSQL = "UPDATE mandiri.tblsetingcopyfolder SET dircpa='" + Replace(Txtdirektori1.Text, "\", "\\") + "', dirdokumen='" + Replace(TxtDirketoridocPendukung.Text, "\", "\\") + "', "
        CMDSQL = CMDSQL + " dirtujuan = '" + Replace(TxtDestinasiFolder.Text, "\", "\\") + "' where id_dir=" + CStr(IIf(IsNull(rs!id_dir), 0, rs!id_dir)) + ""
        M_OBJCONN.Execute CMDSQL
        rs.MoveNext
    Wend
    MsgBox "Update setting Done"
    If rs.RecordCount = 0 Then
    CMDSQL = "insert into mandiri.tblsetingcopyfolder(dircpa, dirdokumen,dirtujuan )values ('" & Replace(Txtdirektori1.Text, "\", "\\") & "', "
    CMDSQL = CMDSQL + " '" & Replace(TxtDirketoridocPendukung.Text, "\", "\\") & "', '" & Replace(TxtDestinasiFolder.Text, "\", "\\") & "')"
    M_OBJCONN.Execute CMDSQL
    MsgBox "Insert setting Done"
    End If
    Set rs = Nothing
End Sub

Private Sub Loadsettingfolder()
    Dim fso As Scripting.FileSystemObject
    Dim rs          As ADODB.Recordset
    Dim slokasi     As String
    Dim ifilenumber As Integer
    
    Set fso = New Scripting.FileSystemObject
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "select * from mandiri.tblsetingcopyfolder", M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    While Not rs.EOF
        Txtdirektori1.Text = IIf(IsNull(rs!dircpa), "", rs!dircpa)
        TxtDirketoridocPendukung.Text = IIf(IsNull(rs!dirdokumen), "", rs!dirdokumen)
        TxtDestinasiFolder.Text = IIf(IsNull(rs!dirtujuan), "", rs!dirtujuan)
        rs.MoveNext
    Wend
    Set rs = Nothing
    
    If FolderExists(Txtdirektori1.Text) Then
        File1.Path = Txtdirektori1.Text
        File2.Path = TxtDirketoridocPendukung.Text
        File3.Path = TxtDestinasiFolder.Text
    End If

    slokasi = File3.Path & "\log.txt"
    If fso.FolderExists(slokasi) = False Then
        ifilenumber = FreeFile
        Open slokasi For Append As #ifilenumber
        Close #ifilenumber
    End If
End Sub

Private Sub Command2_Click()
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    'Path of the list box
    FromPath1 = File1.Path ' "\\169.239.19.44\Speedy Payment\2014\"
    FromPath2 = File2.Path
    topath = File3.Path     '"C:\test"
    
    'File1.Path = FromPath
    If fso.FolderExists(topath) = True Then
        MsgBox ""
    Else
        MsgBox "Folder not found"
        fso.CreateFolder topath
    End If
    
    If CreateLog("log", " & topath & ", False) = True Then
        MsgBox "buat log berasil"
    End If
    
    If Connection = False Or Finished = False Then
        ProgressBar1.Max = File1.ListCount
        DoEvents
        
        For i = 0 To File1.ListCount - 1
            ProgressBar1.Value = i
            OurFile = "\" & File1.list(i)
            
            namefolder1 = Split(File1.list(i), "-")
            If fso.FolderExists(topath & "\" & namefolder1(0) & "\") = False Then
                fso.CreateFolder topath & "\" & namefolder1(0) & "\"
            End If
            'For each file in it
            fso.CopyFile FromPath1 & OurFile, topath & "\" & namefolder1(0) & OurFile, True
                log "(" & OurFile & ") file has been copied from (" & FromPath & ") to (" & topath & "). Success!", False, True, True
                
            'Else
                ''''''''''''''''''''''''''''''' Log Module ''''''''''''''''''''''''''''''''
                ''Usage: LogString, LogDate, LogTime, DateTimeBeforeLog, DateTimeAfterLog''
                ''Log     "Hello" ,  False ,   True ,        True      ,      False      ''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                log "(" & OurFile & ") file could not be copied from (" & FromPath & ") to (" & topath & "). Faliure!", False, True, True
            'End If
        Next

    Else

        End

    End If
    
  ' COpy file ke dua
    If Connection = False Or Finished = False Then
        ProgressBar1.Max = File1.ListCount
        DoEvents
        For i = 0 To File2.ListCount - 1
            ProgressBar1.Value = i
            OurFile = "\" & File2.list(i)
            
            namefolder1 = Split(File2.list(i), "-")
            If fso.FolderExists(topath & "\" & namefolder1(0) & "\") = False Then
                fso.CreateFolder topath & "\" & namefolder1(0) & "\"
            End If
            'For each file in it
            fso.CopyFile FromPath2 & OurFile, topath & "\" & namefolder1(0) & OurFile, True
                log "(" & OurFile & ") file has been copied from (" & FromPath & ") to (" & topath & "). Success!", False, True, True
                
            'Else
                ''''''''''''''''''''''''''''''' Log Module ''''''''''''''''''''''''''''''''
                ''Usage: LogString, LogDate, LogTime, DateTimeBeforeLog, DateTimeAfterLog''
                ''Log     "Hello" ,  False ,   True ,        True      ,      False      ''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                log "(" & OurFile & ") file could not be copied from (" & FromPath & ") to (" & topath & "). Faliure!", False, True, True
            'End If
        Next

    Else
        End
    End If


End Sub

Public Sub ProgressDec(ProgressBarName As ProgressBar, Optional Max As Long, Optional Min As Long, Optional Dec As Long, Optional Continues As Boolean = False)
    Dim Recent As Long

    On Error GoTo ProgressErr

    ProgressBarName.ShowWhatsThis

    DoEvents
    'Maximum ProgressBar Value
    If Max <> 0 Then
        ProgressBarName.Max = Max 'If set use it
    Else
        Max = 100 'If max value is not set then make it 100
        ProgressBarName.Max = Max
    End If

    DoEvents
    'Minimum ProgressBar Value
    If Min <> 0 Then
        ProgressBarName.Min = Min 'If set use it
    Else
        Min = 1 'If minimum value is not set then make it 1
        ProgressBarName.Min = Min
    End If

    If Dec <> 0 Then Dec = Dec Else Dec = 1

    'When the ProgressBar value is at Minimum
    'Return to the Maximum value
    If Continues = True And ProgressBarName.Value = Min Then
        ProgressBarName.Value = Max
    End If

    'Checkout Recent progress (pre calculate bar value)
    Recent = ProgressBarName.Value - Dec

    DoEvents
    If Recent <= Min Then
        'Recent value is lower than or equals to Min value
        'to avoid errors caused by this issue value should equal to Min
        ProgressBarName.Value = Min
    ElseIf Recent > Min Then
        'Recent(pre calculated bar value) is higher than Min
        'So nothing wrong here, proceed..
        ProgressBarName.Value = ProgressBarName.Value - Dec
    End If

    Exit Sub

ProgressErr:

    'ProgressBar is null then create an error report.
    MsgBox "With " & err.number & " number : '" & err.Description & "' error occured. "
    'MsgBox "ProgressBar is not defined or Cant found the ProgressBar.. Please check the name of ProgressBar and re identify it.", vbCritical, "Unidentified ProgressBar!"

End Sub

Public Sub ProgressInc(ProgressBarName As ProgressBar, Optional Max As Long, Optional Min As Long, Optional Inc As Long, Optional Continues As Boolean = False)
    Dim Recent As Long

    On Error GoTo ProgressErr

    ProgressBarName.ShowWhatsThis

    DoEvents
    'Maximum ProgressBar Value
    If Max <> 0 Then
        ProgressBarName.Max = Max 'If set use it
    Else
        Max = 100 'If max value is not set then make it 100
        ProgressBarName.Max = Max
    End If

    DoEvents
    'Minimum ProgressBar Value
    If Min <> 0 Then
        ProgressBarName.Min = Min 'If set use it
    Else
        Min = 1 'If min value is not set then make it 1
        ProgressBarName.Min = Min
    End If

    If Inc <> 0 Then Inc = Inc Else Inc = 1

    'When the ProgressBar value is at Maximum
    'Return to the Minimum value
    If Continues = True And ProgressBarName.Value = Max Then
        ProgressBarName.Value = Min
    End If

    'Checkout Recent progress (pre calculate bar value)
    Recent = ProgressBarName.Value + Inc

    DoEvents
    If Recent >= Max Then
        'Recent value is higher than or equals to Max value
        'to avoid errors caused by this issue Value should equal to Max
        ProgressBarName.Value = Max
    ElseIf Recent < Max Then
        'Recent(pre calculated bar value) is lower than Max
        'So nothing wrong here, proceed..
        ProgressBarName.Value = ProgressBarName.Value + Inc
    End If

    Exit Sub

ProgressErr:

    'ProgressBar error report.
    MsgBox "With " & err.number & " number : '" & err.Description & "' error occured. "
    'MsgBox "ProgressBar is not defined or Cant found the ProgressBar.. Please check the name of ProgressBar and re identify it.", vbCritical, "Unidentified ProgressBar!"

End Sub

Function CheckPath(ByVal Path As String) As String

    If Right(Trim(Path), 1) = "\" Then
        CheckPath = Mid(Trim(Path), 1, Len(Trim(Path)) - 1)
    Else
        CheckPath = Trim(Path)
    End If

End Function

Function log(LogString As String, Optional LogDate As Boolean, Optional LogTime As Boolean, Optional BeforeLogText As Boolean = False, Optional AfterLogText As Boolean = False) As Boolean
    Dim WillBePrinted As String

    On err GoTo LogErr

    If BeforeLogText = True Then

        'Date Time Before Log
        WillBePrinted = "(" & Now & ") " & LogString

    ElseIf AfterLogText = True Then
        'Date Time After Log
        WillBePrinted = LogString & " (" & Now & ")"
    Else
        'No DateTime Included
        WillBePrinted = LogString
    End If

    Print logfile, WillBePrinted

    log = True

LogErr:

    log = False

End Function

Function CreateLog(Optional Name As String, Optional Path As String, Optional DateTimeBeforeName As Boolean = False) As Boolean
    Dim fso As New Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    logfile = FreeFile

    DoEvents
    'Name of Log File
    If Trim(Name) <> "" Then
        Name = Trim(Name)
    Else
        Name = Trim(App.EXEName)
    End If

    DoEvents
    'Path to Log File
    If Trim(Path) <> "" Then
        Path = CheckPath(Path)
    Else
        Path = CheckPath(App.Path)
    End If

    'If the path does not exists create it!
    If fso.FolderExists(Path) = False Then
        fso.CreateFolder Path
    End If

    'DateTimeBeforeName
    If DateTimeBeforeName = True Then

        DoEvents
        FullPath = Path & "\" & TimeMachine & " - " & Name & ".txt"
        'if already exists (Highly unlikely while date time is involved)
        If (fso.FileExists(FullPath) = True) Then
            fso.DeleteFile FullPath, True
            Open Path & "\" & TimeMachine & " - " & Name & ".txt" For Output As #logfile
        Else
            Open Path & "\" & TimeMachine & " - " & Name & ".txt" For Output As #logfile
        End If

    ElseIf DateTimeBeforeName = False Then

        DoEvents
        FullPath = Path & "\" & Name & ".txt"
        'if already exists (Highly posible while date time is not involved)
        If (fso.FileExists(FullPath) = True) Then
            fso.DeleteFile FullPath, True
            Open Path & "\" & Name & ".txt" For Output As #logfile
        Else
            Open Path & "\" & Name & ".txt" For Output As #logfile
        End If

    End If

    DoEvents
    'Now if everything was successfull
    If (fso.FileExists(FullPath) = True) Then
        CreateLog = True
    Else
        CreateLog = False
    End If

End Function

Function TimeMachine(Optional OnlyDate As Boolean = False) As String
    Dim MyDate, MyTime As String

    'Get local date
    For Each Part In Split(Date, ".")
        'Some times 01.01.2012 is shown as 1.1.2012
        'to fix this do a zero check..
        If Len(Part) < 3 And Len(Part) > 0 Then Part = Right("00" & Part, 2) Else Part = Part
        MyDate = MyDate & "." & Part
    Next

    'Get local time
    For Each Part In Split(Time, ":")
        'Some times 01.01.2012 is shown as 1.1.2012
        'to fix this do a zero check..
        If Len(Part) < 3 And Len(Part) > 0 Then
            MyTime = MyTime & "." & Right("00" & Part, 2)
        End If
    Next

    'Clean "." at start
    MyDate = Mid(MyDate, 2, Len(MyDate))
    MyTime = Mid(MyTime, 2, Len(MyTime))

    'Publish
    If OnlyDate = True Then
        TimeMachine = "Date " & MyDate
    Else
        TimeMachine = "Date " & MyDate & " Time " & MyTime
    End If

End Function

Private Sub Command3_Click()
    Call Loadsettingfolder
End Sub

Private Sub Dir1_Change()
    File3.Path = Dir1.Path
End Sub

Private Sub Form_Load()
    ' panggil setingan folder
    Call Loadsettingfolder
End Sub

Private Sub TxtDestinasiFolder_Change()
'    Dir1.Path = TxtDestinasiFolder
End Sub

Private Sub Txtdirektori1_Change()
'    File1.Path = Txtdirektori1.Text
End Sub

Private Sub TxtDirketoridocPendukung_Change()
'    File2.Path = TxtDirketoridocPendukung.Text
End Sub
