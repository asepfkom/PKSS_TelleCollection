VERSION 5.00
Begin VB.Form FrmRemarksRequest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remarks Request Form"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6075
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   3120
      Width           =   1515
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   3120
      Width           =   1515
   End
   Begin VB.TextBox TxtRemarks 
      Appearance      =   0  'Flat
      Height          =   1035
      Left            =   180
      TabIndex        =   9
      Top             =   1860
      Width           =   5835
   End
   Begin VB.TextBox TxtAgent 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   1140
      Width           =   1215
   End
   Begin VB.TextBox TxtCustid 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   780
      Width           =   3315
   End
   Begin VB.TextBox TxtIdForm 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox TxtForm 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   420
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Agent:"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   1140
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "Custid:"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   780
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Id Form:"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Form Request:"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Remarks:"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   1500
      Width           =   1455
   End
End
Attribute VB_Name = "FrmRemarksRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim STRSQL As String
    
    If TxtRemarks.Text = "" Then
        MsgBox "Anda belum menulis remarks!", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If
    
    'Update form Request PUM
    If Trim(TxtForm.Text) = "PUM" Then
        STRSQL = "update mandiri.tbl_req_pum set remarks='"
        STRSQL = STRSQL + Trim(TxtRemarks.Text) + "', status='1' where id='"
        STRSQL = STRSQL + Trim(TxtIdForm.Text) + "'"
        M_OBJCONN.Execute STRSQL
        MsgBox "Data PUM berhasil diupdate!", vbOKOnly + vbInformation, "Informasi"
    End If
    
    'Update Form EC
     If Trim(TxtForm.Text) = "EC" Then
        STRSQL = "update mandiri.tbl_req_ec set remarks='"
        STRSQL = STRSQL + Trim(TxtRemarks.Text) + "', status='1' where id='"
        STRSQL = STRSQL + Trim(TxtIdForm.Text) + "'"
        M_OBJCONN.Execute STRSQL
        MsgBox "Data EC berhasil diupdate!", vbOKOnly + vbInformation, "Informasi"
    End If
    
    'Update Form BS
     If Trim(TxtForm.Text) = "BS" Then
        STRSQL = "update mandiri.tbl_req_bs set remarks='"
        STRSQL = STRSQL + Trim(TxtRemarks.Text) + "', status='1' where id='"
        STRSQL = STRSQL + Trim(TxtIdForm.Text) + "'"
        M_OBJCONN.Execute STRSQL
        MsgBox "Data BS berhasil diupdate!", vbOKOnly + vbInformation, "Informasi"
    End If
    
    'Update form RS
     If Trim(TxtForm.Text) = "RS" Then
        STRSQL = "update mandiri.tbl_req_rs set remarks='"
        STRSQL = STRSQL + Trim(TxtRemarks.Text) + "', status='1' where id='"
        STRSQL = STRSQL + Trim(TxtIdForm.Text) + "'"
        M_OBJCONN.Execute STRSQL
        MsgBox "Data RS berhasil diupdate!", vbOKOnly + vbInformation, "Informasi"
    End If
    
    'Update form OST
     If Trim(TxtForm.Text) = "OST" Then
        STRSQL = "update mandiri.tbl_req_ost set remarks='"
        STRSQL = STRSQL + Trim(TxtRemarks.Text) + "', status='1' where id='"
        STRSQL = STRSQL + Trim(TxtIdForm.Text) + "'"
        M_OBJCONN.Execute STRSQL
        MsgBox "Data OST berhasil diupdate!", vbOKOnly + vbInformation, "Informasi"
    End If
    
    'Update form Problem
     If Trim(TxtForm.Text) = "PROBLEM" Then
        STRSQL = "update mandiri.tbl_req_problem set solve='"
        STRSQL = STRSQL + Trim(TxtRemarks.Text) + "', status='1' where id='"
        STRSQL = STRSQL + Trim(TxtIdForm.Text) + "'"
        M_OBJCONN.Execute STRSQL
        MsgBox "Data Problem berhasil diupdate!", vbOKOnly + vbInformation, "Informasi"
    End If
End Sub
