VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_CARI_HISTORY_CC_mgm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub CARI_STATUS_CUSTOMER(TYPE_PRODUK As String)
Dim NAMA_FORM As Object
Select Case UCase(TYPE_PRODUK)
Case "CREDIT CARD"
    Set NAMA_FORM = FRMCUST_CC_mgm
Case Else
    Exit Sub
End Select
With NAMA_FORM
If .Check2(1).Value = 1 Then
    .KETHSLKERJA = .Combo3(0).Text
    .STATUS_CUSTOMER = "0"
    .STATUS_FIELD_LAMA = .Combo3(0).Text
    .DOK1 = Empty
    Exit Sub
End If
If .Check2(0).Value = 1 Then
.KETHSLKERJA = "1Z"
.STATUS_CUSTOMER = "25"
.STATUS_FIELD_LAMA = "1Z"
.DOK1 = Empty
    If .Option2(0).Value Then
        .STATUS_CUSTOMER = "3"
        .KETHSLKERJA = "1F"
        .STATUS_FIELD_LAMA = "2A"
        .DOK1 = Empty
    Else
        If .Option2(1).Value Then
            .STATUS_CUSTOMER = "2"
            .KETHSLKERJA = "1G"
            .STATUS_FIELD_LAMA = "1B"
            .DOK1 = Empty
        End If
    End If
End If
If .Check2(0).Value = 0 And .Check2(1).Value = 0 Then
    .STATUS_CUSTOMER = "1"
    .KETHSLKERJA = "1A"
    .STATUS_FIELD_LAMA = "1A"
End If
End With
End Sub

Public Function VIEW_HISTORY_KERJA(HISTORY As String, DOKUMEN As String, TYPE_PRODUK As String)
Dim NAMA_FORM As Variant
Select Case UCase(TYPE_PRODUK)
Case "CREDIT CARD"
    Set NAMA_FORM = FRMCUST_CC_mgm
Case Else
    Exit Function
End Select
With NAMA_FORM
Select Case Left(HISTORY, 1)
    Case 0
        .Check2(1).Value = 1
            If UCase(MDIForm1.Text2.Text) = "AGENT" Then
                .Check2(1).Enabled = False
                .Frame25.Enabled = False
            End If
    Case Else
        .Check2(1).Value = 0
End Select
If Left(HISTORY, 1) = "9" Then
    .Check2(1).Value = 1
End If

 Select Case HISTORY
    Case "0B"
            .Check2(1).Value = 1
            If UCase(MDIForm1.Text2.Text) = "AGENT" Then
'                .Check2(1).Enabled = False
 '               .Frame25.Enabled = False
            End If
    Case "1Z"
            .Check2(0).Value = 1
    Case "1F"
            .Check2(0).Value = 1
            .Option2(0).Value = True
    Case "1H"
            .Check2(0).Value = 1
            .Option2(0).Value = True
            .Check2(0).Enabled = False
            .Combo2(0).Enabled = False
            .Option2(0).Enabled = False
            .Option2(1).Enabled = False
    Case "1I"
            .Check2(0).Value = 1
            .Option2(0).Value = True
            .Check2(0).Enabled = False
            .Combo2(0).Enabled = False
            .Option2(0).Enabled = False
            .Option2(1).Enabled = False
    Case "1J"
            .Check2(0).Value = 1
            .Option2(0).Value = True
            .Frame19(1).Visible = True
            .Check2(0).Enabled = False
            .Combo2(0).Enabled = False
            .TDBNumber1(0).Enabled = False
            .Option2(0).Enabled = False
            .Option2(1).Enabled = False
    Case "1G"
            .Check2(0).Value = 1
            .Option2(1).Value = True
    Case "2D"
            .Check2(0).Value = 1
            .Option2(0).Value = True
            .Check2(2).Value = 1
            .Frame19(1).Visible = True
            .Check2(0).Enabled = False
            .Combo2(0).Enabled = False
            .TDBNumber1(0).Enabled = False
            .Option2(0).Enabled = False
            .Option2(1).Enabled = False
            .Check2(2).Enabled = False
    Case "2E"
            .Check2(0).Value = 1
            .Option2(0).Value = True
            .Option3(0).Value = True
            .Check2(2).Value = 1
            .Frame19(1).Visible = True
            .Frame23.Visible = True
            .Check2(3).Visible = True
            .Check2(0).Enabled = False
            .TDBNumber1(0).Enabled = False
            .Combo2(0).Enabled = False
            .Option2(0).Enabled = False
            .Option2(1).Enabled = False
            If .Check2(3).Value = 0 Then
                .Frame23.Enabled = False
            End If
    Case "2F"
            .Check2(0).Value = 1
            .Option2(0).Value = True
            .Option3(1).Value = True
            .Check2(2).Value = 1
            .Frame19(1).Visible = True
            .Check2(0).Enabled = False
            .Combo2(0).Enabled = False
            .TDBNumber1(0).Enabled = False
            .Option2(0).Enabled = False
            .Option2(1).Enabled = False
            .Check2(2).Enabled = False
    Case "30"
            .Check2(0).Value = 1
            .Option2(0).Value = True
            .Option3(0).Value = True
            .Check2(2).Value = 1
            .Check2(3).Value = 1
            .Frame19(1).Visible = True
            .Frame23.Visible = True
            .Check2(3).Visible = True
            .Check2(0).Enabled = False
            .Combo2(0).Enabled = False
            .TDBNumber1(0).Enabled = False
            .Option2(0).Enabled = False
            .Option2(1).Enabled = False
            .Check2(2).Enabled = False
            .Option3(0).Enabled = False
            .Option3(1).Enabled = False
 End Select
If UCase(MDIForm1.Text2.Text) = "AGENT" Then
    .Frame23.Visible = False
    .Check2(3).Visible = False
End If
End With
End Function

