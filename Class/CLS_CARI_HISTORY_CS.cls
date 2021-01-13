VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_CARI_HISTORY_CS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub CARI_STATUS_CUSTOMER(TYPE_PRODUK As String)
Dim NAMA_FORM As Object

Select Case UCase(TYPE_PRODUK)
Case "CREDIT CARD"
    Set NAMA_FORM = FRMCUST_CC
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
        If .Check9(0).Value = 1 Then
                .STATUS_CUSTOMER = "4"
                .KETHSLKERJA = "1H"
                .STATUS_FIELD_LAMA = "2A"
                .DOK1 = Empty
            If .Check9(1).Value = 1 Then
                    .STATUS_CUSTOMER = "4"
                    .KETHSLKERJA = "1I"
                    .STATUS_FIELD_LAMA = "2A"
                    .DOK1 = Empty
                If .Check9(2).Value = 1 Then
                        .STATUS_CUSTOMER = "4"
                        .KETHSLKERJA = "1J"
                        .STATUS_FIELD_LAMA = "2A"
                        .DOK1 = Empty
                End If
            End If
        End If
    Else
        If .Option2(1).Value Then
            .STATUS_CUSTOMER = "2"
            .KETHSLKERJA = "1G"
            .STATUS_FIELD_LAMA = "1B"
            .DOK1 = Empty
        End If
    End If
End If

If .Check2(2).Value = 1 Then
    .STATUS_CUSTOMER = "4"
    .STATUS_FIELD_LAMA = "2A"
    .KETHSLKERJA = "2D"
    
    If .Option3(0).Value Then
        .KETHSLKERJA = "2E"
        .STATUS_CUSTOMER = "5"
        .STATUS_FIELD_LAMA = "2C"
        .DOK1 = Empty
    Else
        If .Option3(1).Value Then
            .KETHSLKERJA = "2F"
            .STATUS_CUSTOMER = "6"
            .STATUS_FIELD_LAMA = "2A"
            If .Check9(3).Value = 1 Then
                .DOK1 = "A"
            End If
                If .Check9(4).Value = 1 Then
                    If Len(.DOK1) > 0 Then
                        .DOK1 = .DOK1 + ";" + "B"
                    Else
                        .DOK1 = "B"
                    End If
                End If
                    If .Check9(5).Value = 1 Then
                        If Len(.DOK1) > 0 Then
                            .DOK1 = .DOK1 + ";" + "C"
                        Else
                            .DOK1 = "C"
                        End If
                    End If
                        If .Check9(6).Value = 1 Then
                            If Len(.DOK1) > 0 Then
                                .DOK1 = .DOK1 + ";" + "D"
                            Else
                                .DOK1 = "D"
                            End If
                        End If
                            If .Check9(7).Value = 1 Then
                                If Len(.DOK1) > 0 Then
                                    .DOK1 = .DOK1 + ";" + "E"
                                Else
                                    .DOK1 = "E"
                                End If
                            End If
                                If .Check9(8).Value = 1 Then
                                    If Len(.DOK1) > 0 Then
                                        .DOK1 = .DOK1 + ";" + "F"
                                    Else
                                        .DOK1 = "F"
                                    End If
                                End If
                                If .Check9(9).Value = 1 Then
                                    If Len(.DOK1) > 0 Then
                                        .DOK1 = .DOK1 + ";" + "G"
                                    Else
                                        .DOK1 = "G"
                                    End If
                                End If
                                If .Check9(10).Value = 1 Then
                                    If Len(.DOK1) > 0 Then
                                        .DOK1 = .DOK1 + ";" + "H"
                                    Else
                                        .DOK1 = "H"
                                    End If
                                End If
                                    If .Check9(11).Value = 1 Then
                                    If Len(.DOK1) > 0 Then
                                        .DOK1 = .DOK1 + ";" + "I"
                                    Else
                                        .DOK1 = "I"
                                    End If
                                    End If
                                        If .Check9(12).Value = 1 Then
                                        If Len(.DOK1) > 0 Then
                                            .DOK1 = .DOK1 + ";" + "J"
                                        Else
                                            .DOK1 = "J"
                                        End If
                                        End If
        End If
    End If
End If
If .Check2(3).Value Then
.STATUS_CUSTOMER = "5"
.KETHSLKERJA = "30"
.STATUS_FIELD_LAMA = "2C"
.DOK1 = Empty
    If .Option5(0).Value Then
        .STATUS_CUSTOMER = "7"
        .KETHSLKERJA = "3A"
        .STATUS_FIELD_LAMA = "3A"
        .DOK1 = Empty
    Else
        If .Option5(1).Value Then
            .STATUS_CUSTOMER = "8"
            .KETHSLKERJA = "3B"
            .STATUS_FIELD_LAMA = "XX"
            .DOK1 = Empty
        Else
            If .Option5(2).Value Then
                .KETHSLKERJA = "3C"
                .STATUS_CUSTOMER = "9"
                .STATUS_FIELD_LAMA = "3A"
                .DOK1 = Empty
'                If .Option6(0).Value Then
'                    .STATUS_CUSTOMER = "9"
 '                   .KETHSLKERJA = "3D"
  '                  .STATUS_FIELD_LAMA = "3A"
   '                 .DOK1 = Empty
    '            Else
     '               If .Option6(1).Value Then
      '                  .STATUS_CUSTOMER = "0"
       '                 .KETHSLKERJA = "3E"
        '                .STATUS_FIELD_LAMA = "XX"
         '               .DOK1 = Empty
          '          End If
           '     End If
            End If
        End If
    End If
End If
If .Check2(0).Value = 0 And .Check2(1).Value = 0 And .Check2(2).Value = 0 And .Check2(3).Value = 0 Then
    .STATUS_CUSTOMER = "1"
    .KETHSLKERJA = "1A"
    .STATUS_FIELD_LAMA = "1A"
End If
End With
End Sub

Public Function VIEW_HISTORY_KERJA(HISTORY As String, DOKUMEN As String, TYPE_PRODUK As String)

Dim NAMA_FORM As Variant

Select Case UCase(TYPE_PRODUK)
Case "KTA"
    Set NAMA_FORM = FRMCUST_KTA
Case "CREDIT CARD"
    Set NAMA_FORM = FRMCUST_CC
Case "KTA - CROSS SELL"
    Set NAMA_FORM = FRMCUST_CS
Case Else
    Exit Function
End Select

With NAMA_FORM
Select Case Left(HISTORY, 1)
    Case 0
        .Check2(1).Value = 1
            If UCase(MDIForm1.Text2.Text) = "AGENT" Then
     '           .Check2(1).Enabled = False
     '           .Frame25.Enabled = False
            End If
    Case Else
        .Check2(1).Value = 0
End Select
 Select Case HISTORY
                Case "0B"
                        .Check2(1).Value = 1
                        If UCase(MDIForm1.Text2.Text) = "AGENT" Then
                            .Check2(1).Enabled = False
                            .Frame25.Enabled = False
                        End If
                Case "1Z"
                        .Check2(0).Value = 1
                Case "1F"
                        .Check2(0).Value = 1
                        .Option2(0).Value = True
                        .Check9(1).Visible = False
                        .Check9(2).Visible = False
                Case "1H"
                        .Check2(0).Value = 1
                        .Option2(0).Value = True
                        .Check9(0).Value = 1
                        .Check2(0).Enabled = False
                        .Combo2(0).Enabled = False
                        .Option2(0).Enabled = False
                        .Option2(1).Enabled = False
                        .Check9(2).Visible = False
                Case "1I"
                        .Check2(0).Value = 1
                        .Option2(0).Value = True
                        .Check9(0).Value = 1
                        .Check9(1).Value = 1
                        .Check2(0).Enabled = False
                        .Combo2(0).Enabled = False
                        .Option2(0).Enabled = False
                        .Option2(1).Enabled = False
                        .Check9(0).Enabled = False
                Case "1J"
                        .Check2(0).Value = 1
                        .Option2(0).Value = True
                        .Check9(0).Value = 1
                        .Check9(1).Value = 1
                        .Check9(2).Value = 1
                        .Frame19(1).Visible = True
                        .Check2(0).Enabled = False
                        .Combo2(0).Enabled = False
                        .TDBNumber1(0).Enabled = False
                        .Option2(0).Enabled = False
                        .Option2(1).Enabled = False
                        .Check9(0).Enabled = False
                        .Check9(1).Enabled = False
                Case "1G"
                        .Check2(0).Value = 1
                        .Option2(1).Value = True
                        .Check9(0).Visible = False
                        .Check9(1).Visible = False
                        .Check9(2).Visible = False
                Case "2D"
                        .Check2(0).Value = 1
                        .Option2(0).Value = True
                        .Check9(0).Value = 1
                        .Check9(1).Value = 1
                        .Check9(2).Value = 1
                        .Check2(2).Value = 1
                        .Frame19(1).Visible = True
                        .Check2(0).Enabled = False
                        .Combo2(0).Enabled = False
                        .TDBNumber1(0).Enabled = False
                        .Option2(0).Enabled = False
                        .Option2(1).Enabled = False
                        .Check9(0).Enabled = False
                        .Check9(1).Enabled = False
                        .Check2(2).Enabled = False
                Case "2E"
                        .Check2(0).Value = 1
                        .Option2(0).Value = True
                        .Option3(0).Value = True
                        .Check9(0).Value = 1
                        .Check9(1).Value = 1
                        .Check9(2).Value = 1
                        .Check2(2).Value = 1
                        .Frame19(1).Visible = True
                        .Frame23.Visible = True
                        .Check2(3).Visible = True
                        .Check2(0).Enabled = False
                        .TDBNumber1(0).Enabled = False
                        .Combo2(0).Enabled = False
                        .Option2(0).Enabled = False
                        .Option2(1).Enabled = False
                        .Check9(0).Enabled = False
                        .Check9(1).Enabled = False
                        .Check2(2).Enabled = False
                        If .Check2(3).Value = 0 Then
                            .Frame23.Enabled = False
                        End If
                Case "2F"
                        .Check2(0).Value = 1
                        .Option2(0).Value = True
                        .Option3(1).Value = True
                        .Check9(0).Value = 1
                        .Check9(1).Value = 1
                        .Check9(2).Value = 1
                        .Check2(2).Value = 1
                            If Len(DOKUMEN) > 0 Then
                                Call VIEW_DOKUMEN(DOKUMEN, TYPE_PRODUK)
                            End If
                        .Frame19(1).Visible = True
'                        Check2(2).Visible = True
                        .Check2(0).Enabled = False
                        .Combo2(0).Enabled = False
                        .TDBNumber1(0).Enabled = False
                        .Option2(0).Enabled = False
                        .Option2(1).Enabled = False
                        .Check9(0).Enabled = False
                        .Check9(1).Enabled = False
                        .Check2(2).Enabled = False
                Case "30"
                        .Check2(0).Value = 1
                        .Option2(0).Value = True
                        .Option3(0).Value = True
                        .Check9(0).Value = 1
                        .Check9(1).Value = 1
                        .Check9(2).Value = 1
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
                        .Check9(0).Enabled = False
                        .Check9(1).Enabled = False
                        .Check2(2).Enabled = False
                        .Option3(0).Enabled = False
                        .Option3(1).Enabled = False
                Case "3A"
                        .Check2(0).Value = 1
                        .Option2(0).Value = True
                        .Option3(0).Value = True
                        .Option5(0).Value = True
                        .Check9(0).Value = 1
                        .Check9(1).Value = 1
                        .Check9(2).Value = 1
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
                        .Check9(0).Enabled = False
                        .Check9(1).Enabled = False
                        .Check2(2).Enabled = False
                        .Option3(0).Enabled = False
                        .Option3(1).Enabled = False
                        .Combo2(0).Enabled = False
                Case "3B"
                        .Check2(0).Value = 1
                        .Option2(0).Value = True
                        .Option3(0).Value = True
                        .Option5(1).Value = True
                        .Check9(0).Value = 1
                        .Check9(1).Value = 1
                        .Check9(2).Value = 1
                        .Check2(2).Value = 1
                        .Check2(3).Value = 1
                        .Frame19(1).Visible = True
                        .Frame23.Visible = True
                        .Check2(3).Visible = True
                        .Check2(0).Enabled = False
                        .Combo2(0).Enabled = False
                        .Option2(0).Enabled = False
                        .Option2(1).Enabled = False
                        .Check9(0).Enabled = False
                        .Check9(1).Enabled = False
                        .Check2(2).Enabled = False
                        .Option3(0).Enabled = False
                        .Option3(1).Enabled = False
                        .Combo2(0).Enabled = False
                Case "3C"
                        .Check2(0).Value = 1
                        .Option2(0).Value = True
                        .Option3(0).Value = True
                        .Option5(2).Value = True
                        .Check9(0).Value = 1
                        .Check9(1).Value = 1
                        .Check9(2).Value = 1
                        .Check2(2).Value = 1
                        .Check2(3).Value = 1
                        .Frame19(1).Visible = True
                        .Frame23.Visible = True
                        .Check2(3).Visible = True
                        .Check2(0).Enabled = False
                        .Combo2(0).Enabled = False
                        .Option2(0).Enabled = False
                        .Option2(1).Enabled = False
                        .Check9(0).Enabled = False
                        .Check9(1).Enabled = False
                        .Check2(2).Enabled = False
                        .Option3(0).Enabled = False
                        .Option3(1).Enabled = False
                        .Combo2(0).Enabled = False
                Case "3D"
                        .Check2(0).Value = 1
                        .Option2(0).Value = True
                        .Option3(0).Value = True
                        .Option5(2).Value = True
                        .Option6(0).Value = True
                        .Check9(0).Value = 1
                        .Check9(1).Value = 1
                        .Check9(2).Value = 1
                        .Check2(2).Value = 1
                        .Check2(3).Value = 1
                        .Frame19(1).Visible = True
                        .Frame23.Visible = True
                        .Check2(3).Visible = True
                        .Combo2(0).Enabled = False
                Case "3E"
                        .Check2(0).Value = 1
                        .Option2(0).Value = True
                        .Option3(0).Value = True
                        .Option5(2).Value = True
                        .Option6(1).Value = True
                        .Check9(0).Value = 1
                        .Check9(1).Value = 1
                        .Check9(2).Value = 1
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
                        .Check9(0).Enabled = False
                        .Check9(1).Enabled = False
                        .Check2(2).Enabled = False
                        .Option3(0).Enabled = False
                        .Option3(1).Enabled = False
                        .Option5(0).Enabled = False
                        .Option5(1).Enabled = False
                        .Option5(2).Enabled = False
                        .Combo2(0).Enabled = False
            End Select

If UCase(MDIForm1.Text1.Text) = "AGENT" Then
    .Frame23.Visible = False
    .Check2(3).Visible = False
End If
End With
End Function

Private Function VIEW_DOKUMEN(DOK As String, TYPE_PRODUK As String)
Dim M_DOK As String
Dim NAMA_FORM As Object
Select Case UCase(TYPE_PRODUK)
Case "KTA"
    Set NAMA_FORM = FRMCUST_KTA
Case "CREDIT CARD"
    Set NAMA_FORM = FRMCUST_CC
Case "KTA - CROSS SELL"
    Set NAMA_FORM = FRMCUST_CS
Case Else
    Exit Function
End Select

With NAMA_FORM

For I = 1 To Len(DOK)
    M_DOK = Mid(DOK, I, 1)
    Select Case M_DOK
        Case "A"
            .Check9(3).Value = 1
        Case "B"
            .Check9(4).Value = 1
        Case "C"
            .Check9(5).Value = 1
        Case "D"
            .Check9(6).Value = 1
        Case "E"
            .Check9(7).Value = 1
        Case "F"
            .Check9(8).Value = 1
        Case "G"
            .Check9(9).Value = 1
        Case "H"
            .Check9(10).Value = 1
        Case "I"
            .Check9(11).Value = 1
        Case "J"
            .Check9(12).Value = 1
    End Select
    I = I + 1
Next I
End With
End Function

