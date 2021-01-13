Attribute VB_Name = "Encrypt1"
Private mstrKey As String
Private mstrText As String

'~~~.DoXor
'Exclusive-or method to encrypt or decrypt
Private Sub DoXor()
On Error Resume Next
     Dim lngC As Long
     Dim intB As Long
     Dim lngN As Long
     For lngN = 1 To Len(mstrText)
         lngC = Asc(Mid(mstrText, lngN, 1))
         intB = Int(Rnd * 256)
         Mid(mstrText, lngN, 1) = Chr(lngC Xor intB)
     Next lngN
End Sub

'~~~.Stretch
'Convert any string to a printable, displayable string
Private Sub Stretch()
On Error Resume Next
     Dim lngC As Long
     Dim lngN As Long
     Dim lngJ As Long
     Dim lngK As Long
     Dim lngA As Long
     Dim strB As String
     lngA = Len(mstrText)
     strB = Space(lngA + (lngA + 2) \ 3)
     For lngN = 1 To lngA
         lngC = Asc(Mid(mstrText, lngN, 1))
         lngJ = lngJ + 1
         Mid(strB, lngJ, 1) = Chr((lngC And 63) + 59)
         Select Case lngN Mod 3
         Case 1
             lngK = lngK Or ((lngC \ 64) * 16)
         Case 2
             lngK = lngK Or ((lngC \ 64) * 4)
         Case 0
             lngK = lngK Or (lngC \ 64)
             lngJ = lngJ + 1
             Mid(strB, lngJ, 1) = Chr(lngK + 59)
             lngK = 0
         End Select
     Next lngN
     If lngA Mod 3 Then
         lngJ = lngJ + 1
         Mid(strB, lngJ, 1) = Chr(lngK + 59)
     End If
     mstrText = strB
End Sub
'~~~.Shrink
'Inverse of the Stretch method;
'result can contain any of the 256-byte values
Private Sub Shrink()
On Error Resume Next
     Dim lngC As Long
     Dim lngD As Long
     Dim lngE As Long
     Dim lngA As Long
     Dim lngB As Long
     Dim lngN As Long
     Dim lngJ As Long
     Dim lngK As Long
     Dim strB As String
     lngA = Len(mstrText)
     lngB = lngA - 1 - (lngA - 1) \ 4
     strB = Space(lngB)
     For lngN = 1 To lngB
         lngJ = lngJ + 1
         lngC = Asc(Mid(mstrText, lngJ, 1)) - 59
         Select Case lngN Mod 3
         Case 1
             lngK = lngK + 4
             If lngK > lngA Then lngK = lngA
             lngE = Asc(Mid(mstrText, lngK, 1)) - 59
             lngD = ((lngE \ 16) And 3) * 64
         Case 2
             lngD = ((lngE \ 4) And 3) * 64
         Case 0
             lngD = (lngE And 3) * 64
             lngJ = lngJ + 1
         End Select
         Mid(strB, lngN, 1) = Chr(lngC Or lngD)
     Next lngN
     mstrText = strB
End Sub
'Initializes random numbers using the key string
Private Sub Initialize()
     Dim lngN As Long
     Randomize Rnd(-1)
     For lngN = 1 To Len(mstrKey)
         Randomize Rnd(-Rnd * Asc(Mid(mstrKey, lngN, 1)))
     Next lngN
End Sub

Public Function Encrypt(ByVal sKey As String, ByVal sPlainText As String) As String 'call encrypt
     mstrKey = sKey 'The key is used to encrypt the text. it can be letters and numbers
     Call Initialize
     mstrText = sPlainText 'Plaintext is a textbox containg what you want to encrypt
     Call DoXor
     Call Stretch
     Encrypt = mstrText 'encryptedtext.text is where the encrypted text will show up
End Function
'use this to decrypt

Public Function Decrypt(ByVal sKey As String, ByVal sEncryptedText As String) As String 'call decrypt
     mstrKey = sKey 'keyusedtodecrypt.text is what is used to decrypt the text. It can be letters or numbers
     Call Initialize
     mstrText = sEncryptedText 'textyouwantdecrypted is a textbox containg the text you want decrypted
     Call Shrink
     Call DoXor
     Decrypt = mstrText 'decryptedtext.text is where the decrypted text will show up
End Function

