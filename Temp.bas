Attribute VB_Name = "Temp"
Sub SucheinDatei()
   Const FILE_NAME = "z:\vbaProject.bin"
   Dim dat As Integer
   Dim bytes() As Byte
   dat = FreeFile
   Open FILE_NAME For Binary Access Read As #dat
   ReDim bytes(LOF(dat))
      Get #dat, , bytes()
   Close #dat
   For i = 0 To UBound(bytes) - 2
      If bytes(i) = Asc("D") And _
         bytes(i + 1) = Asc("P") And _
         bytes(i + 2) = Asc("x") Then
         Pr ("Signature found at: " & Hex(i))
         bytes(i + 2) = Asc("x")
         Pr (bytes(i) & bytes(i + 1))
      End If
   Next i
   Open FILE_NAME For Binary Access Write As #dat
      Put #dat, , bytes()
   Close #dat
End Sub

Function SubArray(arr() As Byte, index1 As Long, index2 As Long) As Variant
   Dim tmp() As Byte
   ReDim tmp(index2 - index1)
   For i = 0 To (index2 - index1)
      tmp(i) = arr(i)
   Next i
   SubArray = tmp
End Function

Sub testsubarr()
   Dim g(15) As Byte
   Dim h() As Byte

   For i = 0 To 15
      g(i) = i + 65
   Next i

   h = SubArray(g, 3, 5)
   Pr StrConv(h, vbUnicode)
End Sub

