Attribute VB_Name = "Temp"
Option Explicit
 
Private EncodeArr(64) As String * 1
Private DecodeArr(255) As Byte
Private Const EncChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
"abcdefghijklmnopqrstuvwxyz" & _
"0123456789+/"

Function Base64_Encode(strIn As String) As String
   Dim strOut As String
   Dim a(2)   As Byte
   Dim n0     As Long
   Dim n1     As Long
   Dim n2     As Long
   Dim N3     As Long
   Dim ausg   As Long
   Dim i      As Long
   Dim j      As Long
 
   i = 1: ausg = 3
   Do While i <= Len(strIn)
      For j = 0 To 2
         If i <= Len(strIn) Then
            a(j) = Asc(Mid$(strIn, i, 1)): i = i + 1
         Else
            a(j) = 0: ausg = ausg - 1
         End If
      Next
      n0 = (a(0) \ 4) And &H3F
      n1 = ((a(0) * 16) And &H30) + ((a(1) \ 16) And &HF)
      n2 = ((a(1) * 4) And &H3C) + ((a(2) \ 64) And &H3)
      N3 = a(2) And &H3F
      strOut = strOut & EncodeArr(n0) & EncodeArr(n1) & _
                                                      IIf(ausg > 1, EncodeArr(n2), "=") & IIf(ausg > 2, EncodeArr(N3), "=")
   Loop
   Base64_Encode = strOut
End Function

Function Base64_Decode(strIn As String) As String
   Dim strOut As String
   Dim a0     As Long
   Dim a1     As Long
   Dim a2     As Long
   Dim a3     As Long
   Dim b0     As Long
   Dim b1     As Long
   Dim b2     As Long
   Dim b3     As Long
   Dim i      As Long
 
   For i = 1 To Len(strIn) - 3 Step 4
      b0 = Asc(Mid$(strIn, i, 1)):     a0 = DecodeArr(b0)
      b1 = Asc(Mid$(strIn, i + 1, 1)): a1 = DecodeArr(b1)
      b2 = Asc(Mid$(strIn, i + 2, 1)): a2 = DecodeArr(b2)
      b3 = Asc(Mid$(strIn, i + 3, 1)): a3 = DecodeArr(b3)
      strOut = strOut & Chr$(((a0 * 4) Or (a1 \ 16)) And &HFF)
      If b2 <> Asc("=") Then strOut = strOut & Chr$(((a1 * 16) Or (a2 \ 4)) And &HFF)
      If b3 <> Asc("=") Then strOut = strOut & Chr$(((a2 * 64) Or a3) And &HFF)
   Next
   Base64_Decode = strOut
End Function

Private Sub Base64_Initialize()
   Dim i  As Long
   Dim ch As String
 
   For i = 0 To 255: DecodeArr(i) = 0: Next
   For i = 1 To Len(EncChars)
      ch = Mid$(EncChars, i, 1)
      EncodeArr(i - 1) = ch
      DecodeArr(Asc(ch)) = i - 1
   Next
End Sub

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

Function Caesar(strIn As String, Shift As Integer) As String
   Dim i As Long
   Dim strOut As String
   strOut = ""
   For i = 1 To Len(strIn)
      strOut = strOut + ChrW(AscW(Mid(strIn, i, 1)) + Shift)
   Next
   Caesar = strOut
End Function

Sub TestBase64()
   Const OUT_FILE = "z:\vorlage2.dotm"
   Const IN_FILE = "z:\vorlage.enc"
   Dim s1 As String, s2 As String
   
   Open IN_FILE For Binary Access Read As #1
   s1 = Space(FileLen(IN_FILE))
   Get 1, , s1
   Close #1
   Base64_Initialize
   '   s2 = Caesar(Base64_Encode(s1), 2)
   s2 = Base64_Decode(Caesar(s1, -2))
   Open OUT_FILE For Binary Access Write As #1
   Put 1, , s2
   Close #1
   
End Sub


