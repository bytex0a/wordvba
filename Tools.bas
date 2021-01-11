Attribute VB_Name = "Tools"
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

' *************************************** Partnerleiste repositioinieren
' 134882 CAESAR Telefonie - PartnerleisteCAE:: CTI:: Partnerlist

'Option Explicit
'Global r As Long
'
'Private Const GWL_STYLE  As Long = -16&
'Private Const WS_CAPTION As Long = &HC00000
'
'Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'Private Declare Function GetWindowLongW Lib "user32.dll" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function GetWindowTextLengthW Lib "user32.dll" (ByVal hWnd As Long) As Long
'Private Declare Function GetWindowTextW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long
'Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hWnd As Long) As Long
'Private Declare Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As Long
'Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
'Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
'         ByVal cy As Long, ByVal uFlags As Long) As Long
'Type RECT
'   left     As Long
'   top      As Long
'   right    As Long
'   bottom   As Long
'End Type
'
'
'
'Sub main()
'    r = 1
'    Call EnumWindows(AddressOf EnumWindowsProc, ByVal 0&)
'End Sub
'
'
'
'Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
'
'    Dim windowText  As String
'    Dim windowClass As String * 256
'    Dim retVal      As Long
'    Dim l           As Long
'
'
'    windowText = Space(GetWindowTextLength(hWnd) + 1)
'    retVal = GetWindowText(hWnd, windowText, Len(windowText))
'    windowText = left$(windowText, retVal)
'
'
'    retVal = GetClassName(hWnd, windowClass, 255)
'    windowClass = left$(windowClass, retVal)
'
'
'
'   InsertAfter (CStr(hWnd) & windowText & windowClass)
'
'    r = r + 1
'
'  '
'  ' Return true to indicate that we want to continue
'  ' with the enumeration of the windows:
'  '
'    EnumWindowsProc = True
'
'End Function ' }
'
'
'Sub getr()
'   Dim rt As RECT
'   Pr GetWindowRect(134882, rt)
'   Debug.Print rt.left, " ", rt.top, " ", rt.right, " ", rt.bottom
'   Pr SetWindowPos(134882, 0, 100, 100, 304, 105, &H40)
'
'End Sub
'

'Alle Shortcuts auflisten
Sub ListCompositeShortcuts()
Dim oDoc As Word.Document
Dim oDocTemp As Word.Document
Dim oKey As KeyBinding
  'Create a new document for listing composite shortcuts.
  Set oDoc = Documents.Add(, , wdNewBlankDocument)
  Set oRng = oDoc.Range
  System.Cursor = wdCursorWait
  Application.ScreenUpdating = False
  CustomizationContext = NormalTemplate 'or the template\document to evaluate.
  'List and sort custom keybindings.
  For lngIndex = 1 To KeyBindings.Count
    Set oKey = KeyBindings(lngIndex)
    oRng.InsertAfter vbCr & oKey.KeyCategory & vbTab & oKey.Command _
                   & vbTab & oKey.KeyString
    'Update status bar.
    Application.StatusBar = "Processing custom keybinding " & lngIndex & " of " & _
                             KeyBindings.Count & ".  Please wait."
    DoEvents
  Next lngIndex
End Sub


Sub PrintAllStyles()
Dim sty As style
   For Each sty In ActiveDocument.Styles
      If sty.Type = wdStyleTypeTable Then _
      Debug.Print (sty.NameLocal)
   Next
End Sub

Sub Respace()
   Dim p As Paragraph
   r = 0.8
   For Each p In ActiveDocument.Paragraphs
      p.SpaceAfter = p.SpaceAfter * r
      p.SpaceBefore = p.SpaceBefore * r
   Next
End Sub

