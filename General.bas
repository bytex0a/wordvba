Attribute VB_Name = "General"
'******************** General Variables and Routines ********************
Option Explicit
Public arr() As Variant
Public Befehle(3290) As String
Public regex As New regexp

Sub AutoExec()                                   'Run automatically
   Einlesen
End Sub

Function Pr(s As Variant)                        'Shortcut for Debug.Print
   Debug.Print s
End Function

Sub PrAst(str As String)
   Debug.Print ("*" & str & "*")
End Sub

Sub PrBy(str As String)
   Dim bytes() As Byte
   bytes = StrConv(str, vbFromUnicode)
   output = ""
   For i = 0 To UBound(bytes)
      output = output & " " & CStr(bytes(i))
   Next i
   Debug.Print output
End Sub

Sub PrH(str As String)
   Dim bytes() As Byte
   bytes = StrConv(str, vbFromUnicode)
   output = ""
   For i = 0 To UBound(bytes)
      If Len(CStr(Hex(bytes(i)))) = 1 Then
         output = output & " 0" & CStr(Hex(bytes(i)))
      Else: output = output & " " & CStr(Hex(bytes(i)))
      End If
   Next i
   Debug.Print output
End Sub


Sub Inc(ByRef ival): ival = ival + 1: End Sub
Sub Dec(ByRef ival): ival = ival - 1: End Sub

'******************** Find-Replace ********************
Sub SearchReplace(fe, re, sty As String)
   With Selection.Find
      .ClearFormatting
      .MatchWildcards = True
      .Text = fe
      .Replacement.ClearFormatting
      .Format = True
      .Replacement.style = ActiveDocument.Styles(sty)
      .Replacement.Text = re
      .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
   End With
End Sub

'******************** RegExp-Functions ********************
Public Function RxTest(ByVal str, ByVal pat As String) As Boolean
   regex.Pattern = pat
   RxTest = regex.test(str)
End Function

Public Function RxReplace(ByVal str, ByVal pat, ByVal rpl As String) As String
   regex.Global = True
   regex.Pattern = pat
   RxReplace = regex.Replace(str, rpl)
End Function

Public Function RxExecute(ByVal str, ByVal pat As String, Optional glob As Boolean = True, Optional multi As Boolean = False) As String
   regex.Global = glob
   regex.MultiLine = multi
   regex.Pattern = pat
   RxExecute = regex.Execute(str)
End Function

'******************** Dialogs & Commands ********************
Function Einlesen()                              'Liest Dialognamen ein
   Dim i As Integer
   i = 0
   ReDim arr(1)
   Open NormalTemplate.Path + "\dlglist.txt" For Input As #1 ' Open file for input.
   Do While Not EOF(1)                           ' Loop until end of file.
      Line Input #1, arr(i)                      ' read next line from file and add text to the array
      i = i + 1
      ReDim Preserve arr(i)                      ' Redim the array for the new element
   Loop
   Close #1                                      ' Close file.
End Function

Sub BefehlslisteLaden()                          'Liest Befehlsliste ein
   CommandForm.ListBox1.Clear
   Dim Count As Integer
   Dim entry As String
   Count = 1
   Open NormalTemplate.Path + "\Befehlsliste.txt" For Input As #1
   While EOF(1) = False
      Input #1, entry
      Befehle(Count) = entry
      CommandForm.ListBox1.AddItem (entry)
      Count = Count + 1
   Wend
   Close #1
   Open NormalTemplate.Path + "\CommandList.txt" For Input As #1
   While EOF(1) = False
      Input #1, entry
      Befehle(Count) = entry
      CommandForm.ListBox1.AddItem (entry)
      Count = Count + 1
   Wend
   Close #1
End Sub

Sub DlgAufrufen()
   Dim inp, liste As String, i, cnt As Integer
   On Error Resume Next
   i = UBound(arr)
   If Err.Number = 9 Then
      Call Einlesen
      StatusBar = "Lese Dialoge ein..."
   End If
   inp = InputBox("Dialog-Nr.:", "Dialoge aufrufen")
   If inp = "" Then Exit Sub
   If IsNumeric(inp) Then
      Dialogs(inp).Show
   Else
      liste = ""
      cnt = 0
      For i = 0 To UBound(arr)
         If InStr(LCase(arr(i)), LCase(inp)) Then
            liste = liste & arr(i) & vbCr
            cnt = cnt + 1
         End If
      Next i
      InStr
      If liste <> "" Then
         If cnt = 1 Then
            Dialogs(Int(Mid(liste, 1, 3))).Show
            StatusBar = "Dialog # " & Mid(liste, 1, 3)
         Else
            Options.EnableSound = False
            MsgBox liste
            DlgAufrufen
         End If
      End If
   End If
End Sub

Function GetPoints(ca() As String) As Single
   If UBound(ca) = 2 Then If ca(2) = "cm" Then GetPoints = CentimetersToPoints(ca(1))
   If UBound(ca) = 1 Then GetPoints = CInt(ca(1))
End Function

Function GetLinePoints(ca() As String) As Single
   If UBound(ca) = 2 Then If ca(2) = "cm" Then GetLinePoints = CentimetersToPoints(ca(1))
   If UBound(ca) = 1 Then GetLinePoints = (CInt(ca(1)))
End Function

Sub Kommandos()
   Dim com, s As String, sp As Integer, par As Paragraph, comarr() As String
   com = InputBox("Kommando eingeben", "Komanndo")
   If InStr(com, " ") = 0 Then
   Select Case com
   Case "hp":
      s = "Horizontale Position : " & vbCr & Round(Application.Selection.Information(wdHorizontalPositionRelativeToTextBoundary) _
                                                   / 72 * 2.54, 2) & "cm / " & Application.Selection.Information(wdHorizontalPositionRelativeToTextBoundary) & "pt. (relative to Text Boundary)" _
                                    & vbCr & Round(Application.Selection.Information(wdHorizontalPositionRelativeToPage) _
                                                   / 72 * 2.54, 2) & "cm / " & Application.Selection.Information(wdHorizontalPositionRelativeToPage) & "pt. (relative to Page)"
      MsgBox s
   Case "rds": Application.Run ("RedefineStyle")
   End Select
   Else
      comarr = Split(com)
      Select Case comarr(0)
      Case "ctp": MsgBox (CentimetersToPoints(CSng(comarr(1))) & " pt.")
      Case "ptc": MsgBox (PointsToCentimeters(CSng(comarr(1))) & " cm")
      Case "pa": For Each par In Selection.Paragraphs: par.SpaceAfter = GetPoints(comarr): Next
      Case "pb": For Each par In Selection.Paragraphs: par.SpaceBefore = GetPoints(comarr): Next
      Case "pse": For Each par In Selection.Paragraphs: par.LineSpacingRule = wdLineSpaceExactly: par.LineSpacing = GetLinePoints(comarr): Next
      Case "psm": For Each par In Selection.Paragraphs: par.LineSpacingRule = wdLineSpaceMultiple: par.LineSpacing = LinesToPoints((CSng(comarr(1)))): Next
      End Select
   End If
End Sub

Sub BefehllisteAnzeigen()
   CommandForm.Show
   CommandForm.ListBox1.SetFocus
End Sub

Sub ShowKeyForm()
   KeyForm.Show
End Sub

Sub RegisterHotkeys()
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, 192), KeyCategory:=wdKeyCategoryCommand, Command:="ShowKeyForm"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyComma), KeyCode2:=wdKeyL, KeyCategory:=wdKeyCategoryCommand, Command:="BefehllisteAnzeigen"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyComma), KeyCode2:=wdKeyD, KeyCategory:=wdKeyCategoryCommand, Command:="DlgAufrufen"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyComma), KeyCode2:=wdKeyK, KeyCategory:=wdKeyCategoryCommand, Command:="Kommandos"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyComma), KeyCode2:=wdKeyE, KeyCategory:=wdKeyCategoryCommand, Command:="LoopEdit"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeySpacebar), KeyCategory:=wdKeyCategoryCommand, Command:="CheckWord"
End Sub

'******************** ListTemplate Functions ********************

Public Function ListTemplateIndex(ListTemplateName As String, _
  Source As Word.Document) As ListTemplate
  Dim lt As ListTemplate
  
  For Each lt In Source.ListTemplates
    If lt.Name = ListTemplateName Then
      Set ListTemplateIndex = lt
      Exit For
    End If
  Next
 
  If ListTemplateIndex Is Nothing Then
    '"True" bedeutet, ListTemplate hat neun Ebenen
    Set ListTemplateIndex = Source.ListTemplates.Add(True)
    ListTemplateIndex.Name = ListTemplateName
  End If
  Set lt = Nothing
End Function

Sub RegisterListtemplateRZ()
   Dim lt As ListTemplate
   Set lt = ListTemplateIndex("RzList", ActiveDocument)
   With lt.ListLevels(1)
      .NumberFormat = "%1"
      .TextPosition = CentimetersToPoints(1)
   End With
End Sub

