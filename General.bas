Attribute VB_Name = "General"
'******************** General Variables and Routines ********************
Option Explicit
Public arr() As Variant
Public regex As New RegExp

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
   Dim bytes() As Byte, soutput As String, i As Long
   bytes = StrConv(str, vbFromUnicode)
   soutput = ""
   For i = 0 To UBound(bytes)
      soutput = soutput & " " & CStr(bytes(i))
   Next i
   Debug.Print soutput
End Sub

Sub PrH(str As String)
   Dim bytes() As Byte
   bytes = StrConv(str, vbFromUnicode)
   Output = ""
   For i = 0 To UBound(bytes)
      If Len(CStr(Hex(bytes(i)))) = 1 Then
         Output = Output & " 0" & CStr(Hex(bytes(i)))
      Else: Output = Output & " " & CStr(Hex(bytes(i)))
      End If
   Next i
   Debug.Print Output
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

Function AscPos(src As String, Pos As Long) As Long
   AscPos = AscW(Mid(src, Pos, 1))
End Function

Function CPos(src As String, Pos As Long) As String
   CPos = Mid(src, Pos, 1)
End Function


Function MakeString(str As String) As String     ' Wandelt String mit #NUM in String mit ChrW(NUM) um
   Dim res As String, num As String
   Dim i As Long, j As Long
   res = ""
   i = 1
   While i <= Len(str)
      If CPos(str, i) = "#" Then
         If CPos(str, i + 1) = "#" Then
            res = res & "#"
            i = i + 1
         End If
         If AscW(CPos(str, i + 1)) > 47 And AscW(CPos(str, i + 1)) < 58 Then
            num = ""
            j = 1
            Do While (i + j <= Len(str)) And AscW(CPos(str, i + j)) > 47 And AscW(CPos(str, i + j)) < 58
               num = num + CStr(AscW(CPos(str, i + j)) - 48)
               j = j + 1
               If i + j > Len(str) Then Exit Do
            Loop
            res = res + ChrW(CInt(num))
            i = i + j - 1
         End If
      Else
         res = res + CPos(str, i)
      End If
      i = i + 1
   Wend
   MakeString = res
End Function

'******************** Dialogs & Commands ********************
Sub Einlesen()                              'Liest Dialognamen ein
   Dim i As Integer
   i = 0
   ReDim arr(1)
   Open NormalTemplate.PATH + "\dlglist.txt" For Input As #1 ' Open file for input.
   Do While Not EOF(1)                           ' Loop until end of file.
      Line Input #1, arr(i)                      ' read next line from file and add text to the array
      i = i + 1
      ReDim Preserve arr(i)                      ' Redim the array for the new element
   Loop
   Close #1                                      ' Close file.
End Sub

Sub BefehlslisteLaden()                          'Liest Befehlsliste ein
   Dim cItem As Variant
   If colCmd.Count = 0 Then FillcolCmd
   With frmCommand.ListBox1
      .Clear
      For Each cItem In colCmd
         .AddItem cItem
      Next cItem
   End With
End Sub

Sub DlgAufrufen()
   Dim inp, liste As String, i, cnt As Integer
   On Error Resume Next
   i = colDlg.Count
   If colDlg.Count = 0 Then FillcolDlg
   inp = InputBox("Dialog-Nr.:", "Dialoge aufrufen")
   If inp = "" Then Exit Sub
   If IsNumeric(inp) Then
      Dialogs(inp).Show
   Else
      liste = ""
      cnt = 0
      For i = 1 To colDlg.Count
         If InStr(LCase(colDlg(i)), LCase(inp)) Then
            liste = liste & colDlg(i) & vbCr
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
      Case "docprop": frmDP.Show
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
   frmCommand.Show
   frmCommand.ListBox1.SetFocus
End Sub

Sub frmRegExp_Anzeigen()
   frmRegExp.Show
End Sub

Sub ShowfrmKey()
   frmKey.Show
End Sub

Sub ToggleStyleInspector()
   If Application.TaskPanes(wdTaskPaneStyleInspector).Visible = True Then _
      Application.TaskPanes(wdTaskPaneStyleInspector).Visible = False Else _
      Application.TaskPanes(wdTaskPaneStyleInspector).Visible = True
End Sub

Sub H2D()
   Dim nr As Integer
   nr = CInt("&h" + InputBox("Hex"))
   StatusBar = nr
End Sub


Sub RegisterHotkeys()
   CustomizationContext = NormalTemplate
   KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, 192), KeyCategory:=wdKeyCategoryCommand, Command:="ShowfrmKey"
   KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyComma), KeyCode2:=wdKeyL, KeyCategory:=wdKeyCategoryCommand, Command:="BefehllisteAnzeigen"
   KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyComma), KeyCode2:=wdKeyR, KeyCategory:=wdKeyCategoryCommand, Command:="frmRegExp_Anzeigen"
   KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyComma), KeyCode2:=wdKeyD, KeyCategory:=wdKeyCategoryCommand, Command:="DlgAufrufen"
   KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyComma), KeyCode2:=wdKeyK, KeyCategory:=wdKeyCategoryCommand, Command:="Kommandos"
   KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyComma), KeyCode2:=wdKeyV, KeyCategory:=wdKeyCategoryCommand, Command:="ViewSettings"
   KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyComma), KeyCode2:=wdKeyC, KeyCategory:=wdKeyCategoryCommand, Command:="CharCode"
   KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeySpacebar), KeyCategory:=wdKeyCategoryCommand, Command:="CheckWord"
   KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyComma), KeyCode2:=wdKeyI, KeyCategory:=wdKeyCategoryCommand, Command:="ToggleStyleInspector"
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

Sub CharCode()                                   ' Zeigt Code des Zeichens links vom Cursor
   Dim c As String
   Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
   c = Selection.Characters(1)
   Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdMove
   StatusBar = """" + c + """ = " & AscW(c) & " = " & Hex(AscW(c)) & "h" & " = U+" & right("0000" & CStr(Hex(AscW(c))), 4)
End Sub

Sub SetParText(ByRef p As Paragraph, txt As String)
   Dim Rng As Range
   Set Rng = p.Range
   Rng.MoveEnd wdCharacter, 0
   Rng.Text = txt
End Sub

Sub Mark_Paragraphs()
' Marks the beginning and the end of a paragraph with 170 � and 186 �
   Dim p As Paragraph
   Dim r As Range
   Dim undo As UndoRecord
   Set undo = Application.UndoRecord
   undo.StartCustomRecord ("pared")
   If Selection.Type = wdSelectionIP Then Set r = ActiveDocument.Range Else Set r = Selection.Range
   For Each p In r.Paragraphs
      p.Range.Select
      Selection.Range.InsertBefore "�"
      Selection.MoveEnd Unit:=wdCharacter, Count:=-1
      Selection.InsertAfter "�"
   Next p
   r.Select
   undo.EndCustomRecord
End Sub

Sub FindReplace(Rng As Range, Fnd As String, rpl As String)
   Dim undo As UndoRecord
   Set undo = Application.UndoRecord
   With Rng.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .MatchWildcards = False
      .Text = Fnd
      .Replacement.Text = rpl
      .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
   End With
   undo.EndCustomRecord
End Sub

Sub DeMark_Paragraphs()
' DeMarks the beginning and the end of a paragraph with 170 � and 186 �
  FindReplace Rng:=ActiveDocument.Range, Fnd:="�", rpl:=""
  FindReplace Rng:=ActiveDocument.Range, Fnd:="�", rpl:=""
End Sub

Private Sub testfind()
  Dim dlg As Dialog
  Set dlg = Dialogs(wdDialogEditReplace)
  With dlg
   .Find = "Stichwort"
   .Replace = "^&"
   .PatternMatch = 1
   .Wrap = 1
   .Show
  End With
  
Function SimpleRegEx(re As String) As String
   Replace("\d",re,"[0-9]")
   Replace("\D",re,"[^0-9]")
   Replace("\w",re," [a-zA-Z0-9_")
   Replace("\W",re," [^a-zA-Z0-9_")
   Replace("\a",re,"[a-z]")
   Replace("\A",re,"[A-Z]")
   Replace("\D",re,"[^0-9]")
End Function

Sub dollarHighlighter()
    Set RegExp = New RegExp
    Set Regexp2 = New RegExp
    Dim objMatch As Match
    Dim colMatches As MatchCollection
    Dim myrange As Range
    Dim offsetStart As Long
    offsetStart = 0 ' Selection.Start

    RegExp.Pattern = "bietet"
    RegExp.Global = True
    RegExp.IgnoreCase = True
    Set colMatches = RegExp.Execute(ActiveDocument.Range.Text)   ' Execute search.
    
    For Each objMatch In colMatches   ' Iterate Matches collection.
      Set myrange = ActiveDocument.Range(objMatch.FirstIndex + offsetStart, End:=offsetStart + objMatch.FirstIndex + objMatch.Length)
      myrange.FormattedText.Text = "TEST"
      myrange.ParagraphFormat.Shading.BackgroundPatternColor = wdColorBlueGray
    Next
End Sub
