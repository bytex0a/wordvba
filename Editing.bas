Attribute VB_Name = "Editing"
Option Explicit
Public lastpat As String

'******************** Editing ******************************


Sub Smallcaps()
   Dim s, s2 As String, i, c As Integer
   s = Selection.Text
   s2 = ""
   For i = 1 To Len(s)
      c = AscW(Mid(s, i, 1))
      If (c >= &H61) And (c <= &H7A) Then
         s2 = s2 & ChrW(c + &HF700)
      Else
         s2 = s2 & ChrW(c)
      End If
   Next i
   Selection.Text = s2
End Sub

Sub Randnummern_Erstellen()
   Dim p As Paragraph
   Dim Rng As Range
   Dim objUndo As UndoRecord
   Set objUndo = Application.UndoRecord
   objUndo.StartCustomRecord ("Undo")
   RegisterListtemplateRZ                        'ListTemplate für Randziffern erstellen
   Application.ScreenUpdating = False
   If Selection.Paragraphs.Count = 1 Then Set Rng = ActiveDocument.Range Else Set Rng = Selection.Range
   For Each p In Rng.Paragraphs
      'If RxTest(p.Range.Text, "^\d+\r")   Then p.Range.Delete
      p.Range.Select
      Selection.Collapse wdCollapseStart
      Selection.Range.InsertBefore ("rz ")
      Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
      Selection.Range.InsertAutoText
   Next p
   Application.ScreenUpdating = True
   
   objUndo.EndCustomRecord
End Sub

Sub Randnummern_Loeschen()
   Dim p As Paragraph
   Dim fr As Frame
   Dim objUndo As UndoRecord
   Set objUndo = Application.UndoRecord
   objUndo.StartCustomRecord ("Undo Remove Marginals")
   Application.ScreenUpdating = False
   For Each fr In ActiveDocument.Frames
      If fr.Range.Paragraphs(1).style = "RzA" Then
         fr.Select
         Selection.Delete
      End If
   Next fr
   Application.ScreenUpdating = True
   
   objUndo.EndCustomRecord
End Sub

Sub Randnummern_Ausrichten()
   For i = 1 To ActiveDocument.Frames.Count
      ActiveDocument.Frames(i).Select
      Selection.MoveRight wdCharacter
      ActiveDocument.Frames(i).Select
      Selection.MoveRight wdCharacter
      b = Selection.Information(wdVerticalPositionRelativeToPage)
      Selection.MoveEnd wdParagraph
      Selection.Collapse wdCollapseEnd
      Selection.MoveLeft wdCharacter
      ed = Selection.Information(wdVerticalPositionRelativeToPage)
      Debug.Print PointsToCentimeters(ed - b)
      ActiveDocument.Frames(i).VerticalPosition = (ed - b + 1)
   Next
End Sub


Sub DeleteUnusedStyles()
   Dim oStyle As style
   For Each oStyle In ActiveDocument.Styles
      'Only check out non-built-in styles
      If oStyle.BuiltIn = False Then
         With ActiveDocument.Content.Find
            .ClearFormatting
            .style = oStyle.NameLocal
            .Execute FindText:="", Format:=True
            If .Found = False Then oStyle.Delete
         End With
      End If
   Next oStyle
End Sub

Sub LoopEdit()
   Dim pat As String, rpl As String, style As String, Rng As Range
   Dim objUndo As UndoRecord, par As Paragraph
   Set objUndo = Application.UndoRecord
   objUndo.StartCustomRecord ("Edit Loop")
   If Selection.Type = wdSelectionIP Then
      Set Rng = ActiveDocument.Range
   Else: Set Rng = Selection.Range
   End If
   pat = InputBox("Suchmuster eingeben" & vbCr & "<< für '" + lastpat + "'", "Suchmuster")
   If pat = "" Then Exit Sub
   If pat = "<<" Then pat = lastpat
   rpl = InputBox("Ersetzen mit", "")
   style = InputBox("Formatvorlage", "")
   For Each par In Rng.Paragraphs
      If RxTest(par.Range.Text, pat) Then
         If rpl = "#del" Then par.Range.Delete Else _
            If rpl <> "" Then par.Range.Text = RxReplace(par.Range.Text, pat, rpl)
         If style <> "" Then par.style = style
      End If
   Next par
   lastpat = pat
   objUndo.EndCustomRecord
End Sub

Sub RemoveHyperlinks()
 Dim l As Hyperlink, i As Integer
 For i = ActiveDocument.Hyperlinks.Count To 1 Step -1
    ActiveDocument.Hyperlinks(i).Range.Select
    Selection.Paragraphs(1).Range.Delete
 Next i
End Sub


Sub Satznummern_Erstellen()
   Dim Rng As Range
   Dim txt As String, txt2 As String
   Dim fld As Field
   Dim sty As style
   Dim objUndo As UndoRecord, par As Paragraph
   Set objUndo = Application.UndoRecord
   objUndo.StartCustomRecord ("Edit Loop")
   On Error Resume Next
   Set sty = ActiveDocument.Styles("SatzNr")
   If Err = 5941 Then
         Set sty = ActiveDocument.Styles.Add("SatzNr", wdStyleTypeCharacter)
         sty.Font.Superscript = True
   End If
   If Selection.Type = wdSelectionIP Then ActiveDocument.Select
   Set Rng = Selection.Range
   txt = RxReplace(Rng.Text, "(" + Chr(13) + "\([0-9]+[a-z]*\) )([A-ZÄÖÜ])", "$1####$2")
   txt = RxReplace(txt, "^(\([0-9]+[a-z]*\) )([A-ZÄÖÜ])", "$1####$2")
   txt2 = RxReplace(txt, "\. ([A-ZÄÖÜ])", ". ####$1")
   Rng.Text = txt2
   Rng.Select
  Set Rng = Selection.Range
   Do While Rng.Find.Execute(FindText:="####", _
        Forward:=True, Format:=False, Wrap:=wdFindStop, ReplaceWith:="", Replace:=wdReplaceOne) = True
        Rng.MoveStart Unit:=wdCharacter, Count:=0
        Set fld = Rng.Fields.Add(Range:=Rng, Type:=wdFieldEmpty, Text:="SEQ sn \n", PreserveFormatting:=True)
        fld.Select
        Selection.Range.style = "SatzNr"
   Loop
   ActiveDocument.Fields.Update
   objUndo.EndCustomRecord
End Sub

Sub Satzummern_Loeschen()
   Dim fld As Field
   Dim Rng As Range
   If Selection.Type = wdSelectionIP Then ActiveDocument.Select
   Set Rng = Selection.Range
   For Each fld In Rng.Fields
      If InStr(fld.Code, "sn") Then fld.Delete
   Next
End Sub

Sub Normalize_Spaces()
   With ActiveDocument.Content.Find
      .ClearFormatting
      .MatchWildcards = True
      .Forward = True
      .Wrap = wdFindContinue
      .Text = "[" + ChrW(8194) + "-" + ChrW(8202) + " ]"
      '.HitHighlight FindText:="[" + ChrW(8194) + "-" + ChrW(8202) + "]", MatchWildcards:=True
      .Execute ReplaceWith:=" ", Replace:=wdReplaceAll
   End With
End Sub

Sub RBB()
  Dim Rng As Range
  Dim Fnd As Boolean
    Set Rng = ActiveDocument.Range
    With Rng.Find
        .ClearFormatting
        .Execute FindText:="Rechtlicher Hinweis:", Forward:=True, _
                 Format:=False, Wrap:=wdFindStop
        Fnd = .Found
    End With

    If Fnd = True Then
      Rng.Select
      Selection.MoveEnd Unit:=wdParagraph, Count:=2
      Selection.Range.Text = "#rbb"
      Selection.MoveRight Unit:=wdWord, Count:=2
      Selection.Range.InsertAutoText
      Selection.InsertAfter vbCrLf
    End If
End Sub

Private Sub MarkPar()
   Dim p As Paragraph, i As Integer
   i = 0
   For Each p In ActiveDocument.Paragraphs
      i = i + 1
      If i = 10 Then Exit For
      p.Range.InsertBefore (ChrW(2302))
   Next
End Sub

Sub RemMark()
   With ActiveDocument.Range.Find
      .ClearFormatting
      .Text = ChrW(2302)
      .Replacement.Text = ""
      .Forward = True
      .Wrap = wdFindContinue
      .Execute Replace:=wdReplaceAll
   End With
End Sub
