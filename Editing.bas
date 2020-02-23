Attribute VB_Name = "Editing"
Option Explicit
Public lastpat As String

'******************** Editing ******************************
Sub CreateStyles()
   Dim ez1, ez2, ez3 As style
   ez1 = ActiveDocument.Styles.Add("ez1", wdStyleTypeParagraph)
   ez2 = ActiveDocument.Styles.Add("ez2", wdStyleTypeParagraph)
   ez3 = ActiveDocument.Styles.Add("ez3", wdStyleTypeParagraph)
   par = ActiveDocument.Styles.Add("par", wdStyleTypeParagraph)
End Sub

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
   Dim rng As Range
   Dim objUndo As UndoRecord
   Set objUndo = Application.UndoRecord
   objUndo.StartCustomRecord ("Undo")
   RegisterListtemplateRZ 'ListTemplate für Randziffern erstellen
   Application.ScreenUpdating = False
   If Selection.Paragraphs.Count = 1 Then Set rng = ActiveDocument.Range Else Set rng = Selection.Range
   For Each p In rng.Paragraphs
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
      If fr.Range.Paragraphs(1).style = "Rz" Then
         fr.Select
         Selection.Delete
      End If
   Next fr
   Application.ScreenUpdating = True
   
   objUndo.EndCustomRecord
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
   Dim pat As String, rpl As String, style As String, rng As Range
   Dim objUndo As UndoRecord, par As Paragraph
   Set objUndo = Application.UndoRecord
   objUndo.StartCustomRecord ("Edit Loop")
   If Selection.Type = wdSelectionIP Then
        Set rng = ActiveDocument.Range
   Else: Set rng = Selection.Range
   End If
   pat = InputBox("Suchmuster eingeben" & vbCr & "<< für '" + lastpat + "'", "Suchmuster")
   If pat = "" Then Exit Sub
   If pat = "<<" Then pat = lastpat
   rpl = InputBox("Ersetzen mit", "")
   style = InputBox("Formatvorlage", "")
   For Each par In rng.Paragraphs
      If RxTest(par.Range.Text, pat) Then
         If rpl = "#del" Then par.Range.Delete Else _
         If rpl <> "" Then par.Range.Text = RxReplace(par.Range.Text, pat, rpl)
         If style <> "" Then par.style = style
      End If
   Next par
   lastpat = pat
   objUndo.EndCustomRecord
End Sub


