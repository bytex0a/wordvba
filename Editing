Attribute VB_Name = "Editing"
Option Explicit

'******************** Editing ******************************
Sub CreateStyles()
   Dim ez1, ez2, ez3 As Style
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
   Dim par As Paragraph, rzp As Paragraph
   Dim rng As Range
   Dim Selrng As Range
   Dim undo As UndoRecord
   Set undo = Application.UndoRecord
   undo.StartCustomRecord ("Randnummern")
   Dim frm As Frame, ate As AutoTextEntry
   Set Selrng = Selection.Range
   For Each par In Selrng.Paragraphs
      Set rng = par.Range
      rng.Select
      rng.Collapse wdCollapseStart
      rng.InsertBefore "Rz "
      rng.InsertAutoText
   Next par
   undo.EndCustomRecord
End Sub

Sub RemoveNumberLines()
   Dim p As Paragraph
   Dim objUndo As UndoRecord
   Set objUndo = Application.UndoRecord
   objUndo.StartCustomRecord ("Undo")
   Application.ScreenUpdating = False
   
   For Each p In ActiveDocument.Paragraphs
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

Sub ShowKeyForm()
   KeyForm.Show
End Sub

