Attribute VB_Name = "IfSG"
Sub Aktenzeichen()
   az = ActiveDocument.Tables(1).Cell(5, 3).Range.Text
   Pos = InStr(az, "ROF")
   If Pos > 1 Then
      ActiveDocument.Tables(1).Cell(5, 3).Range.Text = Mid(az, Pos, Len(az) - Pos - 1)
   End If
   
End Sub


Sub Edit1()
   Dim ud As UndoRecord
   Set ud = Application.UndoRecord
   ud.StartCustomRecord ("IfSG.Edit1")
   Aktenzeichen
   ActiveDocument.Tables(1).Cell(2, 1).Range.Paragraphs(3).Range.Delete
   ActiveDocument.Tables(1).Cell(13, 3).Range.Text = Format(Date, "DD.MM.YYYY ")
   Selection.EndKey Unit:=wdStory
   Selection.MoveRight Unit:=wdCharacter, Count:=1
   Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
   Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
   Selection.Delete Unit:=wdCharacter, Count:=1
   Selection.MoveUp Unit:=wdLine, Count:=6
   Selection.TypeParagraph
   Selection.TypeText Text:="Ihre Regierung von Oberfranken"
   ud.EndCustomRecord
End Sub

Sub Edit2()
   Dim ud As UndoRecord
   Set ud = Application.UndoRecord
   ud.StartCustomRecord ("IfSG.Edit1")
   Aktenzeichen
   ActiveDocument.Tables(1).Cell(13, 3).Range.Text = Format(Date, "DD.MM.YYYY ")
   Selection.EndKey Unit:=wdStory
   Selection.MoveRight Unit:=wdCharacter, Count:=1
   Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
   Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
   Selection.Delete Unit:=wdCharacter, Count:=1
   Selection.MoveUp Unit:=wdLine, Count:=6
   Selection.TypeParagraph
   Selection.TypeText Text:="Ihre Regierung von Oberfranken"
   ud.EndCustomRecord
End Sub

