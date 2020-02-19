Attribute VB_Name = "Complete"
Public dict As Object
Public resultdict As Object

Sub BuildDatabase()
   Dim max As Long, i As Long
   Set dict = CreateObject("Scripting.Dictionary")
   Set resultdict = CreateObject("Scripting.Dictionary")
   max = ActiveDocument.Words.Count
   For i = 1 To max
       wort = ActiveDocument.Words(i)
       If (Len(wort) > 2) Then
          If Not dict.Exists(wort) Then dict.Add wort, i
       End If
   Next i

End Sub

Sub checkword()
    Dim i As Long
    CompleteForm.ListBox1.Clear
    resultdict.RemoveAll
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    wort = Selection.Words(1)
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    
    i = 1
    For Each Item In dict.keys
       If Left(Item, Len(wort)) = wort Then
       resultdict.Add Item, i
       i = i + 1
       End If
    Next Item
    If resultdict.Count > 1 Then
       CompleteForm.ListBox1.List = resultdict.keys
       CompleteForm.Show
    End If
    If resultdict.Count = 1 Then
       Selection.MoveLeft Unit:=wdCharacter, Count:=1
       Selection.Words(1).Select
      ' Selection.Text = resultdict.Keys(1)
      a = resultdict.keys
      Pr a(0)
    End If

End Sub


