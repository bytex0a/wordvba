Attribute VB_Name = "Complete"
Option Explicit
Const ADD_SPACE = ""
Public dict As Object
Public resultdict As Object

Sub BuildDatabase()
   Dim max As Long, i As Long, wort As String
   Set dict = CreateObject("Scripting.Dictionary")
   Set resultdict = CreateObject("Scripting.Dictionary")
   max = ActiveDocument.Words.Count
   For i = 1 To max
       wort = Trim(ActiveDocument.Words(i))
       If (Len(wort) > 2) Then
          If Not dict.Exists(wort) Then dict.Add wort, i
       End If
   Next i
End Sub

Sub CheckWord()
    Dim i As Long, wort As String, Item As Variant, a
    CompleteForm.ListBox1.Clear
    If dict Is Nothing Then BuildDatabase
    resultdict.RemoveAll
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    wort = Selection.Words(1)
    If wort = "" Then Exit Sub
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    
    i = 1
    For Each Item In dict.keys
       If (LCase(Left(Item, Len(wort))) = LCase(wort)) And (LCase(wort) <> LCase(Item)) Then
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
       a = resultdict.keys
       Selection.Range.Text = a(0) + ADD_SPACE
       Selection.MoveRight wdWord, 1, wdMove
    End If

End Sub


