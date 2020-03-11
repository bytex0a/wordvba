Attribute VB_Name = "Complete"
Public Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long
Option Explicit
Global Const ADD_SPACE = " "
Public dict As Object
Public resultdict As Object

Sub BuildDatabase()
   Dim max As Long, i As Long, wort As String
   Dim wrng As Range, t1 As Long, t2 As Long
   t1 = GetTickCount
   i = 1
   StatusBar = "Rebuild Database..."
   Set dict = CreateObject("Scripting.Dictionary")
   dict.Add "ADNAME", ActiveDocument.Name
   Set resultdict = CreateObject("Scripting.Dictionary")
   For Each wrng In ActiveDocument.Words
      wort = wrng.Text
      If Len(wort) < 3 Then GoTo NextContinue
      If InStr(wort, " ") > 0 Then wort = Trim(wrng.Text)
      If (Len(wort) > 2) Then
         If Not dict.Exists(wort) Then dict.Add wort, i
      End If
      i = i + 1
NextContinue:
   Next wrng
   StatusBar = "Database rebuilt in " & (GetTickCount - t1) & " ms (" & i & " words)"
End Sub

Sub CheckWord()
   Dim i As Long, wort As String, Item As Variant, a
   frmComplete.ListBox1.Clear
   If dict Is Nothing Then
      BuildDatabase
   ElseIf (ActiveDocument.Words.Count < 5000) Or (dict("ADNAME") <> ActiveDocument.Name) Then BuildDatabase
   End If
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
      frmComplete.ListBox1.List = resultdict.keys
      frmComplete.Show
   End If
   If resultdict.Count = 1 Then
      Selection.MoveLeft Unit:=wdCharacter, Count:=1
      Selection.Words(1).Select
      a = resultdict.keys
      Selection.Range.Text = a(0) + ADD_SPACE
      Selection.MoveRight wdWord, 1, wdMove
   End If

End Sub

