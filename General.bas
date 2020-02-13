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
      .Replacement.Style = ActiveDocument.Styles(sty)
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

Sub Kommandos()
   Dim com, s As String
   com = InputBox("Kommando eingeben", "Komanndo")
   Select Case com
   Case "hp":
      s = "Horizontale Position : " & vbCr & Round(Application.Selection.Information(wdHorizontalPositionRelativeToTextBoundary) _
                                                   / 72 * 2.54, 2) & "cm / " & Application.Selection.Information(wdHorizontalPositionRelativeToTextBoundary) & "pt. (relative to Text Boundary)" _
                                    & vbCr & Round(Application.Selection.Information(wdHorizontalPositionRelativeToPage) _
                                                   / 72 * 2.54, 2) & "cm / " & Application.Selection.Information(wdHorizontalPositionRelativeToPage) & "pt. (relative to Page)"
      MsgBox s
   Case "rds": Application.Run ("RedefineStyle")
      
   End Select
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
End Sub

'******************** Module bearbeiten ********************
Sub ExportModules()
   Dim proj As VBProject, vbc As VBComponent
   Dim s As String, szFileName As String, pth As String
   pth = "d:\dok\word\makros\"
   For Each vbc In VBE.VBProjects("Normal").VBComponents
      If vbc.Name <> "ThisDocument" Then
      szFileName = vbc.Name
      Select Case vbc.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
      End Select
      vbc.Export pth + szFileName
      End If
   Next vbc
End Sub

Sub ImportModules()

    Dim StrFile As String
    'Debug.Print "in LoopThroughFiles. inputDirectoryToScanForFile: ", inputDirectoryToScanForFile
    Dim proj As VBProject, vbc As VBComponent
    StrFile = Dir("d:\dok\word\makros\*.*")
    Do While Len(StrFile) > 0
        VBE.VBProjects("Project").VBComponents.Import StrFile
        Debug.Print StrFile
        StrFile = Dir
    Loop
End Sub

Sub DeleteModules()
   Dim i As Integer
   Dim sName As String
   For i = 1 To VBE.VBProjects("Project").VBComponents.Count
   sName = VBE.VBProjects("Project").VBComponents.Item(i).Name
   If sName <> "ThisDocument" Then
     With VBE.VBProjects("Project").VBComponents
         .Remove .Item(sName)
     End With
    Exit For
   End If
   Next i
End Sub


