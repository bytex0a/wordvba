Attribute VB_Name = "Initialize"
Sub ViewSettings()
   On Error Resume Next
   With ActiveWindow.View
      .ShowAll = False
      .ShowTabs = False
      .ShowSpaces = False
      .ShowParagraphs = False
      .FieldShading = wdFieldShadingWhenSelected
   End With
   Options.AutoFormatAsYouTypeReplaceQuotes = True
End Sub

Sub AutoExec()   'Run automatically
   FillcolDlg
   FillcolCmd
   Application.OnTime When:=Now + TimeValue("00:00:1"), Name:="ViewSettings"
End Sub

Sub AutoNew()
   ViewSettings
End Sub

Sub AutoOpen()
   ViewSettings
End Sub

Function ConvertFileLF(strFileName As String)
   Open strFileName For Input As #1
   buff = Input$(LOF(1), #1)
   Close #1
   buff = Replace$(buff, vbLf, vbCrLf)
   Open strFileName For Output As #1
   Print #1, buff
   Close #1
End Function

Sub Convert_frm()
   Const PATH = "U:\Dokumente\Sonstiges\Word\Neu\"
   Files = Dir(PATH & "*.frm")
   While Files <> ""
      Pr PATH & Files
      ConvertFileLF (PATH & Files)
      Files = Dir
   Wend
End Sub

Sub FiCONV()
   Dim s As String, s1 As String, s2 As String
   Open "U:\Dokumente\Sonstiges\Word\Tmp\RegExpForm_o.frm" For Input As #1
      s = Input(LOF(1), 1)
   Close #1
      s2 = Replace(s, Chr(10), vbCrLf)
   Open "U:\Dokumente\Sonstiges\Word\Tmp\RegExpForm2.frm" For Output As #1
      Print #1, s2
   Close #1
End Sub

