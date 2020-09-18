Attribute VB_Name = "ModulOp"
Global MODULE_PATH As String
'Global Const MODULE_PATH = "D:\dok\word\makros\"
'Global Const MODULE_PATH = "U:\Dokumente\Sonstiges\Word\"

Sub ExportModules()
   If InputBox("(1) D:\Dok\Word\Makros" & vbCr & "(2) U:\Dokumente\Sonstiges\Word", "Bitte wählen", "1") = "1" Then
      MODULE_PATH = "D:\dok\word\makros\"
      Else: MODULE_PATH = "U:\Dokumente\Sonstiges\Word\"
   End If
   Dim proj As VBProject, vbc As VBComponent
   Dim s, szFileName As String
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
         vbc.Export MODULE_PATH + szFileName
      End If
   Next vbc
End Sub

Sub ImportModules()
   Const PROJECT_NAME = 2 '"Normal"
   Dim StrFile As String
   Dim proj As VBProject, vbc As VBComponent
   On Error Resume Next
   StrFile = Dir(MODULE_PATH & "*.*")
   Do While Len(StrFile) > 0
      VBE.VBProjects(PROJECT_NAME).VBComponents.Import MODULE_PATH & StrFile
      Debug.Print StrFile
      StrFile = Dir
   Loop
End Sub

Sub DeleteModules()
   Const PROJECT_NAME = "Normal"
   Dim i As Integer
   Dim sName As String
   For i = 1 To VBE.VBProjects(PROJECT_NAME).VBComponents.Count
      sName = VBE.VBProjects(PROJECT_NAME).VBComponents.Item(i).Name
      If sName <> "ThisDocument" Then
         With VBE.VBProjects(PROJECT_NAME).VBComponents
            .Remove .Item(sName)
         End With
         Exit For
      End If
   Next i
End Sub
