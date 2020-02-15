Attribute VB_Name = "ModulOp"
Global Const MODULE_PATH = "D:\dok\word\makros\"

Sub ExportModules()
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
    Dim StrFile As String
    Dim proj As VBProject, vbc As VBComponent
    StrFile = Dir(MODULE_PATH & "*.*")
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

