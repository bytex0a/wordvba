VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDP 
   Caption         =   "Custom DocProps"
   ClientHeight    =   5760
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   8940.001
   OleObjectBlob   =   "frmDP.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub DPL_Click()
   TTC (DPL.List(DPL.ListIndex))
End Sub

Private Sub DPL_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   TTC (DPL.List(DPL.ListIndex, 1))
End Sub


Private Sub UserForm_Activate()
  Dim dp As DocumentProperty
  Dim a As Document, b As Document
  Set a = ActiveDocument
  Set b = Documents(1)
  DPL.Clear
  For Each dp In a.CustomDocumentProperties
     ' Debug.Print dp.Name & vbtab & dp.Value & vbCrLf
     If dp.Value <> "" Then
      With frmDP.DPL
         .AddItem dp.Name
         .Column(1, .ListCount - 1) = dp.Value
      End With
     End If
  Next dp
End Sub

Private Sub UserForm_Click()

End Sub

Public Sub TTC(sText As String)
   TextBox1.Text = sText
End Sub

