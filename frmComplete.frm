VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmComplete 
   Caption         =   "Vorschläge"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2655
   OleObjectBlob   =   "frmComplete.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   Set WshShell = CreateObject("WScript.Shell")
   If KeyCode = vbKeyEscape Then Me.Hide
   If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeySpace) Then
      Selection.MoveLeft unit:=wdCharacter, Count:=1
      Selection.Words(1).Select
      Selection.Range.Text = ListBox1.List(ListBox1.ListIndex) + ADD_SPACE
      Me.Hide
      Selection.MoveRight wdWord, 1, wdMove
   End If
   If KeyCode = vbKeyJ Then If ListBox1.ListIndex < ListBox1.ListCount - 1 Then ListBox1.ListIndex = ListBox1.ListIndex + 1
   If KeyCode = vbKeyK Then If ListBox1.ListIndex > 0 Then ListBox1.ListIndex = ListBox1.ListIndex - 1
End Sub

Private Sub UserForm_Activate()
   ListBox1.ListIndex = 0
End Sub

