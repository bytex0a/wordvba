VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KeyForm 
   Caption         =   "Enter Shortcut"
   ClientHeight    =   450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2925
   OleObjectBlob   =   "KeyForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "KeyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************** Modul: KeyForm ********************

Public oldkey As Integer, key As Integer, pressed As Integer

Function CheckKey(coldkey, ckey) As Boolean
   If (key = ckey) And (oldkey = coldkey) Then CheckKey = True Else CheckKey = False
End Function

Private Sub Label1_Click()

End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   key = KeyCode
    Label1.Caption = key & "  " & oldkey
   
   '******** Tab operations *************
   If oldkey = vbKeyTab Then
      pos = Selection.Information(wdHorizontalPositionRelativeToTextBoundary)
      If KeyCode = vbKeyS Then
         Selection.Paragraphs(1).TabStops.Add Selection.Information(wdHorizontalPositionRelativeToTextBoundary), wdAlignTabLeft, 0
         KeyForm.Hide
      End If
      If KeyCode = vbKeyL Then Selection.Paragraphs(1).TabStops(pos).Alignment = wdAlignTabLeft: KeyForm.Hide
      If KeyCode = vbKeyR Then Selection.Paragraphs(1).TabStops(pos).Alignment = wdAlignTabRight: KeyForm.Hide
      If KeyCode = vbKeyC Then Selection.Paragraphs(1).TabStops(pos).Alignment = wdAlignTabCenter: KeyForm.Hide
      If KeyCode = vbKeyD Then Selection.Paragraphs(1).TabStops(pos).Alignment = wdAlignTabDecimal: KeyForm.Hide
      If KeyCode = vbKeyBack Then Selection.Paragraphs(1).TabStops(pos).Clear: KeyForm.Hide
   End If
   
   pressed = 1 - pressed
   oldkey = KeyCode
   If KeyCode = 27 Then KeyForm.Hide
End Sub

