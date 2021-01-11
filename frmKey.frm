VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKey 
   Caption         =   "Enter Shortcut"
   ClientHeight    =   455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2925
   OleObjectBlob   =   "frmKey.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************** Modul: frmKey ********************

Public oldkey As Integer, key As Integer, pressed As Integer

Function CheckKey(coldkey, ckey) As Boolean
   If (key = ckey) And (oldkey = coldkey) Then CheckKey = True Else CheckKey = False
End Function

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   key = KeyCode
   Label1.Caption = key & "  " & oldkey
   On Error Resume Next
   '******** Tab operations *************
   If oldkey = vbKeyTab Then
      Pos = Selection.Information(wdHorizontalPositionRelativeToTextBoundary)
      If KeyCode = vbKeyS Then
         Selection.Paragraphs(1).TabStops.Add Selection.Information(wdHorizontalPositionRelativeToTextBoundary), wdAlignTabLeft, 0
         frmKey.Hide
      End If
      If KeyCode = vbKeyL Then Selection.Paragraphs(1).TabStops(Pos).Alignment = wdAlignTabLeft: frmKey.Hide
      If KeyCode = vbKeyR Then Selection.Paragraphs(1).TabStops(Pos).Alignment = wdAlignTabRight: frmKey.Hide
      If KeyCode = vbKeyC Then Selection.Paragraphs(1).TabStops(Pos).Alignment = wdAlignTabCenter: frmKey.Hide
      If KeyCode = vbKeyD Then Selection.Paragraphs(1).TabStops(Pos).Alignment = wdAlignTabDecimal: frmKey.Hide
      If KeyCode = vbKeyBack Then Selection.Paragraphs(1).TabStops(Pos).Clear: frmKey.Hide
   End If
   
   If oldkey = vbKeyT Then If KeyCode = vbKeyT Then Application.Run ("TabelleTabelleMarkieren"): frmKey.Hide
   If oldkey = vbKeyD Then If KeyCode = vbKeyP Then ActiveDocument.Protect wdAllowOnlyReading: frmKey.Hide
   If oldkey = vbKeyD Then If KeyCode = vbKeyU Then ActiveDocument.Unprotect: frmKey.Hide
   
   
   If KeyCode = vbKeySpace Then
      frmKey.Hide
      BuildDatabase
   End If
   pressed = 1 - pressed
   oldkey = KeyCode
   If KeyCode = 27 Then frmKey.Hide
End Sub





