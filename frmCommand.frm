VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCommand 
   Caption         =   "Befehle"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3825
   OleObjectBlob   =   "frmCommand.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'******************** Modul: frmCommand ********************

Private Sub ListBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
   On Error GoTo ErrorHandler
   If KeyAscii = 13 Then
      cmd = ListBox1.List(ListBox1.ListIndex)
      Application.Run (cmd)
      frmCommand.Hide
   End If
   If KeyAscii = 27 Then frmCommand.Hide
ErrorHandler:
   ' Befehl misslungen: If Err = -2147352573 Then True
End Sub

Private Sub TextBox1_Change()
   Dim part() As String
   Dim Count As Integer, anz As Byte
   Dim cItem As Variant
   ListBox1.Clear
   For Each cItem In colCmd
      st = TextBox1.Text
      If Left(st, 1) = " " Then bflag = True: st = Mid(st, 2) Else bflag = False
      part = Split(st, " ")
      bfc = UCase(cItem)
      Select Case UBound(part)
      Case 0
         If bflag = True Then
            If InStr(bfc, UCase(part(0))) = 1 Then ListBox1.AddItem cItem
         Else
            If InStr(bfc, UCase(part(0))) > 0 Then ListBox1.AddItem cItem
         End If
      Case 1
         If InStr(bfc, UCase(part(0))) > 0 And _
                                       InStr(bfc, UCase(part(1))) > 0 Then ListBox1.AddItem cItem
      Case 2
         If InStr(bfc, UCase(part(0))) > 0 And _
                                       InStr(bfc, UCase(part(1))) > 0 And _
                                       InStr(bfc, UCase(part(2))) > 0 Then ListBox1.AddItem cItem
      Case 3
         If InStr(bfc, UCase(part(0))) > 0 And _
                                       InStr(bfc, UCase(part(1))) > 0 And _
                                       InStr(bfc, UCase(part(2))) > 0 And _
                                       InStr(bfc, UCase(part(3))) > 0 Then ListBox1.AddItem cItem
      End Select
   Next cItem
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   If KeyCode = vbKeyEscape Then frmCommand.Hide
   If KeyCode = vbKeyDown Then ListBox1.Selected(0) = True: ListBox1.SetFocus
   If KeyCode = vbKeyReturn Then
      On Error Resume Next
      If (ListBox1.ListIndex) = -1 Then
         If ListBox1.ListCount > 0 Then ListBox1.ListIndex = 0
         Else: Exit Sub
      End If
      Application.Run (ListBox1.List(ListBox1.ListIndex))
      frmCommand.Hide
   End If
End Sub

Private Sub UserForm_Activate()
   TextBox1.SetFocus
   TextBox1.Text = ""
   BefehlslisteLaden
End Sub

Private Sub UserForm_Initialize()
End Sub






