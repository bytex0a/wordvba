VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegExp 
   Caption         =   "RegExp Suchen - Ersetzen"
   ClientHeight    =   2910
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   4980
   OleObjectBlob   =   "frmRegExp.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************** RegExp Suchen - Ersetzen *****************

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   If KeyCode = vbKeyEscape Then frmRegExp.Hide
End Sub

Private Sub UserForm_Initialize()
   Dim sty As style
   ComboBox3.Clear
   For Each sty In ActiveDocument.Styles
      ComboBox3.AddItem sty.NameLocal
   Next
   
End Sub

Private Sub CommandButton2_Click()               ' Abbrechen
   Me.Hide
End Sub

Private Sub CommandButton3_Click()               ' Rückgängig
   ActiveDocument.Undo
End Sub

Private Sub CommandButton1_Click()               ' Ausführen
   Dim srng As Range
   Dim pat As String, rpl As String, rpl2 As String, style As String, rng As Range
   Dim par As Paragraph
   If OptionButton1.Value = True Then Set rng = ActiveDocument.Range Else Set rng = Selection.Range
   pat = ComboBox1.Text
   rpl = MakeString(ComboBox2.Text)
   rpl2 = Replace(rpl, "\", "$")
   rpl = Replace(rpl2, "$$", "\")
   style = ComboBox3.Text
   Set objUndo = Application.UndoRecord
   objUndo.StartCustomRecord ("RegExp Suchen Ersetzen")
   
   If pat = "" Then MsgBox ("Kein Suchausdruck"): Exit Sub
      
   For Each par In rng.Paragraphs
      If RxTest(par.Range.Text, pat) Then
         If rpl = "#del" Then par.Range.Delete Else _
            If rpl <> "" Then par.Range.Text = RxReplace(par.Range.Text, pat, rpl)
         If style <> "" Then par.style = style
      End If
   Next par
   ComboBox1.AddItem (pat)
   ComboBox2.AddItem (rpl)
   objUndo.EndCustomRecord
End Sub


