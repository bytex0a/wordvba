VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRegExp 
   Caption         =   "RegExp Suchen - Ersetzen"
   ClientHeight    =   3465
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   5265
   OleObjectBlob   =   "frmRegExp.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***************** RegExp Suchen - Ersetzen *****************

Private Sub CheckBox1_Change()
   If CheckBox1.Value = False Then
      ComboBox3.Enabled = False
      ComboBox4.Enabled = False
   Else
      ComboBox3.Enabled = True
      ComboBox4.Enabled = True
   End If
End Sub

Private Sub FillFVCombo()
   Dim sty As style
   ComboBox3.Clear
   ComboBox4.Clear
   For Each sty In ActiveDocument.Styles
      ComboBox3.AddItem sty.NameLocal
      ComboBox4.AddItem sty.NameLocal
   Next
End Sub

Private Sub CommandButton4_Click()
   FillFVCombo
End Sub

Private Sub UserForm_Initialize()
   FillFVCombo
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   If KeyCode = vbKeyEscape Then frmRegExp.Hide
End Sub

Private Sub UserForm_Activate()
   ' FillFVCombo
   ComboBox1.SetFocus
End Sub

Private Sub CommandButton2_Click()               ' Abbrechen
   Me.Hide
End Sub

Private Sub CommandButton3_Click()               ' Rückgängig
   ActiveDocument.Undo
End Sub

Private Sub CommandButton1_Click()               ' Ausführen
   Dim srng As Range
   Dim pat As String, rpl As String, rpl2 As String, style As String, suchstyle As String
   Dim rng As Range
   Dim par As Paragraph
   Dim cnt As Long
   If OptionButton1.Value = True Then Set rng = ActiveDocument.Range Else Set rng = Selection.Range
   pat = MakeString(ComboBox1.Text)
   rpl = MakeString(ComboBox2.Text)
   rpl2 = Replace(rpl, "\", "$")
   rpl = Replace(rpl2, "$$", "\")
   style = ComboBox3.Text
   suchstyle = ComboBox4.Text
   Set objUndo = Application.UndoRecord
   objUndo.StartCustomRecord ("RegExp Suchen Ersetzen")
   
   If pat = "" Then MsgBox ("Kein Suchausdruck"): Exit Sub
      
   If CheckBox1.Value = True Then ' Ersetzung paragrafenweise
      For cnt = rng.Paragraphs.Count To 1 Step -1
         Set par = rng.Paragraphs(cnt)
         If ComboBox4.Value = "" Then ' keine SuchFV
            If RxTest(par.Range.Text, pat) Then
               If rpl = ":del" Then
                  par.Range.Delete
               Else
                  If rpl <> "" Then
                     If rpl = ":e" Then par.Range.Text = RxReplace(par.Range.Text, pat, "") Else
                     par.Range.Text = RxReplace(par.Range.Text, pat, rpl)
                     Set par = rng.Paragraphs(cnt)
                  End If
               End If
            If style <> "" Then par.style = style
            End If
         Else ' SuchFV
            If RxTest(par.Range.Text, pat) And (par.style = suchstyle) And (suchstyle <> "") Then
               If rpl = "#del" Then
                  par.Range.Delete
               Else
               par.Range.Text = RxReplace(par.Range.Text, pat, rpl)
               Set par = rng.Paragraphs(cnt)
            End If
            If style <> "" Then par.style = style
            End If
         End If ' SuchFV
      Next cnt
   Else ' Ersetzung bezogen auf das gesamte Dokument
      If OptionButton1.Value = True Then
         Set rng = ActiveDocument.Range
      Else: Set rng = Selection.Range
      End If
      rng.Text = RxReplace(rng.Text, pat, rpl)
   End If
   ComboBox1.AddItem (pat)
   ComboBox2.AddItem (rpl)
   objUndo.EndCustomRecord
   ComboBox1.SetFocus
End Sub






