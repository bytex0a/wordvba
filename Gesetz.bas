Attribute VB_Name = "Gesetz"
'******************** Gesetz General Functions ********************
Public Sub CleanUp()
   With ActiveDocument.Range.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .MatchWildcards = False
      .Text = "^p^p"
      .Replacement.Text = "^p"
      .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
   End With
   For Each par In ActiveDocument.Paragraphs
      If RxTest(par.Range.Text, "^[ ]+") Then par.Range.Text = RxReplace(par.Range.Text, "^[ ]+", "")
   Next par
End Sub

'******************** Gesetz Routinen ********************
Sub ErstelleStyles()
   Dim st As style
   On Error Resume Next
   st = ActiveDocument.Styles.Add("G_Absatz", wdStyleTypeParagraph)
   st = ActiveDocument.Styles.Add("G_Num1", wdStyleTypeParagraph)
   st = ActiveDocument.Styles.Add("G_Num2", wdStyleTypeParagraph)
   st = ActiveDocument.Styles.Add("G_Num3", wdStyleTypeParagraph): st.BaseStyle = ""
   st = ActiveDocument.Styles.Add("G_Para", wdStyleTypeParagraph)
   st = ActiveDocument.Styles.Add("G_FolgeText", wdStyleTypeParagraph)
   st = ActiveDocument.Styles.Add("G_�Para", wdStyleTypeParagraph)
   st = ActiveDocument.Styles.Add("G_�Teil", wdStyleTypeParagraph)
   st = ActiveDocument.Styles.Add("G_�Abschnitt", wdStyleTypeParagraph)
End Sub

Sub FormatStyles()
   Dim spa, spa_num, spa_para, spa_uepara, spb_uepara, fs As Single, b As Integer
   ErstelleStyles
   b = CentimetersToPoints(0.5)
   fs = ActiveDocument.Styles(wdStyleNormal).Font.Size
   spa_num = 3
   spa_para = 3
   spa_uepara = 3
   spb_uepara = 9
   With ActiveDocument.Styles("G_Num1").ParagraphFormat
      .LeftIndent = b
      .FirstLineIndent = -b
      .SpaceAfter = spa_num
   End With
   With ActiveDocument.Styles("G_Num2").ParagraphFormat
      .LeftIndent = 2 * b
      .FirstLineIndent = -b
      .SpaceAfter = spa_num
   End With
   With ActiveDocument.Styles("G_Num3").ParagraphFormat
      .LeftIndent = 12
      .FirstLineIndent = -b
      .SpaceAfter = spa
      .SpaceAfter = spa_num
   End With
   With ActiveDocument.Styles("G_Para").ParagraphFormat
      .FirstLineIndent = b * 0.8
      .SpaceAfter = spa_para
   End With
   With ActiveDocument.Styles("G_Folgetext").ParagraphFormat
      '[erg�nzen]
   End With
   With ActiveDocument.Styles("G_�Para")
      .ParagraphFormat.SpaceAfter = spa_uepara
      .ParagraphFormat.SpaceBefore = spb_uepara
      .ParagraphFormat.OutlineLevel = wdOutlineLevel3
      .Font.Bold = True
   End With
   With ActiveDocument.Styles("G_�Teil")
      .Font.Bold = True
      .Font.Size = ActiveDocument.Styles(wdStyleNormal).Font.Size
      .ParagraphFormat.Alignment = wdAlignParagraphCenter
      .ParagraphFormat.OutlineLevel = wdOutlineLevel1
   End With
   With ActiveDocument.Styles("G_�Abschnitt")
      .Font.Bold = True
      .Font.Size = ActiveDocument.Styles(wdStyleNormal).Font.Size
      .ParagraphFormat.Alignment = wdAlignParagraphCenter
      .ParagraphFormat.OutlineLevel = wdOutlineLevel2
   End With
End Sub

Sub LoopPara()
   Dim par, ppar, npar As Paragraph
   Dim regex      As New regexp
   Dim str, pat, repls As String
   Dim objUndo As UndoRecord
   
   Application.ScreenUpdating = False
   CleanUp
   ErstelleStyles
   FormatStyles
   Set objUndo = Application.UndoRecord
   objUndo.StartCustomRecord ("LoopPara")
   For Each par In ActiveDocument.Paragraphs
      str = par.Range.Text
      'Links l�schen
      If RxTest(str, "^Nichtamtliches Inhaltsverzeichnis") = True Then par.Range.Delete
      If RxTest(str, "^\(Fundstelle: ") = True Then par.Range.Delete
      
      'Numerierung 1. 2. 3a.
      pat = "^(\d+[a-z]*.)\r"
      If RxTest(str, pat) Then
         par.Range = RxReplace(str, pat, "$1" & vbTab)
         par.style = "G_Num1"
      End If
      
      'Numerierung a) b) c)
      pat = "^([a-z]\))\r"
      If RxTest(str, pat) Then
         rpls = RxReplace(str, pat, "$1" & vbTab)
         par.Range.Text = rpls
         par.style = "G_Num2"
      End If
      
      'Numerierung aa) bb) cc)
      pat = "^([a-z]{2}\))\r"
      If RxTest(str, pat) Then
         par.Range.Text = RxReplace(str, pat, "$1" & vbTab)
         par.style = "G_Num3"
      End If
      
      'Abs�tze
      pat = "^\(\d+[a-z]*\)"
      If RxTest(str, pat) Then
         par.style = "G_Para"
      End If
      
      'Paragrafen-�berschriften
      pat = "^� \d+[a-z]*[" & ChrW(160) & " ][A-Z\(]"
      If RxTest(str, pat) Then
         par.style = "G_�Para"
         par.KeepWithNext = True
      End If
      
      '      Paragrafen-�berschriften (nur � ohne Text)
      '      pat = "^� \d+" + Chr(160) + vbCr + "$"
      '      If RxTest(str, pat) Then
      '         par.Style = "G_�Para"
      '         par.KeepWithNext = True
      '      End If
            
      'Artikel-�berschriften
      pat = "^Art\. .*$"
      If RxTest(str, pat) Then
         par.style = "G_�Para"
         par.KeepWithNext = True
      End If
      
      'Gliederung Teil
      If RxTest(str, "^Art. \d+" & vbCr & "$") Then par.Range.Text = RxReplace(str, vbCr, Chr(11)): par.style = "G_Para"
      If RxTest(str, "^Teil \d+" & vbCr & "$") Then par.Range.Text = RxReplace(str, vbCr, Chr(11)): par.style = "G_�Teil"
      If RxTest(str, "^[A-Z][a-z]+ Teil ") Then
         par.Range.Text = RxReplace(str, "^([A-Z][a-z]+ Teil)", "$1" & Chr(11))
         Set par = par.Previous
         par.style = "G_�Teil"
      End If
      If RxTest(str, "^[A-Z][a-z]+ Teil" & Chr(11)) Then par.style = "G_�Teil"
      If RxTest(str, "^[A-Z][a-z]+ Abschnitt" & Chr(11)) Then par.style = "G_�Abschnitt"
      If RxTest(str, "^Abschnitt \d+" & vbCr & "$") Then par.Range.Text = RxReplace(str, vbCr, Chr(11)): par.style = "G_�Abschnitt"
      If RxTest(str, "^(Abschnitt I+) (.*)") Then
         par.Range.Text = RxReplace(str, "(^Abschnitt I+) ", "$1" & Chr(11))
         Set par = par.Previous
         par.style = "G_�Abschnitt"
      End If
      If RxTest(str, "^Anlage \d+") Then par.style = "G_�Para"
      
      'Text nach Paragrafen/�berschriften
      On Error Resume Next
      Set ppar = par.Previous
      prevsty = ppar.style
      parsty = par.style
      If (parsty = "Standard") And _
                               ((InStr(prevsty, "G_Num") = 1) Or _
                                (InStr(prevsty, "G_Para") = 1)) Then
         par.style = "G_FolgeText"
      End If
   Next par
   objUndo.EndCustomRecord
ErrorHandler:
   If Err = 91 Then Resume Next
End Sub


