Sub ChangeProofingLanguageToEnglish()
    Dim j, k        As Integer
    Dim languageID  As MsoLanguageID
    
    'Set this to your preferred language
    languageID = msoLanguageIDEnglishUK
    
    'Loop all the slides in the document, and change the language
    For j = 1 To ActivePresentation.Slides.Count
        For k = 1 To ActivePresentation.Slides(j).Shapes.Count
            ChangeAllSubShapes ActivePresentation.Slides(j).Shapes(k), languageID
        Next k
    Next j
    
    'Loop all the master slides, and change the language
    For j = 1 To ActivePresentation.SlideMaster.CustomLayouts.Count
        For k = 1 To ActivePresentation.SlideMaster.CustomLayouts(j).Shapes.Count
            ChangeAllSubShapes ActivePresentation.SlideMaster.CustomLayouts(j).Shapes(k), languageID
        Next k
    Next j
    
    'Change the default presentation language, so that all new slides respect the new language
    ActivePresentation.DefaultLanguageID = languageID
End Sub

Sub ChangeAllSubShapes(targetShape As shape, languageID As MsoLanguageID)
    Dim i           As Integer, R As Integer, c As Integer
    
    If targetShape.HasTextFrame Then
        targetShape.TextFrame.TextRange.languageID = languageID
    End If
    
    If targetShape.HasTable Then
        For R = 1 To targetShape.Table.Rows.Count
            For c = 1 To targetShape.Table.Columns.Count
                targetShape.Table.Cell(R, c).shape.TextFrame.TextRange.languageID = languageID
            Next
        Next
    End If
    
    Select Case targetShape.Type
        Case msoGroup, msoSmartArt
            For i = 1 To targetShape.GroupItems.Count
                ChangeAllSubShapes targetShape.GroupItems.Item(i), languageID
            Next i
    End Select
End Sub
