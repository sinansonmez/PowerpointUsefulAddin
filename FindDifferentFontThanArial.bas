Sub FindDifferentFont()
    Dim p           As Presentation: Set p = ActivePresentation
    Dim slide, slide2 As slide
    Dim shape, shape2, shape3 As shape
    Set slideNumberList = CreateObject("System.Collections.ArrayList")
    Dim slideNumberListString As String
    
    For Each slide In p.Slides
        For Each shape In slide.Shapes
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    ' if font size is smaller than 12
                    If Not shape.TextFrame.TextRange.Font.Name = "Arial" Then
                        ' put a circle to highlight smaller font size text
                        Set shape3 = slide.Shapes.AddShape(msoShapeOval, shape.Left - 30, shape.Top, 30, 30)
                        shape3.Fill.ForeColor.RGB = RGB(255, 0, 0)
                        shape3.Line.Visible = msoFalse
                        shape3.TextFrame.MarginLeft = 0
                        shape3.TextFrame.MarginRight = 0
                        shape3.TextEffect.Text = shape.TextFrame.TextRange.Font.Name
                        shape3.Name = "smallFontHighlighter"
                        If Not slideNumberList.Contains(slide.SlideNumber) Then
                            ' include slide number to the list
                            slideNumberList.Add slide.SlideNumber
                        End If
                    End If
                End If
            End If
        Next
    Next
    slideNumberList.sort
    slideNumberListString = Join(slideNumberList.toArray, ", ")
    
    If slideNumberList.Count = 0 Then
        MsgBox "No other font Is found"
    Else
        ' Select the first slide
        Set slide2 = ActivePresentation.Slides(1)
        ' Inside the slide add one box for each colour
        Set shape2 = slide2.Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 100)
        shape2.Fill.ForeColor.RGB = RGB(255, 0, 0)
        shape2.Line.Visible = msoFalse
        shape2.TextFrame.MarginLeft = 0
        shape2.TextFrame.MarginRight = 0
        shape2.TextEffect.Text = "Fonts smaller than 12 found On Slide: " & slideNumberListString
    End If
End Sub
