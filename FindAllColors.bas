' program to loop through all shape colors in the presentation and insert new slide to the back with information on which slide the colur is used
Sub FindAllColors()
    Dim sld, sld2         As PowerPoint.slide
    Dim shp, shp2         As shape
    Dim colour, R, B, G, slideCount As Long
    Set colourList = CreateObject("System.Collections.ArrayList")
    Dim pptLayout As CustomLayout
    
    slideCount = ActivePresentation.Slides.Count
    ' loop through all slides and get the color of the shapes
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = msoAutoShape Then
                colour = shp.Fill.ForeColor.RGB
                ' if colour dont previously included in the list, include
                
                If colourList.Contains(colour) Then
                Else
                    colourList.Add colour
                End If
                
            End If
        Next shp
    Next sld
    
    ' Add new slide
    Set sld2 = ActivePresentation.Slides.AddSlide(slideCount + 1, ActivePresentation.Slides(2).CustomLayout)
    ' Inside the slide add one box for each colour
    colourList.sort
    Dim element     As Variant
    For Each element In colourList
        Set shp2 = sld2.Shapes.AddShape(msoShapeRectangle, 50, 50, 35, 300)
        shp2.Fill.ForeColor.RGB = element
        shp2.Line.Visible = msoFalse
        shp2.TextFrame.MarginLeft = 0
        shp2.TextFrame.MarginRight = 0
        ' inside each element wrote the which slide it is used
        shp2.TextEffect.Text = LookForSlideNumber(element)
    Next element
    
    ' distrbute the boxes horizontally
    DistributeShapes (slideCount)
    
End Sub

Public Function ConvertLongToRGB(colour)
    ' Convert LONG to RGB:
    B = colour \ 65536
    G = (colour - B * 65536) \ 256
    R = colour - B * 65536 - G * 256
    ConvertLongToRGB = R & " " & G & " " & B
End Function

' distribute shapes horizontally
Sub DistributeShapes(slideCount)
    Set myDocument = ActivePresentation.Slides(slideCount + 1)
    With myDocument.Shapes
        numShapes = .Count
        If numShapes > 1 Then
            numAutoShapes = 0
            ReDim autoShpArray(1 To numShapes)
            For i = 1 To numShapes
                If .Item(i).Type = msoAutoShape Then
                    numAutoShapes = numAutoShapes + 1
                    autoShpArray(numAutoShapes) = .Item(i).Name
                End If
            Next
            If numAutoShapes > 1 Then
                ReDim Preserve autoShpArray(1 To numAutoShapes)
                Set asRange = .Range(autoShpArray)
                asRange.Distribute msoDistributeHorizontally, msoCTrue
            End If
        End If
    End With
End Sub

' function to look for on which slide passed colour is used
Public Function LookForSlideNumber(colour) As String
    Dim sld         As PowerPoint.slide
    Dim shp        As shape
    Set slideNumberList = CreateObject("System.Collections.ArrayList")
    
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = msoAutoShape Then
                ' if given colour is matching with shape colour include the slide number to slide number list
                If shp.Fill.ForeColor.RGB = colour Then
                    If slideNumberList.Contains(sld.SlideNumber) Then
                    Else
                        If sld.SlideNumber = ActivePresentation.Slides.Count Then
                        Else
                            slideNumberList.Add sld.SlideNumber
                        End If
                    End If
                End If
            End If
        Next shp
    Next sld
    
    ' convert Array List to string and return this string
    LookForSlideNumber = Join(slideNumberList.toArray, ", ")
    
End Function
