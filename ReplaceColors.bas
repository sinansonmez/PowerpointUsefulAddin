' function asks user to provide to RGB color
' first one is the old color to be replaced and second one is the new color to be used
Sub ReplaceColors()
    Dim lFindColor As Long
    Dim lReplaceColor As Long
    Dim oSl As slide
    Dim oSh As shape
    Dim userColor As String
    
    userColor = InputBox("Please enter old RGB color code to be replaced as 255,255,255", "Enter Color")
    splittedUserColor = Split(userColor, ",")
    lFindColor = RGB(CInt(splittedUserColor(0)), CInt(splittedUserColor(1)), CInt(splittedUserColor(2)))
    
    userColor = InputBox("Please enter new RGB color code to be used as 255,255,255", "Enter Color")
    splittedUserColor = Split(userColor, ",")
    lReplaceColor = RGB(CInt(splittedUserColor(0)), CInt(splittedUserColor(1)), CInt(splittedUserColor(2)))
    
    For Each oSl In ActivePresentation.Slides
        For Each oSh In oSl.Shapes
            With oSh
            ' Fill
            If .Fill.ForeColor.RGB = lFindColor Then
                .Fill.ForeColor.RGB = lReplaceColor
            End If
            End With
        Next
    Next
    
End Sub
