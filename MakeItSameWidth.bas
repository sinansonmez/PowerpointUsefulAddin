Sub MakeItSameWidth()
' Copies the property of the first selected chart / shape to the second

' No variables are saved into the memory

' IMPORTANT: YOU NEED TO SELECT THE SHAPES IN ORDER

Dim shp1, oSh As shape

Dim objHeigh As Double
Dim objWidth As Double

If ActiveWindow.Selection.ShapeRange.Count < 2 Then
    MsgBox "Please select at least two shapes.", vbCritical
Else
    Set shp1 = ActiveWindow.Selection.ShapeRange(1)
    objHeigh = shp1.Height
    objWidth = shp1.Width
    
    For Each oSh In ActiveWindow.Selection.ShapeRange
    oSh.Width = objWidth
    Next
End If

End Sub
