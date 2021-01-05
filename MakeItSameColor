' Make the fill color same for all selected shapes based on firs selected shape color
' First you need to select reference shape 
Sub MakeItSameColor()
    
    Dim shp1, oSh   As shape
    Dim color       As MsoThemeColorSchemeIndex
    
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please Select at least two shapes.", vbCritical
    Else
        Set shp1 = ActiveWindow.Selection.ShapeRange(1)
        color = shp1.Fill.ForeColor.RGB
        
        'Debug.Print "color:" & color
        For Each oSh In ActiveWindow.Selection.ShapeRange
            oSh.Fill.ForeColor.RGB = color
        Next
    End If
    
End Sub
