' Moves the second shape to bottom of the first shape
Sub MoveShapeBottomOfSelectedShape()
    Dim shp1, shp2  As shape
    
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please Select at least two shapes.", vbCritical
    Else
        Set shp1 = ActiveWindow.Selection.ShapeRange(1)
        Set shp2 = ActiveWindow.Selection.ShapeRange(2)
        
        shp2.Top = shp1.Top + shp1.Height + 5
    End If
    
End Sub
