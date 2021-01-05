' Moves the second shape next to first shape
Sub MoveShapeNextToSelectedShape()
    Dim shp1, shp2  As shape
    
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please Select at least two shapes.", vbCritical
    Else
        Set shp1 = ActiveWindow.Selection.ShapeRange(1)
        Set shp2 = ActiveWindow.Selection.ShapeRange(2)
        
        shp2.Left = shp1.Left + shp1.Width + 5
        
    End If
    
End Sub
