' Replace two shapes with each other
Sub SwapShapes()
    Dim shp1, shp2  As shape
    Dim shp1Left, shp1Top, shp2Left, shp2Top As Single
    
    Debug.Print ActiveWindow.Selection.ShapeRange.Count
    
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please Select two shapes.", vbCritical
    ElseIf ActiveWindow.Selection.ShapeRange.Count > 2 Then
        MsgBox "Please Select only two shapes.", vbCritical
    Else
        Set shp1 = ActiveWindow.Selection.ShapeRange(1)
        Set shp2 = ActiveWindow.Selection.ShapeRange(2)
        
        shp1Left = shp1.Left
        shp1Top = shp1.Top
        shp2Left = shp2.Left
        shp2Top = shp2.Top
        
        shp1.Left = shp2Left
        shp1.Top = shp2Top
        shp2.Left = shp1Left
        shp2.Top = shp1Top
    End If
    
End Sub
