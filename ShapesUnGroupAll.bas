Sub ShapesUnGroupAll()
    Dim sld         As slide
    Dim shp         As shape
    Dim intCount    As Integer
    intCount = 0
    Dim groupsExist As Boolean: groupsExist = True
    If MsgBox("Are you sure you want to ungroup every level of grouping on every slide?", (vbYesNo + vbQuestion), "Ungroup Everything?") = vbYes Then
        For Each sld In ActivePresentation.Slides        ' iterate slides
            Do While (groupsExist = True)
                groupsExist = False
                For Each shp In sld.Shapes
                    If shp.Type = msoGroup Then
                        shp.Ungroup
                        intCount = intCount + 1
                        groupsExist = True
                    End If
                Next shp
            Loop
            groupsExist = True
        Next sld
    End If
    MsgBox "All Done! " & intCount & " groups are now ungrouped."
End Sub
