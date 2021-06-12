Sub ModifyShapesGroup()
  'reference: http://www.pptools.com/, https://docs.microsoft.com/en-us/, Richard Mansfield. Mastering VBA for Microsoft Office. Wiley
    Dim x As Long
    With ActivePresentation.Slides(1).Shapes("Group 1")
        For x = 1 To .GroupItems.Count
            If .GroupItems(x).Name = "TextBox 1" Then
                With .GroupItems(x)
                    ' do something with it, for example:
                    .TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                    .Fill.ForeColor.RGB = RGB(0, 0, 0)
                End With
            End If
        Next
    End With
End Sub

