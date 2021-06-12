Sub PrintAllSlidesName()
  'reference: https://stackoverflow.com/, Paul McFedries. Absolute Beginner's Guide to VBA. Que
Dim slide As slide

For Each slide In ActivePresentation.Slides

Debug.Print slide.Name
Next

End Sub
