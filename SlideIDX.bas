Sub SlideIDX()
'Determine the current slide in the Slide View mode
'reference:stackoverflow.com
MsgBox "The slide index of the current slide is:" & _

ActiveWindow.View.Slide.SlideIndex

End Sub
