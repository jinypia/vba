Sub countdown()

Dim time As Date
time = Now()

Dim count As Integer
count = ActivePresentation.Slides(1).Shapes("timelimit").TextFrame.TextRange
time = DateAdd("s", count, time)

Do Until time < Now()

  DoEvents
  
  For i = 1 To 5
    ActivePresentation.Slides(i).Shapes("countdown").TextFrame.TextRange = Format((time - Now()), "ss")
  Next i
    
  If time < Now() Then
    For i = 1 To 5
      ActivePresentation.Slides(i).Shapes("countdown").TextFrame.TextRange = "Time up!"
    Next i
    ActivePresentation.SlideShowWindow.View.GotoSlide (6)
  End If
             
Loop

End Sub 
