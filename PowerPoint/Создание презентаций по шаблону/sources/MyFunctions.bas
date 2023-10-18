Attribute VB_Name = "MyFunctions"
Public Sub Delay(seconds As Currency)
    startTime = Timer
    Do While Timer < startTime + seconds
        DoEvents
    Loop
End Sub
