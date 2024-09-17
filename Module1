Dim currentIndex As Integer

Sub Initialize()
    ' Start at the first word
    currentIndex = 2
    ShowWord
End Sub

Sub ShowWord()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim nextDate As Date
    Dim todaysDate As Date
    
    Set ws1 = ThisWorkbook.Sheets("Sheet1") ' Adjust name as necessary
    Set ws2 = ThisWorkbook.Sheets("Sheet2") ' Adjust name as necessary
    
    ' Get the Next Date and Today's Date
    nextDate = ws1.Cells(currentIndex, 5).Value ' Assuming "Next Date" is in column 5
    todaysDate = ws1.Cells(currentIndex, 3).Value ' Assuming "Today's Date" is in column 3 (using =TODAY())
    
    ' Check if Next Date is greater than Today's Date
    If nextDate > todaysDate Then
        ' Skip this word and move to the next one
        MoveNext
    Else
        ' Show Spanish word on Sheet2 in merged cell
        ws2.Range("A1").Value = ws1.Cells(currentIndex, 1).Value ' Spanish word
    End If
End Sub

Sub SkipWord()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Set ws2 = ThisWorkbook.Sheets("Sheet2")
        
    ' Show English word in merged cell
    ws2.Range("A1").Value = ws1.Cells(currentIndex, 2).Value ' English word
    
    ' Set SPS Increment to 0
    ws1.Cells(currentIndex, 4).Value = 0
    
    ' Wait for user to press Next manually before moving to the next word
End Sub

Sub GotIt()
    Dim ws1 As Worksheet
    Dim increment As Integer
    Dim currentDate As Date
    Dim newDate As Date
    
    Set ws1 = ThisWorkbook.Sheets("Sheet1")
 
    ' Get current SPS increment value
    increment = ws1.Cells(currentIndex, 4).Value
    If increment = 0 Then increment = 1 ' Set initial increment if zero
    
    ' Update SPS Increment
    ws1.Cells(currentIndex, 4).Value = increment * 2
    
    ' Calculate next review date
    currentDate = ws1.Cells(currentIndex, 3).Value ' Today's date
    newDate = currentDate + ws1.Cells(currentIndex, 4).Value ' Add increment to today's date
    ws1.Cells(currentIndex, 5).Value = newDate ' Set next review date
    
    ' Show the English word but wait for user to press Next before advancing
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("Sheet2")
    ws2.Range("A1").Value = ws1.Cells(currentIndex, 2).Value ' English word
End Sub

Sub NextWord()
    MoveNext
End Sub

Sub MoveNext()
    Dim ws1 As Worksheet
    Dim lastRow As Long
    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    
    ' Get the last row with data in column A
    lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    
    ' Increment current index
    currentIndex = currentIndex + 1
    
    ' Check if we are past the last row
    If currentIndex > lastRow Then
        ' Display message and end the program if there are no more flashcards to review
        MsgBox "There are no flashcards to review today."
        Exit Sub
    End If
    
    ' Show the next word (this will also skip any that shouldn't be shown)
    ShowWord
End Sub

