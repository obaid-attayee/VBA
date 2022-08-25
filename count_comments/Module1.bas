Option Explicit

Sub comment_counter()

    Dim c
    Dim i As Integer, counter As Integer, ws_count As Integer
    
    ws_count = ActiveWorkbook.Worksheets.Count
    
    counter = 0
    
    For i = 1 To ws_count
        For Each c In ActiveWorkbook.Worksheets(i).Comments
            counter = counter + 1
        Next c
    Next i
    
    MsgBox "There are " & counter & " comments in this workbook."

End Sub
