Option Explicit

'this macro creates a table of contents in a given range in the first sheet
'it automatically uses the value of cell A1 of every sheet as the title and include a link to that sheet


Sub table_of_contents()

    Dim startCell As Range
    Dim sh As Worksheet
    Dim shName As String
    Dim msgConfirm As VBA.VbMsgBoxResult 'to be confirmed by user
    Dim endCell As Range
    
    
    On Error Resume Next
    
'   asking the user for a range to insert the table of contents
    Set startCell = Excel.Application.InputBox("Where do you want to insert the table of contents?" _
    & vbNewLine & "Please select a cell", "Insert Table of Contents", , , , , , 8)
    
    If Err.Number = 424 Then Exit Sub
    
    On Error GoTo Leave

    Set startCell = startCell.Cells(1, 1)

    Set endCell = startCell.Offset(Worksheets.Count - 2, 1)

'   user must confirm if ok to insert ToC in the given range in case the cells contain values
    msgConfirm = VBA.MsgBox("The values in these cells:" & vbNewLine & startCell.Address & " to " & endCell.Address & _
    " will be overwritten. Would you like to continue?", vbOKCancel + vbDefaultButton2)
    If msgConfirm = vbCancel Then Exit Sub

'   looping through every sheet and getting the value of the first cell aka the title
    For Each sh In Worksheets
    
        shName = sh.Name
    
        If ActiveSheet.Name <> sh.Name Then
            If sh.Visible = xlSheetVisible Then
            ' invisible sheets will not be added to the ToC
                    sh.Hyperlinks.Add Anchor:=startCell, Address:="", SubAddress:= _
                    "'" & shName & "'" & "!A1", TextToDisplay:=shName
                    startCell.Offset(0, 1).Value = sh.Range("A1").Value
                    Set startCell = startCell.Offset(1, 0)
            End If
        End If
    Next sh
    
    Exit Sub

Leave:
    MsgBox "An error has occurred!"


End Sub
