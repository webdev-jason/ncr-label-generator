Sub GenerateLabels()
    Dim wsInput As Worksheet
    Dim wsLabel As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim labelRow As Long
    Dim labelCol As Long
    Dim labelCounter As Long
    
    ' Set worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsLabel = ThisWorkbook.Sheets("Labels")
    
    ' Find the last row of data in Input
    lastRow = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row
    
    ' Prevent screen flickering
    Application.ScreenUpdating = False
    
    ' Clear previous contents on Label sheet (keep formatting)
    wsLabel.Cells.ClearContents
    
    ' Initialize variables
    labelRow = 1
    labelCol = 1
    labelCounter = 1
    
    ' Loop through input data starting at Row 2 (skipping headers)
    For i = 2 To lastRow
        
        ' Check if Part # exists, otherwise skip
        If wsInput.Cells(i, 1).Value <> "" Then
            
            ' Format the Label Cell
            With wsLabel.Cells(labelRow, labelCol)
                .Value = "NCR #: " & wsInput.Cells(i, 4).Value & "   |   Part #: " & wsInput.Cells(i, 1).Value & vbNewLine & _
                         "Lot #: " & wsInput.Cells(i, 2).Value & "   |   Serial #: " & wsInput.Cells(i, 3).Value & vbNewLine & _
                         "Reason: " & wsInput.Cells(i, 5).Value & vbNewLine & _
                         "Insp By: " & wsInput.Cells(i, 6).Value & vbNewLine & _
                         "Comments: " & wsInput.Cells(i, 7).Value
                
                ' Apply text wrapping and alignment
                .WrapText = True
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlLeft
                .Font.Name = "Arial"
                .Font.Size = 10
            End With
            
            ' Logic to move to next label position
            If labelCounter Mod 2 <> 0 Then
                ' If Left label, move to Right label (Column C)
                labelCol = 3
            Else
                ' If Right label, move down to next row and back to Left (Column A)
                labelCol = 1
                labelRow = labelRow + 1
            End If
            
            labelCounter = labelCounter + 1
            
        End If
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "Labels generated successfully!", vbInformation
    
End Sub