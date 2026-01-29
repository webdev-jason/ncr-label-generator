Sub GenerateLabels()
    Dim wsInput As Worksheet
    Dim wsLabel As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startRow As Long
    Dim colOffset As Long
    Dim labelCounter As Long
    Dim isBlankMode As Boolean
    Dim loopLimit As Long
    
    ' Set worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsLabel = ThisWorkbook.Sheets("Labels")
    
    ' Find the last row of data in Input
    lastRow = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row
    
    ' Prevent screen flickering
    Application.ScreenUpdating = False
    
    ' Clear previous contents
    wsLabel.Cells.Clear
    
    ' --- CHECK MODE ---
    If lastRow < 2 Then
        isBlankMode = True
        loopLimit = 10 
    Else
        isBlankMode = False
        loopLimit = lastRow
    End If
    
    ' Initialize Variables
    startRow = 1
    labelCounter = 1
    
    Dim loopStart As Long
    If isBlankMode Then loopStart = 1 Else loopStart = 2
    
    ' --- MAIN LOOP ---
    For i = loopStart To loopLimit
        
        ' 1. Determine Data Variables
        Dim tPart As String, tLot As String, tSerial As String
        Dim tNCR As String, tInsp As String, tReason As String, tComm As String
        
        ' UPDATED: Removed the manual leading space " " from these strings
        ' We will use .IndentLevel instead to handle alignment for wrapped text
        If isBlankMode Then
            tPart = "Part #:"
            tLot = "Lot #:"
            tSerial = "Serial #:"
            tNCR = "NCR #:"
            tInsp = "Inspected By:"
            tReason = "Reason for Failure:"
            tComm = "Comments:"
        Else
            If wsInput.Cells(i, 1).Value = "" Then GoTo NextIteration
            
            tPart = "Part #: " & wsInput.Cells(i, 1).Value
            tLot = "Lot #: " & wsInput.Cells(i, 2).Value
            tSerial = "Serial #: " & wsInput.Cells(i, 3).Value
            tNCR = "NCR #: " & wsInput.Cells(i, 4).Value
            tInsp = "Inspected By: " & wsInput.Cells(i, 6).Value
            tReason = "Reason for Failure: " & wsInput.Cells(i, 5).Value
            tComm = "Comments: " & wsInput.Cells(i, 7).Value
        End If

        ' 2. Determine Column Offset
        If labelCounter Mod 2 <> 0 Then
            colOffset = 1 ' Left Label (Col A)
        Else
            colOffset = 4 ' Right Label (Col D)
        End If
        
        ' 3. WRITE TO GRID
        
        ' Row 1: Part & Lot
        With wsLabel.Cells(startRow, colOffset)
            .Value = tPart
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
        End With
        With wsLabel.Cells(startRow, colOffset + 1)
            .Value = tLot
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
        End With
        
        ' Row 2: Serial & NCR
        With wsLabel.Cells(startRow + 1, colOffset)
            .Value = tSerial
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
        End With
        With wsLabel.Cells(startRow + 1, colOffset + 1)
            .Value = tNCR
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
        End With
        
        ' Row 3: Inspected By (Merged)
        With wsLabel.Range(wsLabel.Cells(startRow + 2, colOffset), wsLabel.Cells(startRow + 2, colOffset + 1))
            .Merge
            .Value = tInsp
            .VerticalAlignment = xlTop 
            .HorizontalAlignment = xlLeft
            .WrapText = True
        End With
        
        ' Row 4: Reason (Merged)
        With wsLabel.Range(wsLabel.Cells(startRow + 3, colOffset), wsLabel.Cells(startRow + 3, colOffset + 1))
            .Merge
            .Value = tReason
            .VerticalAlignment = xlTop
            .HorizontalAlignment = xlLeft
            .WrapText = True
        End With
        
        ' Row 5: Comments (Merged)
        With wsLabel.Range(wsLabel.Cells(startRow + 4, colOffset), wsLabel.Cells(startRow + 4, colOffset + 1))
            .Merge
            .Value = tComm
            .VerticalAlignment = xlTop 
            .HorizontalAlignment = xlLeft
            .WrapText = True
        End With
        
        ' 4. Formatting (Font & INDENT)
        With wsLabel.Range(wsLabel.Cells(startRow, colOffset), wsLabel.Cells(startRow + 4, colOffset + 1))
            .Font.Name = "Arial"
            .Font.Size = 10
            .IndentLevel = 1  ' <--- THIS FIXES THE WRAPPING ALIGNMENT
        End With

        ' 5. Move Logic
        If labelCounter Mod 2 = 0 Then
            startRow = startRow + 5
        End If
        
        labelCounter = labelCounter + 1
        
NextIteration:
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "Labels generated successfully!", vbInformation
End Sub