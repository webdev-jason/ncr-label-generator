Sub GenerateLabels()
    Dim wsInput As Worksheet
    Dim wsLabel As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim c As Long
    Dim startRow As Long
    Dim colOffset As Long
    Dim labelCounter As Long
    Dim isBlankMode As Boolean
    Dim loopLimit As Long
    
    ' Set worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsLabel = ThisWorkbook.Sheets("Labels")
    
    ' --- 1. DETERMINE LAST ROW (Smart Check) ---
    lastRow = 1
    For c = 1 To 8
        Dim colLast As Long
        colLast = wsInput.Cells(wsInput.Rows.Count, c).End(xlUp).Row
        If colLast > lastRow Then lastRow = colLast
    Next c
    
    ' Prevent screen flickering
    Application.ScreenUpdating = False
    
    ' Clear previous contents
    wsLabel.Cells.Clear
    
    ' --- 2. CHECK MODE ---
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
        Dim tNCR As String, tDisp As String
        Dim tInsp As String, tReason As String, tComm As String
        
        If isBlankMode Then
            tPart = "Part #:"
            tLot = "Lot #:"
            tSerial = "Serial #:"
            tNCR = "NCR #:"
            tDisp = "Disposition:"
            tInsp = "Insp By:"
            tReason = "Reason for Failure:"
            tComm = "Comments:"
        Else
            ' DATA MODE: Smart Skip
            If Application.WorksheetFunction.CountA(wsInput.Range(wsInput.Cells(i, 1), wsInput.Cells(i, 8))) = 0 Then GoTo NextIteration
            
            tPart = "Part #: " & wsInput.Cells(i, 1).Value
            tLot = "Lot #: " & wsInput.Cells(i, 2).Value
            tSerial = "Serial #: " & wsInput.Cells(i, 3).Value
            tNCR = "NCR #: " & wsInput.Cells(i, 4).Value
            tDisp = "Disposition: " & wsInput.Cells(i, 5).Value
            tReason = "Reason for Failure: " & wsInput.Cells(i, 6).Value
            tInsp = "Insp By: " & wsInput.Cells(i, 7).Value
            tComm = "Comments: " & wsInput.Cells(i, 8).Value
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
        
        ' Row 3: Insp By & Disposition
        With wsLabel.Cells(startRow + 2, colOffset)
            .Value = tInsp
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
        End With
        With wsLabel.Cells(startRow + 2, colOffset + 1)
            .Value = tDisp
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
        End With
        
        ' Row 4: Reason (Merged & CENTER ALIGNED)
        With wsLabel.Range(wsLabel.Cells(startRow + 3, colOffset), wsLabel.Cells(startRow + 3, colOffset + 1))
            .Merge
            .Value = tReason
            .VerticalAlignment = xlCenter ' <--- UPDATED to Center
            .HorizontalAlignment = xlLeft
            .WrapText = True
        End With
        
        ' Row 5: Comments (Merged & Top Aligned)
        With wsLabel.Range(wsLabel.Cells(startRow + 4, colOffset), wsLabel.Cells(startRow + 4, colOffset + 1))
            .Merge
            .Value = tComm
            .VerticalAlignment = xlTop 
            .HorizontalAlignment = xlLeft
            .WrapText = True
        End With
        
        ' 4. Formatting
        With wsLabel.Range(wsLabel.Cells(startRow, colOffset), wsLabel.Cells(startRow + 4, colOffset + 1))
            .Font.Name = "Arial"
            .Font.Size = 10
            .IndentLevel = 1
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