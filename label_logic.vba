Sub GenerateLabels()
    Dim wsInput As Worksheet
    Dim wsLabel As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim c As Long
    Dim isBlankMode As Boolean
    
    ' Position Calculation Variables
    Dim startRow As Long
    Dim colOffset As Long
    Dim pairIndex As Long
    
    ' Set worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsLabel = ThisWorkbook.Sheets("Labels")
    
    ' --- 1. SMART EMPTY CHECK ---
    ' Check if the entire page (Rows 2-11) is empty.
    ' If completely empty, we trigger "Blank Mode" to generate 10 blanks for handwriting.
    Dim hasData As Boolean
    hasData = False
    
    For i = 2 To 11
        If Application.WorksheetFunction.CountA(wsInput.Range(wsInput.Cells(i, 1), wsInput.Cells(i, 8))) > 0 Then
            hasData = True
            Exit For
        End If
    Next i
    
    ' Prevent screen flickering
    Application.ScreenUpdating = False
    
    ' Clear previous contents
    wsLabel.Cells.Clear
    
    isBlankMode = Not hasData
    
    ' --- MAIN LOOP (Fixed to 10 Labels) ---
    ' We loop from Row 2 to Row 11 (The 10 slots on the input sheet)
    For i = 2 To 11
        
        ' --- A. CALCULATE POSITION ---
        ' 1. Determine Row Block (0 to 4)
        '    (Row 2&3 -> Index 0)
        '    (Row 4&5 -> Index 1)
        pairIndex = Int((i - 2) / 2)
        
        '    Multiply by 5 rows per label, add 1 to start at Excel Row 1
        startRow = (pairIndex * 5) + 1
        
        ' 2. Determine Column (Left vs Right)
        '    Even Rows (2, 4, 6...) -> Left (Column A/1)
        '    Odd Rows (3, 5, 7...)  -> Right (Column D/4)
        If i Mod 2 = 0 Then
            colOffset = 1
        Else
            colOffset = 4
        End If
        
        ' --- B. GATHER DATA ---
        Dim tPart As String, tLot As String, tSerial As String
        Dim tNCR As String, tDisp As String
        Dim tInsp As String, tReason As String, tComm As String
        Dim rowIsEmpty As Boolean
        
        ' Check if THIS specific row has data
        If Application.WorksheetFunction.CountA(wsInput.Range(wsInput.Cells(i, 1), wsInput.Cells(i, 8))) = 0 Then
            rowIsEmpty = True
        Else
            rowIsEmpty = False
        End If

        If isBlankMode Then
            ' MODE 1: All Empty -> Generate Fill-in-the-blank forms
            tPart = "Part #:"
            tLot = "Lot #:"
            tSerial = "Serial #:"
            tNCR = "NCR #:"
            tDisp = "Disposition:"
            tInsp = "Insp By:"
            tReason = "Reason for Failure:"
            tComm = "Comments:"
        Else
            ' MODE 2: Exact Mapping
            If rowIsEmpty Then
                ' If this specific row is empty, SKIP writing to the label sheet.
                ' This leaves the label spot blank so you can re-use the sticker paper.
                GoTo NextIteration
            End If
            
            tPart = "Part #: " & wsInput.Cells(i, 1).Value
            tLot = "Lot #: " & wsInput.Cells(i, 2).Value
            tSerial = "Serial #: " & wsInput.Cells(i, 3).Value
            tNCR = "NCR #: " & wsInput.Cells(i, 4).Value
            tDisp = "Disposition: " & wsInput.Cells(i, 5).Value
            tReason = "Reason for Failure: " & wsInput.Cells(i, 6).Value
            tInsp = "Insp By: " & wsInput.Cells(i, 7).Value
            tComm = "Comments: " & wsInput.Cells(i, 8).Value
        End If

        ' --- C. WRITE TO GRID ---
        
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
        
        ' Row 4: Reason (Merged, Center)
        With wsLabel.Range(wsLabel.Cells(startRow + 3, colOffset), wsLabel.Cells(startRow + 3, colOffset + 1))
            .Merge
            .Value = tReason
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
            .WrapText = True
        End With
        
        ' Row 5: Comments (Merged, Top)
        With wsLabel.Range(wsLabel.Cells(startRow + 4, colOffset), wsLabel.Cells(startRow + 4, colOffset + 1))
            .Merge
            .Value = tComm
            .VerticalAlignment = xlTop 
            .HorizontalAlignment = xlLeft
            .WrapText = True
        End With
        
        ' --- D. FORMATTING ---
        With wsLabel.Range(wsLabel.Cells(startRow, colOffset), wsLabel.Cells(startRow + 4, colOffset + 1))
            .Font.Name = "Arial"
            .Font.Size = 10
            .IndentLevel = 1
        End With
        
NextIteration:
    Next i
    
    Application.ScreenUpdating = True
    
    If isBlankMode Then
        MsgBox "Generated blank forms (Page Full).", vbInformation
    Else
        MsgBox "Labels generated successfully at specific positions!", vbInformation
    End If
End Sub