Sub GenerateLabels()
    Dim wsInput As Worksheet
    Dim wsLabel As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim isBlankMode As Boolean
    
    ' Position Calculation Variables
    Dim labelIndex As Long
    Dim pageIndex As Long
    Dim slotIndex As Long
    Dim pairIndex As Long
    Dim startRow As Long
    Dim colOffset As Long
    
    ' Headers
    Const hPart As String = "Part #: "
    Const hLot As String = "Lot #: "
    Const hSerial As String = "Serial #: "
    Const hNCR As String = "NCR #: "
    Const hDisp As String = "Disposition: "
    Const hInsp As String = "Insp By: "
    Const hReason As String = "Reason for Failure: "
    Const hComm As String = "Comments: "
    
    Dim vPart As String, vLot As String, vSerial As String, vNCR As String
    Dim vDisp As String, vInsp As String, vReason As String, vComm As String
    
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsLabel = ThisWorkbook.Sheets("Labels")
    
    ' --- 1. DETERMINE RANGE ---
    ' Find the last row with data in the Input sheet
    lastRow = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row
    
    ' Even if data ends early, we always scan at least to Row 11 (1 full page)
    If lastRow < 11 Then lastRow = 11
    
    ' Check if empty (Blank Mode trigger)
    Dim hasData As Boolean
    hasData = False
    ' Check A2 down to the last row
    If Application.WorksheetFunction.CountA(wsInput.Range("A2:H" & lastRow)) > 0 Then
        hasData = True
    End If
    
    Application.ScreenUpdating = False
    
    ' Clear Labels sheet and Reset Page Breaks
    wsLabel.Cells.Clear
    wsLabel.ResetAllPageBreaks
    
    isBlankMode = Not hasData
    
    ' --- MAIN LOOP ---
    ' Loop from Input Row 2 to the End of Data
    For i = 2 To lastRow
        vPart = "": vLot = "": vSerial = "": vNCR = ""
        vDisp = "": vInsp = "": vReason = "": vComm = ""
        
        ' --- A. CALCULATE MULTI-PAGE POSITION ---
        ' 0-based index of the label (Row 2 is 0, Row 3 is 1...)
        labelIndex = i - 2
        
        ' Which Page is this label on? (0 = Page 1, 1 = Page 2...)
        pageIndex = Int(labelIndex / 10)
        
        ' Which Slot on that page? (0 to 9)
        slotIndex = labelIndex Mod 10
        
        ' Which Row Pair on that page? (0 to 4)
        pairIndex = Int(slotIndex / 2)
        
        ' Calculate the starting Excel Row for this label
        ' Each page is 25 Excel rows tall (5 labels * 5 rows each)
        startRow = (pageIndex * 25) + (pairIndex * 5) + 1
        
        ' Determine Column (Left vs Right)
        If i Mod 2 = 0 Then colOffset = 1 Else colOffset = 4
        
        ' --- B. DATA GATHERING ---
        Dim rowIsEmpty As Boolean
        If Application.WorksheetFunction.CountA(wsInput.Range(wsInput.Cells(i, 1), wsInput.Cells(i, 8))) = 0 Then
            rowIsEmpty = True
        Else
            rowIsEmpty = False
        End If

        If isBlankMode Then
            ' Blank Mode: Just print headers (handled below)
        Else
            If rowIsEmpty Then GoTo NextIteration
            
            ' CHECK FOR "BLANK" KEYWORD
            If LCase(Trim(wsInput.Cells(i, 1).Value)) = "blank" Then
                ' Leave variables empty
            Else
                vPart = wsInput.Cells(i, 1).Value
                vLot = wsInput.Cells(i, 2).Value
                vSerial = wsInput.Cells(i, 3).Value
                vNCR = wsInput.Cells(i, 4).Value
                vDisp = wsInput.Cells(i, 5).Value
                vReason = wsInput.Cells(i, 6).Value
                vInsp = wsInput.Cells(i, 7).Value
                vComm = wsInput.Cells(i, 8).Value
            End If
        End If

        ' --- C. FORMATTING (CRITICAL FOR MULTI-PAGE) ---
        ' We must apply Row Heights dynamically because the Python script 
        ' only formatted the first 25 rows.
        
        ' Apply Row Heights for this specific 5-row block
        wsLabel.Rows(startRow).RowHeight = 29.64
        wsLabel.Rows(startRow + 1).RowHeight = 29.64
        wsLabel.Rows(startRow + 2).RowHeight = 20
        wsLabel.Rows(startRow + 3).RowHeight = 20
        wsLabel.Rows(startRow + 4).RowHeight = 48.92
        
        ' Apply Font and Indent
        With wsLabel.Range(wsLabel.Cells(startRow, colOffset), wsLabel.Cells(startRow + 4, colOffset + 1))
            .Font.Name = "Arial"
            .Font.Size = 10
            .IndentLevel = 1
        End With
        
        ' Merges
        With wsLabel.Range(wsLabel.Cells(startRow + 3, colOffset), wsLabel.Cells(startRow + 3, colOffset + 1))
            .Merge: .WrapText = True
        End With
        With wsLabel.Range(wsLabel.Cells(startRow + 4, colOffset), wsLabel.Cells(startRow + 4, colOffset + 1))
            .Merge: .WrapText = True
        End With

        ' --- D. WRITE CONTENT ---
        WriteCell wsLabel.Cells(startRow, colOffset), hPart, vPart, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow, colOffset + 1), hLot, vLot, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow + 1, colOffset), hSerial, vSerial, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow + 1, colOffset + 1), hNCR, vNCR, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow + 2, colOffset), hInsp, vInsp, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow + 2, colOffset + 1), hDisp, vDisp, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow + 3, colOffset), hReason, vReason, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow + 4, colOffset), hComm, vComm, xlTop, xlLeft
        
        ' --- E. ADD PAGE BREAK ---
        ' If this is the last label on a page (Slot 9 = Bottom Right), insert a break after it
        If slotIndex = 9 Then
            wsLabel.HPageBreaks.Add Before:=wsLabel.Rows(startRow + 5)
        End If
        
NextIteration:
    Next i
    
    Application.ScreenUpdating = True
    
    If isBlankMode Then
        MsgBox "Generated blank forms (Data range empty).", vbInformation
    Else
        MsgBox "Labels generated successfully!", vbInformation
    End If
End Sub

' ---------------------------------------------------------
' MACRO: Clear Input Button
' ---------------------------------------------------------
Sub ClearInputForm()
    ' Clears the data entry area from A2 down to H200
    Sheets("Input").Range("A2:H200").ClearContents
End Sub

Sub WriteCell(target As Range, header As String, val As String, vAlign As Variant, hAlign As Variant)
    With target
        .Value = header & val
        .VerticalAlignment = vAlign
        .HorizontalAlignment = hAlign
        .Characters(Start:=1, Length:=Len(header)).Font.Bold = True
        If Len(val) > 0 Then
            .Characters(Start:=Len(header) + 1, Length:=Len(val)).Font.Bold = False
        End If
    End With
End Sub