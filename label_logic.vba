Sub GenerateLabels()
    Dim wsInput As Worksheet
    Dim wsLabel As Worksheet
    Dim i As Long
    Dim isBlankMode As Boolean
    
    ' Position Calculation Variables
    Dim startRow As Long
    Dim colOffset As Long
    Dim pairIndex As Long
    
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
    
    ' Smart Empty Check
    Dim hasData As Boolean
    hasData = False
    
    For i = 2 To 11
        If Application.WorksheetFunction.CountA(wsInput.Range(wsInput.Cells(i, 1), wsInput.Cells(i, 8))) > 0 Then
            hasData = True
            Exit For
        End If
    Next i
    
    Application.ScreenUpdating = False
    wsLabel.Cells.Clear
    
    isBlankMode = Not hasData
    
    ' MAIN LOOP
    For i = 2 To 11
        vPart = "": vLot = "": vSerial = "": vNCR = ""
        vDisp = "": vInsp = "": vReason = "": vComm = ""
        
        pairIndex = Int((i - 2) / 2)
        startRow = (pairIndex * 5) + 1
        If i Mod 2 = 0 Then colOffset = 1 Else colOffset = 4
        
        Dim rowIsEmpty As Boolean
        If Application.WorksheetFunction.CountA(wsInput.Range(wsInput.Cells(i, 1), wsInput.Cells(i, 8))) = 0 Then
            rowIsEmpty = True
        Else
            rowIsEmpty = False
        End If

        If isBlankMode Then
            ' Blank Mode
        Else
            If rowIsEmpty Then GoTo NextIteration
            
            ' CHECK FOR "BLANK" KEYWORD
            If LCase(Trim(wsInput.Cells(i, 1).Value)) = "blank" Then
                vPart = "": vLot = "": vSerial = "": vNCR = ""
                vDisp = "": vReason = "": vInsp = "": vComm = ""
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

        ' Formatting
        With wsLabel.Range(wsLabel.Cells(startRow, colOffset), wsLabel.Cells(startRow + 4, colOffset + 1))
            .Font.Name = "Arial"
            .Font.Size = 10
            .IndentLevel = 1
        End With
        
        With wsLabel.Range(wsLabel.Cells(startRow + 3, colOffset), wsLabel.Cells(startRow + 3, colOffset + 1))
            .Merge: .WrapText = True
        End With
        With wsLabel.Range(wsLabel.Cells(startRow + 4, colOffset), wsLabel.Cells(startRow + 4, colOffset + 1))
            .Merge: .WrapText = True
        End With

        ' Write Headers
        WriteCell wsLabel.Cells(startRow, colOffset), hPart, vPart, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow, colOffset + 1), hLot, vLot, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow + 1, colOffset), hSerial, vSerial, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow + 1, colOffset + 1), hNCR, vNCR, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow + 2, colOffset), hInsp, vInsp, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow + 2, colOffset + 1), hDisp, vDisp, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow + 3, colOffset), hReason, vReason, xlCenter, xlLeft
        WriteCell wsLabel.Cells(startRow + 4, colOffset), hComm, vComm, xlTop, xlLeft
        
NextIteration:
    Next i
    
    Application.ScreenUpdating = True
    If isBlankMode Then
        MsgBox "Generated blank forms (Page Full).", vbInformation
    Else
        MsgBox "Labels generated successfully!", vbInformation
    End If
End Sub

' ---------------------------------------------------------
' NEW MACRO: Clear Input Button
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