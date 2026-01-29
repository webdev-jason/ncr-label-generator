Sub GenerateLabels()
    Dim wsInput As Worksheet
    Dim wsLabel As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim labelRow As Long
    Dim labelCol As Long
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
    
    ' --- THE FIX IS HERE ---
    ' .Clear wipes text, borders, and formatting so no "ghost" boxes remain
    wsLabel.Cells.Clear
    
    ' Initialize variables
    labelRow = 1
    labelCol = 1
    labelCounter = 1
    
    ' --- CHECK MODE ---
    ' If lastRow is 1 (only headers), we are in Blank Mode
    If lastRow < 2 Then
        isBlankMode = True
        loopLimit = 10 ' Generate 1 full page (10 labels)
    Else
        isBlankMode = False
        loopLimit = lastRow
    End If
    
    ' --- MAIN LOOP ---
    Dim startLoop As Long
    If isBlankMode Then startLoop = 1 Else startLoop = 2
    
    For i = startLoop To loopLimit
        
        Dim partText As String
        Dim lotText As String
        Dim serialText As String
        Dim ncrText As String
        Dim reasonText As String
        Dim inspText As String
        Dim commText As String
        
        ' Variable to control the horizontal "Center" point
        Dim centerPoint As Integer
        centerPoint = 35
        
        ' Determine content based on mode
        If isBlankMode Then
            ' BLANK MODE:
            partText = "Part #: " & Space(centerPoint - Len("Part #: "))
            lotText = "Lot #: "
            
            serialText = "Serial #: " & Space(centerPoint - Len("Serial #: "))
            ncrText = "NCR #: "
            
            inspText = "Inspected By:"
            reasonText = "Reason for Failure:"
            commText = "Comments:"
        Else
            ' DATA MODE:
            If wsInput.Cells(i, 1).Value = "" Then GoTo NextIteration
            
            Dim rawPart As String, rawLot As String, rawSerial As String, rawNCR As String
            
            rawPart = "Part #: " & wsInput.Cells(i, 1).Value
            rawLot = "Lot #: " & wsInput.Cells(i, 2).Value
            rawSerial = "Serial #: " & wsInput.Cells(i, 3).Value
            rawNCR = "NCR #: " & wsInput.Cells(i, 4).Value
            
            ' Calculate needed padding
            Dim padPart As Integer, padSerial As Integer
            padPart = centerPoint - Len(rawPart)
            If padPart < 1 Then padPart = 1
            
            padSerial = centerPoint - Len(rawSerial)
            If padSerial < 1 Then padSerial = 1
            
            partText = rawPart & Space(padPart)
            lotText = rawLot
            serialText = rawSerial & Space(padSerial)
            ncrText = rawNCR
            
            inspText = "Inspected By: " & wsInput.Cells(i, 6).Value
            reasonText = "Reason for Failure: " & wsInput.Cells(i, 5).Value
            commText = "Comments: " & wsInput.Cells(i, 7).Value
        End If

        ' Format the Label Cell
        With wsLabel.Cells(labelRow, labelCol)
            .Value = partText & lotText & vbNewLine & _
                     serialText & ncrText & vbNewLine & vbNewLine & _
                     inspText & vbNewLine & vbNewLine & _
                     reasonText & vbNewLine & vbNewLine & _
                     commText
            
            ' Apply text wrapping and alignment
            .WrapText = True
            .VerticalAlignment = xlTop 
            .HorizontalAlignment = xlLeft
            .Font.Name = "Arial"
            .Font.Size = 10
            .Borders.LineStyle = xlContinuous
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

NextIteration:
    Next i
    
    Application.ScreenUpdating = True
    
    If isBlankMode Then
        MsgBox "Generated blank forms.", vbInformation
    Else
        MsgBox "Labels generated successfully!", vbInformation
    End If
    
End Sub