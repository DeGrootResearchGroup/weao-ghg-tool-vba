Attribute VB_Name = "AddPlantColumnModule"

Private Sub AddPlantColumnInputs(plantType as String, ByRef newPlantName As String, ByRef newPlantCellName As String)
    ' Set the worksheet to the Inputs sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Inputs")

    ' Set the header row where plant headers are located
    Dim headerRow As Integer
    headerRow = 17

    ' Find the last column with data in the header row
    Dim maxCol As Integer
    maxCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    ' Set the first plant column to default value to avoid searching columns with no labels;
    ' then, if the type is WTP, set the first plant column to the correct value
    Dim firstPlantCol As Integer
    firstPlantCol = 5
    If plantType = "WTP" Then
        For col = firstPlantCol To maxCol
            If Left(ws.Cells(headerRow, col).Name.Name, 3) = "WTP" Then
                firstPlantCol = col
                Exit For
            End If
        Next col
    End If

    ' Set the last row with plant data
    Dim maxRow As Integer
    maxRow = 206

    ' Find the last plant column.
    ' plantType can be either "WWTP" or "WTP", according to the naming of the cells
    Dim lastPlantCol As Integer
    lastPlantCol = firstPlantCol
    For col = firstPlantCol To maxCol
        If Not ws.Cells(headerRow, col).Name Is Nothing Then
            If plantType = "WWTP" Then
                If Left(ws.Cells(headerRow, col).Name.Name, 4) = "WWTP" Then
                    lastPlantCol = col
                End If
            ElseIf plantType = "WTP" Then
                If Left(ws.Cells(headerRow, col).Name.Name, 3) = "WTP" Then
                    lastPlantCol = col
                End If
            Else
                MsgBox "Invalid plant type specified. Please use 'WWTP' or 'WTP' as the plant type."
                Exit Sub
            End If
        End If
    Next col

    ' New plant column is the next column after the current last plant column
    Dim newPlantCol As Integer
    newPlantCol = lastPlantCol + 1

    ' Insert a new column at the new plant position
    ws.Columns(newPlantCol).Insert Shift:=xlToRight

    ' Determine the new plant name
    If plantType = "WWTP" Then
        newPlantName = "WWTP" & newPlantCol - firstPlantCol + 1
        newPlantCellName = newPlantName
    ElseIf plantType = "WTP" Then
        newPlantName = "WTP" & newPlantCol - firstPlantCol + 1
        newPlantCellName = newPlantName & "_"
    Else
        ' This error should already be handled above
        Exit Sub
    End If

    ' Set the value of the new header cell to the new plant name
    ws.Cells(headerRow, newPlantCol).Value = newPlantName

    ' Name the new header cell
    ws.Cells(headerRow, newPlantCol).Name = newPlantCellName

    ' Copy the data and formulas from the last plant column to the new column
    ws.Range(ws.Cells(headerRow + 1, lastPlantCol), ws.Cells(maxRow, lastPlantCol)).Copy Destination:=ws.Cells(headerRow + 1, newPlantCol)

    ' Adjust the formulas in the new column
    Dim cell As Range
    Dim lastPlantHeader As String
    lastPlantHeader = ws.Cells(headerRow, lastPlantCol).Name.Name

    For Each cell In ws.Range(ws.Cells(headerRow + 1, newPlantCol), ws.Cells(maxRow, newPlantCol))
        If cell.HasFormula Then
            cell.Formula = Replace(cell.Formula, lastPlantHeader, newPlantCellName)
            cell.Formula = Replace(cell.Formula, "@" & newPlantCellName, newPlantCellName)
        End If
    Next cell
End Sub

Private Sub AddPlantColumnToSheet(plantType as String, sheetName As String, headerRow As Integer, firstPlantCol As Integer, newPlantName As String, newPlantCellName As String, Optional rowsToSkip As Variant, Optional copyTable As Boolean = True, Optional rowsToShiftBack As Variant)
    ' Set the worksheet based on the provided sheet name
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Initialize the last plant column
    Dim lastPlantCol As Integer
    lastPlantCol = firstPlantCol

    ' Find the first and last plant columns
    Dim maxPlantNum As Integer
    maxPlantNum = 0
    Dim nChar As Integer
    If plantType = "WWTP" Then
        nChar = 4
    ElseIf plantType = "WTP" Then
        nChar = 3
    End If
    For Each nm In ThisWorkbook.Names
        If Left(nm.Name, nChar) = plantType Then
            ' Extract the numeric part of the plant name (WTP strings may contain underscore so they are removed)
            Dim plantNum As Integer
            plantNum = CInt(Replace(Mid(nm.Name, nChar + 1), "_", ""))

            ' Check if the named range is referenced in the header row of the specified sheet
            For col = firstPlantCol To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
                Dim formulaText As String
                formulaText = ws.Cells(headerRow, col).Formula
                If InStr(formulaText, nm.Name & "=") > 0 Then
                    If plantNum > maxPlantNum Then
                        lastPlantCol = col
                        maxPlantNum = plantNum
                    End If
                    If plantNum = 1 Then
                        firstPlantCol = col
                    End If
                End If
            Next col
        End If
    Next nm

    ' New plant column is the next column after the last plant column
    Dim newPlantCol As Integer
    newPlantCol = lastPlantCol + 1

    ' Determine the cell name for the last plant column
    If plantType = "WWTP" Then
        lastPlantCellName = "WWTP" & lastPlantCol - firstPlantCol + 1
    ElseIf plantType = "WTP" Then
        lastPlantCellName = "WTP" & lastPlantCol - firstPlantCol + 1 & "_"
    Else
        ' This error should already be handled above
        Exit Sub
    End If

    ' Insert a new column at the new WWTP position
    ws.Columns(newPlantCol).Insert Shift:=xlToRight

    ' Create the formula for the new header cell
    ws.Cells(headerRow, newPlantCol).Formula = "=IF(" & newPlantCellName & "="" "",""""," & newPlantCellName & ")"

    ' Copy the data and formulas from the last plant column to the new column
    Dim sourceRange As Range
    Set sourceRange = ws.Range(ws.Cells(headerRow + 1, lastPlantCol), ws.Cells(ws.Rows.Count, lastPlantCol).End(xlUp))

    ' Determine the range, excluding the skipped cells, if any
    Dim combinedSourceRange as Range
    Dim combinedTargetRange as Range
    If IsArray(rowsToSkip) Then
        For Each cell In sourceRange
            If IsError(Application.Match(cell.Row, rowsToSkip, 0)) Then
                If combinedSourceRange Is Nothing Then
                    Set combinedSourceRange = cell
                    Set combinedTargetRange = ws.Cells(cell.Row, newPlantCol)
                Else
                    Set combinedSourceRange = Union(combinedSourceRange, cell)
                    Set combinedTargetRange = Union(combinedTargetRange, ws.Cells(cell.Row, newPlantCol))
                End If
            End If
        Next cell
        ' Iterate over individual areas and copy them
        Dim area As Range
        For Each area In combinedSourceRange.Areas
            area.Copy Destination:=ws.Cells(area.Row, newPlantCol)
        Next area
    Else
        ' If no rows to skip, copy the entire range
        sourceRange.Copy Destination:=ws.Cells(headerRow + 1, newPlantCol)
    End If

    ' Adjust the formulas in the new column
    For Each cell In ws.Range(ws.Cells(headerRow + 1, newPlantCol), ws.Cells(ws.Rows.Count, newPlantCol).End(xlUp))
        If cell.HasFormula Then
            cell.Formula = Replace(cell.Formula, lastPlantCellName, newPlantCellName)
        End If
    Next cell

    ' Optionally copy the values and formatting from the last column of the table on lines 4-5 to the new column
    ' This should only be specified for WWTP columns, not WTP columns
    If copyTable Then
        ws.Range(ws.Cells(4, lastPlantCol), ws.Cells(5, lastPlantCol)).Copy
        ws.Cells(4, newPlantCol).PasteSpecial Paste:=xlPasteFormats
        ws.Cells(4, newPlantCol).PasteSpecial Paste:=xlPasteValues
    End If

    ' Shift specified rows back to their original position
    If Not IsMissing(rowsToShiftBack) Then
        Dim row As Variant
        For Each row In rowsToShiftBack
            ws.Cells(row, newPlantCol).Cut Destination:=ws.Cells(row, newPlantCol - 1)
        Next row
    End If

    ' Clear clipboard to remove the marching ants border
    Application.CutCopyMode = False
End Sub

Private Function GenerateConsecutiveNumbersArray(startNum As Integer, endNum As Integer) As Variant
    Dim arr() As Variant
    ReDim arr(0 To endNum)
    Dim i As Integer
    For i = startNum To endNum
        arr(i - startNum) = i
    Next i
    GenerateConsecutiveNumbersArray = arr
End Function

Private Function GenerateTwoConsecutiveNumbersArrays(startNum1 As Integer, endNum1 As Integer, startNum2 As Integer, endNum2 As Integer) As Variant
    Dim arr1() As Variant
    Dim arr2() As Variant
    ReDim arr1(0 To endNum1 - startNum1)
    ReDim arr2(0 To endNum2 - startNum2)
    Dim i As Integer
    For i = startNum1 To endNum1
        arr1(i - startNum1) = i
    Next i
    For i = startNum2 To endNum2
        arr2(i - startNum2) = i
    Next i
    GenerateTwoConsecutiveNumbersArrays = Split(Join(arr1, ",") & "," & Join(arr2, ","), ",")
End Function

Private Sub AddWWTPColumnProcesses(newPlantName As String, newPlantCellName As String)
    AddPlantColumnToSheet "WWTP", "Scope 1 - Process", 12, 7, newPlantName, newPlantCellName, , True
End Sub

Private Sub AddWWTPColumnCombustion(newPlantName As String, newPlantCellName As String)
    Dim rowsToSkip As Variant
    rowsToSkip = GenerateConsecutiveNumbersArray(117, 122)
    AddPlantColumnToSheet "WWTP", "Scope 1 - Combustion", 10, 7, newPlantName, newPlantCellName, rowsToSkip, True
End Sub

Private Sub AddWTPColumnCombustion(newPlantName As String, newPlantCellName As String)
    Dim rowsToSkip As Variant
    rowsToSkip = GenerateConsecutiveNumbersArray(117, 122)
    AddPlantColumnToSheet "WTP", "Scope 1 - Combustion", 73, 7, newPlantName, newPlantCellName, rowsToSkip, True
End Sub

Private Sub AddWWTPColumnElectricity(newPlantName As String, newPlantCellName As String)
    AddPlantColumnToSheet "WWTP", "Scope 2 - Electricity", 6, 7, newPlantName, newPlantCellName, , False
End Sub

Private Sub AddWTPColumnElectricity(newPlantName As String, newPlantCellName As String)
    AddPlantColumnToSheet "WTP", "Scope 2 - Electricity", 6, 7, newPlantName, newPlantCellName, , False
End Sub

Private Sub AddWWTPColumnElectricityUpstream(newPlantName As String, newPlantCellName As String)
    AddPlantColumnToSheet "WWTP", "Scope 3 - Electricity", 6, 7, newPlantName, newPlantCellName, , False
End Sub

Private Sub AddWTPColumnElectricityUpstream(newPlantName As String, newPlantCellName As String)
    AddPlantColumnToSheet "WTP", "Scope 3 - Electricity", 6, 7, newPlantName, newPlantCellName, , False
End Sub

Private Sub AddWWTPColumnFuelUpstream(newPlantName As String, newPlantCellName As String)
    Dim rowsToSkip As Variant
    rowsToSkip = GenerateConsecutiveNumbersArray(45, 48)
    AddPlantColumnToSheet "WWTP", "Scope 3 - Fuel upstream", 6, 7, newPlantName, newPlantCellName, rowsToSkip, False
End Sub

Private Sub AddWTPColumnFuelUpstream(newPlantName As String, newPlantCellName As String)
    Dim rowsToSkip As Variant
    rowsToSkip = GenerateConsecutiveNumbersArray(45, 48)
    AddPlantColumnToSheet "WTP", "Scope 3 - Fuel upstream", 6, 7, newPlantName, newPlantCellName, rowsToSkip, False
End Sub

Private Sub AddWWTPColumnBiosolids(newPlantName As String, newPlantCellName As String)
    AddPlantColumnToSheet "WWTP", "Scope 3 - Biosolids", 8, 7, newPlantName, newPlantCellName, , False
End Sub

Private Sub AddWWTPColumnChemicals(newPlantName As String, newPlantCellName As String)
    Dim rowsToSkip As Variant
    rowsToSkip = GenerateConsecutiveNumbersArray(76, 79)
    AddPlantColumnToSheet "WWTP", "Scope 3 - Chemicals", 30, 7, newPlantName, newPlantCellName, rowsToSkip, False
End Sub

Private Sub AddWTPColumnChemicals(newPlantName As String, newPlantCellName As String)
    Dim rowsToSkip As Variant
    rowsToSkip = GenerateConsecutiveNumbersArray(76, 79)
    AddPlantColumnToSheet "WTP", "Scope 3 - Chemicals", 30, 7, newPlantName, newPlantCellName, rowsToSkip, False
End Sub

Private Sub AddWWTPColumnSummary(newPlantName As String, newPlantCellName As String)
    Dim firstWWTPCol As Integer
    Dim lastWWTPCol As Integer
    Dim newWWTPCol As Integer
    Dim col As Integer
    Dim i As Integer
    Dim formulaPart1 As String
    Dim formulaPart2 As String
    Dim booleanCheckCellRef As String
    Dim booleanCheckCol As Integer

    ' Set the worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Summary")

    ' Set the header row and first WWTP column
    Dim headerRow As Integer
    headerRow = 3
    firstWWTPCol = 6

    ' Find the last WWTP column by checking for "WWTP" in the header formula
    lastWWTPCol = firstWWTPCol
    For col = firstWWTPCol To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        If InStr(ws.Cells(headerRow, col).Formula, "WWTP") > 0 Then
            lastWWTPCol = col
        End If
    Next col

    ' New WWTP column is the next column after the last WWTP column
    newWWTPCol = lastWWTPCol + 1

    ' Capture the boolean check cell reference in row 50 before adding the column
    Dim formulaRow51 As String
    formulaRow51 = ws.Cells(51, firstWWTPCol).Formula
    booleanCheckCellRef = ExtractBooleanCheckCellRef(formulaRow51)
    booleanCheckCol = Range(booleanCheckCellRef).Column

    ' Unmerge cells in the WWTP columns for rows 46, 51, and 52
    ws.Range(ws.Cells(46, firstWWTPCol), ws.Cells(46, lastWWTPCol)).UnMerge
    ws.Range(ws.Cells(51, firstWWTPCol), ws.Cells(51, lastWWTPCol)).UnMerge
    ws.Range(ws.Cells(52, firstWWTPCol), ws.Cells(52, lastWWTPCol)).UnMerge

    ' Call the generic function to add the column
    AddPlantColumnToSheet "WWTP", "Summary", 3, 6, newPlantName, newPlantCellName, , False

    ' Merge the new cells in rows 46, 51, and 52
    ws.Range(ws.Cells(46, firstWWTPCol), ws.Cells(46, newWWTPCol)).Merge
    ws.Range(ws.Cells(51, firstWWTPCol), ws.Cells(51, newWWTPCol)).Merge
    ws.Range(ws.Cells(52, firstWWTPCol), ws.Cells(52, newWWTPCol)).Merge

    ' Update the formula in row 46
    ws.Cells(46, firstWWTPCol).Formula = "=SUM(" & ws.Cells(45, firstWWTPCol).Address(False, False) & ":" & ws.Cells(45, newWWTPCol).Address(False, False) & ")"

    ' Cut and paste the formula in row 47 from the previous last WWTP column to the new WWTP column
    ws.Cells(47, lastWWTPCol).Cut Destination:=ws.Cells(47, newWWTPCol)

    ' Update the formula in row 51
    formulaPart1 = "=IF(" & ws.Cells(50, booleanCheckCol + 1).Address(False, False) & "=TRUE,("
    formulaPart2 = ") / SUM("
    For i = firstWWTPCol To newWWTPCol
        formulaPart1 = formulaPart1 & ws.Cells(50, i).Address(False, False) & "*" & ws.Cells(49, i).Address(False, False) & "+"
        formulaPart2 = formulaPart2 & ws.Cells(49, i).Address(False, False) & ","
    Next i
    formulaPart1 = Left(formulaPart1, Len(formulaPart1) - 1) ' Remove the last "+"
    formulaPart2 = Left(formulaPart2, Len(formulaPart2) - 1) ' Remove the last ","
    formulaPart2 = formulaPart2 & "),0)"
    ws.Cells(51, firstWWTPCol).Formula = formulaPart1 & formulaPart2

    ' Update the formula in row 52
    ws.Cells(52, firstWWTPCol).Formula = "=IF(F48=0,0,(IF(F46=0,0,F46/F48)))"

    ' Clear clipboard to remove the marching ants border
    Application.CutCopyMode = False
End Sub

Private Function ExtractBooleanCheckCellRef(formula As String) As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim cellRef As String

    startPos = InStr(formula, "(") + 1
    endPos = InStr(startPos, formula, "=") - 1

    If startPos > 0 And endPos > 0 Then
        cellRef = Mid(formula, startPos, endPos - startPos)
        ExtractBooleanCheckCellRef = Trim(cellRef)
    Else
        MsgBox "Error extracting boolean check cell reference from formula: " & formula
    End If
End Function

Private Sub AddWTPColumnSummary(newPlantName As String, newPlantCellName As String)
    Dim firstWTPCol As Integer
    Dim lastWTPCol As Integer
    Dim newWTPCol As Integer
    Dim col As Integer
    Dim i As Integer
    Dim formulaPart1 As String
    Dim formulaPart2 As String
    Dim booleanCheckCellRef As String
    Dim booleanCheckCol As Integer

    ' Set the worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Summary")

    ' Set the header row and first WTP column
    Dim headerRow As Integer
    headerRow = 3

    ' Find the last column with data in the header row
    Dim maxCol As Integer
    maxCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    ' Find the first WTP column
    firstWTPCol = 6
    Dim found As Boolean
    found = False
    For col = firstWTPCol To maxCol
        cellFormula = ws.Cells(headerRow, col).Formula
        ' Check if the cell contains a formula
        If Len(cellFormula) > 0 Then
            ' Loop through all named ranges in the workbook
            For Each nm In ThisWorkbook.Names
                ' Check if the named range starts with "WTP" and is referenced in the cell's formula
                If Left(nm.Name, 3) = "WTP" Then
                    If InStr(cellFormula, nm.Name) > 0 Then
                        firstWTPCol = col
                        found = True
                        Exit For
                    End If
                End If
            Next nm
            ' Exit the outer loop if the column is found
            If found Then Exit For
        End If
    Next col

    ' Find the last WTP column by checking for "WTP" in the header formula
    lastWTPCol = firstWTPCol
    For col = firstWTPCol To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        If InStr(ws.Cells(headerRow, col).Formula, "WTP") > 0 Then
            lastWTPCol = col
        End If
    Next col

    ' New WTP column is the next column after the last WTP column
    newWTPCol = lastWTPCol + 1

    ' Capture the boolean check cell reference in row 50 before adding the column
    Dim formulaRow51 As String
    formulaRow51 = ws.Cells(51, firstWTPCol).Formula
    booleanCheckCellRef = ExtractBooleanCheckCellRef(formulaRow51)
    booleanCheckCol = Range(booleanCheckCellRef).Column

    ' Unmerge cells in the WTP columns for rows 46, 51, and 52
    ws.Range(ws.Cells(46, firstWTPCol), ws.Cells(46, lastWTPCol)).UnMerge
    ws.Range(ws.Cells(51, firstWTPCol), ws.Cells(51, lastWTPCol)).UnMerge
    ws.Range(ws.Cells(52, firstWTPCol), ws.Cells(52, lastWTPCol)).UnMerge

    ' Call the generic function to add the column
    AddPlantColumnToSheet "WTP", "Summary", 3, 6, newPlantName, newPlantCellName, , False

    ' Merge the new cells in rows 46, 51, and 52
    ws.Range(ws.Cells(46, firstWTPCol), ws.Cells(46, newWTPCol)).Merge
    ws.Range(ws.Cells(51, firstWTPCol), ws.Cells(51, newWTPCol)).Merge
    ws.Range(ws.Cells(52, firstWTPCol), ws.Cells(52, newWTPCol)).Merge

    ' Update the formula in row 46
    ws.Cells(46, firstWTPCol).Formula = "=SUM(" & ws.Cells(45, firstWTPCol).Address(False, False) & ":" & ws.Cells(45, newWTPCol).Address(False, False) & ")"

    ' Cut and paste the formula in row 47 from the previous last WTP column to the new WTP column
    ws.Cells(47, lastWTPCol).Cut Destination:=ws.Cells(47, newWTPCol)

    ' Update the formula in row 51
    formulaPart1 = "=IF(" & ws.Cells(50, booleanCheckCol + 1).Address(False, False) & "=TRUE,("
    formulaPart2 = ") / SUM("
    For i = firstWTPCol To newWTPCol
        formulaPart1 = formulaPart1 & ws.Cells(50, i).Address(False, False) & "*" & ws.Cells(49, i).Address(False, False) & "+"
        formulaPart2 = formulaPart2 & ws.Cells(49, i).Address(False, False) & ","
    Next i
    formulaPart1 = Left(formulaPart1, Len(formulaPart1) - 1) ' Remove the last "+"
    formulaPart2 = Left(formulaPart2, Len(formulaPart2) - 1) ' Remove the last ","
    formulaPart2 = formulaPart2 & "),0)"
    ws.Cells(51, firstWTPCol).Formula = formulaPart1 & formulaPart2

    ' Update the formula in row 52
    ws.Cells(52, firstWTPCol).Formula = "=IF(F48=0,0,(IF(J46=0,0,J46/F48)))"

    ' Clear clipboard to remove the marching ants border
    Application.CutCopyMode = False
End Sub

Sub AddWWTPColumn()
    ' Disable screen updating, events, and set calculation to manual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim newPlantName As String
    Dim newPlantCellName As String

    AddPlantColumnInputs "WWTP", newPlantName, newPlantCellName
    AddWWTPColumnProcesses newPlantName, newPlantCellName
    AddWWTPColumnCombustion newPlantName, newPlantCellName
    AddWWTPColumnElectricity newPlantName, newPlantCellName
    AddWWTPColumnElectricityUpstream newPlantName, newPlantCellName
    AddWWTPColumnFuelUpstream newPlantName, newPlantCellName
    AddWWTPColumnBiosolids newPlantName, newPlantCellName
    AddWWTPColumnChemicals newPlantName, newPlantCellName
    AddWWTPColumnSummary newPlantName, newPlantCellName

    ' Re-enable screen updating, events, and set calculation back to automatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False ' Clear the clipboard
End Sub

Sub AddWTPColumn()
    ' Disable screen updating, events, and set calculation to manual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim newPlantName As String
    Dim newPlantCellName As String

    AddPlantColumnInputs "WTP", newPlantName, newPlantCellName
    AddWTPColumnCombustion newPlantName, newPlantCellName
    AddWTPColumnElectricity newPlantName, newPlantCellName
    AddWTPColumnElectricityUpstream newPlantName, newPlantCellName
    AddWTPColumnFuelUpstream newPlantName, newPlantCellName
    AddWTPColumnChemicals newPlantName, newPlantCellName
    AddWTPColumnSummary newPlantName, newPlantCellName

    ' Re-enable screen updating, events, and set calculation back to automatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False ' Clear the clipboard
End Sub