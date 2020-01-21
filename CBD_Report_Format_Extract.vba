Option Explicit

Sub CBD_Report_Format_Extract()
'
' CBD_Report_Format_Extract Macro
'

'
    Dim ExtractTable As ListObject
    Dim newCol As Range
    Dim newHeader As Range
    Dim blankCells As Range
    
    ' Delete the first 3 rows
    Rows("1:3").Select
    Selection.Delete shift:=xlUp
    
    ' Resize all columns
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    ' Create a table from the data for ease of manipulation - allows use of column
    ' header names for flexibility
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1").CurrentRegion, , xlYes).Name = _
        "ExtractTable"
    Set ExtractTable = ActiveSheet.ListObjects("ExtractTable")

    ' Rename the values in Entrustment / Overall Category
    ExtractTable.ListColumns("Entrustment / Overall Category").Range.Select
    Selection.Replace What:="Excellence", Replacement:="5. Excellence", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Autonomy", Replacement:="4. Autonomy", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Support", Replacement:="3. Support", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Direction", Replacement:="2. Direction", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Intervention", Replacement:="1. Intervention", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False
        
    ' Create column EPA Code and Name adjacent to Type of Assessment Form
    ExtractTable.ListColumns("Type of Assessment Form").Range.Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "EPA Code and Name"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP([@[Assessment Form Code]],'[VLOOKUP MASTER - 2019-2020 (version 1).xlsb]VLOOKUP MASTER'!C1:C11,11,FALSE)"
    ExtractTable.ListColumns("EPA Code and Name").Range.EntireColumn.AutoFit
    
    ' Create column Site adjacent to Type of Assessment Form
    ExtractTable.ListColumns("Type of Assessment Form").Range.Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "Site"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP([@[CV ID 9533 : Site]],'[VLOOKUP MASTER - 2019-2020 (version 1).xlsb]Site'!C1:C2,2,FALSE)"
    ExtractTable.ListColumns("Site").Range.EntireColumn.AutoFit
    
    ' Create column Block adjacent to Type of Assessment Form
    ExtractTable.ListColumns("Type of Assessment Form").Range.Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "Block"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP([@[Date of encounter]],'[VLOOKUP MASTER - 2019-2020 (version 1).xlsb]BLOCK'!C2:C6,3,TRUE)"
    ExtractTable.ListColumns("Block").Range.EntireColumn.AutoFit
    
    ' Create column Resident from Assessee Lastname and Assessee Firstname
    ExtractTable.ListColumns("Assessee Lastname").Range.EntireColumn.Offset(0, 1).Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "Resident"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = _
        "=UPPER([@[Assessee Lastname]])&"", ""&[@[Assessee Firstname]]"
    ExtractTable.ListColumns("Resident").Range.EntireColumn.AutoFit
    
    ' Delete Rows with blank cells in Date of Assessment Form Submission
    Set blankCells = ExtractTable.ListColumns("Date of Assessment Form Submission").Range.SpecialCells(xlCellTypeBlanks)
    If Not blankCells Is Nothing Then
        blankCells.Delete xlUp
    End If
End Sub
