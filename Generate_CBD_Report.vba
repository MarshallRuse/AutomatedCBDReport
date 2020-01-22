Option Explicit

Sub Generate_CBD_Report()

    Call CBD_Report_Format_Extract
    Call CBD_Report_Create_Pivot_Table

End Sub

Sub CBD_Report_Format_Extract()
'
' CBD_Report_Format_Extract Macro
'

'
    Dim ExtractTable As ListObject
    Dim newCol As Range
    Dim newHeader As Range
    Dim blankCells As Range
    Dim fd As FileDialog
    Dim fileWasChosen As Boolean
    Dim DataWorkbook As Workbook
    Dim LookupTable As Workbook
    
    MsgBox "Choose Extract Dat for the report."
    Set fd = Application.FileDialog(msoFileDialogOpen)
    
    fd.Filters.Clear
    fd.Filters.Add "CSV Files", "*.csv"
    fd.FilterIndex = 1
    
    fd.AllowMultiSelect = False
    fd.InitialFileName = "N:\DOM1\Post Graduate Program"
    fd.Title = "Choose an Extract Data File"
    
    fileWasChosen = fd.Show
    
    If Not fileWasChosen Then
        MsgBox "You didn't choose an Extract Data File. Report generation was terminated."
        Exit Sub
    End If
    
    Set DataWorkbook = Workbooks.Open(Filename:=fd.SelectedItems(1))
    
    MsgBox "Choose a VLOOKUP Table for the report."
    Set fd = Application.FileDialog(msoFileDialogOpen)
    
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xl*"
    fd.FilterIndex = 1
    
    fd.AllowMultiSelect = False
    fd.InitialFileName = "N:\DOM1\Post Graduate Program"
    fd.Title = "Choose a CBD Lookup Table"
    
    fileWasChosen = fd.Show
    
    If Not fileWasChosen Then
        MsgBox "You didn't choose a Lookup Table. Report generation was terminated."
        Exit Sub
    End If
    
    Set LookupTable = Workbooks.Open(Filename:=fd.SelectedItems(1))
    DataWorkbook.Activate
    
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
        "=VLOOKUP([@[Assessment Form Code]],'[" & LookupTable.Name & "]VLOOKUP MASTER'!C1:C11,11,FALSE)"
    ExtractTable.ListColumns("EPA Code and Name").Range.EntireColumn.AutoFit
    
    ' Create column Site adjacent to Type of Assessment Form
    ExtractTable.ListColumns("Type of Assessment Form").Range.Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "Site"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP([@[CV ID 9533 : Site]],'[" & LookupTable.Name & "]Site'!C1:C2,2,FALSE)"
    ExtractTable.ListColumns("Site").Range.EntireColumn.AutoFit
    
    ' Create column Block adjacent to Type of Assessment Form
    ExtractTable.ListColumns("Type of Assessment Form").Range.Select
    Selection.Insert shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "Block"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP([@[Date of encounter]],'[" & LookupTable.Name & "]BLOCK'!C2:C6,3,TRUE)"
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

Sub CBD_Report_Create_Pivot_Table()
'
' CBD_Report_Create_Pivot_Table Macro
'

'
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ExtractTable", Version:=6).CreatePivotTable TableDestination:= _
        "Sheet1!R3C1", TableName:="CompletedEPAsPivotTable", DefaultVersion:=6
    Sheets("Sheet1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("CompletedEPAsPivotTable").PivotFields("Resident")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("CompletedEPAsPivotTable").PivotFields("EPA Code and Name")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("CompletedEPAsPivotTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("CompletedEPAsPivotTable").AddDataField ActiveSheet.PivotTables( _
        "CompletedEPAsPivotTable").PivotFields("Entrustment / Overall Category"), _
        "Count of Entrustment / Overall Category", xlCount
    Range("A4").Select
    ActiveSheet.PivotTables("CompletedEPAsPivotTable").CompactLayoutRowHeader = ""
    Range("B3").Select
    ActiveSheet.PivotTables("CompletedEPAsPivotTable").CompactLayoutColumnHeader = ""
    Range("A5").Select
    ActiveSheet.PivotTables("CompletedEPAsPivotTable").TableStyle2 = "PivotStyleMedium2"
    Range("A6").Select
    Selection.CurrentRegion.Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$A$83").AutoFilter Field:=1, Operator:= _
        xlFilterNoFill
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = -1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.ColumnWidth = 83.43
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveSheet.Range("$A$1:$A$83").AutoFilter Field:=1
    Range("A2").End(xlToRight).Select
    Selection.FormulaR1C1 = "Total Completed EPAs"
    Range("A2").End(xlDown).Select
    Selection.FormulaR1C1 = "Total Completed EPAs"
End Sub


