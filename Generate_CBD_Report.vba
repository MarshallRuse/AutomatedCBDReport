Option Explicit

    Dim DataWorkbook As Workbook
    Dim DataSheet As Worksheet
    Dim LookupTable As Workbook
    Dim ExtractTable As ListObject
    Dim CommentsErrorLog As String

Sub Generate_CBD_Report()
    Application.ScreenUpdating = False
    Call CBD_Report_Format_Extract
    Call CBD_Report_Create_Pivot_Table
    Call CBD_Report_Create_Comments_Lookup
    Call CBD_Report_Add_Comments_to_Pivot_Copy
    Call CBD_Report_Create_Site_Pivot_Table_And_Copy
    Call CBD_Report_Create_Block_Pivot_Table_And_Copy
    Call CBD_Report_Cleanup
    Application.ScreenUpdating = True
    DataWorkbook.Worksheets("ResidentAnalysis").Activate
    Range("A1").Select
    Application.CutCopyMode = False

End Sub

Sub CBD_Report_Format_Extract()
'
' CBD_Report_Format_Extract Macro
'

'
    Dim fso As Scripting.FileSystemObject
    Dim ExtractCopy As String
    Dim newCol As Range
    Dim newHeader As Range
    Dim blankCells As Range
    Dim fd As FileDialog
    Dim fileWasChosen As Boolean
    
    Set fso = New Scripting.FileSystemObject
    
    ' Message and File Dialog box to choose the Extract Data
    MsgBox "Choose Extract Data for the report."
    Set fd = Application.FileDialog(msoFileDialogOpen)
    
    fd.Filters.Clear
    fd.Filters.Add "Excel, CSV Files", "*.xl*, *.csv"
    fd.FilterIndex = 1
    
    fd.AllowMultiSelect = False
    fd.InitialFileName = "N:\DOM1\Post Graduate Program"
    fd.Title = "Choose an Extract Data File"
    
    fileWasChosen = fd.Show
    
    If Not fileWasChosen Then
        MsgBox "You didn't choose an Extract Data File. Report generation was terminated."
        End
    End If
    
    ' Creates a copy of the extract data to preserve the original. Stored in a variable for reference when
    ' opening the new workbook
    ExtractCopy = Left(fd.SelectedItems(1), InStrRev(fd.SelectedItems(1), ".") - 1) & " copy." & _
        Right(fd.SelectedItems(1), Len(fd.SelectedItems(1)) - InStrRev(fd.SelectedItems(1), "."))
    
    fso.CopyFile fd.SelectedItems(1), ExtractCopy
    
    Set DataWorkbook = Workbooks.Open(Filename:=ExtractCopy)
    
    ' Rename the workbook copy with the appropriate division name and date range
    DataWorkbook.SaveAs DataWorkbook.Path & "\" & DataWorkbook.ActiveSheet.Range("B3").Value & _
        "_" & DataWorkbook.ActiveSheet.Range("B2").Value & ".xlsx", 51
    Set DataSheet = DataWorkbook.Worksheets(1)
    DataSheet.Name = "DataExtract"
    
    
    
    ' Message Box and File Dialog to choose the VLOOKUP Sheet
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
        End
    End If
    
    Set LookupTable = Workbooks.Open(Filename:=fd.SelectedItems(1))
    
    
    
    ' Start formatting the Extract Data
    DataWorkbook.Activate
    
    ' Delete the first 3 rows
    Rows("1:3").Select
    Selection.Delete Shift:=xlUp
    
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
        
    ' Create column EPA Code and Name adjacent to Type of Assessment Form column
    ExtractTable.ListColumns("Type of Assessment Form").Range.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "EPA Code and Name"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP([@[Assessment Form Code]],'[" & LookupTable.Name & "]VLOOKUP MASTER'!C1:C11,11,FALSE), ""NONEXISTANT_FORM_ID"")"
    ExtractTable.ListColumns("EPA Code and Name").Range.EntireColumn.AutoFit
    
    'Loop through and make sure all EPA Codes are accounted for
        ' Filter
    ExtractTable.ListColumns("EPA Code and Name").Range.AutoFilter Field:=ExtractTable.ListColumns("EPA Code and Name").Range.Column, Criteria1 _
        :="NONEXISTANT_FORM_ID"
    If ExtractTable.ListColumns("EPA Code and Name").Range.Cells(1, 0).Count > 1 Then
        MsgBox Prompt:="Some Assessment Form Codes could not be found in the Lookup Table.  To find the erroneous codes, filter" & _
            " Assessment Form Code to NONEXISTANT_FORM_ID, and fix the corresponding Assessment Form Codes in the Lookup Table, or else delete the offending rows."
        End
    End If
    ActiveSheet.ListObjects("ExtractTable").Range.AutoFilter Field:=ExtractTable.ListColumns("EPA Code and Name").Range.Column ' Unfilter

    
    ' Create column Site adjacent to Type of Assessment Form
    ExtractTable.ListColumns("Type of Assessment Form").Range.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "Site"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP([@[CV ID 9533 : Site]],'[" & LookupTable.Name & "]Site'!C1:C2,2,FALSE)"
    ExtractTable.ListColumns("Site").Range.EntireColumn.AutoFit
    
    ' Create column Block adjacent to Type of Assessment Form
    ExtractTable.ListColumns("Type of Assessment Form").Range.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "Block"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP([@[Date of encounter]],'[" & LookupTable.Name & "]BLOCK'!C2:C6,3,TRUE)"
    ExtractTable.ListColumns("Block").Range.EntireColumn.AutoFit
    
    ' Create column Resident from Assessee Lastname and Assessee Firstname
    ExtractTable.ListColumns("Assessee Lastname").Range.EntireColumn.Offset(0, 1).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
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
    DataWorkbook.Activate
End Sub

Sub CBD_Report_Create_Pivot_Table()
'
' CBD_Report_Create_Pivot_Table Macro
'

'
    Dim PivotTableSheet As Worksheet
    Dim ResAnSheet As Worksheet
    Dim pivotTableLastRow As Range
    Dim firstCell As Range
    Dim lastCell As Range
    
    DataWorkbook.Activate
    Sheets.Add.Name = "PivotTable"
    Set PivotTableSheet = Worksheets("PivotTable")
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ExtractTable", Version:=6).CreatePivotTable TableDestination:= _
        "PivotTable!R3C1", TableName:="CompletedEPAsPivotTable", DefaultVersion:=6
    Sheets("PivotTable").Select
    Cells(3, 1).Select
    
    ' Residents as first layer of rows
    With ActiveSheet.PivotTables("CompletedEPAsPivotTable").PivotFields("Resident")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    ' EPA Code and Name as second layer of rows
    With ActiveSheet.PivotTables("CompletedEPAsPivotTable").PivotFields("EPA Code and Name")
        .Orientation = xlRowField
        .Position = 2
    End With
    
    ' Entrustment / Overall Category as columns
    With ActiveSheet.PivotTables("CompletedEPAsPivotTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    ' Count of Entrustment / Overall Category as values
    With ActiveSheet.PivotTables("CompletedEPAsPivotTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlDataField
    End With
    
    ' Style the pivot table
    Range("A4").Select
    ActiveSheet.PivotTables("CompletedEPAsPivotTable").CompactLayoutRowHeader = ""
    Range("B3").Select
    ActiveSheet.PivotTables("CompletedEPAsPivotTable").CompactLayoutColumnHeader = ""
    Range("A5").Select
    ActiveSheet.PivotTables("CompletedEPAsPivotTable").TableStyle2 = "PivotStyleMedium2"
    Range("A6").Select
    
    ' Copy the Pivot Tables values
    
    Selection.CurrentRegion.Select
    Selection.Copy
    Set ResAnSheet = Sheets.Add(After:=ActiveSheet)
    ResAnSheet.Name = "ResAn"
    
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
    
    ' Add the Number of Entrustments Column
    Range("A2").End(xlToRight).Offset(0, 1).Select
    Selection.Value = "Number of Entrustments"
    Set firstCell = Cells(Selection.Offset(1, 0).Row, Selection.Column)
    Set lastCell = Cells(Range("A2").End(xlDown).Row, Selection.Column)
    Range(firstCell, lastCell).FormulaR1C1 = "=SUM(RC[-3]:RC[-2])"
    Selection.Offset(0, -1).EntireColumn.Select
    Selection.Copy
    Selection.Offset(0, 1).EntireColumn.Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.EntireColumn.AutoFit
        
    ' Add the Target Entrustments Column
    Range("A2").End(xlToRight).Offset(0, 1).Select
    Selection.Value = "Target Entrustments"
    Set firstCell = Cells(Selection.Offset(1, 0).Row, Selection.Column)
    Set lastCell = Cells(Range("A2").End(xlDown).Row, Selection.Column)
    Range(firstCell, lastCell).Select
    Selection.FormulaR1C1 = _
        "=VLOOKUP(RC1,'[" & LookupTable.Name & "]VLOOKUP MASTER'!C11:C12,2,FALSE)"
    Selection.SpecialCells(xlCellTypeFormulas, xlErrors).Clear
    Selection.Offset(0, -1).EntireColumn.Select
    Selection.Copy
    Selection.Offset(0, 1).EntireColumn.Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.EntireColumn.AutoFit
    
    Selection.Offset(0, -1).EntireColumn.Select
    Dim c As Range
    For Each c In Range(Cells(3, Selection.Column), Cells(Range("A2").End(xlDown).Row, Selection.Column))
        If c.Offset(0, 1).Value > c.Value Then
            c.Style = "Bad"
        End If
    Next c
    
    ExtractTable.ListColumns("2 - 3 Strengths").Range.Replace What:="=-", Replacement:="-.", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    ExtractTable.ListColumns("2 - 3 Actions or areas for improvement").Range.Replace What:="=-", Replacement:="-.", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
End Sub

Sub CBD_Report_Create_Comments_Lookup()

    Dim newSheet As Worksheet
    Dim lngLastRow As String
    Dim lastRow As Long
    Dim lastcolumn As Long
    Dim lookupCol As Range
    Dim lookupHeader As Range
    
    DataWorkbook.Activate
    ' Create a new worksheet for the results
    Set newSheet = DataWorkbook.Worksheets.Add
    newSheet.Name = "CommentsLookup"
    
    ' Get all values from Residents Names
    ExtractTable.ListColumns("Resident").Range.Copy
    newSheet.Columns("A:A").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    ' Get all values from EPAs Code and Name
    ExtractTable.ListColumns("EPA Code and Name").Range.Copy
    newSheet.Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    ' Get all values from 2-3 Strengths
    ExtractTable.ListColumns("2 - 3 Strengths").Range.Copy
    newSheet.Columns("C:C").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ' Get all values from 2-3 Actions or areas for improvement
    ExtractTable.ListColumns("2 - 3 Actions or areas for improvement").Range.Copy
    newSheet.Columns("D:D").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    lastRow = ActiveSheet.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.Count
    lastcolumn = ActiveSheet.UsedRange.Column - 1 + ActiveSheet.UsedRange.Columns.Count
    
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range(Cells(2, 1), Cells(lastRow, 1)), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range(Cells(2, 2), Cells(lastRow, 2)), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With ActiveSheet.Sort
    .SetRange Range(Cells(1, 1), Cells(lastRow, lastcolumn))
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
    End With
    
    CommentsErrorLog = ""
    Dim i As Long
    ' currentColumn used for Error Handling with CommentsError
    Dim currentColumn As Integer
    

    For i = lastRow To 2 Step -1
        
        If Not ActiveWorkbook.Name = DataWorkbook.Name Then
            DataWorkbook.Activate
            newSheet.Activate
        End If
        
    
        ' Ugly Proactive Error Checking (for #NAME errors in comments mostly)
        ' If the errors on the next line of the strengths comments
        If IsError(Cells(i - 1, 3).Value) Then
            CommentsErrorLog = CommentsErrorLog & Cells(i - 1, 1).Value & " - " & Cells(i - 1, 2).Value & " - Strengths" & vbNewLine
            Cells(i - 1, 3).Value = "Value Omitted: Excel Error"
        End If
        
        ' If the Error wasn't caught by the previous iteration and is on this line of the strengths comments
        If IsError(Cells(i, 3).Value) Then
            CommentsErrorLog = CommentsErrorLog & Cells(i, 1).Value & " - " & Cells(i, 2).Value & " - Strengths" & vbNewLine
            Cells(i, 3).Value = "Value Omitted: Excel Error"
        End If
        
        ' If the errors on the next line of the weaknesses comments
        If IsError(Cells(i - 1, 4).Value) Then
            CommentsErrorLog = CommentsErrorLog & Cells(i - 1, 1).Value & " - " & Cells(i - 1, 2).Value & " - Weaknesses" & vbNewLine
            Cells(i - 1, 4).Value = "Value Omitted: Excel Error"
        End If

        ' If the Error wasn't caught by the previous iteration and is on this line of the weaknesses comments
        If IsError(Cells(i, 4).Value) Then
            CommentsErrorLog = CommentsErrorLog & Cells(i, 1).Value & " - " & Cells(i, 2).Value & " - Weaknesses" & vbNewLine
            Cells(i, 4).Value = "Value Omitted: Excel Error"
        End If
        
        ' If the previous row contains the same Resident name and EPA, then append this rows comment values
        If Cells(i, 1).Value = Cells(i - 1, 1).Value And Cells(i, 2).Value = Cells(i - 1, 2).Value Then
            
            currentColumn = 3
            
            Cells(i - 1, 3).Value = Cells(i - 1, 3).Value & vbNewLine & Cells(i, 3).Value

            currentColumn = 4
            
            Cells(i - 1, 4).Value = Cells(i - 1, 4).Value & vbNewLine & Cells(i, 4).Value

            Rows(i).EntireRow.Delete
        End If
    Next
    On Error GoTo 0
    
    ' Insert a column before 2-3 Strengths with Resident Name & EPA Code and Name
    newSheet.Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Set lookupCol = Selection.EntireColumn
    Set lookupHeader = lookupCol.Cells(1)
    lookupHeader.FormulaR1C1 = "Resident & EPA Code and Name"
    ActiveCell.Offset(1, 0).Select
    
    lastRow = ActiveSheet.UsedRange.Row - 1 + ActiveSheet.UsedRange.Rows.Count
    
    Range(ActiveCell, Cells(lastRow, ActiveCell.Column)).FormulaR1C1 = "=R[0]C[-2]&R[0]C[-1]"
    
    Exit Sub
    

End Sub

Sub CBD_Report_Add_Comments_to_Pivot_Copy()
    
    Dim CommentsLookupSheet As Worksheet
    Dim ResAnSheet As Worksheet
    Dim ResidentAnalysisSheet As Worksheet
    Dim currCLSCell As Range ' Current CommentsLookupSheet Cell
    Dim currRASCell As Range ' Current ResAnSheet Cell
    Dim namesCol As Range
    Dim tmp As String
    Dim cell As Range
    Dim nameCounter As Long
    Dim names() As String
    Dim lastCol As Long
    
    Set CommentsLookupSheet = DataWorkbook.Worksheets("CommentsLookup")
    Set ResAnSheet = DataWorkbook.Worksheets("ResAn")
    
    ' Get a unique list of the Residents names and store in an array
    CommentsLookupSheet.Activate
    Set namesCol = ActiveSheet.Range("A2", Range("A2").End(xlDown))
    
    If Not namesCol Is Nothing Then
        For Each cell In namesCol
            If (cell.Value <> "") And (InStr(tmp, cell.Value) = 0) Then
                tmp = tmp & cell & "|"
            End If
        Next cell
    End If
    
    If Len(tmp) > 0 Then tmp = Left(tmp, Len(tmp) - 1)
    
    names = Split(tmp, "|")

    lastCol = ResAnSheet.Cells(2, Columns.Count).End(xlToLeft).Column
    
    Set currRASCell = ResAnSheet.Cells(2, lastCol + 1)
    currRASCell.Value = "Strengths"
    currRASCell.Offset(0, 1) = "Weaknesses"
    
    
    ' While I'm here, might as well change the font and resize the row
    Set currRASCell = currRASCell.Offset(1, 0)
    With currRASCell.EntireRow
        .Font.ColorIndex = xlAutomatic
        .Font.TintAndShade = 0
        .Font.Bold = True
        .RowHeight = 24
    End With
    Set currRASCell = currRASCell.Offset(1, 0)
    
    Set currCLSCell = CommentsLookupSheet.Range("A2")
    
    For nameCounter = LBound(names) To UBound(names)
        Do While currCLSCell.Value = names(nameCounter)
            currRASCell.Value = currCLSCell.Offset(0, 3).Value ' Copy the strengths
            currRASCell.Offset(0, 1).Value = currCLSCell.Offset(0, 4).Value ' Copy the weaknesses
            Set currRASCell = currRASCell.Offset(1, 0)
            Set currCLSCell = currCLSCell.Offset(1, 0)
        Loop
        With currRASCell.EntireRow
            .Font.ColorIndex = xlAutomatic
            .Font.TintAndShade = 0
            .Font.Bold = True
            .RowHeight = 24
        End With
        Set currRASCell = currRASCell.Offset(1, 0) ' Need to skip a line to ralign on table
    Next nameCounter
    
    
    ' Copy the style format to the new columns
    ResAnSheet.Activate
    currRASCell.EntireColumn.Offset(0, -1).Select
    Selection.Copy
    currRASCell.EntireColumn.Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    currRASCell.EntireColumn.Select
    Selection.Copy
    currRASCell.Offset(0, 1).EntireColumn.Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    ' Resize and reformat the new columns
    currRASCell.EntireColumn.Select
    With Selection
        .ColumnWidth = 80
        .HorizontalAlignment = xlGeneral
        .Cells(2, 0).HorizontalAlignment = xlCenter ' The header
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .VerticalAlignment = xlCenter
    End With
    
    currRASCell.EntireColumn.Offset(0, 1).Select
    With Selection
        .ColumnWidth = 80
        .HorizontalAlignment = xlGeneral
        .Cells(2, 1).HorizontalAlignment = xlCenter ' The header
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .VerticalAlignment = xlCenter
    End With
    
    ' Reformat the other columns
    Range("A1", currRASCell.Offset(0, -1)).EntireColumn.Select
    With Selection
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B1", currRASCell.Offset(0, -1)).EntireColumn.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' Copy by values to new sheet
    Range("A4").CurrentRegion.Select
    Selection.Copy
    
    Set ResidentAnalysisSheet = DataWorkbook.Worksheets.Add
    ResidentAnalysisSheet.Name = "ResidentAnalysis"
    DataWorkbook.Worksheets("ResidentAnalysis").Select
    
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells.Select
    Cells.EntireColumn.AutoFit
    Selection.Rows.AutoFit
    
    Range("B2").End(xlToRight).EntireColumn.ColumnWidth = 80
    Range("B2").End(xlToRight).EntireColumn.Offset(0, -1).ColumnWidth = 80
    Range("B2").End(xlToRight).EntireColumn.Offset(0, -2).AutoFit
    Columns("A:A").ColumnWidth = 80
    
End Sub

Sub CBD_Report_Create_Site_Pivot_Table_And_Copy()

    Dim PivotTableSheet As Worksheet
    Dim SiteSheet As Worksheet
    Dim cell As Range
    Dim i As Long
    
    DataWorkbook.Activate
    Sheets.Add.Name = "PivotTableSite"
    Set PivotTableSheet = Worksheets("PivotTableSite")
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ExtractTable", Version:=6).CreatePivotTable TableDestination:= _
        "PivotTableSite!R3C9", TableName:="SitePivotTable", DefaultVersion:=6
        
    Sheets("PivotTableSite").Select
    Cells(3, 9).Select
    
    With ActiveSheet.PivotTables("SitePivotTable").PivotFields("Site")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("SitePivotTable").PivotFields("Block")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("SitePivotTable").PivotFields("Entrustment / Overall Category")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("SitePivotTable").PivotFields("Entrustment / Overall Category")
        .Orientation = xlDataField
    End With
    
    With ActiveSheet.PivotTables("SitePivotTable")
        .CompactLayoutRowHeader = ""
        .CompactLayoutColumnHeader = ""
        .TableStyle2 = "PivotStyleMedium2"
    End With
    
    Range("I3").Select
    Selection.CurrentRegion.Select
    Selection.Copy
    
    Set SiteSheet = Sheets.Add(After:=Worksheets("ResidentAnalysis"))
    SiteSheet.Name = "SiteAnalysis"
    
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Cells.EntireColumn.AutoFit
    
    For i = 0 To Range("A3", Range("A3").End(xlDown)).Count
        Set cell = Range("A3").Offset(i, 0)
        If (Left(cell.Value, 6) <> "Block ") Then
            With cell.EntireRow
                .Font.ColorIndex = xlAutomatic
                .Font.Bold = True
            End With
        End If
    Next i
    
    Range("A1").Select
End Sub


Sub CBD_Report_Create_Block_Pivot_Table_And_Copy()

    Dim PivotTableSheet As Worksheet
    Dim BlockSheet As Worksheet
    Dim cell As Range
    Dim i As Long
    
    DataWorkbook.Activate
    Sheets.Add.Name = "PivotTableBlock"
    Set PivotTableSheet = Worksheets("PivotTableBlock")
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ExtractTable", Version:=6).CreatePivotTable TableDestination:= _
        "PivotTableBlock!R3C18", TableName:="BlockPivotTable", DefaultVersion:=6
        
    Sheets("PivotTableBlock").Select
    Cells(3, 18).Select
    
    With ActiveSheet.PivotTables("BlockPivotTable").PivotFields("Resident")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("BlockPivotTable").PivotFields("Block")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("BlockPivotTable").PivotFields("Entrustment / Overall Category")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("BlockPivotTable").PivotFields("Entrustment / Overall Category")
        .Orientation = xlDataField
    End With
    
    With ActiveSheet.PivotTables("BlockPivotTable")
        .CompactLayoutRowHeader = ""
        .CompactLayoutColumnHeader = ""
        .TableStyle2 = "PivotStyleMedium2"
    End With
    
    Range("R3").Select
    Selection.CurrentRegion.Select
    Selection.Copy
    
    Set BlockSheet = Sheets.Add(After:=Worksheets("SiteAnalysis"))
    BlockSheet.Name = "BlockAnalysis"
    
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Cells.EntireColumn.AutoFit
    
    For i = 0 To Range("A3", Range("A3").End(xlDown)).Count
        Set cell = Range("A3").Offset(i, 0)
        If (Left(cell.Value, 6) <> "Block ") Then
            With cell.EntireRow
                .Font.ColorIndex = xlAutomatic
                .Font.Bold = True
            End With
        End If
    Next i
    
    Range("A1").Select
End Sub

Sub CBD_Report_Cleanup()
    Application.DisplayAlerts = False
    DataWorkbook.Worksheets("PivotTable").Delete
    DataWorkbook.Worksheets("PivotTableSite").Delete
    DataWorkbook.Worksheets("PivotTableBlock").Delete
    DataWorkbook.Worksheets("CommentsLookup").Delete
    DataWorkbook.Worksheets("ResAn").Delete
    DataWorkbook.Worksheets("DataExtract").Delete
    LookupTable.Close
    Application.DisplayAlerts = True
    If Len(CommentsErrorLog) > 0 Then
        MsgBox Prompt:="There were Excel Errors (i.e. #NAME) that could not be removed.  They have been ommitted and " _
            & "replaced with 'Value Omitted: Excel Error'. The following Residents, EPAs, and Comment columns are responsible:" _
            & CommentsErrorLog
    End If
End Sub
