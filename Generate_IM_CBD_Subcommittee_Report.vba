Option Explicit

    Dim DataWorkbook As Workbook
    Dim MacroLaunchWorkbook As Workbook
    Dim DataSheet As Worksheet
    Dim LookupTable As Workbook
    Dim ExtractTable As ListObject
    Dim proceed As Boolean
    
Sub procede_Button_Clicked()
    proceed = True
End Sub
    

Sub Generate_IM_CBD_Subcommittee_Report()
    Application.ScreenUpdating = False
    Set MacroLaunchWorkbook = ThisWorkbook
    Call CBD_Report_Format_Extract
    Call CBD_Report_Create_Pivot_Table
    Application.ScreenUpdating = True
    DataWorkbook.Worksheets("COD5 by Cohort").Activate
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
    
    MsgBox "Choose Extract Data for the report."
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
        End
    End If
    
    ' Creates a copy of the extract data to preserve the original. Stored in a variable for reference when
    ' opening the new workbook
    ExtractCopy = Left(fd.SelectedItems(1), InStrRev(fd.SelectedItems(1), ".") - 1) & " copy." & _
        Right(fd.SelectedItems(1), Len(fd.SelectedItems(1)) - InStrRev(fd.SelectedItems(1), "."))
    
    fso.CopyFile fd.SelectedItems(1), ExtractCopy
    
    Set DataWorkbook = Workbooks.Open(Filename:=ExtractCopy)
    
    ' Rename the workbook copy with the appropriate division name and date range
    DataWorkbook.SaveAs DataWorkbook.Path & "\IM_CBD_Subcommittee_Report" & _
        "_" & DataWorkbook.ActiveSheet.Range("B2").Value & ".xlsx", 51
    Set DataSheet = DataWorkbook.Worksheets(1)
    DataSheet.name = "DataExtract"
    
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
    
    DataWorkbook.Activate
    
    ' Delete the first 3 rows
    Rows("1:3").Select
    Selection.Delete Shift:=xlUp
    
    ' Resize all columns
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    ' Create a table from the data for ease of manipulation - allows use of column
    ' header names for flexibility
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1").CurrentRegion, , xlYes).name = _
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
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "EPA Code and Name"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP([@[Assessment Form Code]],'[" & LookupTable.name & "]VLOOKUP MASTER'!C1:C11,10,FALSE), ""NONEXISTANT_FORM_ID"")"
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
        "=VLOOKUP([@[CV ID 9533 : Site]],'[" & LookupTable.name & "]Site'!C1:C2,2,FALSE)"
    ExtractTable.ListColumns("Site").Range.EntireColumn.AutoFit
    
    ' Create column Block adjacent to Type of Assessment Form
    ExtractTable.ListColumns("Type of Assessment Form").Range.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "Block"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP([@[Date of encounter]],'[" & LookupTable.name & "]BLOCK'!C2:C6,3,TRUE)"
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
    On Error GoTo NoBlanks
    Set blankCells = ExtractTable.ListColumns("Date of Assessment Form Submission").Range.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0 ' Turn off the Error handler
    If Not blankCells Is Nothing Then
        blankCells.Delete xlUp
    End If
    DataWorkbook.Activate
    
    ' Fill in Assessee Training Level
    ExtractTable.ListColumns("Assessee Training Level").Range.Cells(1).Select
    ActiveCell.Offset(1, 0).Select
    Range(ActiveCell, Cells(Range("A3").End(xlDown).Row, ExtractTable.ListColumns("Assessee Training Level").Range.Column)).FormulaR1C1 = _
        "=IFERROR(VLOOKUP([@[Assessee Email]],'[" & LookupTable.name & "]IM Trainees'!C4:C5,2,FALSE), ""TRAINEE_LEVEL_NOT_FOUND"")"
    ExtractTable.ListColumns("Assessee Training Level").Range.EntireColumn.AutoFit
    
    'Loop through and make sure all EPA Codes are accounted for
        ' Filter
    ExtractTable.ListColumns("Assessee Training Level").Range.AutoFilter Field:=ExtractTable.ListColumns("Assessee Training Level").Range.Column, Criteria1 _
        :="TRAINEE_LEVEL_NOT_FOUND"
    If ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 > 1 Then
        
        Application.ScreenUpdating = True
        ' Get a unique list of Resident names to filter by in the loop
        Dim tmp As String
        Dim arr() As String
        Dim names As Range
        Dim cell As Range
        
        Set names = ExtractTable.ListColumns("Resident").Range.Offset(1, 0).SpecialCells(xlCellTypeVisible).Cells
        If Not names Is Nothing Then
           For Each cell In names
              If (cell <> "") And (InStr(tmp, cell) = 0) Then
                tmp = tmp & cell & "|"
              End If
           Next cell
        End If
        
        If Len(tmp) > 0 Then tmp = Left(tmp, Len(tmp) - 1)
        
        arr = Split(tmp, "|")
        
        ' MAKE A STRING OF NAMES TO PUT IN THE MESSAGE BOX
        Dim nameStr As String
        Dim name As Variant
        For Each name In arr
            nameStr = nameStr & name & "," & vbNewLine
        Next name
        nameStr = Left(nameStr, Len(nameStr) - 2)
        
        MsgBox Prompt:="Training Levels could not be found for the Following Residents: " & nameStr _
            & vbNewLine & vbNewLine & "You will be prompted to find the Level of each Resident.  Click the PROCEED button after you have found each level."
        
        Dim trainingLevel As String
        For Each name In arr
            ' Filter the table to only show the resident
            ExtractTable.ListColumns("Resident").Range.AutoFilter Field:=ExtractTable.ListColumns("Resident").Range.Column, Criteria1 _
                :=name
                
            MsgBox Prompt:="Find the Training Level for: " & name & vbNewLine & vbNewLine & "Click the PROCEED button once you have found the level."
            ' Ask the user what the Training Level is for this Resident
            trainingLevel = ""
            proceed = False
            'MacroLaunchWorkbook.Activate
            While Not proceed
                DoEvents
            Wend
            proceed = False
            trainingLevel = InputBox(Prompt:="What is the training level for " & name & "?  Please input the level as ""PGY#"".", _
                Default:="PGY")
            
            With ExtractTable.ListColumns("Assessee Training Level")
                DataWorkbook.Worksheets("DataExtract").Range(DataWorkbook.Worksheets("DataExtract").Cells(2, .Index), _
                    DataWorkbook.Worksheets("DataExtract").Cells(.Range.End(xlDown).Row, .Index)).Cells.SpecialCells(xlCellTypeVisible).Value = trainingLevel
            End With
        Next name
        
        
        'Dim errorCellsCount As Long
        'errorCellsCount = ExtractTable.ListColumns("Assessee Training Level").Range.Cells(1, 0).Count
        '
        'Do While errorCellsCount > 1
        '
        'Loop
    End If
    DataWorkbook.Worksheets("DataExtract").ListObjects("ExtractTable").AutoFilter.ShowAllData ' Unfilter
    Application.ScreenUpdating = False
    
    Exit Sub
NoBlanks:
    Resume Next
    
End Sub

Sub CBD_Report_Create_Pivot_Table()


    Dim COD5byCohortSheet As Worksheet
    Dim COD5byRotationSheet As Worksheet
    Dim EPATypebyBlockSheet As Worksheet
    Dim EPATypebySiteSheet As Worksheet
    Dim CompletionBySiteAndBlockSheet As Worksheet
    Dim CompletionByRotationAndBlockSheet As Worksheet
    Dim CompletionByBlockAndCohortSheet As Worksheet
    Dim CompletionByRotationAndEPASheet As Worksheet
    
    Dim ResAnSheet As Worksheet
    Dim pivotTableLastRow As Range
    Dim firstCell As Range
    Dim lastCell As Range
    
    ' Create the COD5 by Cohort Pivot Table
    
    DataWorkbook.Activate
    Sheets.Add.name = "COD5 by Cohort"
    Set COD5byCohortSheet = Worksheets("COD5 by Cohort")
    COD5byCohortSheet.Move Before:=DataSheet
    DataWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ExtractTable", Version:=6).CreatePivotTable TableDestination:= _
        "'COD5 by Cohort'!R3C1", TableName:="COD5byCohortTable", DefaultVersion:=6
    Sheets("COD5 by Cohort").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("COD5byCohortTable").PivotFields("EPA Code and Name")
        .Orientation = xlRowField
        .Position = 1
        .ClearAllFilters
        Dim i As Long
        For i = 1 To .PivotItems.Count
            If .PivotItems(i).name <> "COD-05 Performing the procedures of Internal Medicine" Then
                .PivotItems(.PivotItems(i).name).Visible = False
            End If
        Next i
    End With
    With ActiveSheet.PivotTables("COD5byCohortTable").PivotFields("CV ID 9539 : Type of Case/Procedure")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("COD5byCohortTable").PivotFields("Assessee Training Level")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("COD5byCohortTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlColumnField
        .Position = 1
        For i = 1 To .PivotItems.Count
            If .PivotItems(i).name = "(blank)" Then
                .PivotItems(.PivotItems(i).name).Visible = False
            End If
        Next i
    End With
    With ActiveSheet.PivotTables("COD5byCohortTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlDataField
    End With
    ' Add the title to the Pivot Table
    ActiveSheet.PivotTables("COD5byCohortTable").DataPivotField.PivotItems( _
        "Count of Entrustment / Overall Category").Caption = _
        "COD-05 Performing the procedures of Internal Medicine"
    ' Remove the default labels
    With ActiveSheet.PivotTables("COD5byCohortTable")
        .CompactLayoutRowHeader = ""
        .CompactLayoutColumnHeader = ""
        .TableStyle2 = "PivotStyleMedium2"
    End With
    ' Add bold font and vertical borders for legibility
    With Range("A3").CurrentRegion
        .Font.Bold = True
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    ActiveSheet.PivotTables("COD5byCohortTable").PivotSelect _
        "'Assessee Training Level'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.Bold = False
    
    ' Create the COD5 by Rotation Pivot Table
    
    DataWorkbook.Activate
    Sheets.Add.name = "COD5 by Rotation"
    Set COD5byRotationSheet = Worksheets("COD5 by Rotation")
    COD5byRotationSheet.Move Before:=DataSheet
    DataWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ExtractTable", Version:=6).CreatePivotTable TableDestination:= _
        "'COD5 by Rotation'!R3C1", TableName:="COD5byRotationTable", DefaultVersion:=6
    Sheets("COD5 by Rotation").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("COD5byRotationTable").PivotFields("EPA Code and Name")
        .Orientation = xlRowField
        .Position = 1
        .ClearAllFilters
        For i = 1 To .PivotItems.Count
            If .PivotItems(i).name <> "COD-05 Performing the procedures of Internal Medicine" Then
                .PivotItems(.PivotItems(i).name).Visible = False
            End If
        Next i
    End With
    With ActiveSheet.PivotTables("COD5byRotationTable").PivotFields("CV ID 9530 : Rotation Service")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("COD5byRotationTable").PivotFields("CV ID 9539 : Type of Case/Procedure")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("COD5byRotationTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlColumnField
        .Position = 1
        For i = 1 To .PivotItems.Count
            If .PivotItems(i).name = "(blank)" Then
                .PivotItems(.PivotItems(i).name).Visible = False
            End If
        Next i
    End With
    With ActiveSheet.PivotTables("COD5byRotationTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlDataField
    End With
    ' Add the title to the Pivot Table
    ActiveSheet.PivotTables("COD5byRotationTable").DataPivotField.PivotItems( _
        "Count of Entrustment / Overall Category").Caption = _
        "COD-05 Performing the procedures of Internal Medicine"
    ' Remove the default labels
    With ActiveSheet.PivotTables("COD5byRotationTable")
        .CompactLayoutRowHeader = ""
        .CompactLayoutColumnHeader = ""
        .TableStyle2 = "PivotStyleMedium2"
    End With
    ' Add bold font and vertical borders for legibility
    With Range("A3").CurrentRegion
        .Font.Bold = True
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    ActiveSheet.PivotTables("COD5byRotationTable").PivotSelect _
        "'CV ID 9539 : Type of Case/Procedure'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.Bold = False
    Columns("A:A").Select
    Selection.ColumnWidth = 85
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
    
    ' Create the EPA Type by Block Pivot Table
    
    DataWorkbook.Activate
    Sheets.Add.name = "EPA Type by Block"
    Set EPATypebyBlockSheet = Worksheets("EPA Type by Block")
    EPATypebyBlockSheet.Move Before:=DataSheet
    DataWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ExtractTable", Version:=6).CreatePivotTable TableDestination:= _
        "'EPA Type by Block'!R3C1", TableName:="EPATypebyBlockTable", DefaultVersion:=6
    Sheets("EPA Type by Block").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("EPATypebyBlockTable").PivotFields("EPA Code and Name")
        .Orientation = xlRowField
        .Position = 1
        .ClearAllFilters
    End With
    With ActiveSheet.PivotTables("EPATypebyBlockTable").PivotFields("Block")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("EPATypebyBlockTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlColumnField
        .Position = 1
        For i = 1 To .PivotItems.Count
            If .PivotItems(i).name = "(blank)" Then
                .PivotItems(.PivotItems(i).name).Visible = False
            End If
        Next i
    End With
    With ActiveSheet.PivotTables("EPATypebyBlockTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlDataField
    End With
    ' Add the title to the Pivot Table
    ActiveSheet.PivotTables("EPATypebyBlockTable").DataPivotField.PivotItems( _
        "Count of Entrustment / Overall Category").Caption = _
        "EPA Type by Block"
    ' Remove the default labels
    With ActiveSheet.PivotTables("EPATypebyBlockTable")
        .CompactLayoutRowHeader = ""
        .CompactLayoutColumnHeader = ""
        .TableStyle2 = "PivotStyleMedium2"
    End With
    ' Add bold font and vertical borders for legibility
    With Range("A3").CurrentRegion
        .Font.Bold = True
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    ActiveSheet.PivotTables("EPATypebyBlockTable").PivotSelect _
        "'Block'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.Bold = False
    ActiveSheet.PivotTables("EPATypebyBlockTable").PivotSelect _
        "'EPA Code and Name'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.ColorIndex = xlAutomatic
    Selection.Font.TintAndShade = 0
    Columns("A:A").Select
    Selection.ColumnWidth = 85
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
    
    
    ' Create the EPA Type by Site Pivot Table
    
    DataWorkbook.Activate
    Sheets.Add.name = "EPA Type by Site"
    Set EPATypebySiteSheet = Worksheets("EPA Type by Site")
    EPATypebySiteSheet.Move Before:=DataSheet
    DataWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ExtractTable", Version:=6).CreatePivotTable TableDestination:= _
        "'EPA Type by Site'!R3C1", TableName:="EPATypebySiteTable", DefaultVersion:=6
    Sheets("EPA Type by Site").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("EPATypebySiteTable").PivotFields("Site")
        .Orientation = xlRowField
        .Position = 1
        .ClearAllFilters
    End With
    With ActiveSheet.PivotTables("EPATypebySiteTable").PivotFields("EPA Code and Name")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("EPATypebySiteTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlColumnField
        .Position = 1
        For i = 1 To .PivotItems.Count
            If .PivotItems(i).name = "(blank)" Then
                .PivotItems(.PivotItems(i).name).Visible = False
            End If
        Next i
    End With
    With ActiveSheet.PivotTables("EPATypebySiteTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlDataField
    End With
    ' Add the title to the Pivot Table
    ActiveSheet.PivotTables("EPATypebySiteTable").DataPivotField.PivotItems( _
        "Count of Entrustment / Overall Category").Caption = _
        "EPA Type by Site"
    ' Remove the default labels
    With ActiveSheet.PivotTables("EPATypebySiteTable")
        .CompactLayoutRowHeader = ""
        .CompactLayoutColumnHeader = ""
        .TableStyle2 = "PivotStyleMedium2"
    End With
    ' Add bold font and vertical borders for legibility
    With Range("A3").CurrentRegion
        .Font.Bold = True
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    ActiveSheet.PivotTables("EPATypebySiteTable").PivotSelect _
        "'EPA Code and Name'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.Bold = False
    ActiveSheet.PivotTables("EPATypebySiteTable").PivotSelect _
        "'Site'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.ColorIndex = xlAutomatic
    Selection.Font.TintAndShade = 0
    Columns("A:A").Select
    Selection.ColumnWidth = 85
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
    
    ' Create the Completion by Site and Block Pivot Table
    
    DataWorkbook.Activate
    Sheets.Add.name = "Completion by Site & Block"
    Set CompletionBySiteAndBlockSheet = Worksheets("Completion by Site & Block")
    CompletionBySiteAndBlockSheet.Move Before:=DataSheet
    DataWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ExtractTable", Version:=6).CreatePivotTable TableDestination:= _
        "'Completion by Site & Block'!R3C1", TableName:="CompletionbySiteAndBlockTable", DefaultVersion:=6
    Sheets("Completion by Site & Block").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("CompletionbySiteAndBlockTable").PivotFields("Site")
        .Orientation = xlRowField
        .Position = 1
        .ClearAllFilters
    End With
    With ActiveSheet.PivotTables("CompletionbySiteAndBlockTable").PivotFields("Block")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("CompletionbySiteAndBlockTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlColumnField
        .Position = 1
        For i = 1 To .PivotItems.Count
            If .PivotItems(i).name = "(blank)" Then
                .PivotItems(.PivotItems(i).name).Visible = False
            End If
        Next i
    End With
    With ActiveSheet.PivotTables("CompletionbySiteAndBlockTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlDataField
    End With
    ' Add the title to the Pivot Table
    ActiveSheet.PivotTables("CompletionbySiteAndBlockTable").DataPivotField.PivotItems( _
        "Count of Entrustment / Overall Category").Caption = _
        "Completion by Site and Block"
    ' Remove the default labels
    With ActiveSheet.PivotTables("CompletionbySiteAndBlockTable")
        .CompactLayoutRowHeader = ""
        .CompactLayoutColumnHeader = ""
        .TableStyle2 = "PivotStyleMedium2"
    End With
    ' Add bold font and vertical borders for legibility
    With Range("A3").CurrentRegion
        .Font.Bold = True
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    ActiveSheet.PivotTables("CompletionbySiteAndBlockTable").PivotSelect _
        "'Block'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.Bold = False
    ActiveSheet.PivotTables("CompletionbySiteAndBlockTable").PivotSelect _
        "'Site'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.ColorIndex = xlAutomatic
    Selection.Font.TintAndShade = 0
    
    
    ' Create the Completion by Rotation and Block Pivot Table
    
    DataWorkbook.Activate
    Sheets.Add.name = "Completion by Rotation & Block"
    Set CompletionByRotationAndBlockSheet = Worksheets("Completion by Rotation & Block")
    CompletionByRotationAndBlockSheet.Move Before:=DataSheet
    DataWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ExtractTable", Version:=6).CreatePivotTable TableDestination:= _
        "'Completion by Rotation & Block'!R3C1", TableName:="CompletionbyRotationAndBlockTable", DefaultVersion:=6
    Sheets("Completion by Rotation & Block").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("CompletionbyRotationAndBlockTable").PivotFields("CV ID 9530 : Rotation Service")
        .Orientation = xlRowField
        .Position = 1
        .ClearAllFilters
    End With
    With ActiveSheet.PivotTables("CompletionbyRotationAndBlockTable").PivotFields("Block")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("CompletionbyRotationAndBlockTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlColumnField
        .Position = 1
        For i = 1 To .PivotItems.Count
            If .PivotItems(i).name = "(blank)" Then
                .PivotItems(.PivotItems(i).name).Visible = False
            End If
        Next i
    End With
    With ActiveSheet.PivotTables("CompletionbyRotationAndBlockTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlDataField
    End With
    ' Add the title to the Pivot Table
    ActiveSheet.PivotTables("CompletionbyRotationAndBlockTable").DataPivotField.PivotItems( _
        "Count of Entrustment / Overall Category").Caption = _
        "Completion by Rotation and Block"
    ' Remove the default labels
    With ActiveSheet.PivotTables("CompletionbyRotationAndBlockTable")
        .CompactLayoutRowHeader = ""
        .CompactLayoutColumnHeader = ""
        .TableStyle2 = "PivotStyleMedium2"
    End With
    ' Add bold font and vertical borders for legibility
    With Range("A3").CurrentRegion
        .Font.Bold = True
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    ActiveSheet.PivotTables("CompletionbyRotationAndBlockTable").PivotSelect _
        "'Block'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.Bold = False
    ActiveSheet.PivotTables("CompletionbyRotationAndBlockTable").PivotSelect _
        "'CV ID 9530 : Rotation Service'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.ColorIndex = xlAutomatic
    Selection.Font.TintAndShade = 0
    
    ' Create the Completion by Block and Cohort Pivot Table
    
    DataWorkbook.Activate
    Sheets.Add.name = "Completion by Block & Cohort"
    Set CompletionByBlockAndCohortSheet = Worksheets("Completion by Block & Cohort")
    CompletionByBlockAndCohortSheet.Move Before:=DataSheet
    DataWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ExtractTable", Version:=6).CreatePivotTable TableDestination:= _
        "'Completion by Block & Cohort'!R3C1", TableName:="CompletionbyBlockAndCohortTable", DefaultVersion:=6
    Sheets("Completion by Block & Cohort").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("CompletionbyBlockAndCohortTable").PivotFields("Block")
        .Orientation = xlRowField
        .Position = 1
        .ClearAllFilters
    End With
    With ActiveSheet.PivotTables("CompletionbyBlockAndCohortTable").PivotFields("Assessee Training Level")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("CompletionbyBlockAndCohortTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlColumnField
        .Position = 1
        For i = 1 To .PivotItems.Count
            If .PivotItems(i).name = "(blank)" Then
                .PivotItems(.PivotItems(i).name).Visible = False
            End If
        Next i
    End With
    With ActiveSheet.PivotTables("CompletionbyBlockAndCohortTable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlDataField
    End With
    ' Add the title to the Pivot Table
    ActiveSheet.PivotTables("CompletionbyBlockAndCohortTable").DataPivotField.PivotItems( _
        "Count of Entrustment / Overall Category").Caption = _
        "Completion by Block and Cohort"
    ' Remove the default labels
    With ActiveSheet.PivotTables("CompletionbyBlockAndCohortTable")
        .CompactLayoutRowHeader = ""
        .CompactLayoutColumnHeader = ""
        .TableStyle2 = "PivotStyleMedium2"
    End With
    ' Add bold font and vertical borders for legibility
    With Range("A3").CurrentRegion
        .Font.Bold = True
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    ActiveSheet.PivotTables("CompletionbyBlockAndCohortTable").PivotSelect _
        "'Assessee Training Level'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.Bold = False
    ActiveSheet.PivotTables("CompletionbyBlockAndCohortTable").PivotSelect _
        "'Block'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.ColorIndex = xlAutomatic
    Selection.Font.TintAndShade = 0
    
    ' Create the Completion by Rotation and EPA Pivot Table
    
    DataWorkbook.Activate
    Sheets.Add.name = "Completion by Rotation & EPA"
    Set CompletionByRotationAndEPASheet = Worksheets("Completion by Rotation & EPA")
    CompletionByRotationAndEPASheet.Move Before:=DataSheet
    DataWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ExtractTable", Version:=6).CreatePivotTable TableDestination:= _
        "'Completion by Rotation & EPA'!R3C1", TableName:="CompletionbyRotationAndEPATable", DefaultVersion:=6
    Sheets("Completion by Rotation & EPA").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("CompletionbyRotationAndEPATable").PivotFields("CV ID 9530 : Rotation Service")
        .Orientation = xlRowField
        .Position = 1
        .ClearAllFilters
    End With
    With ActiveSheet.PivotTables("CompletionbyRotationAndEPATable").PivotFields("EPA Code and Name")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("CompletionbyRotationAndEPATable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlColumnField
        .Position = 1
        For i = 1 To .PivotItems.Count
            If .PivotItems(i).name = "(blank)" Then
                .PivotItems(.PivotItems(i).name).Visible = False
            End If
        Next i
    End With
    With ActiveSheet.PivotTables("CompletionbyRotationAndEPATable").PivotFields( _
        "Entrustment / Overall Category")
        .Orientation = xlDataField
    End With
    ' Add the title to the Pivot Table
    ActiveSheet.PivotTables("CompletionbyRotationAndEPATable").DataPivotField.PivotItems( _
        "Count of Entrustment / Overall Category").Caption = _
        "Completion by Rotation and EPA"
    ' Remove the default labels
    With ActiveSheet.PivotTables("CompletionbyRotationAndEPATable")
        .CompactLayoutRowHeader = ""
        .CompactLayoutColumnHeader = ""
        .TableStyle2 = "PivotStyleMedium2"
    End With
    ' Add bold font and vertical borders for legibility
    With Range("A3").CurrentRegion
        .Font.Bold = True
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    ActiveSheet.PivotTables("CompletionbyRotationAndEPATable").PivotSelect _
        "'EPA Code and Name'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.Bold = False
    ActiveSheet.PivotTables("CompletionbyRotationAndEPATable").PivotSelect _
        "'CV ID 9530 : Rotation Service'[All]", xlLabelOnly + xlFirstRow, True
    Selection.Font.ColorIndex = xlAutomatic
    Selection.Font.TintAndShade = 0
    Columns("A:A").Select
    Selection.ColumnWidth = 85
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
End Sub
