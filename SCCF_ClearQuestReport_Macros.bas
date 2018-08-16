Attribute VB_Name = "SCCF_CQ_Report_Macros"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr) 'MS Office 64 Bit
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long) 'MS Office 32 Bit
#End If

Sub DoFullFormat()
'
' Macro to format an SCCF ClearQuest file export

    ' Create file system object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create an excel application object
    Set xlApp = CreateObject("Excel.Application")
      
    ' Get the current filename and path
    currentPath = fso.GetParentFolderName(ThisWorkbook.FullName)
    currentFile = fso.GetBaseName(ThisWorkbook.FullName)
 
    ' Get the current filename extension
    extension = fso.GetExtensionName(ThisWorkbook.FullName)
    
    ' Cannot save and format in one step, so just save as new extension first
    ' If file is already saved as XLSX, then format the file
    If extension = "xls" Then
    
        ' Define the name of the new file to open
        newExtension = ".xlsx"
        newFileName = fso.BuildPath(currentPath, currentFile & newExtension)
           
        ' Save the current workbook as an XLSX file
        SaveAsXLSX
        
        ' Quit the application to force re-opening of XLSX *not* in compatibility mode
        Application.Quit
        
    Else
        ' Only format XLSX file types
        If extension = "xlsx" Then
            
            ' Format the data
            Format_Data
            
            ' Save the workbook
            Save_Workbook
            
        Else
            ' This file is not an xls or xlsx file type, so declare it invalid and let someone know
            MsgBox "Invalid file type", , "Unable to Process File"
            
        End If
        
    End If

End Sub

Sub Format_Data()
'
' Formats a ClearQuest export to the report data pivot charts defined
'
    'Copy the raw export data to a new sheet
    Copy_PivotWorksheet ("Data")

    ' Select the origin for the active sheet
    Range("A1").Select
    
    ' Set the new sheet as the active sheet
    SetActiveSheetName ("Data")

    ' Turn off screen updates
    Application.ScreenUpdating = False
    Application.Calculation = xlAutomatic

    ' Create Table1 as the table containing the export data
    Create_Table

    ' Remove any duplicate rows from the PTR Column
    RemoveDuplicates ("PTR_Number")

    ' Convert the PTR numbers stored as text to numbers
    ConvertPTRvaluesToNumber

    ' Create Table2 as the table containing the lookup data for Escape Definitions
    Add_EscapeDefinitions
    
    ' Activate Sheet with Table1 to complete remaining functions
    Table1_SheetName = Range("Table1").Parent.Name
    SetActiveSheetName (Table1_SheetName)

    ' Add the columns to Table1
    Add_Report_Date
    Add_CreationDateCalc
    Add_DueDateCalc
    Add_HoldDateCalc
    Add_SeverityCalc
    Add_Escape
    Add_EscapeType

    '
    ' Add the Worksheets for each metric chart
    '

    ' Add AllCreated Worksheet - this is the base worksheet for all subsequent worksheets
    Add_AllCreated_Worksheet

    ' Create the AllCreatedBySeverity Worksheet from the AllCreated Worksheet
    Create_AllCreatedBySeverity

    ' Create the OpenBySeverity Worksheet from the AllCreatedBySeverity Worksheet
    Create_OpenBySeverity

    ' Create the OpenBySeverityCI Worksheet from the OpenBySeverity Worksheet
    Create_OpenBySeverityCI

    ' Create the OpenByState Worksheet from the OpenBySeverityCI Worksheet
    Create_OpenByState

    ' Create the OpenByState_Current Worksheet from the OpenByState Worksheet
    Create_OpenByState_Current

    ' Create the HoldDate Worksheet from the OpenByState_Current Worksheet
    Create_HoldDate

    ' Create the DueDate Worksheet from the HoldDate Worksheet
    Create_PastDue

    ' Create the PastDueByState Worksheet from the PastDue Worksheet
    Create_PastDueByState

    ' Create the CorePtrs Worksheet
    Create_CorePtrs

    ' Create the AllEscapes Worksheet from the OpenByCI Worksheet
    Create_AllEscapes

    ' Create the EscapesByType Worksheet from the AllEscapes Worksheet
    Create_EscapesByType

    ' Create the ExpiredHoldDate Worksheet from the HoldDate Worksheet
    Create_ExpiredHoldDate

    ' Create the ExpiredDueDate Worksheet from the DueDate Worksheet
    Create_ExpiredDueDate

    ' Create the Notes Worksheet from the Data Worksheet
    CreateNotesWorksheet

    ' Jesus Saves...so should you
    'Save_Workbook

    ' Turn on screen updates
    Application.ScreenUpdating = True

End Sub

Sub Create_Table()
'
' Creates a table from the current worksheet
'
ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1").CurrentRegion, , xlYes).Name = "Table1"

'Apply style to table
Range("Table1[#All]").Select
ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"

End Sub

Sub Add_NewColumnToTableEnd()
'
' Adds a new column to the end of a table
'
ActiveSheet.ListObjects("Table1").ListColumns.Add


End Sub
Sub Add_NewColumnToTableEndWithName(ColumnName)
'
' Adds a new column to the end of a table with a column name specified
'
    ActiveSheet.ListObjects(1).ListColumns.Add

    ' Get the last column object
    LastColumn = ActiveSheet.ListObjects(1).Range.Columns.Count

    ' Add the Column Name
    ActiveSheet.ListObjects(1).ListColumns(LastColumn).Name = ColumnName


End Sub
Sub Add_Report_Date()
Attribute Add_Report_Date.VB_Description = "Adds the ReportDate column for every record and populates the current date"
Attribute Add_Report_Date.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Add_Report_Date Macro
' Adds the ReportDate column for every record and populates the current date
'

    ' Add a new column to the end of Table1
    Add_NewColumnToTableEnd

    ' Get the last column object
    LastColumn = ActiveSheet.ListObjects("Table1").Range.Columns.Count

    ' Add the Column Name
    ActiveSheet.ListObjects("Table1").ListColumns(LastColumn).Name = "ReportDate"

    ' Add the date to every cell in this column
    LastColumnDataRange = ActiveSheet.ListObjects("Table1").ListColumns(LastColumn).DataBodyRange.Select
    ActiveCell.FormulaR1C1 = Date
    Selection.FillDown

End Sub
Sub Add_CreationDateCalc()
Attribute Add_CreationDateCalc.VB_Description = "Adds the CreationDateCalc column and parses the creation date time stamp to be just the CreationDate in excel format"
Attribute Add_CreationDateCalc.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Add_CreationDateCalc Macro
' Adds the CreationDateCalc column and parses the creation date time stamp to be just the CreationDate in excel format
'

   ' Add a new column to the end of Table1 with the column name
    Add_NewColumnToTableEndWithName ("CreationDateCalc")

    ' Get the last column object
    LastColumn = ActiveSheet.ListObjects("Table1").Range.Columns.Count

    ' Add the data or formula to every cell in this column
    LastColumnDataRange = ActiveSheet.ListObjects("Table1").ListColumns(LastColumn).DataBodyRange.Select
    ActiveCell.FormulaR1C1 = "=DATEVALUE([@[Creation_Date]])"
    Selection.FillDown

End Sub
Sub Add_HoldDateCalc()
Attribute Add_HoldDateCalc.VB_Description = "Adds the HoldDateCalc column\n"
Attribute Add_HoldDateCalc.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Add_HoldDateCalc Macro
' Adds the HoldDateCalc column
'

   ' Add a new column to the end of Table1 with the column name
    Add_NewColumnToTableEndWithName ("HoldDateCalc")

    ' Get the last column object
    LastColumn = ActiveSheet.ListObjects("Table1").Range.Columns.Count

    ' Add the data or formula to every cell in this column
    LastColumnDataRange = ActiveSheet.ListObjects("Table1").ListColumns(LastColumn).DataBodyRange.Select
    ActiveCell.FormulaR1C1 = "=IF([@[Hold_Until_Date]]="""",TODAY(),DATEVALUE([@[Hold_Until_Date]]))"
    Selection.FillDown

End Sub
Sub Add_DueDateCalc()
'
' Add_DueDateCalc Macro
' Adds the DueDateCalc column
'

   ' Add a new column to the end of Table1 with the column name
    Add_NewColumnToTableEndWithName ("DueDateCalc")

    ' Get the last column object
    LastColumn = ActiveSheet.ListObjects("Table1").Range.Columns.Count

    ' Add the data or formula to every cell in this column
    LastColumnDataRange = ActiveSheet.ListObjects("Table1").ListColumns(LastColumn).DataBodyRange.Select
    ActiveCell.FormulaR1C1 = "=IF([@[DueDate]]="""",TODAY(),DATEVALUE([@[DueDate]]))"
    Selection.FillDown

End Sub

Sub Add_SeverityCalc()
'
' Add_SeverityCalc Macro
' Adds the SeverityCalc column
'

   ' Add a new column to the end of Table1 with the column name
    Add_NewColumnToTableEndWithName ("SeverityCalc")

    ' Get the last column object
    LastColumn = ActiveSheet.ListObjects("Table1").Range.Columns.Count

    ' Add the data or formula to every cell in this column
    LastColumnDataRange = ActiveSheet.ListObjects("Table1").ListColumns(LastColumn).DataBodyRange.Select
    ActiveCell.FormulaR1C1 = "=IF([@Type]=""Defect"",[@Severity],""Enhancement"")"
    Selection.FillDown

End Sub
Sub Add_Escape()
'
' Add_Escape Macro
' Adds the Escape column
' Checks to see if a PTR is an escape or not
'

   ' Add a new column to the end of Table1 with the column name
    Add_NewColumnToTableEndWithName ("Escape")

    ' Get the last column object
    LastColumn = ActiveSheet.ListObjects("Table1").Range.Columns.Count

    ' Add the data or formula to every cell in this column
    LastColumnDataRange = ActiveSheet.ListObjects("Table1").ListColumns(LastColumn).DataBodyRange.Select
    ActiveCell.FormulaR1C1 = "=IF([@Type]=""Defect"",IF([@[Phase_Injected]]=[@[Phase_Discovered]],""No"",""Yes""),""No"")"
    Selection.FillDown

End Sub
Sub Add_EscapeType()
'
' Add_EscapeType Macro
' Adds the EscapeType column
' Inserts formula for looking up the EscapeType from Table2
'

   ' Add a new column to the end of Table1 with the column name
    Add_NewColumnToTableEndWithName ("EscapeType")

    ' Get the last column object
    LastColumn = ActiveSheet.ListObjects("Table1").Range.Columns.Count

    ' Add the data or formula to every cell in this column
    LastColumnDataRange = ActiveSheet.ListObjects("Table1").ListColumns(LastColumn).DataBodyRange.Select
    ActiveCell.FormulaR1C1 = "=IF([@Escape]=""Yes"",INDEX(Table2,MATCH([@[Defect_Category]],Table2[DefectCategory],0),Column(Table2[EscapeType])),""Non-Escape"")"
    Selection.FillDown

End Sub
Sub Add_New_WS(worksheetName)
'
' Add_New_WS Macro
'

' Adds a new worksheet after the active sheet with the name specified by the caller
    Dim ws As Worksheet
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = worksheetName

End Sub
Sub Add_AllCreated_Worksheet()
'
' Add_AllCreated_Worksheet Macro
'
' Adds the "AllCreated" Worksheet as the base worksheet for deriving all worksheets and pivot charts

    ' Create the AllCreated Worksheet
    Add_New_WS ("AllCreated")

    ' Create a PivotTable from Table1
    Add_PivotTable

    ' Add the Pivot Table Fields to create a pivot table for the AllCreated PTRs Pivot Chart
    Add_AllCreated_PivotTableFields

    ' Create sub fields from the CreationDateCalc fields by using the 'Group' function in Excel
    Create_GroupsForCreationDate

    ' Create a Pivot Chart from the current pivot table
    ' This chart will be the base Pivot Chart to be copied and modified
    Create_PivotChart

    ' Format the Pivot Chart for All Created PTRs
    Format_PivotChart ("All Created PTRs")

End Sub

Sub Add_PivotTable()
Attribute Add_PivotTable.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Add_PivotTable Macro
'
' Adds a Pivot Table to the "All Created" worksheet
'
    Range("A1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Table1", Version:=xlPivotTableVersion15).CreatePivotTable TableDestination _
        :="AllCreated!R1C1", TableName:="PivotTable1", DefaultVersion:= _
        xlPivotTableVersion15
    Sheets("AllCreated").Select
    Cells(1, 1).Select

End Sub
Sub Add_AllCreated_PivotTableFields()
Attribute Add_AllCreated_PivotTableFields.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Add_AllCreated_PivotTableFields Macro
'
' Adds the Pivot Table Fields to the AllCreated Pivot Table

    Sheets("AllCreated").Select

    'Create Filters
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CI")
        .Orientation = xlPageField
        .Position = 1
    End With

    'Create Rows
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CreationDateCalc")
        .Orientation = xlRowField
        .Position = 1
    End With

    'No Columns to create

    'Create Values
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("SeverityCalc"), "Count of SeverityCalc", xlCount


End Sub
Sub Create_GroupsForCreationDate()
Attribute Create_GroupsForCreationDate.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Create_GroupsForCreationDate Macro
'
    ' Create the sub fields for the CreationDateCalc field
   ActiveSheet.PivotTables("PivotTable1").PivotFields("CreationDateCalc").DataRange.Cells(1, 1).Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        True, True, True, True)

    ' Year_Created
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Years")
        .Caption = "Year_Created"
        .IncludeNewItemsInFilter = True
        .Orientation = xlRowField
    End With

    ' Quarter_Created
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Quarters")
        .Caption = "Quarter_Created"
        .IncludeNewItemsInFilter = True
        .Orientation = xlHidden
    End With

    ' Month_Created
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Months")
        .Caption = "Month_Created"
        .IncludeNewItemsInFilter = True
        .Orientation = xlHidden
    End With

    ' CreationDateCalc_Day
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CreationDateCalc")
        .Caption = "CreationDateCalc_Day"
        .IncludeNewItemsInFilter = True
        .Orientation = xlHidden
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, _
        False, False, False)
    End With

End Sub
Sub Remove_Unused_Fields()
Attribute Remove_Unused_Fields.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Remove_Unused_Fields Macro
'
' Removes the CreationDate_Calc sub fields that are not used in the Pivot Table

    ActiveSheet.PivotTables("PivotTable1").PivotFields("Quarter_Created"). _
        Orientation = xlHidden
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Month_Created"). _
        Orientation = xlHidden
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CreationDateCalc_Day"). _
        Orientation = xlHidden
End Sub
Sub Create_PivotChart()
Attribute Create_PivotChart.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Create_PivotChart Macro
'
    ' Creates the Pivot Chart
    ActiveSheet.Shapes.AddChart2(297, xlColumnStacked).Select

    ' Set the Position of the Pivot Chart and Chart Size
    With ActiveSheet.Shapes("Chart 1")
        .Top = Range("C2").Top
        .Left = Range("D2").Left
        .Height = Range("C2:C22").Height
        .Width = Range("C2:M2").Width
    End With

    ' Set the Chart to have a Title
    ActiveChart.SetElement (msoElementChartTitleAboveChart)

    ' Set the Chart to have a Data Table and Legend
    ActiveChart.SetElement (msoElementDataTableWithLegendKeys)

    ' Set the Chart Style
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 304

End Sub
Sub Format_PivotChart(Title)
Attribute Format_PivotChart.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Format_PivotChart Macro
'
    ' Set the Chart Title
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "SCCF CQ Database" & Chr(10) & Title & " - " & Date

    'Format the Chart Title to Centered
    With Selection.Format.TextFrame2.TextRange.Characters(1, 11).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With

    ActiveChart.ChartArea.Select

End Sub
Sub Copy_PivotWorksheet(worksheetName)
Attribute Copy_PivotWorksheet.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Copy_PivotWorksheet Macro
'
    ' Copy the current worksheet to the end
    ActiveSheet.Select
    ActiveSheet.Copy After:=Sheets(Sheets.Count)

    ' Rename the worksheet to the name specified
    ActiveSheet.Select
    ActiveSheet.Name = worksheetName

End Sub
Sub Create_AllCreatedBySeverity()
'
' Create_AllCreatedBySeverity Macro
'

    ' Copy the current pivot worksheet and rename it
    Copy_PivotWorksheet ("AllCreatedBySeverity")

    ' Set the Chart Title
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "SCCF CQ Database" & Chr(10) & "All Created By Severity - " & Date

    ' Add the Severity Calc to the Pivot Column
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("SeverityCalc")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ' Update the Sort direction
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.PivotLayout.PivotTable.PivotFields("SeverityCalc").AutoSort _
        xlDescending, "SeverityCalc"

    ' Change the Chart size
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveSheet.Shapes("Chart 1").ScaleWidth 1.2152380952, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart 1").ScaleHeight 1.1493055556, msoFalse, _
        msoScaleFromTopLeft



End Sub
Sub Create_OpenBySeverity()
Attribute Create_OpenBySeverity.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Create_OpenBySeverity Macro
'
    ' Copy the current pivot worksheet and rename it
    Copy_PivotWorksheet ("OpenBySeverity")

    ' Remove Unused Fields
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Year_Created"). _
        Orientation = xlHidden

    ' Add the Fields for OpenBySeverity
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("State")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Type")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("PTR_Number")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("SeverityCalc")
        .Orientation = xlRowField
        .Position = 1
    End With


    ' Set the Chart Title
    ActiveSheet.ChartObjects("Chart 1").Activate
    Format_PivotChart ("Open By Severity")

    ' Update the Sort direction for SeverityCalc
    ActiveChart.PivotLayout.PivotTable.PivotFields("SeverityCalc").AutoSort _
        xlAscending, "SeverityCalc"

    ' Set the State Field Filter for just the open PTRs
    SetStateToOpenPTRs

End Sub
Sub Create_OpenBySeverityCI()
Attribute Create_OpenBySeverityCI.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Create_OpenBySeverityCI Macro
'
    ' Copy the current pivot worksheet and rename it
    Copy_PivotWorksheet ("OpenBySeverityCI")

    ' Set the Chart Title
    ActiveSheet.ChartObjects("Chart 1").Activate
    Format_PivotChart ("Open By Severity CI")

    ' Update the Pivot Table Fields
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("SeverityCalc")
        .Orientation = xlColumnField
        .Position = 1
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HWCI_CSCI")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' Set the HWCI_CSCI field filter to show the iCON components
    SetHwciToIcon

    ' Update the sorting for SeverityCalc
    ActiveSheet.PivotTables("PivotTable1").PivotFields("SeverityCalc").AutoSort _
        xlDescending, "SeverityCalc"

    ' Update the Pivot Chart Size
    ActiveSheet.Shapes("Chart 1").ScaleWidth 1.0461095101, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart 1").ScaleHeight 1.151057719, msoFalse, _
        msoScaleFromTopLeft

End Sub

Sub Create_OpenByState()
Attribute Create_OpenByState.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Create_OpenByState Macro
'
    ' Copy the current pivot worksheet and rename it
    Copy_PivotWorksheet ("OpenByState")

    ' Set the Chart Title
    ActiveSheet.ChartObjects("Chart 1").Activate
    Format_PivotChart ("Open By State")

    ' Update the Pivot Table Fields
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("SeverityCalc")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("State")
        .Orientation = xlColumnField
        .Position = 1
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HWCI_CSCI")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' Set the State field filter for just the open PTRs
    SetStateToOpenPTRs

    ' Set the sort order for the State field
    ActiveChart.PivotLayout.PivotTable.PivotFields("State").AutoSort xlAscending, _
        "State"

    ' Set the HWCI_CSCI field filter to show the iCON components
    SetHwciToIcon

    ' Set the sort order for the HWCI_CSCI field
    ActiveChart.PivotLayout.PivotTable.PivotFields("HWCI_CSCI").AutoSort _
        xlAscending, "HWCI_CSCI"

End Sub

Sub Create_OpenByState_Current()
Attribute Create_OpenByState_Current.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Create_OpenByState_Current Macro
'

    ' Copy the current pivot worksheet and rename it
    Copy_PivotWorksheet ("OpenByState_Current")

    ' Set the Chart Title
    ActiveSheet.ChartObjects("Chart 1").Activate
    Format_PivotChart ("Open By State (Current)")

    ' Update the Pivot Table Fields
    With ActiveChart.PivotLayout.PivotTable.PivotFields("State")
        .Orientation = xlRowField
        .Position = 2
    End With

    With ActiveChart.PivotLayout.PivotTable.PivotFields("HWCI_CSCI")
        .Orientation = xlPageField
        .Position = 1
    End With

    'Set State to Open PTRs
    SetStateToOpenPTRs

    ' Update the Chart Size
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveSheet.Shapes("Chart 1").ScaleWidth 1.1911111111, msoFalse, _
        msoScaleFromTopLeft

End Sub

Sub Create_HoldDate()
Attribute Create_HoldDate.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Create_HoldDate Macro
'

    ' Create the Worksheet
    Add_New_WS ("HoldDate")

    'Define the source data
    SrcData = "Table1"

    'Define the worksheet to operate on
    Set sht = ActiveSheet

    'Set the Pivot Table start location to A1
    StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

    'Create Pivot Cache from Source Data
    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

    'Create Pivot table from Pivot Cache
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")

    ' Add the Pivot Table Filters
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CI")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("State")
        .Orientation = xlPageField
        .Position = 2
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HWCI_CSCI")
        .Orientation = xlPageField
        .Position = 2
    End With

    ' Add the Pivot Table Columns
    ' -- No Columns in this table --

    ' Add the Pivot Table Rows
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HoldDateCalc")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' Add the Values
    Activate_CountofSeverityCalc

    ' Create Hold Date Groups
    Set rngGroup = ActiveSheet.PivotTables("PivotTable1").PivotFields("HoldDateCalc").DataRange
    rngGroup.Cells(1).Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        True, True, True, True)

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Years")
        .Caption = "Year_Hold"
        .IncludeNewItemsInFilter = True
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Quarters")
        .Caption = "Quarter_Hold"
        .IncludeNewItemsInFilter = True
    End With

      With ActiveSheet.PivotTables("PivotTable1").PivotFields("Months")
        .Caption = "Month_Hold"
        .IncludeNewItemsInFilter = True
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HoldDateCalc")
        .Caption = "HoldDateCalc_Day"
        .IncludeNewItemsInFilter = True
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, _
        False, False, False)
    End With

    ' Remove the unused fields
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Quarter_Hold").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Month_Hold").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("PivotTable1").PivotFields("HoldDateCalc_Day"). _
        Orientation = xlHidden

    ' Create the Chart
    Create_PivotChart

    ' Set the default Chart format
    Format_PivotChart ("iCON PTR Hold Dates")

    ' Set the "State" to only show HOLD
   SetStateToHold

    ' Set the "HWCI_CSCI" to All
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HWCI_CSCI")
        .CurrentPage = "(All)"
        .EnableMultiplePageItems = True
    End With

    ' Set the "CI" to iCON
    IconSet = SetPivotItemArrayVisibleTrue("PivotTable1", "CI", Array("iCON"))

End Sub

Sub Create_PastDue()
Attribute Create_PastDue.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Create_PastDue Macro
'
    ' Create the Worksheet
    Add_New_WS ("DueDate")

    'Define the source data
    SrcData = "Table1"

    'Define the worksheet to operate on
    Set sht = ActiveSheet

    'Set the Pivot Table start location to A1
    StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

    'Create Pivot Cache from Source Data
    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

    'Create Pivot table from Pivot Cache
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")

    ' Add the Pivot Table Filters
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CI")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("State")
        .Orientation = xlPageField
        .Position = 2
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HWCI_CSCI")
        .Orientation = xlPageField
        .Position = 2
    End With

    ' Add the Pivot Table Columns
    ' -- No Columns in this table --

    ' Add the Pivot Table Rows
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("DueDateCalc")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' Add the Values
    Activate_CountofSeverityCalc

    ' Create Hold Date Groups
    Set rngGroup = ActiveSheet.PivotTables("PivotTable1").PivotFields("DueDateCalc").DataRange
    rngGroup.Cells(1).Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        True, True, True, True)

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Years")
        .Caption = "Year_Due"
        .IncludeNewItemsInFilter = True
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Quarters")
        .Caption = "Quarter_Due"
        .IncludeNewItemsInFilter = True
    End With

      With ActiveSheet.PivotTables("PivotTable1").PivotFields("Months")
        .Caption = "Month_Due"
        .IncludeNewItemsInFilter = True
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("DueDateCalc")
        .Caption = "DueDateCalc_Day"
        .IncludeNewItemsInFilter = True
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, _
        False, False, False)
    End With

    ' Remove the unused fields
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Quarter_Due").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Month_Due").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("PivotTable1").PivotFields("DueDateCalc_Day"). _
        Orientation = xlHidden

    ' Create the Chart
    Create_PivotChart

    ' Set the default Chart format
    Format_PivotChart ("iCON PTR Due Dates")

    ' Set the "State" to show open past due PTRs
    SetStateToPastDuePTRs

    ' Set the "HWCI_CSCI" to All
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HWCI_CSCI")
        .CurrentPage = "(All)"
        .EnableMultiplePageItems = True
    End With

    ' Set the "CI" to iCON
    IconSet = SetPivotItemArrayVisibleTrue("PivotTable1", "CI", Array("iCON"))

End Sub

Sub Create_PastDueByState()
'
' Create_PastDueByState Macro
'

    ' Copy the current pivot worksheet and rename it
    Copy_PivotWorksheet ("PastDueByState")

    ' Set the Chart Title
    ActiveSheet.ChartObjects("Chart 1").Activate
    Format_PivotChart ("Open PTR Due Dates By State")

    ' Update the Pivot Table Fields
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Year_Due")
        .Orientation = xlColumnField
        .Position = 1
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("State")
        .Orientation = xlRowField
        .Position = 1
    End With

    ActiveSheet.ChartObjects("Chart 1").Activate
    With ActiveChart.Parent
         .Height = 300
         .Width = 500
         .Top = 10
         .Left = 300
    End With

    ' Set the "State" to show open past due PTRs
    SetStateToPastDuePTRs


End Sub

Sub Create_CorePtrs()
Attribute Create_CorePtrs.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Create_CorePtrs Macro
'
    ' Create the Worksheet
    Add_New_WS ("Core_PTRs")

    'Define the source data
    SrcData = "Table1"

    'Define the worksheet to operate on
    Set sht = ActiveSheet

    'Set the Pivot Table start location to A1
    StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

    'Create Pivot Cache from Source Data
    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

    'Create Pivot table from Pivot Cache
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")
    
    ' Add the Pivot Table Filters
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CI")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("State")
        .Orientation = xlPageField
        .Position = 2
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HWCI_CSCI")
        .Orientation = xlPageField
        .Position = 3
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Type")
        .Orientation = xlPageField
        .Position = 4
    End With

    ' Add the Pivot Table Columns
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("SeverityCalc")
        .Orientation = xlColumnField
        .Position = 1
    End With


    ' Add the Pivot Table Rows
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CSC")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' Add the Values
    Activate_CountOfPTR_Number

    ' Create the Chart
    Create_PivotChart

    ' Set the default Chart format
    Format_PivotChart ("Open PTRs for Core")

    ' Update the chart format
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.ChartType = xlBarStacked
    ActiveChart.SetElement (msoElementDataLabelInsideEnd)
    With ActiveChart.Parent
         .Height = 500
         .Width = 700
         .Top = 10
         .Left = 300
    End With

    ' Set the "State" to show open PTRs
    SetStateToOpenPTRs

    ' Set the HWCI_CSCI field filter to show the iCON components
    'SetHwciToIcon

    ' Set the "HWCI_CSCI" to Core
    CoreSet = SetPivotItemArrayVisibleTrue("PivotTable1", "HWCI_CSCI", Array("Core"))

    ' Set the "CI" to iCON
    IconSet = SetPivotItemArrayVisibleTrue("PivotTable1", "CI", Array("iCON"))

    ' Set the sort order for the CSC
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CSC").AutoSort xlDescending, "CSC"

    ' Set the sort order for the Severity Calc
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CSC").AutoSort xlAscending, "SeverityCalc"


End Sub
Sub Create_AllEscapes()
'
' Create_AllEscapes Macro
'

    ' Copy the current pivot worksheet and rename it
    Copy_PivotWorksheet ("AllEscapes")

    ' Set the Chart Title
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.ChartTitle.Select
    Format_PivotChart ("All Escapes")
    ActiveChart.ChartType = xlColumnStacked
    ActiveChart.SetElement (msoElementDataLabelNone)
    ActiveChart.SetElement (msoElementDataTableWithLegendKeys)

    ' Remove all PivotTable Fields
    ResetPivotTable

    ' Update the Pivot Table Filters
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CI")
        .Orientation = xlPageField
        .Position = 1
    End With

    ' Update the Pivot Table Columns
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("EscapeType")
        .Orientation = xlColumnField
        .Position = 1
        .PivotItems("Non-Escape").Visible = False
    End With

    ' Update the Pivot Table Row
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Defect_Category")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' Update the Values
    Activate_CountOfPTR_Number

    ' Update the Sort direction
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.PivotLayout.PivotTable.PivotFields("SeverityCalc").AutoSort _
        xlDescending, "SeverityCalc"

    ' Set the Chart size
    ActiveSheet.ChartObjects("Chart 1").Activate
    With ActiveChart.Parent
         .Height = 325
         .Width = 500
         .Top = 10
         .Left = 300
    End With

End Sub
Sub Create_EscapesByType()
'
' Create_EscapesByType Macro
'

    ' Copy the current pivot worksheet and rename it
    Copy_PivotWorksheet ("EscapesByType")

    ' Set the Chart Title
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.ChartTitle.Select
    Format_PivotChart ("Escapes By Type")
    ActiveChart.ChartType = xlColumnStacked
    ActiveChart.SetElement (msoElementDataLabelNone)
    ActiveChart.SetElement (msoElementDataTableWithLegendKeys)

    ' Remove all PivotTable Fields
    ResetPivotTable

    ' Update the Pivot Table Filters
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CI")
        .Orientation = xlPageField
        .Position = 1
    End With

    ' Update the Pivot Table Columns
    ' -- No Columns for this chart --

    ' Update the Pivot Table Row
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("EscapeType")
        .Orientation = xlRowField
        .Position = 1
        .PivotItems("Non-Escape").Visible = False
    End With

    ' Update the Values
    Activate_CountOfPTR_Number

End Sub

Sub Add_EscapeDefinitions()
'
' Add_EscapeDefinitions Macro
'
    ' Add a new worksheet
    Add_New_WS ("DataLookups")

    ' Create Header Row
    Range("A1:C1") = Array("Index", "EscapeType", "DefectCategory")

    ' Entry 1
    Range("A2:C2") = Array("1", "Planning", "Process")

    ' Entry 2
    Range("A3:C3") = Array("2", "Planning", "Requirements")

    ' Entry 3
    Range("A4:C4") = Array("3", "Planning", "Documentation")

    ' Entry 4
    Range("A5:C5") = Array("4", "Planning", "Analyses")

    ' Entry 5
    Range("A6:C6") = Array("5", "Planning", "Architecture")

    ' Entry 6
    Range("A7:C7") = Array("6", "Execution", "Implementation")

    ' Entry 7
    Range("A8:C8") = Array("7", "Execution", "Logic/Design")

    ' Entry 8
    Range("A9:C9") = Array("8", "Synthesis", "Test Environment")

    ' Entry 9
    Range("A10:C10") = Array("9", "Synthesis", "COTS")

    ' Entry 10
    Range("A11:C11") = Array("10", "Interface", "Interface")

    ' Center the text in all the columns
    Columns("A:C").Select
    Columns("A:C").EntireColumn.AutoFit
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    'Create a table with the definitions
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1").CurrentRegion, , xlYes).Name = "Table2"

    Range("A1:A1").Select

End Sub
Sub ResetPivotTable()
'
'  ResetPivotTable Macro
'
    Dim pvt As PivotTable
    Dim pvf As PivotField

    ' Activate the Pivot Table to clear
    Set pvt = ActiveSheet.PivotTables(1)

    ' Set Manual Updates
    pvt.ManualUpdate = True

    ' Clear each of the RowFields
    For Each rowFld In pvt.RowFields
        rowFld.Orientation = xlHidden
    Next rowFld

    ' Clear each of the ColumnFields
    For Each colFld In pvt.ColumnFields
        colFld.Orientation = xlHidden
    Next colFld

    ' Clear each of the PageFields
    For Each pageFld In pvt.PageFields
        pageFld.Orientation = xlHidden
    Next pageFld

    ' Clear each of the DataFields
    For Each dataFld In pvt.DataFields
        dataFld.Orientation = xlHidden
    Next dataFld

    ' Clear Manual Updates
    pvt.ManualUpdate = False

    ' Refresh the PivotTable
    pvt.RefreshTable

End Sub
Sub Activate_CountOfState()
'
'  Activate_CountOfState Macro
'
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("State"), "Count of State", xlCount

End Sub
Sub Activate_CountOfPTR_Number()
'
'  Activate_CountOfPTR_Number Macro
'
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("PTR_Number"), "Count of PTR_Number", xlCount

End Sub
Sub Activate_CountofSeverityCalc()
'
'  Activate_CountofSeverityCalc Macro
'
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("SeverityCalc"), "Count of SeverityCalc", xlCount

End Sub

Sub Save_Workbook()
'
'  Save_Workbook Macro
'

    ' Jesus saves...so should you
    ActiveWorkbook.Save

End Sub

Sub RemoveDuplicates(ColumnName)
'
'  RemoveDuplicates Macro
'
    ' Removes duplicates from Table 1 using the column name as an index
    columnIndex = ActiveSheet.ListObjects("Table1").ListColumns(ColumnName).Index
    ActiveSheet.Range("Table1[#All]").RemoveDuplicates Columns:=columnIndex, Header:=xlYes

End Sub

Sub SetActiveSheetName(sheetName)
'
'  SetActiveSheetName Macro
'
    ' Sets the Active sheet to the sheet name specified
    Worksheets(sheetName).Activate

End Sub

Sub Cleanup()
'
' Cleanup Macro
'
'  Use in a contingency if the top-level macro fails for any reason.
'  A failure will leave the screen updating set to False.
'
    ' Turn On screen updates and make sure auto calc is functioning
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic

End Sub

Sub Create_ExpiredHoldDate()
'
' Create_ExpiredHoldDate Macro
'

    ' Create the Worksheet
    Add_New_WS ("ExpiredHoldDate")

    'Define the source data
    SrcData = "Table1"

    'Define the worksheet to operate on
    Set sht = ActiveSheet

    'Set the Pivot Table start location to A1
    StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

    'Create Pivot Cache from Source Data
    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

    'Create Pivot table from Pivot Cache
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")

    ' Add the Pivot Table Filters
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CI")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("State")
        .Orientation = xlPageField
        .Position = 2
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HWCI_CSCI")
        .Orientation = xlPageField
        .Position = 2
    End With

    ' Add the Pivot Table Columns
    ' -- No Columns in this table --

    ' Add the Pivot Table Rows
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HoldDateCalc")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' Add the Values
    Activate_CountofSeverityCalc

    ' Create Hold Date Groups
    Set rngGroup = ActiveSheet.PivotTables("PivotTable1").PivotFields("HoldDateCalc").DataRange
    rngGroup.Cells(1).Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        True, True, True, True)

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Years")
        .Caption = "Year_Hold"
        .IncludeNewItemsInFilter = True
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Quarters")
        .Caption = "Quarter_Hold"
        .IncludeNewItemsInFilter = True
    End With

      With ActiveSheet.PivotTables("PivotTable1").PivotFields("Months")
        .Caption = "Month_Hold"
        .IncludeNewItemsInFilter = True
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HoldDateCalc")
        .Caption = "HoldDateCalc_Day"
        .IncludeNewItemsInFilter = True
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, _
        False, False, False)
    End With

    ' Remove the unused fields
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Quarter_Hold").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Month_Hold").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("PivotTable1").PivotFields("HoldDateCalc_Day"). _
        Orientation = xlHidden

    ' Create the Chart
    Create_PivotChart

    ' Set the default Chart format
    'ActiveSheet.ChartObjects("Chart 1").Activate
    Format_PivotChart ("iCON PTR Expired Hold Dates")

    ' Set the "State" to only show HOLD
   SetStateToHold

    ' Set the "HWCI_CSCI" to All
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HWCI_CSCI")
        .CurrentPage = "(All)"
        .EnableMultiplePageItems = True
    End With

    ' Set the "CI" to iCON
    IconSet = SetPivotItemArrayVisibleTrue("PivotTable1", "CI", Array("iCON"))

    ' Set the Date Filter to show data before and equal to today
    ActiveSheet.PivotTables("PivotTable1").PivotFields("HoldDateCalc_Day").PivotFilters _
        .Add2 Type:=xlBeforeOrEqualTo, Value1:=Date



End Sub
Sub Create_ExpiredDueDate()
'
' Create_ExpiredDueDate Macro
'

    ' Create the Worksheet
    Add_New_WS ("ExpiredDueDate")

    'Define the source data
    SrcData = "Table1"

    'Define the worksheet to operate on
    Set sht = ActiveSheet

    'Set the Pivot Table start location to A1
    StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)

    'Create Pivot Cache from Source Data
    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

    'Create Pivot table from Pivot Cache
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")

    ' Add the Pivot Table Filters
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CI")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("State")
        .Orientation = xlPageField
        .Position = 2
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HWCI_CSCI")
        .Orientation = xlPageField
        .Position = 2
    End With

    ' Add the Pivot Table Columns
    ' -- No Columns in this table --

    ' Add the Pivot Table Rows
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("DueDateCalc")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' Add the Values
    Activate_CountofSeverityCalc

    ' Create Hold Date Groups
    Set rngGroup = ActiveSheet.PivotTables("PivotTable1").PivotFields("DueDateCalc").DataRange
    rngGroup.Cells(1).Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        True, True, True, True)

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Years")
        .Caption = "Year_Due"
        .IncludeNewItemsInFilter = True
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Quarters")
        .Caption = "Quarter_Due"
        .IncludeNewItemsInFilter = True
    End With

      With ActiveSheet.PivotTables("PivotTable1").PivotFields("Months")
        .Caption = "Month_Due"
        .IncludeNewItemsInFilter = True
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("DueDateCalc")
        .Caption = "DueDateCalc_Day"
        .IncludeNewItemsInFilter = True
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, _
        False, False, False)
    End With

    ' Remove the unused fields
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Quarter_Due").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Month_Due").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("PivotTable1").PivotFields("DueDateCalc_Day"). _
        Orientation = xlHidden

    ' Create the Chart
    Create_PivotChart

    ' Set the default Chart format
    Format_PivotChart ("iCON PTR Expired Due Dates")

    ' Set the "State" to show open past due PTRs
    SetStateToPastDuePTRs

    ' Set the "HWCI_CSCI" to All
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("HWCI_CSCI")
        .CurrentPage = "(All)"
        .EnableMultiplePageItems = True
    End With

    ' Set the "CI" to iCON
    IconSet = SetPivotItemArrayVisibleTrue("PivotTable1", "CI", Array("iCON"))

    ' Set the Date Filter to show data before and equal to today
    ActiveSheet.PivotTables("PivotTable1").PivotFields("DueDateCalc_Day").PivotFilters _
        .Add2 Type:=xlBeforeOrEqualTo, Value1:=Date



End Sub

Sub SetNone()

    ' Set the Chart "State" to None
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveSheet.PivotTables("PivotTable1").PivotFields("State").ClearAllFilters


End Sub


Function SetPivotItemArrayVisibleTrue(PivotTableName, FieldName, PivotItemNameArray)
'
'  Macro to set the names in a supplied array to a visibility of true
'
    ' Assume a false return value
    SetPivotItemArrayVisibleTrue = False

    ' Set variable for Pivot Fields
    Set PivotField = ActiveSheet.PivotTables(PivotTableName).PivotFields(FieldName)

    ' Setup a collection to capture the settings for the Pivot Items visibility
    Set PivotItemCollection = New Collection
    
    ' Setup a counter for the number of false matches
    CountFalseMatches = 0

    'Loop through the PivotItems
    For Each PivotItem In PivotField.PivotItems

        ' Set a variable to match the name in the pivot field names
        PivotNameMatch = False

            'Loop through the Names in the array
            For Each Name In PivotItemNameArray

                'Check to see if the PivotItem Name matches a name in the array
                If PivotItem.Name = Name Then
                    'If it matches, set visibility to True
                    PivotNameMatch = True
                Else
                    'If it doesn't match, do nothing
                End If
            Next

            ' If this pivot item name is not a match, update the count o false matches
            If PivotNameMatch = False Then
                CountFalseMatches = CountFalseMatches + 1
            End If
            
            ' Add the match result to the collection
            PivotItemCollection.Add PivotNameMatch

        Next

    ' Get the count for the pivot items
    PivotItemCount = PivotField.PivotItems.Count()

    ' Compare the Count of False Matches to the count of pivot items
    ' Setting the pivot field filter will not succeed if all the pivot items are set to false
    If PivotItemCount > CountFalseMatches Then
    
        ' Loop through the pivot field filters and set them according to the collection results
        For i = 1 To PivotItemCount
            ThePivotItemName = PivotField.PivotItems(i).Name
            PivotField.PivotItems(ThePivotItemName).Visible = PivotItemCollection(i)
        Next

        ' Operation succeeded so the set the variable to true
        SetPivotItemArrayVisibleTrue = True
        
    Else
        ' The operation of setting the filter will not cuceed because the array of filters to select
        ' leaves no filters set.  There is nothing to do except exit out with the false return value.
    End If

End Function


Sub SetStateToHold()
'
'  Macro to set the state field in a pivot table to select the HOLD PTRs
'
    'Declare the table variables
    PivotTableName = "PivotTable1"
    FieldName = "State"
    PivotItemNameArray = Array("Hold")

    'Clear all the filters to start from a known configuration
    ActiveSheet.PivotTables(PivotTableName).PivotFields(FieldName).ClearAllFilters

    'Allow multiple pivot items
    ActiveSheet.PivotTables(PivotTableName).PivotFields(FieldName).EnableMultiplePageItems = True

    'Set each Pivot Item to achieve the desired selection
    result = SetPivotItemArrayVisibleTrue(PivotTableName, FieldName, PivotItemNameArray)

End Sub

Sub SetStateToOpenPTRs()
'
'  Macro to set the state field in a pivot table to select the Open PTRs
'
    'Declare the table variables
    PivotTableName = "PivotTable1"
    FieldName = "State"
    PivotItemNameArray = Array("Analysis_Complete", "Completed", "Hold", "In_Analysis", "In_Review", "In_Test", "In_Work", "Reviewed", "Submitted", "Verified")

    'Clear all the filters to start from a known configuration
    ActiveSheet.PivotTables(PivotTableName).PivotFields(FieldName).ClearAllFilters

    'Allow multiple pivot items
    ActiveSheet.PivotTables(PivotTableName).PivotFields(FieldName).EnableMultiplePageItems = True

    'Set each Pivot Item to achieve the desired selection
    result = SetPivotItemArrayVisibleTrue(PivotTableName, FieldName, PivotItemNameArray)

End Sub

Sub SetStateToPastDuePTRs()
'
'  Macro to set the state field in a pivot table to select the Open PTRs
'    that are not in the Hold state
'
    'Declare the table variables
    PivotTableName = "PivotTable1"
    FieldName = "State"
    PivotItemNameArray = Array("Analysis_Complete", "Completed", "In_Analysis", "In_Review", "In_Test", "In_Work", "Reviewed", "Submitted", "Verified")

    'Clear all the filters to start from a known configuration
    ActiveSheet.PivotTables(PivotTableName).PivotFields(FieldName).ClearAllFilters

    'Allow multiple pivot items
    ActiveSheet.PivotTables(PivotTableName).PivotFields(FieldName).EnableMultiplePageItems = True

    'Set each Pivot Item to achieve the desired selection
    result = SetPivotItemArrayVisibleTrue(PivotTableName, FieldName, PivotItemNameArray)

End Sub

Sub SetStateArray()
'
'  Macro to set the state field in a pivot table to select the names in the defined array
'
' Used to test SetPivotItemArrayVisibleTrue function

    'Declare the table variables
    PivotTableName = "PivotTable1"
    FieldName = "State"
    'PivotItemNameArray = Array("Hold")
    'PivotItemNameArray = Array("Hold", "Migrated")
    PivotItemNameArray = Array("Analysis_Complete", "Completed", "In_Analysis", "In_Review", "In_Test", "In_Work", "Reviewed", "Submitted", "Verified")

    'Clear all the filters to start from a known configuration
    ActiveSheet.PivotTables(PivotTableName).PivotFields(FieldName).ClearAllFilters

    'Allow multiple pivot items
    ActiveSheet.PivotTables(PivotTableName).PivotFields(FieldName).EnableMultiplePageItems = True

    'Set each Pivot Item to achieve the desired selection
    result = SetPivotItemArrayVisibleTrue(PivotTableName, FieldName, PivotItemNameArray)

End Sub

Sub HideColumnsOnNotesWorksheet()
'
' Hide_Column Macro
'
    'Define the Table to reference
    Set Table = ActiveSheet.ListObjects(1)

    ' Hide the Description column
    Table.ListColumns("Description").Range.Select
    Selection.EntireColumn.Hidden = True

    ' Hide the bulk of the worksheet
    Range(Table.Name & "[[#All],[Originator]:[CreationDateCalc]]").Select
    Selection.EntireColumn.Hidden = True

    ' Hide the Escape-related columns
    Range(Table.Name & "[[#All],[SeverityCalc]:[EscapeType]]").Select
    Selection.EntireColumn.Hidden = True
    
    ' Autofit selected columns
    Range(Table.Name & "[[#All],[DueDateCalc]:[HoldDateCalc]]").Select
    Selection.EntireColumn.AutoFit

    'Select the PTR_Number as the origin
    Range(Table.Name & "[[#Headers],[PTR_Number]]").Select
    
End Sub

Function DoesPivotItemExist(PivotTableName, FieldName, PivotItemName)
'
'  Macro to see if a pivot item exists
'
    Set PivotField = ActiveSheet.PivotTables(PivotTableName).PivotFields(FieldName)
    DoesPivotItemExist = False

        For Each PivotItem In PivotField.PivotItems
            If PivotItem = PivotItemName Then
                DoesPivotItemExist = True
                Exit For
            End If
        Next

End Function

Sub CreateNotesWorksheet()

    ' Copy the Data pivot worksheet and rename it
    Sheets("Data").Select
    Copy_PivotWorksheet ("Notes")

    ' Add the new columns to the Notes worksheet
    Add_NewColumnToTableEndWithName ("Action")
    Add_NewColumnToTableEndWithName ("Notes")
    Add_NewColumnToTableEndWithName ("Select")

    ' Add formulas to the Notes Worksheet
    'AddFormulas
    
    ' Hide selected columns on the worksheet
    HideColumnsOnNotesWorksheet
    
    ' Set the Row Height
    SetRowHeight

End Sub

Sub ConvertPTRvaluesToNumber()

    ' Select the first row of Table 1 that contains the PTR values as text
    ActiveSheet.ListObjects("Table1").ListColumns("PTR_Number").DataBodyRange.Select
    With Selection
        Selection.NumberFormat = "General"
        .value = .value
    End With
End Sub

Sub SetHwciToIcon()
'
'  Macro to set the HWCI_CSCI field in a pivot table to select the names in the defined array
'
    'Declare the table variables
    PivotTableName = "PivotTable1"
    FieldName = "HWCI_CSCI"
    PivotItemNameArray = Array("Behavior", "Core", "Device", "Developer_Library")

    'Clear all the filters to start from a known configuration
    ActiveSheet.PivotTables(PivotTableName).PivotFields(FieldName).ClearAllFilters

    'Allow multiple pivot items
    ActiveSheet.PivotTables(PivotTableName).PivotFields(FieldName).EnableMultiplePageItems = True

    'Set each Pivot Item to achieve the desired selection
    result = SetPivotItemArrayVisibleTrue(PivotTableName, FieldName, PivotItemNameArray)

End Sub

Sub RemoveActionFormulas()
'
'  Macro to remove the formulas from the Action column
'

    ' Turn off Screen Updates
    Application.ScreenUpdating = False

    ' Clear any filters
    If ActiveSheet.ListObjects(1).ShowAutoFilter Then
        ActiveSheet.ListObjects(1).Range.AutoFilter
        ActiveSheet.ListObjects(1).Range.AutoFilter
    End If

    'Define the Table to reference
    Set Table = ActiveSheet.ListObjects(1)

    ' Select the data in the Action column
    Set Rng = Range(Table.Name & "[Action]")

    ' Convert the formulas to values
    With Rng
        .NumberFormat = "General"
        .value = .value
    End With

    ' Filter Action column for values that are zero
    actionColumn = Table.ListColumns("Action").Index
    With Rng
        .AutoFilter Field:=actionColumn, Criteria1:="=0"
    End With

    ' Make sure the filter has something in it
    visibleRows = ActiveSheet.AutoFilter.Range.Columns(actionColumn).SpecialCells(xlCellTypeVisible).Count

    ' Clear the contents of the filtered range
    If visibleRows > 1 Then
        With Rng
            .ClearContents
        End With
    End If

    ' Toggle filter on Action column to remove filter criteria
    With Rng
        .AutoFilter Field:=actionColumn
    End With

    'Select the PTR_Number as the origin
    Range(ActiveSheet.ListObjects(1).Name & "[[#Headers],[PTR_Number]]").Select

    ' Turn on Screen Updates
    Application.ScreenUpdating = True

End Sub
Sub RemoveNotesFormulas()
'
'  Macro to remove the formulas from the Notes column
'

    ' Turn off Screen Updates
    Application.ScreenUpdating = False

    ' Clear any filters
    If ActiveSheet.ListObjects(1).ShowAutoFilter Then
        ActiveSheet.ListObjects(1).Range.AutoFilter
        ActiveSheet.ListObjects(1).Range.AutoFilter
    End If

    'Define the Table to reference
    Set Table = ActiveSheet.ListObjects(1)

    ' Select the data in the Notes column
    Set Rng = Range(Table.Name & "[Notes]")

    ' Convert the formulas to values
    With Rng
        .NumberFormat = "General"
        .value = .value
    End With

    ' Filter Notes column for values that are zero
    notesColumn = Table.ListColumns("Notes").Index
    With Rng
        .AutoFilter Field:=notesColumn, Criteria1:="=0"
    End With

    ' Make sure the filter has something in it
    visibleRows = ActiveSheet.AutoFilter.Range.Columns(notesColumn).SpecialCells(xlCellTypeVisible).Count

    ' Clear the contents of the filtered range
    If visibleRows > 1 Then
        With Rng
            .ClearContents
        End With
    End If

    ' Toggle filter on Notes column to remove filter criteria
    With Rng
        .AutoFilter Field:=notesColumn
    End With

    'Select the PTR_Number as the origin
    Range(ActiveSheet.ListObjects(1).Name & "[[#Headers],[PTR_Number]]").Select

    ' Turn on Screen Updates
    Application.ScreenUpdating = True

End Sub

Sub UnhideColumn()

    'Define the Table to reference
    Set Table = ActiveSheet.ListObjects(1)
    
    'Get the Column name to unhide
    theColumnName = InputBox("Enter Column Name to Unhide", "Unhide Column")
    
    'Determine if the column name is valid
    If theColumnName <> "" Then
        
        'Check that the column exists in the table
        columnNameIsValid = ColumnExists(Table, theColumnName)
        
    End If
    
    'Unhide valid column name
    If columnNameIsValid Then
        
        'Select the column and unhide it
        Table.ListColumns(theColumnName).Range.Select
        Selection.EntireColumn.Hidden = False
        
    Else
        'Tell user column name is not valid
        result = MsgBox("The column name '" & theColumnName & "' was not found", vbOKOnly, "Error")
    End If
        
End Sub
Sub SetRowHeight(Optional RowHeightValue = 20)

    'Define the Table to reference
    Set Table = ActiveSheet.ListObjects(1)
   
    'Set the Row Height if the value is greater than zero
    If RowHeightValue > 0 Then
        Table.Range.RowHeight = RowHeightValue
    Else
        ' Do nothing because the Row Height is zero or less
    End If

End Sub
Sub removeActionAndNotesFormulas()

    'Select the PTR_Number as the origin
    Range(ActiveSheet.ListObjects(1).Name & "[[#Headers],[PTR_Number]]").Select
    
    RemoveActionFormulas
    RemoveNotesFormulas

End Sub

Function ColumnExists(Table, ColumnName)

    'Assume the Column Exists
    ColumnExists = True
    
    On Error GoTo DoesNotExist
    
        'Attempt to set a variable to the column name
        Set Column = Table.ListColumns(ColumnName)

    'If the line above executes, there was no error, so exit
    Exit Function
    
DoesNotExist:
    'Failed the variable assignment, so column name does not exist
    ColumnExists = False

End Function

Sub SaveAsXLSX()
'
' SaveAsXLSX Macro
' Saves the current file as an XSLX file
'

    ' Create Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    currentFullPath = ThisWorkbook.FullName
    currentPath = fso.GetParentFolderName(ThisWorkbook.FullName)
    currentFileName = ThisWorkbook.Name
    extension = fso.GetExtensionName(ThisWorkbook.FullName)
    currentFile = fso.GetBaseName(ThisWorkbook.FullName)
    
    newExtension = ".xlsx"
    newFileName = fso.BuildPath(currentPath, currentFile & newExtension)
    
    If extension = "xls" Then
        'Hide Display Alerts
        Application.DisplayAlerts = False
        
        ' Save the File
        ActiveWorkbook.SaveAs filename:=newFileName _
            , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            
        ' Restore Display Alerts
        Application.DisplayAlerts = True
    End If
End Sub

Private Function DTG(Optional theDate)
'******************************************************************************
' This function creates a string for the date-time group(DTG).
' Optional parameter to accept a date value or it will use the current time.
' Input format is MM/DD/YYYY HH:mm:ss am/pm     Example: 3/22/2018 4:03:07pm
' Output format is YYYYMMDD_HHmmss              Example: 20180322_160307
'******************************************************************************

    ' Check the Date parameter
    If IsMissing(theDate) Then
        theDate = Now()
    End If
    
    ' Format for 4-digit year
    YY = DatePart("yyyy", theDate)
    
    ' Format for 2-digit month
    MM = DatePart("m", theDate)
    If (MM < 10) Then
        MM = "0" & MM
    End If

    ' Format for 2-digit day
    DD = DatePart("d", theDate)
    If (DD < 10) Then
        DD = "0" & DD
    End If

    ' Format for 2-digit hour
    HH = DatePart("h", theDate)
    If (HH < 10) Then
        HH = "0" & HH
    End If

    ' Format for 2-digit minute
    Min = DatePart("n", theDate)
    If (Min < 10) Then
        Min = "0" & Min
    End If

    ' Format for 2-digit second
    ss = DatePart("s", theDate)
    If (ss < 10) Then
        ss = "0" & ss
    End If

    ' Return the DTG
    DTG = YY & MM & DD & "_" & HH & Min & ss
    
End Function

Private Function DateStamp(Optional theDate)
'******************************************************************************
' This function creates a string for the DateStamp.
' Optional parameter to accept a date value or it will use the current date.
' Input format is MM/DD/YYYY                    Example: 3/22/2018
' Output format is YYYYMMDD                     Example: 20180322
'******************************************************************************

    ' Check the Date parameter
    If IsMissing(theDate) Then
        theDate = Now()
    End If
    
    ' Format for 4-digit year
    YY = DatePart("yyyy", theDate)
    
    ' Format for 2-digit month
    MM = DatePart("m", theDate)
    If (MM < 10) Then
        MM = "0" & MM
    End If

    ' Format for 2-digit day
    DD = DatePart("d", theDate)
    If (DD < 10) Then
        DD = "0" & DD
    End If

    ' Return the DateStamp
    DateStamp = YY & MM & DD
    
End Function

Sub AddActionFormula()

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get last week's date
    fqLastWeekFilename = FindRecentFilePath
    lastWeekFileName = fso.GetFileName(fqLastWeekFilename)
    
    ' Add a formula to the Action column
    actionFormula = "=INDEX('" & lastWeekFileName & "'!Table14[#Data],MATCH([PTR_Number],'" & lastWeekFileName & "'!Table14[PTR_Number],0),COLUMN('" & lastWeekFileName & "'!Table14[Action]))"

    'Define the Table to reference
    Set Table = ActiveSheet.ListObjects(1)

    ' Select the data in the Action column
    Set Rng = Range(Table.Name & "[Action]")
    Rng.Select
   
    ' Apply the formula
    With Selection
        Selection.NumberFormat = "General"
        .value = actionFormula
   End With

End Sub

Sub AddNotesFormula()

    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Get last week's file name
    fqLastWeekFilename = FindRecentFilePath
    lastWeekFileName = fso.GetFileName(fqLastWeekFilename)
    
    ' Add a formula to the Notes column
    notesFormula = "=INDEX('" & lastWeekFileName & "'!Table14[#Data],MATCH([PTR_Number],'" & lastWeekFileName & "'!Table14[PTR_Number],0),COLUMN('" & lastWeekFileName & "'!Table14[Notes]))"

    'Define the Table to reference
    Set Table = ActiveSheet.ListObjects(1)

    ' Select the data in the Notes column
    Set Rng = Range(Table.Name & "[Notes]")
    Rng.Select
   
    ' Apply the formula
    With Selection
        Selection.NumberFormat = "General"
        .value = notesFormula
    End With

End Sub
Function OpenWorkbook()
' Uses the FindRecentFilePath to get the most recent file that is not today's date

    ' Assume the OpenWorkbook is not returned
    OpenWorkbook = -1
    
    ' Create Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Define the start point
    rootDir = "\\gs.myharris.net\data\General\iCON\92_SoftwareSystemsEngineering\Chief Software Architect Data\Report Data\iCON Quarterly\2018\Notes"
    
    ' Define today's file
    todayFilename = DateStamp(Date) & "_SCCF_PTRs_Notes.xlsx"
    fqTodayFilename = fso.BuildPath(rootDir, todayFilename)
    
    ' Determine the filename for last week's file...will return -1 if not valid
    fqLastWeekFilename = FindRecentFilePath
    
    ' Open last week's file if it is valid
    If fso.FileExists(fqLastWeekFilename) Then
        Workbooks.Open fqLastWeekFilename
        OpenWorkbook = fqLastWeekFilename
        Wait 3
    Else
        MsgBox ("File does not exist:" & vbLf & fqLastWeekFilename)
        Exit Function
    End If
    
    ' Set the focus on the current file
    If fso.FileExists(fqTodayFilename) Then
        Workbooks(todayFilename).Activate
    Else
        MsgBox ("File does not exist:" & vbLf & todayFilename)
        Exit Function
    End If
    
End Function
Sub AddFormulas()
' Opens the most recent workbook and copies the Action and Notes to the current workbook

    ' Open the previous workbook
    If (OpenWorkbook <> -1) Then
    
        ' Add a formula to look at the previous workbook Action and Notes
        AddActionFormula
        AddNotesFormula
        
        ' Keep the values and format them while removing the formulas
        removeActionAndNotesFormulas
    
    Else
        ' Unable to open the previous workbook
        MsgBox ("Unable to open the previous workbook")
    End If

End Sub

Function FindRecentFilePath()
' Looks for the filename containing the most recent DateStamp in a folder.
' Returns the FilePath.

    ' Create object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Hard code the root directory until we feel the pain of hard coding
    rootDir = "\\gs.myharris.net\data\General\iCON\92_SoftwareSystemsEngineering\Chief Software Architect Data\Report Data\iCON Quarterly\2018\Notes"

    ' Latest file is not known
    Set LatestFile = Nothing
    FindRecentFilePath = -1
    
    ' Cannot be greater than today
    todayDateStamp = DateStamp

    ' Get the folder object to process
    Set objFolder = fso.GetFolder(rootDir)
    
    ' Loop through the files in the folder
    For Each file In objFolder.Files
    
        ' Get the DateStamp from the file
        theFileDateStamp = RegExDateStamp(file.Name)
        
        ' If the file has a valid DateStamp <> -1, then process it
        If (theFileDateStamp <> -1) Then
            
            ' If the Latest file is nothing, then initialize it with the first file
            If (LatestFile Is Nothing) And (theFileDateStamp < todayDateStamp) Then
                Set LatestFile = file
                LatestFileDateStamp = theFileDateStamp
                FindRecentFilePath = file.Path
            Else
                ' Check the current file against the latest file
                ' If this file date stamp is newer (greater) than the last one
                ' and less than todays date stamp, then update the latest
                ' Otherwise, do nothing and go to the next file
                If (theFileDateStamp > LatestFileDateStamp) And (theFileDateStamp < todayDateStamp) Then
                    Set LatestFile = file
                    LatestFileDateStamp = theFileDateStamp
                    FindRecentFilePath = file.Path
                End If
            End If
        Else
            ' Do nothing, the date stamp is not valid
        End If
        
    Next
    
End Function

Private Function RegExDateStamp(s As String) As String
' Looks for a DateStamp in the format YYYYMMDD at the beginning of a string
' Returns: the Date Stamp in the format YYYYMMDD if found
'           -1 if not found

    Dim re, match, allMatches
    Set re = CreateObject("vbscript.regexp")
    
    ' Looking for YYYYMMDD
    ' Pattern starts at 20100101 and goes to 20991231
    re.Pattern = "^(20[1-9][0-9]([0][1-9]|[1][0-2])([0-2][0-9]|[3][0-1]))"
    re.Global = True
    Set allMatches = re.Execute(s)

    If allMatches.Count <> 0 Then
       
        For Each match In allMatches
            'MsgBox match.Value
            RegExDateStamp = match.value
            Exit For
        Next
        
    Else
        RegExDateStamp = -1
    End If
    
    Set re = Nothing
    
    
End Function
Sub Wait(Optional parameterInSeconds)
' Provides a wait in 1 second increments
    
    ' Set optional parameter
    If IsMissing(parameterInSeconds) Then
        theValueInSeconds = 1
    End If
    
    ' Make usage less dangerous
    If valueInSeconds > 300 Then
        theValueInSeconds = parameterInSeconds / 100
        MsgBox ("Resetting wait time from " & parameterInSeconds & " to " & theValueInSeconds)
    Else
        theValueInSeconds = parameterInSeconds
    End If
    
    ' Use the application wait
    Application.Wait DateAdd("s", valueInSeconds, Now)
    
End Sub

Sub CloseActiveWorkbook()
'******************************************************************************
' This function closes the active workbook without saving changes.
'******************************************************************************

    Application.DisplayAlerts = False
    ActiveWorkbook.Close savechanges:=False
    Application.DisplayAlerts = True
    
End Sub

Sub TellMeTime()
    
    Application.Speech.Speak ("The time is " & Format(Time, "h nn AM/PM"))
    'Wait 5
    
End Sub

Sub MyScratchSub()

'Dim pws As Worksheet, sws As String
sws = Range("Table1").Parent.Name

wsName = ActiveSheet.ListObjects("Table1").Parent.Name

End Sub


