Attribute VB_Name = "Module1"
'DECLARE VARS WITH SCOPE PUBLIC IN CASE WE WANT TO USE ACROSS SUBS OR FUNCS

'FUNCTIONAL VARIABLES
Public rng As Range
Public cell As Range
Public ws As Worksheet

'COLUMN CONSTANTS - copied from hidden worksheet: vba_vars
Public Const stats_column As String = "A"
Public Const pcn_column As String = "B"
Public Const title_column As String = "C"
Public Const burden_column As String = "D"
Public Const burden_aviation_column As String = "E"
Public Const object_code_column As String = "F"
Public Const description_column As String = "G"
Public Const quantity_column As String = "H"
Public Const cost_column As String = "I"
Public Const cost_aviation_column As String = "J"
Public Const rural_airport_column As String = "K"

'CELL CONSTANTS - copied from hidden worksheet: vba_vars
Public Const stats_name_cell As String = stats_column & "3"
Public Const stats_through_miles_title_cell As String = stats_column & "5"
Public Const stats_through_miles_cell As String = stats_column & "6"
Public Const stats_lane_miles_title_cell As String = stats_column & "7"
Public Const stats_lane_miles_cell As String = stats_column & "8"
Public Const stats_sidewalk_miles_title_cell As String = stats_column & "9"
Public Const stats_sidewalk_miles_cell As String = stats_column & "10"
Public Const stats_airport_surface_area_title_cell As String = stats_column & "11"
Public Const stats_airport_surface_area_cell As String = stats_column & "12"
Public Const stats_fed_cip_title_cell As String = stats_column & "13"
Public Const stats_fed_cip_cell As String = stats_column & "14"
Public Const stats_aviation_title_cell As String = stats_column & "17"
Public Const stats_aviation_cell As String = stats_column & "18"
Public Const stats_aviation_percent_title_cell As String = stats_column & "19"
Public Const stats_aviation_percent_cell As String = stats_column & "20"
Public Const stats_total_title_cell As String = stats_column & "21"
Public Const stats_total_cell As String = stats_column & "22"
Public Const stats_district_cell As String = stats_column & "25"
Public Const stats_region_cell As String = stats_column & "26"


'FUNCTION TO CREATE A RANGE FROM COLUMN WITH ROW 2 TO EXCEL ROW LIMIT.
Private Function Row2toEnd(ByRef column As String) As String
    Row2toEnd = column & "2:" & column & "1048576"
End Function
'FUNCTION TO CREATE A RANGE FROM COLUMN WITH ROW 2 TO EXCEL ROW LIMIT WITH $TATIC ADDRESSING.
Private Function Row2toEndStatic(ByRef column As String) As String
    Row2toEndStatic = column & "$2:" & column & "$1048576"
End Function
Private Function statsCell(row)
    statsCell = stats_column & row
End Function
'For creating maintenance stations sheets and links to/from, operates on a user selection. -DP
Sub xCreateSheets()

'___  ___      _       _                                    _____ _        _  ._
'|  \/  |     (_)     | |                                  /  ___| |      | | (_)
'| .  . | __ _ _ _ __ | |_ ___ _ __   __ _ _ __   ___ ___  \ `--.| |_ __ _| |_ _  ___  _ __  ___
'| |\/| |/ _` | | '_ \| __/ _ \ '_ \ / _` | '_ \ / __/ _ \  `--. \ __/ _` | __| |/ _ \| '_ \/ __|
'| |  | | (_| | | | | | ||  __/ | | | (_| | | | | (_|  __/ /\__/ / || (_| | |_| | (_) | | | \__ \
'\_|  |_/\__,_|_|_| |_|\__\___|_| |_|\__,_|_| |_|\___\___| \____/ \__\__,_|\__|_|\___/|_| |_|___/
                                                                                                
                                                                                                
Set rng = Application.InputBox(Prompt:="Select cell range:", Title:="Create sheets", Default:=Selection.Address, Type:=8)

'USER SELECTS ALL MAINTSTATION CELLS
For Each cell In rng
    If cell <> "" Then
        'ADD A WORKSHEET
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        
        'NAME WORKSHEET CURRENT MAINTENANCE STATION
        ws.Name = cell
        
        'CREATE BACK BUTTON THAT WORKS EVEN IF MAIN PAGE IS SORTED
        With ws.Range("A1")
            .Formula = "=HYPERLINK(""#"" & CELL(""address"", XLOOKUP( """ & cell & """, 'Budget Overview'!C:C, 'Budget Overview'!C:C)),""[     <-BACK      ]"")"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        'CREATE STATS AREA
        With ws.Range(stats_name_cell)
            .Value = cell
            .Font.Bold = True
            .Interior.ColorIndex = 11
            .Font.ColorIndex = 2
        End With
        
        ws.Range(stats_through_miles_title_cell).Formula = "Through Miles:"
        ws.Range(stats_lane_miles_title_cell).Formula = "Lane Miles:"
        ws.Range(stats_sidewalk_miles_title_cell).Formula = "Sidewalk Miles:"
        ws.Range(stats_airport_surface_area_title_cell).Formula = "Airport Surface Area:"
        
        ws.Range(stats_fed_cip_title_cell).Formula = "FED/CIP:"
        ws.Range(stats_fed_cip_cell).NumberFormat = "_($* #,##0.0_);_($* (#,##0.0);_($* ""-""??_);_(@_)"
        
        ws.Range(stats_aviation_title_cell).Formula = "Aviation:"
        ws.Range(stats_aviation_cell).NumberFormat = "_($* #,##0.0_);_($* (#,##0.0);_($* ""-""??_);_(@_)"
        ws.Range(stats_aviation_cell).Formula = "=SUMPRODUCT(" & Row2toEnd(burden_column) & "*" & Row2toEnd(burden_aviation_column) & ")+SUMPRODUCT(" & Row2toEnd(cost_column) & "*" & Row2toEnd(cost_aviation_column) & ")"
        
        ws.Range(stats_aviation_percent_title_cell).Formula = "Aviation (%):"
        ws.Range(stats_aviation_percent_cell).NumberFormat = "0%"
        ws.Range(stats_aviation_percent_cell).Formula = "=" & stats_aviation_cell & "/" & stats_total_cell
        
        ws.Range(stats_total_title_cell).Formula = "Total:"
        ws.Range(stats_total_cell).NumberFormat = "_($* #,##0.0_);_($* (#,##0.0);_($* ""-""??_);_(@_)"
        ws.Range(stats_total_cell).Formula = "=SUM(" & Row2toEndStatic(burden_column) & ")+SUM(" & Row2toEndStatic(cost_column) & ")"
        
        ws.Range(stats_district_cell).Formula = Range("'Budget Overview'!B" & cell.row()).Value
        ws.Range(stats_district_cell).HorizontalAlignment = xlRight
        ws.Range(stats_region_cell).Formula = Range("'Budget Overview'!A" & cell.row()).Value

        
        'USER INPUT CELLS, YELLOW BACKGROUND
        ws.Range(stats_through_miles_cell).Interior.ColorIndex = 19
        ws.Range(stats_lane_miles_cell).Interior.ColorIndex = 19
        ws.Range(stats_sidewalk_miles_cell).Interior.ColorIndex = 19
        ws.Range(stats_airport_surface_area_cell).Interior.ColorIndex = 19
        ws.Range(stats_fed_cip_cell).Interior.ColorIndex = 19
        
        'COMPUTED CELLS, YELLOW BACKGROUND
        ws.Range(stats_total_title_cell).Interior.ColorIndex = 15
        ws.Range(stats_total_cell).Interior.ColorIndex = 15
        ws.Range(stats_aviation_percent_title_cell).Interior.ColorIndex = 15
        ws.Range(stats_aviation_percent_cell).Interior.ColorIndex = 15
        ws.Range(stats_aviation_title_cell).Interior.ColorIndex = 15
        ws.Range(stats_aviation_cell).Interior.ColorIndex = 15
        ws.Range(stats_district_cell).Interior.ColorIndex = 15
        ws.Range(stats_region_cell).Interior.ColorIndex = 15
                                  
        'CREATE HEADERS
        ws.Range(pcn_column & "1").Value = "PCN"
        ws.Range(title_column & "1").Value = "Class/Title"
        ws.Range(burden_column & "1").Value = "Full Burden"
        ws.Range(burden_aviation_column & "1").Value = "(%) Aviation"
        ws.Range(object_code_column & "1").Value = "Object Code"
        ws.Range(description_column & "1").Value = "Description"
        ws.Range(quantity_column & "1").Value = "Quantity"
        ws.Range(cost_column & "1").Value = "Cost"
        ws.Range(cost_aviation_column & "1").Value = "(%) Aviation"
        ws.Range(rural_airport_column & "1").Value = "Rural Airports"
        
        'SET WIDTHS
        ws.Columns(stats_column).ColumnWidth = 20
        ws.Columns(pcn_column).ColumnWidth = 12
        ws.Columns(title_column).ColumnWidth = 30
        ws.Columns(burden_column).ColumnWidth = 12
        ws.Columns(burden_aviation_column).ColumnWidth = 12
        ws.Columns(object_code_column).ColumnWidth = 12
        ws.Columns(description_column).ColumnWidth = 30
        ws.Columns(quantity_column).ColumnWidth = 12
        ws.Columns(cost_column).ColumnWidth = 12
        ws.Columns(cost_aviation_column).ColumnWidth = 12
        ws.Columns(rural_airport_column).ColumnWidth = 30
        
        'OBJECT CODE VALIDATION
        With Range(object_code_column & "$2:" & object_code_column & "$1048576").Validation
            .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="2000", Formula2:="5999"
            .ErrorMessage = "You must enter a 4 digit object code"
            .ErrorTitle = "Must be an object code"
        End With
        
        'SET COLUMN FORMATS
        ws.Range(burden_column & "2:" & burden_column & "1048576").NumberFormat = "_($* #,##0.0_);_($* (#,##0.0);_($* ""-""??_);_(@_)"
        ws.Range(cost_column & "2:" & cost_column & "1048576").NumberFormat = "_($* #,##0.0_);_($* (#,##0.0);_($* ""-""??_);_(@_)"
        ws.Range(burden_aviation_column & "2:" & burden_aviation_column & "1048576").NumberFormat = "0%"
        ws.Range(cost_aviation_column & "2:" & cost_aviation_column & "1048576").NumberFormat = "0%"
        
        'SET HEADER COLORS
        ws.Range(pcn_column & "1:" & rural_airport_column & "1").Interior.ColorIndex = 1
        ws.Range(pcn_column & "1:" & rural_airport_column & "1").Font.ColorIndex = 2
        
        'HIDE SPARE COLUMNS
        Range(Cells(1, 12), Cells(1, Columns.Count)).EntireColumn.Hidden = True
        
        'RED BORDERS
        With ws.Range(stats_column & ":" & stats_column).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 3
        End With
        
        With ws.Range(burden_aviation_column & ":" & burden_aviation_column).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 3
        End With
        
        With ws.Range(cost_aviation_column & ":" & cost_aviation_column).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 3
        End With

'______           _            _     ______                    ._
'| ___ \         | |          | |   |  _  |                    (_)
'| |_/ /_   _  __| | __ _  ___| |_  | | | |_   _____ _ ____   ___  _____      __
'| ___ \ | | |/ _` |/ _` |/ _ \ __| | | | \ \ / / _ \ '__\ \ / / |/ _ \ \ /\ / /
'| |_/ / |_| | (_| | (_| |  __/ |_  \ \_/ /\ V /  __/ |   \ V /| |  __/\ V  V /
'\____/ \__,_|\__,_|\__, |\___|\__|  \___/  \_/ \___|_|    \_/ |_|\___| \_/\_/
'                    __/ |
'                   |___/
'
     
        'Set Formulas for row on Budget Overview
        
        'Rural Airports
        Range("'Budget Overview'!D" & cell.row()).Formula = "=COUNTA('" & cell & "'!" & Row2toEndStatic(rural_airport_column) & ")"
        'Airport Surface Area
        Range("'Budget Overview'!E" & cell.row()).Formula = "='" & cell & "'!" & stats_airport_surface_area_cell
        'Through Miles
        Range("'Budget Overview'!F" & cell.row()).Formula = "='" & cell & "'!" & stats_through_miles_cell
        'Lane Miles
        Range("'Budget Overview'!G" & cell.row()).Formula = "='" & cell & "'!" & stats_lane_miles_cell
        'Sidewalk Miles
        Range("'Budget Overview'!H" & cell.row()).Formula = "='" & cell & "'!" & stats_sidewalk_miles_cell
        'Positions
        Range("'Budget Overview'!I" & cell.row()).Formula = "=COUNTA('" & cell & "'!" & Row2toEndStatic(pcn_column) & ")"

        'Total
        Range("'Budget Overview'!J" & cell.row()).Formula = "=SUM('" & cell & "'!" & Row2toEndStatic(burden_column) & ")+SUM('" & cell & "'!" & Row2toEndStatic(cost_column) & ")"
        '1000
        Range("'Budget Overview'!K" & cell.row()).Formula = "=SUM('" & cell & "'!" & Row2toEndStatic(burden_column) & ")"
        '2000
        Range("'Budget Overview'!L" & cell.row()).Formula = "=SUMIFS('" & cell & "'!" & Row2toEndStatic(cost_column) & ", '" & cell & "'!" & Row2toEndStatic(object_code_column) & ", "">1999"", '" & cell & "'!" & Row2toEndStatic(object_code_column) & ", ""<3000"")"
        '3000
        Range("'Budget Overview'!M" & cell.row()).Formula = "=SUMIFS('" & cell & "'!" & Row2toEndStatic(cost_column) & ", '" & cell & "'!" & Row2toEndStatic(object_code_column) & ", "">2999"", '" & cell & "'!" & Row2toEndStatic(object_code_column) & ", ""<4000"")"
        '4000
        Range("'Budget Overview'!N" & cell.row()).Formula = "=SUMIFS('" & cell & "'!" & Row2toEndStatic(cost_column) & ", '" & cell & "'!" & Row2toEndStatic(object_code_column) & ", "">3999"", '" & cell & "'!" & Row2toEndStatic(object_code_column) & ", ""<5000"")"
        '5000
        Range("'Budget Overview'!O" & cell.row()).Formula = "=SUMIFS('" & cell & "'!" & Row2toEndStatic(cost_column) & ", '" & cell & "'!" & Row2toEndStatic(object_code_column) & ", "">4999"", '" & cell & "'!" & Row2toEndStatic(object_code_column) & ", ""<6000"")"
       
        'FED/CIP
        Range("'Budget Overview'!P" & cell.row()).Formula = "='" & cell & "'!" & stats_fed_cip_cell
        'Rural Amount
        Range("'Budget Overview'!Q" & cell.row()).Formula = "='" & cell & "'!" & stats_aviation_cell
        
        'Always do this last because cell is used above, Hyperlink to worksheet created above
        cell.Formula = "=HYPERLINK(""#'" & cell & "'!A1"",""" & cell & """)"
    End If
Next cell
End Sub
Sub yGenerateReport()
'TO DO: CREATE A WAY USING ADBO OR SQL OR PLAIN VBA TO AGGREGATE ALL 1000 AND SERVICE LINES TO SINGLE SHEET FOR REVIEW.
'THIS CAN BE A START
'https://ourcodeworld.com/articles/read/1534/how-to-run-a-sql-query-with-vba-on-excel-spreadsheets-data

End Sub
Sub zCopySheets()
'TO DO: CREATE A WAY TO MERGE BACK FROM REGIONAS FILLING OUT. BELOW IS A SHELL BORROWED FROM ANOTHER PROJECT AS A BASE.
Dim Source As String
Dim Destination As String

Source = "Workbook1.xlsx"
Destination = "Workbook2.xlsm"

Dim Worksheets As Variant
ReDim Worksheets(3)

Worksheets(1) = "January"
Worksheets(2) = "February"
Worksheets(3) = "March"

Dim i As Variant
For i = 1 To UBound(Worksheets)
    Workbooks(Source).Sheets(Worksheets(i)).Copy _
        After:=Workbooks(Destination).Sheets(Workbooks(Destination).Sheets.Count)
Next i

End Sub

Sub zzUnHideCheckColumn()
ThisWorkbook.Sheets("Budget Overview").Columns("R:R").EntireColumn.Hidden = False
End Sub

Sub zzzDeletSheets()
'NEEDS WORK
'TO DO: BE ABLE TO DELETE SHEETS FROM SELECTION ON BUDGET OVERVIEW, CREATE SHEETS BUT IN REVDERSE.
Dim exists As Boolean

For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "MySheet" Then
        exists = True
    End If
Next i

If Not exists Then
    Worksheets.Add.Name = "MySheet"
    Sheets("Data").Delete
End If


End Sub



