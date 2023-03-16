Sub Transfer_WPtoMO()
    ' Declare variables for Start time, array of dates and counting seconds
    Dim StartTime As Double
    Dim LResult(1 To 6) As Date
    Dim CountingSeconds As Long

    ' Set the start time of the macro
    StartTime = Timer

    ' Temporarily turn off screen updating and manual calculation
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual

        ' Call subroutine IW and TWM
        Call IW(LResult)
        Call TWM

        ' Re-enable screen updating and automatic calculation
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With

    ' Calculate the total time of the macro
    CountingSeconds = Round(Timer - StartTime, 2)

    ' Display the total time in a message box
    MsgBox "Total time: " & CountingSeconds & " sec" & Chr(10), vbInformation
End Sub

Sub TWM()
    ' Declare variables for workbook, worksheets, and ranges
    Dim wb As Workbook, ws_data As Worksheet, ws_criteria As Worksheet, ws_output As Worksheet
    Dim rgData As Range, rgCriteria As Range, rgOutput As Range
    Dim Row_Last As Long

    ' Set the workbook to the current workbook
    Set wb = ThisWorkbook

    ' Set the worksheets to the relevant worksheets in the workbook
    Set ws_data = wb.Worksheets("Update Workplan")
    Set ws_criteria = wb.Worksheets("Dashboard")
    Set ws_output = wb.Worksheets("MO")

    ' Set the data range in the "Update Workplan" worksheet
    With ws_data
        Set rgData = .Range("A5:CX" & .Cells(.Rows.Count, 1).End(xlUp).Row)
    End With

    ' Set the criteria range in the "Dashboard" worksheet
    Set rgCriteria = ws_criteria.Range("A12:A13")

    ' Set the output range in the "MO" worksheet
    Set rgOutput = ws_output.Range("A4:bj4")

    ' Show all data in the worksheets if they are in filter mode
    If ws_data.FilterMode Then ws_data.ShowAllData
    If ws_output.FilterMode Then ws_output.ShowAllData

    ' Clear any previous data in the output worksheet
    With ws_output
        Row_Last = .Cells(.Rows.Count, 1).End(xlUp).Row
        If Row_Last > 4 Then .Rows("5:" & Row_Last).Delete
    End With

    ' Copy data from the "Update Workplan" worksheet to the "MO" worksheet based on the criteria
    rgData.AdvancedFilter xlFilterCopy, rgCriteria, rgOutput

    ' Call the ApplyCF subroutine to apply conditional formatting to the data in the "MO" worksheet
    ApplyCF ws_output.Range("V5:V" & Row_Last), _
        Array("Waiting", "Ongoing", "Confirmed", "Implemented", "Submitted"), _
        Array(RGB(244, 176, 132), RGB(201, 201, 201), RGB(255, 217, 102), RGB(169, 208, 142), RGB(155, 194, 230))

    ApplyCF ws_output.Range("AF5:AF" & Row_Last), _
            Array("Drafts to BO", "QC Finalised", "Draft Outputs Submitted", "EC Comments Received", "Final Outputs Submitted", "Preliminary Drafts submitted"), _
            Array(RGB(244, 176, 132), RGB(201, 201, 201), RGB(255, 217, 102), RGB(169, 208, 142), RGB(155, 194, 230), RGB(125, 50, 50))

    ' Update the timestamp in cell B2 of the "MO" worksheet
    ws_output.Range("B2") = Now()

End Sub

Sub ApplyCF(ByVal Target As Range, StatusList As Variant, ColorList As Variant)
    ' Declare a variable for the loop counter
    Dim i As Long
    ' Delete any previous format conditions
    Target.FormatConditions.Delete

        ' Loop through the status values and colors
        For i = LBound(StatusList) To UBound(StatusList)
        ' Add a new format condition for each status value
        With Target.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & StatusList(i) & """")
        ' Set the interior color and font color
        .Interior.Color = ColorList(i)
        .Font.ColorIndex = vbBlack
        End With
    Next i
End Sub