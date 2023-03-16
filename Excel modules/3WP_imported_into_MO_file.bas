' The code provided consists of several subroutines and functions that work together to import workplans from multiple Excel files stored on SharePoint into a single workbook. The main subroutine Import_Workplans is responsible for executing the overall process. Here's a brief overview of the code's purpose and functionality:

' Import_Workplans: This subroutine stops screen updating and automatic calculation for speed gains, calls the IW subroutine to import workplans, and then re-establishes screen updating and automatic calculation.

' IW: This subroutine is responsible for preparing the "Update workplan" sheet for data import, iterating through the workbooks specified in the "Dashboard" sheet, calling the IW_ImportWP subroutine for each workbook, and finally formatting the "Update workplan" sheet.

' IW_ImportWP: This subroutine calls a series of other subroutines (Date_File, OpenWB, Copy_wb, and CloseWB) to import data from workbooks specified in the "Dashboard" sheet.

' Date_File: Retrieves the last modified date of a file located at a given SharePoint path.

' OpenWB: Opens an Excel workbook stored on SharePoint.

' Copy_wb: Copies a range of data from the "Workplan" sheet of one workbook to the "Update workplan" sheet of the main workbook.

' CloseWB: Closes the specified Excel workbook without saving changes.

' SharePointURLtoUNC: Converts a SharePoint URL to a UNC (Uniform Naming Convention) path, which is used to access the Excel files on SharePoint.

' The code is designed to import workplans from various workbooks and consolidate them into a single workbook. This is done by iterating through the specified workbooks, opening them, copying relevant data, and pasting it into the main workbook. After importing the data, the code formats the "Update workplan" sheet and calculates the date of updates.

'Press button "Import Workplans" in another document's worksheet "Update workplan"
Sub Import_Workplans()
    Dim IW_StartTime As Double
    Dim IW_Sec As Long
    Dim LResult(1 To 6) As Date
    IW_StartTime = Timer
    'Stop updating screen and automatic calculation (for speed gains)
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    'Call subroutine for Importing Workplans and return the time needed to do so
    Call IW(LResult)
    'Re-establish updating screen and automatic calculation
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub
' Import workplans to the other file
Sub IW(LResult)
   ' This sub is named IW and takes one parameter, LResult
   Sub IW(LResult)
   
   ' Declare the variables as Variant, Date, and Integer data types
   Dim Second_workbook, wbpath As Variant
   Dim Ldate As Date
   Dim Dash_row(1 To 3) As Integer
   Dim IW_i, IW_i_R As Integer
   
   ' Initialize the variable IW_i_R to 1
   IW_i_R = 1
   
   ' Set the values of three elements in the array Dash_row to 6, 7, and 8 respectively
   Dash_row(1) = 6
   Dash_row(2) = 7
   Dash_row(3) = 8
   
   ' Store name of the workbook that contains this macro in the variable Second_workbook
   Second_workbook = ThisWorkbook.Name
   
   ' Remove filters from sheet "Update workplan"
   With Workbooks(Second_workbook).Sheets("Update workplan")
       If .AutoFilterMode Then
           .AutoFilterMode = False
       End If
   End With
   
   ' Delete contents of the sheet "Update workplan" from row A6 to the last used row in column CW
   With Workbooks(Second_workbook).Sheets("Update workplan")
       .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
   End With
   
   ' Store the Sharepoint path of the workbooks in wbpath variable
   wbpath = Workbooks(Second_workbook).Sheets("Dashboard").Cells(5, 2).Text
   
   ' Call a function named IW_ImportWP for each element in the Dash_row array
   For Each IW_i In Dash_row
       Call IW_ImportWP(Second_workbook, wbpath, IW_i, LResult(IW_i_R))
       IW_i_R = IW_i_R + 1
   Next IW_i
   
   ' Turn off filters on the sheet "Update workplan"
   Workbooks(Second_workbook).Sheets("Update workplan").AutoFilterMode = False  
    
    'The first section of the code formats cells in the "Update workplan" sheet up to the last filled cell:
    'It sets the row height, applies border styles and font formatting, and centers the data.
    With Workbooks(Second_workbook).Sheets("Update workplan")
        .Range("A5:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).AutoFilter
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).RowHeight = 25
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Borders.LineStyle = xlContinuous
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Borders.ThemeColor = 1
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Borders.TintAndShade = -0.499984740745262
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Borders.Weight = xlThin
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Font.Name = "Calibri"
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).Font.Size = 8
        .Range("A6:CW" & .Cells(.Rows.Count, 1).End(xlUp).Row).HorizontalAlignment = xlCenter
    End With
    
    'The second section calculates the date of updates and writes it to cell (1,4) in the "Update workplan" sheet:
    ' Calculate Formula date of updates 
    Ldate = Now()
    With Workbooks(Second_workbook).Sheets("Update workplan")
        .Cells(1, 4).Value = "Last update:" & Ldate
    End With
End Sub

'The sub IW_ImportWP is defined to import three workbooks specified in the "Dashboard" sheet:
'It takes four inputs: Second_workbook, wbpath, IW_Dash_row_WP, and LResult_WP. 
'It uses the Call keyword to execute three other subroutines: Date_File, OpenWB, and Copy_wb before calling CloseWB. 
'The purpose of these subroutines is not provided in the code, but they are used to open and copy sheets from other workbooks.
Sub IW_ImportWP(Second_workbook, wbpath, IW_Dash_row_WP, LResult_WP)
    Dim IW_Name_WB_WP, IW_Name_WB_WP_path As String
    Dim FilePAth As String
    IW_Name_WB_WP = Workbooks(Second_workbook).Sheets("Dashboard").Cells(IW_Dash_row_WP, 2).Text
    IW_Name_WB_WP_path = wbpath & IW_Name_WB_WP
    FilePAth = wbpath
    Call Date_File(FilePAth, IW_Name_WB_WP, LResult_WP)
    Call OpenWB(IW_Name_WB_WP, IW_Name_WB_WP_path)
    Call Copy_wb(Second_workbook, IW_Name_WB_WP)
    Call CloseWB(IW_Name_WB_WP)
End Sub

' This subroutine retrieves the last modified date of a file located at a given SharePoint path.
Sub Date_File(File_Path, File_Name, LResult)
' Declare variables
    Dim New_Path As String
    Dim filespec As String
    Dim fs, f As Variant

   ' Error handling
On Error GoTo 20

' Convert SharePoint URL to UNC path
New_Path = SharePointURLtoUNC(File_Path)

' Combine UNC path and file name
filespec = New_Path & File_Name

' Create FileSystemObject and get file reference
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFile(filespec)

' Get the file's last modified date
LResult = f.DateLastModified

' Jump to line 30
GoTo 30
' Error handling section
20 LResult = "00:00:00"

' Exit section
30
End Sub

' This subroutine opens an Excel workbook, at a specified path (later on in the code, the path is a SharePoint path)
Sub OpenWB(WBFilename, WBFilename_path)
   ' Error handling
    On Error GoTo 10

    'Open the workbook at the specified path
    Workbooks.Open Filename:=WBFilename_path, Password:="", UpdateLinks:=0, ReadOnly:=True
Exit Sub

' Error handling section
10 filetoopen = Application.GetOpenFilename("Excel files (.xl), .xls")
    ' Check if the user canceled the open file dialog
    If filetoopen = False Then End

    ' Open the selected workbook
    Workbooks.Open filetoopen, Password:="", UpdateLinks:=0, ReadOnly:=True
    BFilename = ActiveWorkbook.Name
End Sub

' This subroutine copies a range of data from the "Workplan" sheet of one of the WP to the "Update workplan" sheet of the MO file.
Sub Copy_wb(Second_workbook, WBFilename)
    ' Copy the data from the source workbook's "Workplan" sheet
    With Workbooks(WBFilename).Sheets("Workplan")
        .AutoFilterMode = False
        .Columns("A:CW").EntireColumn.Hidden = False
        .Range("A7:CW" & Workbooks(WBFilename).Sheets("Workplan").Cells(.Rows.Count, 1).End(xlUp).Row).Copy
    End With

   ' Paste the copied data to the destination workbook's "Update workplan" sheet
    With Workbooks(Second_workbook).Sheets("Update workplan")
        .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteFormats
        .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
    End With
End Sub

' This subroutine closes the specified Excel workbook without saving changes.
Sub CloseWB(WBFilename)
    Workbooks.Application.CutCopyMode = False
    Workbooks(WBFilename).Close SaveChanges:=False
    Workbooks.Application.CutCopyMode = True
End Sub

' This function converts a SharePoint URL to a UNC path.
Public Function SharePointURLtoUNC(sURL)
    Dim bIsSSL As Boolean
        ' Check if the URL uses HTTPS (SSL)
        bIsSSL = InStr(1, sURL, "https:") > 0

        'Replace forward slashes with backslashes, and replace "%20" with spaces
        sURL = Replace(Replace(sURL, "/", "\"), "%20", " ")
        ' Remove "https:" or "http:" from the URL
        sURL = Replace(Replace(sURL, "https:", vbNullString), "http:", vbNullString)

        ' Add "@SSL\DavWWWRoot" to the server name portion of the URL
        sURL = Replace(sURL, Split(sURL, "\")(2), Split(sURL, "\")(2) & "@SSL\DavWWWRoot")

        ' If the URL does not use SSL, remove "@SSL\" from the UNC path
        If Not bIsSSL Then sURL = Replace(sURL, "@SSL\", vbNullString)

        ' Return the converted UNC path
        SharePointURLtoUNC = sURL
End Function

 