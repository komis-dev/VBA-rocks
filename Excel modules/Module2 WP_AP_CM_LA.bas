Option Explicit

Sub Process_MO_Workplan()
    Dim wksht As Worksheet
    Dim rangeval1 As Range

    Set wksht = ThisWorkbook.Worksheets("Workplan")
    Set rangeval1 = wksht.Range("AP7:AP2000")

    With rangeval1.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlEqual, Formula1:="AF,AS,CS,ET,FT,GP,JL,LD,MK,DS,TE,IP"
        .ErrorTitle = "MO Initials"
        .ErrorMessage = "Please enter valid MO Initials"
        .InputTitle = " "
        .InputMessage = " "
        .ShowInput = False
        .ShowError = False
    End With

    wksht.Columns("A:CP").EntireColumn.Hidden = False
    wksht.Columns("A:CP").EntireRow.Hidden = False
    wksht.Cells(6, 8).Select

    Dim wpVariableNames As Variant
    wpVariableNames = Array("wOrgStart", "wBriefDate", "wDeskStart", "wFieldTravelS", "wFieldWorkS", "wFieldWorkE", "wFieldTravelE", _
                            "wDraftsDue", "wDraftsReal", "wQCDue", "wQCReal", "wDebrief", "wDraftReport", "wDraftReportReal", "wDeadlineECDraftComments", "wFinalPlanned", _
                            "wFinalReal", "wDeadlineECComments", "wFieldPhase", "wNoReport", "wNoMQ", "wStatusROM", "wIntTitle", "wType", "wEntity", "wOM", "wImplement", "wImplementType", _
                            "wProjectKE", "wCoreTeam", "wBriefExp", "wOrgStatus", "wNameExp", "wTypeExp", "wOutStatus", "wNameQC", "wTypeQC", "wDebriefExp", "wMonComment", "wQCComment", _
                            "wMissionID", "wContractNo", "wMO", "wCountry", "wStrand", "wDSExpert", "wDelivered", "IW_Start", "IW_Sec", "IW_SecT", "LResult")

    Dim wpData As Dictionary
    Set wpData = New Dictionary

    Dim i As Long
    For i = LBound(wpVariableNames) To UBound(wpVariableNames)
        wpData(wpVariableNames(i)) = Empty
    Next i

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
End Sub

Sub MO()
    Dim wData As Dictionary
    Set wData = New Dictionary
    
    Dim mData(1 To 12) As Dictionary
    Dim i As Integer
    For i = 1 To 12
        Set mData(i) = New Dictionary
    Next i
    
    Dim MO_initials(1 To 12) As Variant
    MO_initials(1) = "ET"
    MO_initials(2) = "JL"
    MO_initials(3) = "AS"
    MO_initials(4) = "FT"
    MO_initials(5) = "GP"
    MO_initials(6) = "LD"
    MO_initials(7) = "MK"
    MO_initials(8) = "CS"
    MO_initials(9) = "AF"
    MO_initials(10) = "TE"
    MO_initials(11) = "IP"
    MO_initials(12) = "DS"

    ' Update Workplan with MO data
Sub UpdateWorkplan(wData As Dictionary, mData As Dictionary, kst As Long)

Dim key As Variant
Dim tempData As Dictionary

' Reference the Workplan worksheet and the MO worksheet
Dim wp As Worksheet
Set wp = ThisWorkbook.Worksheets("Workplan")

' Iterate through the Workplan data
For key = 1 To kst
    ' Check if the Workplan row's MissionID matches the MO row's MissionID
    If wData("row_" & key)("wMissionID") = mData("row_" & key)("mMissionID") Then
        ' Update Workplan data with MO data
        Set tempData = wData("row_" & key)

        With tempData
            .Item("wStatusROM") = mData("row_" & key)("mOrgStatus")
            .Item("wMO") = mData("row_" & key)("mMO")
            .Item("wDSExpert") = mData("row_" & key)("mOM")
            .Item("wOrgStart") = mData("row_" & key)("mOrgStart")
            .Item("wBriefDate") = mData("row_" & key)("mBriefDate")
            .Item("wDeskStart") = mData("row_" & key)("mDeskStart")
            .Item("wFieldTravelS") = mData("row_" & key)("mFieldTravelS")
            .Item("wFieldWorkS") = mData("row_" & key)("mFieldWorkS")
            .Item("wFieldWorkE") = mData("row_" & key)("mFieldWorkE")
            .Item("wFieldTravelE") = mData("row_" & key)("mFieldTravelE")
            .Item("wDraftsDue") = mData("row_" & key)("mDraftsDue")
            .Item("wDraftsReal") = mData("row_" & key)("mDraftsReal")
            .Item("wQCDue") = mData("row_" & key)("mQCDue")
            .Item("wQCReal") = mData("row_" & key)("mQCReal")
            .Item("wDebrief") = mData("row_" & key)("mDebrief")
            .Item("wFinalDue") = mData("row_" & key)("mFinalDue")
            .Item("wFinalReal") = mData("row_" & key)("mFinalReal")
            .Item("wFinalPlanned") = mData("row_" & key)("mFinalPlanned")
            .Item("wFinalDelivered") = mData("row_" & key)("mFinalDelivered")
            .Item("wDelivered") = mData("row_" & key)("mDelivered")
        End With

        ' Update the Workplan data dictionary with the updated row
        wData("row_" & key) = tempData
    End If
Next key

End Sub

' Write Workplan data to the worksheet
Sub WriteWorkplan(wData As Dictionary, kst As Long)

Dim key As Variant
Dim row As Long

' Reference the Workplan worksheet
Dim wp As Worksheet
Set wp = ThisWorkbook.Worksheets("Workplan")

' Write the updated Workplan data to the worksheet
For key = 1 To kst
    row = key + 6
    With wData("row_" & key)
        wp.Cells(row, 3).Value = .Item("wStrand")
        wp.Cells(row, 4).Value = .Item("wStatusROM")
        wp.Cells(row, 5).Value = .Item("wMO")
        wp.Cells(row, 6).Value = .Item("wDSExpert")
        wp.Cells(row, 7).Value = .Item("wOrgStart")
        wp.Cells(row, 8).Value = .Item("wBriefDate")
        wp.Cells(row, 9).Value = .Item("wDeskStart")
        wp.Cells(row, 10).Value = .Item("wFieldTravelS")
        wp.Cells(row, 11).Value = .Item("wFieldWorkS")
        wp.Cells(row, 12).Value = .Item("wFieldWorkE")
        wp.Cells(row, 13).Value = .Item("wFieldTravelE")
        wp.Cells(row, 14).Value = .Item("wDraftsDue")
        wp.Cells(row, 15).Value = .Item("wDraftsReal")
        wp.Cells(row, 16).Value = .Item("wQCDue")
        wp.Cells(row, 17).Value = .Item("wQCReal")
        wp.Cells(row, 18).Value = .Item("wDebrief")
        wp.Cells(row, 19).Value = .Item("wFinalDue")
        wp.Cells(row, 20).Value = .Item("wFinalReal")
        wp.Cells(row, 21).Value = .Item("wFinalPlanned")
        wp.Cells(row, 22).Value = .Item("wFinalDelivered")
        wp.Cells(row, 23).Value = .Item("wDelivered")
    End With
Next key

End Sub

' Read Workplan data from the worksheet
Sub ReadWorkplan(ByRef wData As Dictionary, ByVal kst As Long)
    Dim wp As Worksheet
    Set wp = ThisWorkbook.Worksheets("Workplan")
    
    Dim row As Long
    Dim key As String
    
    For row = 2 To kst
        key = wp.Cells(row, 1).Value
        If Not wData.Exists(key) Then
            wData.Add key, New Dictionary
            With wData(key)
                .Add "wProject", wp.Cells(row, 2).Value
                .Add "wMO", wp.Cells(row, 3).Value
                .Add "wDS", wp.Cells(row, 4).Value
                .Add "wDSEmail", wp.Cells(row, 5).Value
                .Add "wDSExpert", wp.Cells(row, 6).Value
                .Add "wOrgStart", wp.Cells(row, 7).Value
                .Add "wBriefDate", wp.Cells(row, 8).Value
                .Add "wDeskStart", wp.Cells(row, 9).Value
                .Add "wFieldTravelS", wp.Cells(row, 10).Value
                .Add "wFieldWorkS", wp.Cells(row, 11).Value
                .Add "wFieldWorkE", wp.Cells(row, 12).Value
                .Add "wFieldTravelE", wp.Cells(row, 13).Value
                .Add "wDraftsDue", wp.Cells(row, 14).Value
                .Add "wDraftsReal", wp.Cells(row, 15).Value
                .Add "wQCDue", wp.Cells(row, 16).Value
                .Add "wQCReal", wp.Cells(row, 17).Value
                .Add "wDebrief", wp.Cells(row, 18).Value
                .Add "wFinalDue", wp.Cells(row, 19).Value
                .Add "wFinalReal", wp.Cells(row, 20).Value
                .Add "wFinalPlanned", wp.Cells(row, 21).Value
                .Add "wFinalDelivered", wp.Cells(row, 22).Value
                .Add "wDelivered", wp.Cells(row, 23).Value
            End With
        End If
    Next row
End Sub

' ReadMO() subroutine reads the data from the currently opened MO workbook.
Sub ReadMO(mData As Dictionary, kst As Long, MOWorkbook As Workbook)
    Dim mo As Worksheet
    Set mo = ThisWorkbook.Worksheets("MO")

    Dim row As Long
    Dim key As String

    For row = 2 To kst
        key = mo.Cells(row, 1).Value 
        If Not mData.Exists(key) Then
            mData.Add key, New Dictionary
            With mData(key)
                ' Add MO data for each column
                .Add "mMissionID", mo.Cells(row, 1).Value
                .Add "mOrgStatus", mo.Cells(row, 2).Value
                .Add "mMO", mo.Cells(row, 3).Value
                .Add "mDS", mo.Cells(row, 4).Value
                .Add "mDSEmail", mo.Cells(row, 5).Value
                .Add "mDSExpert", mo.Cells(row, 6).Value
                .Add "mOrgStart", mo.Cells(row, 7).Value
                .Add "mBriefDate", mo.Cells(row, 8).Value
                .Add "mDeskStart", mo.Cells(row, 9).Value
                .Add "mFieldTravelS", mo.Cells(row, 10).Value
                .Add "mFieldWorkS", mo.Cells(row, 11).Value
                .Add "mFieldWorkE", mo.Cells(row, 12).Value
                .Add "mFieldTravelE", mo.Cells(row, 13).Value
                .Add "mDraftsDue", mo.Cells(row, 14).Value
                .Add "mDraftsReal", mo.Cells(row, 15).Value
                .Add "mQCDue", mo.Cells(row, 16).Value
                .Add "mQCReal", mo.Cells(row, 17).Value
                .Add "mDebrief", mo.Cells(row, 18).Value
                .Add "mFinalDue", mo.Cells(row, 19).Value
                .Add "mFinalReal", mo.Cells(row, 20).Value
                .Add "mFinalPlanned", mo.Cells(row, 21).Value
                .Add "mFinalDelivered", mo.Cells(row, 22).Value
                .Add "mDelivered", mo.Cells(row, 23).Value
            End With
        End If
    Next row
End Sub

'  UpdateWorkplanMain() subroutine continues processing the data and then closes the currently opened MO workbook.    
Sub UpdateWorkplanMain()

    Dim MO_initials(1 To 12) As Variant
    MO_initials(1) = "ET"
    MO_initials(2) = "JL"
    MO_initials(3) = "AS"
    MO_initials(4) = "FT"
    MO_initials(5) = "GP"
    MO_initials(6) = "LD"
    MO_initials(7) = "MK"
    MO_initials(8) = "CS"
    MO_initials(9) = "AF"
    MO_initials(10) = "TE"
    MO_initials(11) = "IP"
    MO_initials(12) = "DS"

    ' moves on to the next MO file and repeats the process
    Dim i As Long
    For i = 1 To 12
        ' Open the MO file with the current MO initials
        Dim MOWorkbook As Workbook
        Set MOWorkbook = Workbooks.Open("Path\to\your\folder\MO_" & MO_initials(i) & ".xlms")

        ' Declare and initialize variables
        Dim wData As Dictionary
        Set wData = New Dictionary

        Dim mData As Dictionary
        Set mData = New Dictionary

        Dim kst As Long
        kst = GetRowCount("Workplan")

        ' Read Workplan and MO data
        ReadWorkplan wData, kst
        ReadMO mData, kst, MOWorkbook

        ' Update Workplan with MO data
        UpdateWorkplan wData, mData, kst

        ' Write updated Workplan data to the worksheet
        WriteWorkplan wData, kst

        ' Close the MO workbook without saving changes
        MOWorkbook.Close SaveChanges:=False
    Next i

    ' Inform the user that the update is complete
    MsgBox "Workplan has been successfully updated!", vbInformation, "Update Complete"

End Sub