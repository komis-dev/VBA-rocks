Option Explicit

Sub Process_MO_Workplan()
    Dim wksht As Worksheet
    Dim rangeval1 As Range

    ' MO_Validation
    Set wksht = ThisWorkbook.Worksheets("Workplan")
    Set rangeval1 = wksht.Range("AP7:AP2000")

    With rangeval1.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlEqual, Formula1:="AF,AS,CS,ET,FT,GP,JL,LD,MK,NR,TE,IP"
        .ErrorTitle = "MO Initials"
        .ErrorMessage = "Please enter valid MO Initials"
        .InputTitle = " "
        .InputMessage = " "
        .ShowInput = False
        .ShowError = False
    End With

    ' Unhide_All
    With wksht
        .Columns("A:CP").EntireColumn.Hidden = False
        .Columns("A:CP").EntireRow.Hidden = False
        .Cells(6, 8).Select
    End With

    ' WP variables
    Dim wOrgStart(), wBriefDate() As Date
	Dim wDeskStart(), wFieldTravelS(), wFieldWorkS(), wFieldWorkE(), wFieldTravelE() As Date
	Dim wDraftsDue(), wDraftsReal(), wQCDue(), wQCReal(), wDebrief() As Date
	Dim wDraftReport(), wDraftReportReal(), wDeadlineECDraftComments(), wFinalPlanned() As Date
	Dim wFinalReal(), wDeadlineECComments() As Date

	Dim wFieldPhase(), wNoReport(), wNoMQ() As Integer

	Dim wStatusROM(), wIntTitle(), wType(), wEntity() As String
	Dim wOM(), wImplement(), wImplementType(), wProjectKE() As String
	Dim wCoreTeam(), wBriefExp(), wOrgStatus() As String
	Dim wNameExp(), wTypeExp(), wOutStatus(), wNameQC(), wTypeQC() As String
	Dim wDebriefExp(), wMonComment(), wQCComment() As String
	Dim wMissionID(), wContractNo(), wMO(), wCountry() As String
	Dim wStrand(), wDSExpert() As String
	Dim wDelivered() As Variant

	Dim IW_Start() As Double
	Dim IW_Sec(), IW_SecT() As Long
	Dim LResult() As Date

	With Application
		.ScreenUpdating = False
		.Calculation = xlCalculationManual
	End With

	' Set the start time of the macro
	Dim defaultTime As Date
	defaultTime = TimeValue("00:00:00")

	Dim dateVariables As Object
	Set dateVariables = CreateObject("Scripting.Dictionary")

	
	dateVariables.Add "wOrgStart", defaultTime
	dateVariables.Add "wBriefDate", defaultTime
	dateVariables.Add "wDeskStart", defaultTime
	dateVariables.Add "wFieldTravelS", defaultTime
	dateVariables.Add "wFieldWorkS", defaultTime
	dateVariables.Add "wFieldWorkE", defaultTime
	dateVariables.Add "wFieldTravelE", defaultTime
	dateVariables.Add "wDraftsDue", defaultTime
	dateVariables.Add "wDraftsReal", defaultTime
	dateVariables.Add "wQCDue", defaultTime
	dateVariables.Add "wQCReal", defaultTime
	dateVariables.Add "wDebrief", defaultTime
	dateVariables.Add "wDraftReport", defaultTime
	dateVariables.Add "wDraftReportReal", defaultTime
	dateVariables.Add "wDeadlineECDraftComments", defaultTime
	dateVariables.Add "wFinalPlanned", defaultTime
	dateVariables.Add "wFinalReal", defaultTime
	dateVariables.Add "wDeadlineECComments", defaultTime
	dateVariables.Add "wQCDue", defaultTime
	mDateUpdated = TimeValue("00:00:00")



    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    ' Perform the required operations here...

    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With

End Sub

Sub MO()

    ' Workplan variables
    Dim wData As Dictionary
    Set wData = New Dictionary
    
    ' MO file variables
    Dim mData(12) As Dictionary
    Dim i As Integer
    For i = 1 To 12
        Set mData(i) = New Dictionary
    Next i
    
    ' General variables
    Dim LRWorkplan As Long
    Dim kst As Long
    
    ' MO initials
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

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    ' Read Workplan data
    Call ReadWorkplan(wData, kst, LRWorkplan)

    ' Update Workplan with MO data
    Dim MO_i As Variant
    Dim IW_i2 As Integer
    IW_i2 = 1
    For Each MO_i In MO_initials
        Call ReadMO(mData(IW_i2), MO_i)
        Call UpdateWorkplan(wData, mData(IW_i2), kst, CountMO(mData(IW_i2), MO_initials(IW_i2)))
        IW_i2 = IW_i2 + 1
    Next MO_i

    ' Write Workplan data to the worksheet
    Call WriteWorkplan(wData, kst)

    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With

    MsgBox "The Workplan PP is updated."

End Sub

' Read Workplan data
Public Sub ReadWorkplan(ByRef wData As Dictionary, ByRef kst As Long, ByRef LRWorkplan As Long)

    Dim MyFile As String
    Dim i As Long
    Dim key As String

    MyFile = ThisWorkbook.Name
    Application.ScreenUpdating = False

    LRWorkplan = Worksheets("Workplan").Cells.SpecialCells(xlCellTypeLastCell).Row

    ' Read workplan into dictionary
    kst = 0

    For i = 7 To LRWorkplan
        kst = kst + 1
        key = "row_" & kst

        wData(key) = New Dictionary
        With wData(key)
            .Item("wStrand") = Worksheets("Workplan").Cells(i, 3).Value
            .Item("wStatusROM") = Worksheets("Workplan").Cells(i, 4).Value
            .Item("wMissionID") = Worksheets("Workplan").Cells(i, 5).Value
            .Item("wContractNo") = Worksheets("Workplan").Cells(i, 6).Value
			.Item("wCountry") = Worksheets("Workplan").Cells(i, 7).Value
			.Item("wMO") = Worksheets("Workplan").Cells(i, 8).Value
			.Item("wDSExpert") = Worksheets("Workplan").Cells(i, 9).Value
			.Item("wOrgStart") = Worksheets("Workplan").Cells(i, 10).Value
			.Item("wBriefDate") = Worksheets("Workplan").Cells(i, 11).Value
			.Item("wDeskStart") = Worksheets("Workplan").Cells(i, 12).Value
			.Item("wFieldTravelS") = Worksheets("Workplan").Cells(i, 13).Value
			.Item("wFieldWorkS") = Worksheets("Workplan").Cells(i, 14).Value
			.Item("wFieldWorkE") = Worksheets("Workplan").Cells(i, 15).Value
			.Item("wFieldTravelE") = Worksheets("Workplan").Cells(i, 16).Value
			.Item("wDraftsDue") = Worksheets("Workplan").Cells(i, 17).Value
			.Item("wDraftsReal") = Worksheets("Workplan").Cells(i, 18).Value
			.Item("wQCDue") = Worksheets("Workplan").Cells(i, 19).Value
			.Item("wQCReal") = Worksheets("Workplan").Cells(i, 20).Value
			.Item("wDebrief") = Worksheets("Workplan").Cells(i, 21).Value
			.Item("wFinalDue") = Worksheets("Workplan").Cells(i, 22).Value
			.Item("wFinalReal") = Worksheets("Workplan").Cells(i, 23).Value
			.Item("wFinalPlanned") = Worksheets("Workplan").Cells(i, 24).Value
			.Item("wFinalDelivered") = Worksheets("Workplan").Cells(i, 25).Value
            .Item("wDelivered") = Worksheets("Workplan").Cells(i, 102).Value
        End With
    Next i

    Application.ScreenUpdating = True

End Sub

' Read MO data
Sub ReadMO(mData As Dictionary, MO_initial As String)
	Dim moWorksheet As Worksheet
	Dim lastRow As Long
	Dim i As Long
	Dim key As String

	' Set the appropriate worksheet based on the MO_initial
	Set moWorksheet = ThisWorkbook.Worksheets(MO_initial)

	' Find the last row with data
	lastRow = moWorksheet.Cells(moWorksheet.Rows.Count, "A").End(xlUp).Row

	' Read MO data into the dictionary
	For i = 2 To lastRow
		key = "row_" & i - 1

		mData(key) = New Dictionary
		With mData(key)
			.Item("mMissionID") = moWorksheet.Cells(i, 3).Value
			.Item("mContractNo") = moWorksheet.Cells(i, 4).Value
			.Item("mMO") = moWorksheet.Cells(i, 11).Value
			.Item("mCountry") = moWorksheet.Cells(i, 19).Value
			.Item("mOM") = moWorksheet.Cells(i, 8).Value
			.Item("mImplement") = moWorksheet.Cells(i, 9).Value
			.Item("mDelivered") = moWorksheet.Cells(i, 17).Value
			.Item("mOrgStart") = moWorksheet.Cells(i, 21).Value
			.Item("mBriefExp") = moWorksheet.Cells(i, 24).Value
			.Item("mBriefDate") = moWorksheet.Cells(i, 25).Value
			.Item("mOrgStatus") = moWorksheet.Cells(i, 22).Value
			.Item("mOutStatus") = moWorksheet.Cells(i, 32).Value
			.Item("mDraftsReal") = moWorksheet.Cells(i, 36).Value
			.Item("mQCReal") = moWorksheet.Cells(i, 38).Value
			.Item("mDebrief") = moWorksheet.Cells(i, 30).Value
			.Item("mDebriefExp") = moWorksheet.Cells(i, 31).Value
			.Item("mDraftReportReal") = moWorksheet.Cells(i, 40).Value
			.Item("mFinalReal") = moWorksheet.Cells(i, 44).Value
			.Item("mDeskStart") = moWorksheet.Cells(i, 23).Value
			.Item("mDateUpdated") = moWorksheet.Cells(i, 7).Value
			.Item("mFinalPlanned") = moWorksheet.Cells(i, 26).Value
			.Item("mFinalDelivered") = moWorksheet.Cells(i, 27).Value
			.Item("mDelivered") = moWorksheet.Cells(i, 28).Value

		End With
	Next i

End Sub

' Update Workplan with MO data
Sub UpdateWorkplan(wData As Dictionary, mData As Dictionary, kst As Long)
    
End Sub

' Workplan data to the worksheet
Sub WriteWorkplan(wData As Dictionary, kst As Long)
    
End Sub

' Count the occurrences of a specific MO
Function CountMO(mData As Dictionary, MO_initial As String) As Long
    Dim moCount As Long
    Dim key As Variant

    moCount = 0
    For Each key In mData
        If mData(key) = MO_initial Then
            moCount = moCount + 1
        End If
    Next key
    CountMO = moCount
End Function