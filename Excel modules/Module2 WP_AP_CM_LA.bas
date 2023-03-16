Sub TP()

 Call Unhide_All
Columns("A:C").Select
Selection.EntireColumn.Hidden = True
Columns("G").Select
Selection.EntireColumn.Hidden = True
Columns("I:O").Select
Selection.EntireColumn.Hidden = True
Columns("Q:T").Select
Selection.EntireColumn.Hidden = True
Columns("V:AL").Select
Selection.EntireColumn.Hidden = True
Columns("AO").Select
Selection.EntireColumn.Hidden = True
Columns("AU:AV").Select
Selection.EntireColumn.Hidden = True
Columns("BK:BK").Select
Selection.EntireColumn.Hidden = True
Columns("BM:BM").Select
Selection.EntireColumn.Hidden = True
Columns("BO").Select
Selection.EntireColumn.Hidden = True
Columns("BQ").Select
Selection.EntireColumn.Hidden = True
Columns("BU:CA").Select
Selection.EntireColumn.Hidden = True
Cells(6, 21).Select
End Sub

Sub Filter_Off()
'
' Unlock_Filters Macro
'
    Range("A6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Cells(6, 8).Select

End Sub

Sub MO_Validation()
    Dim wksht As Worksheet
    Dim rangeval1 As Range

    Set wksht = ThisWorkbook.Worksheets("Workplan")
    Set rangeval1 = wksht.Range("AP7:AP2000")

    With rangeval1.Validation
        .Delete
        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlEqual, Formula1:="AF,AS,CS,ET,FT,GP,JL,LD,MK,NR,TE,IP", Formula2:=""
            .ErrorTitle = "MO Initials"
            .ErrorMessage = "Please enter valid MO Initials"
            .InputTitle = " "
            .InputMessage = " "
            .ShowInput = False
            .ShowError = False
    End With

End Sub

Sub Unhide_All()
'
' Unhide Columns
'
    Columns("A:CP").Select
    Selection.EntireColumn.Hidden = False
    Selection.EntireRow.Hidden = False
    Cells(6, 8).Select

End Sub

Sub MO()

'Workplan variables

Dim wOrgStart(5000), wBriefDate(5000) As Date
Dim wDeskStart(5000), wFieldTravelS(5000), wFieldWorkS(5000), wFieldWorkE(5000), wFieldTravelE(5000) As Date
Dim wDraftsDue(5000), wDraftsReal(5000), wQCDue(5000), wQCReal(5000), wDebrief(5000) As Date
Dim wDraftReport(5000), wDraftReportReal(5000), wDeadlineECDraftComments(5000), wFinalPlanned(5000) As Date
Dim wFinalReal(5000), wDeadlineECComments(5000) As Date

Dim wFieldPhase(5000), wNoReport(5000), wNoMQ(5000) As Integer

Dim wStatusROM(5000), wIntTitle(5000), wType(5000), wEntity(5000) As String
Dim wOM(5000), wImplement(5000), wImplementType(5000), wProjectKE(5000) As String
Dim wCoreTeam(5000), wBriefExp(5000), wOrgStatus(5000) As String
Dim wNameExp(5000), wTypeExp(5000), wOutStatus(5000), wNameQC(5000), wTypeQC(5000) As String
Dim wDebriefExp(5000), wMonComment(5000), wQCComment(5000) As String
Dim wMissionID(5000), wContractNo(5000), wMO(5000), wCountry(5000) As String
Dim wStrand(5000), wDSExpert(5000) As String
Dim wDelivered(5000) As Variant

Dim IW_Start(12) As Double
Dim IW_Sec(12) As Long
Dim IW_SecT As Long
Dim LResult(12) As Date

With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
End With

wOrgStart(5000) = "00:00:00"
wBriefDate(5000) = "00:00:00"
wDeskStart(5000) = "00:00:00"
wFieldTravelS(5000) = "00:00:00"
wFieldWorkS(5000) = "00:00:00"
wFieldWorkE(5000) = "00:00:00"
wFieldTravelE(5000) = "00:00:00"
wDraftsDue(5000) = "00:00:00"
wDraftsReal(5000) = "00:00:00"
wQCDue(5000) = "00:00:00"
wQCReal(5000) = "00:00:00"
wDebrief(5000) = "00:00:00"
wDraftReport(5000) = "00:00:00"
wDraftReportReal(5000) = "00:00:00"
wDeadlineECDraftComments(5000) = "00:00:00"
wFinalPlanned(5000) = "00:00:00"
wFinalReal(5000) = "00:00:00"
wDeadlineECComments(5000) = "00:00:00"
wQCDue(5000) = "00:00:00"

'General variables

Dim LRWorkplan As Integer

Dim kst As Integer

'MO File variables

Dim mOrgStart(12, 5000), mBriefDate(12, 5000) As Date
Dim mDraftsReal(12, 5000), mDebrief(12, 5000) As Date
Dim mDraftReportReal(12, 5000) As Date
Dim mFinalReal(12, 5000), mQCReal(12, 5000), mDeskStart(12, 5000) As Date
Dim mDateUpdated As Date
Dim mDelivered(12, 5000) As Variant

Dim mImplement(12, 5000), mImplementType(12, 5000), mOutStatus(12, 5000) As String
Dim mOrgStatus(12, 5000), mOM(12, 5000), mBriefExp(12, 5000) As String
Dim mDebriefExp(12, 5000), mMonComment(12, 5000), mQCComment(12, 5000) As String
Dim mMissionID(12, 5000), mContractNo(12, 5000), mMO(12, 5000), mCountry(12, 5000) As String
'
'mOrgStart(5000) = "00:00:00"
'mBriefDate(5000) = "00:00:00"
'mDraftsReal(5000) = "00:00:00"
'mDebrief(5000) = "00:00:00"
'mDraftReportReal(5000) = "00:00:00"
'mFinalReal(5000) = "00:00:00"
'mQCReal(5000) = "00:00:00"
'mDeskStart(5000) = "00:00:00"
mDateUpdated = "00:00:00"

'General variables

Dim LRMO As Integer
Dim kwt As Integer

Dim Counter1, Counter2, Counter3, XX As Integer

Counter1 = 0
Counter2 = 0
Counter3 = 0
XX = 0

Dim StartTime As Double
Dim SecondsElapsed As Double
'Remember time when macro starts

'Stop updating screen

StartTime = Timer
With Application
        .ScreenUpdating = False
End With

'Application.EnableEvents = False

'MO file location
fpath = "https://komisbrussels.sharepoint.com/sites/operations/ROMGlobal/A0_TQM/A0_03_Operations/T0_03.1_Project%20Planning/Mission%20Organisers/"

'Desk study file location
DSFile = "https://komisbrussels.sharepoint.com/sites/operations/ROMGlobal/A0_TQM/A0_03_Operations/T0_03.1_Project%20Planning/Workplan_Updates/Desk_Study_PP.xlsm"
Call ReadWorkplan(wStrand, wStatusROM, wMissionID, wContractNo, wIntTitle, wType, wEntity, wOM, wImplement, wImplementType, wProjectKE, wCoreTeam, wMO, wOrgStart, wBriefExp, wBriefDate, wOrgStatus, wDeskStart, wCountry, wFieldTravelS, wFieldWorkS, wFieldWorkE, wFieldTravelE, wNameExp, wTypeExp, wOutStatus, wNameQC, wTypeQC, wNoReport, wNoMQ, wDraftsDue, wDraftsReal, wQCDue, wQCReal, wDebrief, wDebriefExp, wDraftReport, wDraftReportReal, wDeadlineECDraftComments, wFinalPlanned, wFinalReal, wDeadlineECComments, wMonComment, wQCComment, wFieldPhase, wDelivered, kst, LRWorkplan)

Application.ScreenUpdating = False

Call MOView

'Read MO Files
Dim fNameAndPath As Variant
Dim thisfname As Variant
Dim Current As Long
Dim rnSource As Range
Dim wbk As Variant

thisfname = ThisWorkbook.Name

Dim MO_initials(1 To 12) As Variant
Dim IW_i As Variant
Dim IW_i2 As Integer
IW_i2 = 1

'MO initials
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
'
'For each MO file insert info to the Workplan

Call MO_Validation

For Each IW_i In MO_initials
    IW_Start(IW_i2) = Timer
    fNameAndPath = _
    "https://komisbrussels.sharepoint.com/sites/operations/ROMGlobal/A0_TQM/A0_03_Operations/T0_03.1_Project%20Planning/Mission%20Organisers/MO_" & IW_i & ".xlsm"
    aPath = Split(fNameAndPath, "\")
    fname = "MO_" & IW_i & ".xlsm"
    Call Date_File(fNameAndPath, LResult(IW_i2))
    Set wkb = Workbooks.Open(fNameAndPath, ReadOnly:=True, Password:="", UpdateLinks:=0)
    Workbooks(fname).Worksheets("MO").Activate
    
    'Copy into workplan
    
    Call ReadMO(mMissionID, mContractNo, mMO, mCountry, mOM, mImplement, mImplementType, _
                mOrgStart, mBriefExp, mBriefDate, mOrgStatus, mOutStatus, mDraftsReal, _
                mQCReal, mDebrief, mDebriefExp, mDraftReportReal, mFinalReal, mMonComment, _
                mQCComment, mDeskStart, IW_i, IW_i2, kwt, mDateUpdated, mDelivered)

    For Counter1 = 1 To kst
        For Counter2 = 1 To kwt
            If (wMissionID(Counter1) = mMissionID(IW_i2, Counter2) And _
                wContractNo(Counter1) = mContractNo(IW_i2, Counter2) And _
                wMO(Counter1) = mMO(IW_i2, Counter2) And _
                wDeskStart(Counter1) = mDeskStart(IW_i2, Counter2) And _
                wCountry(Counter1) = mCountry(IW_i2, Counter2)) Then

                wOM(Counter1) = mOM(IW_i2, Counter2)

                wImplement(Counter1) = mImplement(IW_i2, Counter2)
                'wImplementType(Counter1) = mImplementType(IW_i2, Counter2)
                wOrgStart(Counter1) = mOrgStart(IW_i2, Counter2)
                wDelivered(Counter1) = mDelivered(IW_i2, Counter2)

                wOrgStatus(Counter1) = mOrgStatus(IW_i2, Counter2)
                wOutStatus(Counter1) = mOutStatus(IW_i2, Counter2)
                wDraftsReal(Counter1) = mDraftsReal(IW_i2, Counter2)
                wQCReal(Counter1) = mQCReal(IW_i2, Counter2)

                wDraftReportReal(Counter1) = mDraftReportReal(IW_i2, Counter2)
                wFinalReal(Counter1) = mFinalReal(IW_i2, Counter2)
                'wMonComment(Counter1) = mMonComment(IW_i2, Counter2)
                'wQCComment(Counter1) = mQCComment(IW_i2, Counter2)
            End If
        Next Counter2
    Next Counter1

    ThisWorkbook.Worksheets("Workplan").Activate

    For Counter3 = 1 To kst
         Cells(Counter3 + 6, 25).Value = wOM(Counter3)
         Cells(Counter3 + 6, 26).Value = wImplement(Counter3)
         'Cells(Counter3 + 6, 27).Value = wImplementType(Counter3)
        
         If wOrgStart(Counter3) = "00:00:00" Then wOrgStart(Counter3) = ""
         Cells(Counter3 + 6, 44).Value = wOrgStart(Counter3)
        
        
         Cells(Counter3 + 6, 49).Value = wOrgStatus(Counter3)
         Cells(Counter3 + 6, 58).Value = wOutStatus(Counter3)
        
         If wDraftsReal(Counter3) = "00:00:00" Then wDraftsReal(Counter3) = ""
         Cells(Counter3 + 6, 67).Value = wDraftsReal(Counter3)
        
         If wQCReal(Counter3) = "00:00:00" Then wQCReal(Counter3) = ""
         Cells(Counter3 + 6, 69).Value = wQCReal(Counter3)
        
         If wDraftReportReal(Counter3) = "00:00:00" Then wDraftReportReal(Counter3) = ""
         Cells(Counter3 + 6, 73).Value = wDraftReportReal(Counter3)
        
         If wFinalReal(Counter3) = "00:00:00" Then wFinalReal(Counter3) = ""
         Cells(Counter3 + 6, 76).Value = wFinalReal(Counter3)
        
         'Cells(Counter3 + 6, 78).Value = wMonComment(Counter3)
         'Cells(Counter3 + 6, 79).Value = wQCComment(Counter3)
         Cells(Counter3 + 6, 102).Value = wDelivered(Counter3)
         
    Next Counter3
    
    Workbooks(fname).Close SaveChanges:=False
    IW_Sec(IW_i2) = Round(Timer - IW_Start(IW_i2), 2)
    IW_i2 = IW_i2 + 1
Next IW_i

'Reset search default values

Cells.Find(What:="", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False).Activate

    'Application.EnableEvents = True
  
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    
    IW_SecT = IW_Sec(1) + IW_Sec(2) + IW_Sec(3) + IW_Sec(4) + IW_Sec(5) + IW_Sec(6) + IW_Sec(7) + IW_Sec(8) + IW_Sec(9) + IW_Sec(10) + IW_Sec(11) + IW_Sec(12)



    MsgBox _
    "The Workplan PP is updated as follows:" & Chr(10) & _
    MO_initials(1) & ": " & LResult(1) & " IT: " & IW_Sec(1) & Chr(10) & _
    MO_initials(2) & ": " & LResult(2) & " IT: " & IW_Sec(2) & Chr(10) & _
    MO_initials(3) & ": " & LResult(3) & " IT: " & IW_Sec(3) & Chr(10) & _
    MO_initials(4) & ": " & LResult(4) & " IT: " & IW_Sec(4) & Chr(10) & _
    MO_initials(5) & ": " & LResult(5) & " IT: " & IW_Sec(5) & Chr(10) & _
    MO_initials(6) & ": " & LResult(6) & " IT: " & IW_Sec(6) & Chr(10) & _
    MO_initials(7) & ": " & LResult(7) & " IT: " & IW_Sec(7) & Chr(10) & _
    MO_initials(8) & ": " & LResult(8) & " IT: " & IW_Sec(8) & Chr(10) & _
    MO_initials(9) & ": " & LResult(9) & " IT: " & IW_Sec(9) & Chr(10) & _
    MO_initials(10) & ": " & LResult(10) & " IT: " & IW_Sec(10) & Chr(10) & _
    MO_initials(11) & ": " & LResult(11) & " IT: " & IW_Sec(11) & Chr(10) & _
    MO_initials(12) & ": " & LResult(12) & " IT: " & IW_Sec(12) & Chr(10) & _
    "------------------------------------------------------" & Chr(10) & _
    " Total time for updates" & Chr(9) & " : " & IW_SecT & " sec" & Chr(10) _
    , vbInformation
On Error Resume Next

End Sub
Sub JB()

    Call Unhide_All

    Columns("B:B").Select
    Selection.EntireColumn.Hidden = True


    Columns("L:O").Select
    Selection.EntireColumn.Hidden = True

    Columns("R:T").Select
    Selection.EntireColumn.Hidden = True

    Columns("V:AB").Select
    Selection.EntireColumn.Hidden = True

    Columns("AD:AD").Select
    Selection.EntireColumn.Hidden = True

    Columns("AF:AI").Select
    Selection.EntireColumn.Hidden = True

    Columns("AK:AL").Select
    Selection.EntireColumn.Hidden = True

    Columns("AO:AO").Select
    Selection.EntireColumn.Hidden = True

    Columns("AS:AV").Select
    Selection.EntireColumn.Hidden = True

    Columns("BK:BM").Select
    Selection.EntireColumn.Hidden = True


    Columns("BQ").Select
    Selection.EntireColumn.Hidden = True

    Columns("BR").Select
    Selection.EntireColumn.Hidden = True

    Columns("BU").Select
    Selection.EntireColumn.Hidden = True

    Columns("BX").Select
    Selection.EntireColumn.Hidden = True

    Columns("BZ").Select
    Selection.EntireColumn.Hidden = True

    Columns("CA").Select
    Selection.EntireColumn.Hidden = True


    Cells(6, 9).Select

End Sub
Public Sub ReadWorkplan(wStrand, wStatusROM, wMissionID, wContractNo, wIntTitle, wType, wEntity, wOM, wImplement, wImplementType, wProjectKE, wCoreTeam, wMO, wOrgStart, wBriefExp, wBriefDate, wOrgStatus, wDeskStart, wCountry, wFieldTravelS, wFieldWorkS, wFieldWorkE, wFieldTravelE, wNameExp, wTypeExp, wOutStatus, wNameQC, wTypeQC, wNoReport, wNoMQ, wDraftsDue, wDraftsReal, wQCDue, wQCReal, wDebrief, wDebriefExp, wDraftReport, wDraftReportReal, wDeadlineECDraftComments, wFinalPlanned, wFinalReal, wDeadlineECComments, wMonComment, wQCComment, wFieldPhase, wDelivered, kst, LRWorkplan)

MyFile = ThisWorkbook.Name

Application.ScreenUpdating = False


LRWorkplan = Worksheets("Workplan").Cells.SpecialCells(xlCellTypeLastCell).Row

'Read workplan into array

kst = 0

    For i = 7 To LRWorkplan
            kst = kst + 1
            
            wStrand(kst) = Cells(i, 3).Value
            wStatusROM(kst) = Cells(i, 4).Value
            wMissionID(kst) = Cells(i, 5).Value
            wContractNo(kst) = Cells(i, 6).Value
            wIntTitle(kst) = Cells(i, 8).Value
            wType(kst) = Cells(i, 9).Value
            wEntity(kst) = Cells(i, 21).Value
            wOM(kst) = Cells(i, 25).Value
            wImplement(kst) = Cells(i, 26).Value
            wImplementType(kst) = Cells(i, 27).Value
            wProjectKE(kst) = Cells(i, 40).Value
            wCoreTeam(kst) = Cells(i, 41).Value
            wMO(kst) = Cells(i, 42).Value
           
            
            If IsDate(Cells(i, 44).Value) Then wOrgStart(kst) = Cells(i, 44).Value
        
            wBriefExp(kst) = Cells(i, 45).Value
            
            If IsDate(Cells(i, 46).Value) Then wBriefDate(kst) = Cells(i, 46).Value

            wOrgStatus(kst) = Cells(i, 49).Value
            
            If IsDate(Cells(i, 50).Value) Then wDeskStart(kst) = Cells(i, 50).Value
                
            wCountry(kst) = Cells(i, 51).Value
            
            If IsDate(Cells(i, 52).Value) Then wFieldTravelS(kst) = Cells(i, 52).Value
            
            If IsDate(Cells(i, 53).Value) Then wFieldWorkS(kst) = Cells(i, 53).Value
            
            If IsDate(Cells(i, 54).Value) Then wFieldWorkE(kst) = Cells(i, 54).Value
            
            If IsDate(Cells(i, 55).Value) Then wFieldTravelE(kst) = Cells(i, 55).Value
            
            wNameExp(kst) = Cells(i, 56).Value
            wTypeExp(kst) = Cells(i, 57).Value
            wOutStatus(kst) = Cells(i, 58).Value
            wNameQC(kst) = Cells(i, 59).Value
            wTypeQC(kst) = Cells(i, 60).Value
            wNoReport(kst) = Cells(i, 61).Value
            wNoMQ(kst) = Cells(i, 62).Value
            
            If IsDate(Cells(i, 66).Value) Then wDraftsDue(kst) = Cells(i, 66).Value
            
            If IsDate(Cells(i, 67).Value) Then wDraftsReal(kst) = Cells(i, 67).Value
            
            If IsDate(Cells(i, 68).Value) Then wQCDue(kst) = Cells(i, 68).Value
                
            If IsDate(Cells(i, 69).Value) Then wQCReal(kst) = Cells(i, 69).Value
            
            If IsDate(Cells(i, 70).Value) Then wDebrief(kst) = Cells(i, 70).Value
            
            wDebriefExp(kst) = Cells(i, 71).Value
            
            If IsDate(Cells(i, 72).Value) Then wDraftReport(kst) = Cells(i, 72).Value
            
            If IsDate(Cells(i, 73).Value) Then wDraftReportReal(kst) = Cells(i, 73).Value
            
            If IsDate(Cells(i, 74).Value) Then wDeadlineECDraftComments(kst) = Cells(i, 74).Value

            
            If IsDate(Cells(i, 75).Value) Then wFinalPlanned(kst) = Cells(i, 75).Value
            
            If IsDate(Cells(i, 76).Value) Then wFinalReal(kst) = Cells(i, 76).Value
            
            If IsDate(Cells(i, 77).Value) Then wDeadlineECComments(kst) = Cells(i, 77).Value
            
            wMonComment(kst) = Cells(i, 78).Value
            wQCComment(kst) = Cells(i, 79).Value
            wFieldPhase(kst) = Cells(i, 83).Value
            wDelivered(kst) = Cells(i, 102).Value
            
    Next i
            
            
End Sub
Public Sub ReadMO(mMissionID, mContractNo, mMO, mCountry, mOM, mImplement, mImplementType, _
                  mOrgStart, mBriefExp, mBriefDate, mOrgStatus, mOutStatus, mDraftsReal, mQCReal, _
                  mDebrief, mDebriefExp, mDraftReportReal, mFinalReal, mMonComment, mQCComment, mDeskStart, MO_initials, iMO, kwt, mDateUpdated, mDelivered)

LRMO = Worksheets("MO").Cells.SpecialCells(xlCellTypeLastCell).Row

'Read MO file into array

kwt = 0

    For i = 5 To LRMO
        'If Worksheets("MO").Cells(i, 11).Value = MO_initials Then
            kwt = kwt + 1
            mMissionID(iMO, kwt) = Worksheets("MO").Cells(i, 3).Value
            mContractNo(iMO, kwt) = Worksheets("MO").Cells(i, 4).Value
            mMO(iMO, kwt) = Worksheets("MO").Cells(i, 11).Value
            mCountry(iMO, kwt) = Worksheets("MO").Cells(i, 19).Value
            mOM(iMO, kwt) = Worksheets("MO").Cells(i, 8).Value
            mImplement(iMO, kwt) = Worksheets("MO").Cells(i, 9).Value
            'mImplementType(iMO, kwt) = Worksheets("MO").Cells(i, 10).Value
            mDelivered(iMO, kwt) = Worksheets("MO").Cells(i, 17).Value
            
            If IsDate(Worksheets("MO").Cells(i, 21).Value) Then
                mOrgStart(iMO, kwt) = Worksheets("MO").Cells(i, 21).Value
            Else
                mOrgStart(iMO, kwt) = "00:00:00"
            End If
            
            mBriefExp(iMO, kwt) = Worksheets("MO").Cells(i, 24).Value
            
            If IsDate(Worksheets("MO").Cells(i, 25).Value) Then
                mBriefDate(iMO, kwt) = Worksheets("MO").Cells(i, 25).Value
            Else
                mBriefDate(iMO, kwt) = "00:00:00"
            End If
            
            mOrgStatus(iMO, kwt) = Worksheets("MO").Cells(i, 22).Value
            mOutStatus(iMO, kwt) = Worksheets("MO").Cells(i, 32).Value
            
            If IsDate(Worksheets("MO").Cells(i, 36).Value) Then
                mDraftsReal(iMO, kwt) = Worksheets("MO").Cells(i, 36).Value
            Else
                mDraftsReal(iMO, kwt) = "00:00:00"
            End If
            
            If IsDate(Worksheets("MO").Cells(i, 38).Value) Then
                mQCReal(iMO, kwt) = Worksheets("MO").Cells(i, 38).Value
            Else
                mQCReal(iMO, kwt) = "00:00:00"
            End If
            
            If IsDate(Worksheets("MO").Cells(i, 30).Value) Then
                mDebrief(iMO, kwt) = Worksheets("MO").Cells(i, 30).Value
            Else
                mDebrief(iMO, kwt) = "00:00:00"
            End If
            
            mDebriefExp(iMO, kwt) = Worksheets("MO").Cells(i, 31).Value
            
            If IsDate(Worksheets("MO").Cells(i, 40).Value) Then
                mDraftReportReal(iMO, kwt) = Worksheets("MO").Cells(i, 40).Value
            Else
                mDraftReportReal(iMO, kwt) = "00:00:00"
            End If
            
            If IsDate(Worksheets("MO").Cells(i, 44).Value) Then
                mFinalReal(iMO, kwt) = Worksheets("MO").Cells(i, 44).Value
            Else
                mFinalReal(iMO, kwt) = "00:00:00"
            End If
            
            'mMonComment(iMO, kwt) = Worksheets("MO").Cells(i, 43).Value
            'mQCComment(iMO, kwt) = Worksheets("MO").Cells(i, 44).Value

            If IsDate(Worksheets("MO").Cells(i, 23).Value) Then
                mDeskStart(iMO, kwt) = Cells(i, 23).Value
            Else
                mDeskStart(iMO, kwt) = "00:00:00"
            End If
            
            If IsDate(Worksheets("MO").Cells(1, 7).Value) Then
                mDateUpdated = Worksheets("MO").Cells(1, 7).Value
            End If
         'End If
    Next i
                      
End Sub


Public Sub ReadWorkplanDS(wpMissionID, wpContractNo, wpDSExpert, wpMO, wpBriefExp, wpBriefDate, wpDeskStart, wpCountry, wpNameExp, wpTypeExp, wpQCDue, wpDebrief, wpDebriefExp, wpFieldTravelS, kst, LRWorkplan)

MyFile = ThisWorkbook.Name

Application.ScreenUpdating = False


LRWorkplan = Worksheets("Workplan").Cells.SpecialCells(xlCellTypeLastCell).Row

'Read workplan into array

kst = 0

    For i = 7 To LRWorkplan
            kst = kst + 1

            wpMissionID(kst) = Cells(i, 5).Value
            wpContractNo(kst) = Cells(i, 6).Value
            wpDSExpert(kst) = Cells(i, 39).Value
            wpMO(kst) = Cells(i, 42).Value
            wpBriefExp(kst) = Cells(i, 45).Value

            If IsDate(Cells(i, 46).Value) Then wpBriefDate(kst) = Cells(i, 46).Value


            If IsDate(Cells(i, 50).Value) Then wpDeskStart(kst) = Cells(i, 50).Value

            wpCountry(kst) = Cells(i, 51).Value

            If IsDate(Cells(i, 52).Value) Then wpFieldTravelS(kst) = Cells(i, 52).Value


            wpNameExp(kst) = Cells(i, 59).Value
            wpTypeExp(kst) = Cells(i, 60).Value


            If IsDate(Cells(i, 68).Value) Then wpQCDue(kst) = Cells(i, 68).Value

            If IsDate(Cells(i, 70).Value) Then wpDebrief(kst) = Cells(i, 70).Value

            wpDebriefExp(kst) = Cells(i, 71).Value

    Next i

End Sub

Public Sub ReadDS(dMissionID, dContractNo, dDSExpert, dMO, dBriefExp, dBriefDate, dDeskStart, dCountry, dNameExp, dTypeExp, dQCDue, dDebrief, dDebriefExp, dFieldTravelS, kwt, LRWorkplan)

Application.ScreenUpdating = False

LRWorkplan = Worksheets("Workplan").Cells.SpecialCells(xlCellTypeLastCell).Row

'Read workplan into array

kwt = 0

    For i = 7 To LRWorkplan
            kwt = kwt + 1

            dMissionID(kwt) = Cells(i, 5).Value
            dContractNo(kwt) = Cells(i, 6).Value
            dDSExpert(kwt) = Cells(i, 39).Value
            dMO(kwt) = Cells(i, 42).Value
            dBriefExp(kwt) = Cells(i, 45).Value

            If IsDate(Cells(i, 46).Value) Then dBriefDate(kwt) = Cells(i, 46).Value


            If IsDate(Cells(i, 50).Value) Then dDeskStart(kwt) = Cells(i, 50).Value

            dCountry(kwt) = Cells(i, 51).Value
            If IsDate(Cells(i, 52).Value) Then dFieldTravelS(kwt) = Cells(i, 52).Value


            dNameExp(kwt) = Cells(i, 59).Value
            dTypeExp(kwt) = Cells(i, 60).Value


            If IsDate(Cells(i, 68).Value) Then dQCDue(kwt) = Cells(i, 68).Value

            If IsDate(Cells(i, 70).Value) Then dDebrief(kwt) = Cells(i, 70).Value

            dDebriefExp(kwt) = Cells(i, 71).Value


    Next i


End Sub

Public Sub UpdateDS()

'Desk study variables

Dim dMissionID(5000), dContractNo(5000), dDSExpert(5000), dMO(5000), dBriefExp(5000), dCountry(5000), dNameExp(5000), _
dTypeExp(5000), dDebriefExp(5000) As String

Dim dQCDue(5000), dDeskStart(5000), dBriefDate(5000), dDebrief(5000), dFieldTravelS(5000) As Date

dQCDue(5000) = "00:00:00"
dBriefDate(5000) = "00:00:00"
dDeskStart(5000) = "00:00:00"
dDebrief(5000) = "00:00:00"
dFieldTravelS(5000) = "00:00:00"

Dim wpMissionID(5000), wpContractNo(5000), wpDSExpert(5000), wpMO(5000), wpBriefExp(5000), wpCountry(5000), wpNameExp(5000), _
wpTypeExp(5000), wpDebriefExp(5000) As String

Dim wpQCDue(5000), wpBriefDate(5000), wpDebrief(5000), wpDeskStart(5000), wpFieldTravelS(5000) As Date

wpQCDue(5000) = "00:00:00"
wpBriefDate(5000) = "00:00:00"
wpDeskStart(5000) = "00:00:00"
wpDebrief(5000) = "00:00:00"
wpFieldTravelS(5000) = "00:00:00"


'General variables

Dim LRWorkplan As Integer

Dim kst As Integer

Dim fNameAndPath As Variant, Wb As Workbook
Dim thisfname As Variant
Dim Current As Long
Dim Rcount1 As Variant
Dim Rcount2 As Variant
Dim ThisWkb As String


Dim Counter1, Counter2, Counter3, XX As Integer

Counter1 = 0
Counter2 = 0
Counter3 = 0
XX = 0
ThisWkb = ThisWorkbook.Name


Dim StartTime As Double
Dim SecondsElapsed As Double
'Remember time when macro starts

'Stop updating screen


StartTime = Timer
With Application
        .ScreenUpdating = False
End With

Application.EnableEvents = False

thisfname = ThisWorkbook.Name

DSFile = "https://komisbrussels.sharepoint.com/sites/operations/ROMGlobal/A0_TQM/A0_03_Operations/T0_03.1_Project%20Planning/Workplan_Updates/Desk_Study_PP.xlsm"

Set wkb = Workbooks.Open(DSFile, ReadOnly:=True)

fNameAndPath = DSFile

aPath = Split(fNameAndPath, "\")

fname = "Desk_Study_PP.xlsm"

With wkb.Worksheets("Update Workplan")
    .Range("A1").AutoFilter
    .Activate
End With

Call ReadWorkplanDS(wpMissionID, wpContractNo, wpDSExpert, wpMO, wpBriefExp, wpBriefDate, wpDeskStart, wpCountry, wpNameExp, wpTypeExp, wpQCDue, wpDebrief, wpDebriefExp, wpFieldTravelS, kst, LRWorkplan)

Call ReadDS(dMissionID, dContractNo, dDSExpert, dMO, dBriefExp, dBriefDate, dDeskStart, dCountry, dNameExp, dTypeExp, dQCDue, dDebrief, dDebriefExp, dFieldTravelS, kwt, LRWorkplan)


'Copy into workplan

For Counter1 = 1 To kst

    For Counter2 = 1 To kwt

    If (wpMissionID(Counter1) = dMissionID(Counter2) And _
    wpDeskStart(Counter1) = dDeskStart(Counter2) And _
    wpFieldTravelS(Counter1) = dFieldTravelS(Counter2) And wpMissionID(Counter1) <> "") Then


ThisWorkbook.Worksheets("Workplan").Activate

    With wkb.Worksheets("Update Workplan")
        Rcount1 = 3000
        .AutoFilterMode = False
        .Columns("A:CR").EntireColumn.Hidden = False
        .Cells(Counter2 + 6, 39).Copy ThisWorkbook.Sheets("Workplan").Cells(Counter2 + 6, 39)
        .Cells(Counter2 + 6, 45).Copy ThisWorkbook.Sheets("Workplan").Cells(Counter2 + 6, 45)
        .Cells(Counter2 + 6, 46).Copy ThisWorkbook.Sheets("Workplan").Cells(Counter2 + 6, 46)
        .Cells(Counter2 + 6, 59).Copy ThisWorkbook.Sheets("Workplan").Cells(Counter2 + 6, 59)
        .Cells(Counter2 + 6, 60).Copy ThisWorkbook.Sheets("Workplan").Cells(Counter2 + 6, 60)
        .Cells(Counter2 + 6, 68).Copy ThisWorkbook.Sheets("Workplan").Cells(Counter2 + 6, 68)
        .Cells(Counter2 + 6, 70).Copy ThisWorkbook.Sheets("Workplan").Cells(Counter2 + 6, 70)
        .Cells(Counter2 + 6, 71).Copy ThisWorkbook.Sheets("Workplan").Cells(Counter2 + 6, 71)
        .Cells(Counter2 + 6, 81).Copy ThisWorkbook.Sheets("Workplan").Cells(Counter2 + 6, 81)
        .Cells(Counter2 + 6, 82).Copy ThisWorkbook.Sheets("Workplan").Cells(Counter2 + 6, 82)
    End With



          End If
    Next Counter2

Next Counter1




wkb.Close SaveChanges:=False

'Reset search default values

Cells.Find(What:="", After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False).Activate



With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
End With

SecondsElapsed = Round(Timer - StartTime, 2)

Application.EnableEvents = True

MsgBox "Desk Study successfully updated (in " & SecondsElapsed & " sec)", vbInformation




End Sub

Sub Date_File(File_Path, LResult)
    On Error GoTo 20
    New_Path = SharePointURLtoUNC(File_Path)
    filespec = New_Path
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    LResult = f.DateLastModified
    GoTo 30
20 LResult = "00:00:00"
30
End Sub

Public Function SharePointURLtoUNC(sURL)
  Dim bIsSSL As Boolean

  bIsSSL = InStr(1, sURL, "https:") > 0
  sURL = Replace(Replace(sURL, "/", "\"), "%20", " ")
  sURL = Replace(Replace(sURL, "https:", vbNullString), "http:", vbNullString)

  sURL = Replace(sURL, Split(sURL, "\")(2), Split(sURL, "\")(2) & "@SSL\DavWWWRoot")
  If Not bIsSSL Then sURL = Replace(sURL, "@SSL\", vbNullString)
  SharePointURLtoUNC = sURL
  End Function

