Sub MergeC()
Dim rChange As Range
Dim MasterBook As Workbook
Dim MasterData As Workbook
Dim MasterDataFrozen As Workbook
Dim CallBook As Workbook
Dim pt As PivotTable
Dim rng As Range
Dim rowRng As Range
Dim TestBook As Workbook
Dim DataFileName As String
Dim DataFileIsOpen As Boolean
Set MasterBook = ThisWorkbook
CurrentPath = ThisWorkbook.Path

t = Now()

Select Case Range("Dev!B1")
    Case 1
        On Error GoTo 0
    Case 0
        On Error GoTo MergeC_Error
End Select

DataFileName = "ApolloMasterData.xlsx" 'Set Name of master file to a string
DataFileIsOpen = False 'set boolean value to not open
For Each TestBook In Application.Workbooks 'loop through open books to see if master file is open
    If UCase(TestBook.Name) = UCase(DataFileName) Then
        DataFileIsOpen = True
        Exit For
    End If
Next TestBook
Debug.Print "Find book open:" & Format(Now() - t, "hh:mm:ss")
Select Case DataFileIsOpen
    Case False 'Since DataFile is not open, open it as MasterDataFrozen
        Set MasterDataFrozen = Workbooks.Open(CurrentPath & "\ApolloMasterData.xlsx") 'Define MasterData and open DataFile as MasterData
    Case Else ' Since DataFile is already open, just define it as DataFileBook
        Set MasterDataFrozen = Workbooks(DataFileName)
End Select
Yr = Right(Year(Now), 2)
Mm = Format(Month(Now), "00")
Dd = Format(Day(Now), "00")
MasterDataFrozen.SaveAs (CurrentPath & "\" & Yr & Mm & Dd & "ApolloMasterData.xlsx") 'save frozen copy with todays date
MasterDataFrozen.Close
Set MasterData = Workbooks.Open(CurrentPath & "\ApolloMasterData.xlsx") 'Open MasterData for editing

MasterBook.Activate
With Application 'disable calculation, screenupdating and events to speed up process
    .Calculation = xlCalculationManual
    .ScreenUpdating = False
    .EnableEvents = False
End With
    
CallDate = Sheets("PM").Range("D8")
For c = 1 To Sheets("PivotTables").Range("G1")
    Caller = Sheets("PivotTables").Range("F5").Offset(c)
    Select Case Caller
        Case "", "(blank)"
            GoTo NoCaller
        Case Else
    End Select
    CallerFileName = Replace(Caller, " ", "")
    Set CallBook = Workbooks.Open(CurrentPath & "\" & CallerFileName & "Data.xlsx")
Debug.Print "Open data:" & Format(Now() - t, "hh:mm:ss")
'    For i = 1 To NumChanged
        With CallBook.Sheets("Data")
            '.ListObjects("Data").Range.AutoFilter 'clear any existing filters
            .ListObjects("Data").Range.AutoFilter field:=33, Criteria1:=CallDate 'filter for the selected call date
            '.ListObjects("Data").Range.AutoFilter field:=17, Criteria1:=Caller 'filter for this caller only
            NumChange = .ListObjects("Data").Range.Resize(, 1).SpecialCells(xlCellTypeVisible).Count - 1 'count number of changed records
            Select Case NumChange
                Case 0 'no calls for this date
                    GoTo NoCalls
                Case Else
                    'records found do nothing
            End Select
            Set rng = .ListObjects("Data").DataBodyRange.Columns(1).SpecialCells(xlCellTypeVisible)
'         For Each rowRng In rng.Areas
            For Each rowRng In rng
                ChangeID = rowRng.Cells(, 1)
                ChangeRow = rowRng.Row
                Set rLoaded = MasterData.Sheets("Data").Columns(1).Find(What:=ChangeID, After:=.Cells(1, 1), LookIn:=xlValues, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                    , SearchFormat:=False)
                LoadedRow = rLoaded.Row
    '            .Rows(ChangeRow).Replace "=", "#="
                .Range("B" & ChangeRow & ":BW" & ChangeRow).Copy MasterData.Sheets("Data").Range("B" & LoadedRow)
    '            MasterBook.Sheets("Queue").Rows(LoadedRow).Replace "#=", "="
            Next rowRng
        End With
NoCalls:
        Application.DisplayAlerts = False
        CallBook.Close
        Application.DisplayAlerts = True
'    Next i
    With MasterBook.Sheets("Log")
        '.Unprotect ("davis1")
        'Make a log entry
        .Range("C6:H6").Insert xlShiftDown 'Push down previous log entries to add one at the top
        .Range("C6") = FormatDateTime(Now, vbGeneralDate) 'DateTime stamp
        .Range("D6") = Caller
        .Range("E6") = CallDate
        .Range("F6") = NumChange
        .Range("G6") = Range("LoginX!C1")
    End With
    
    Debug.Print Caller & ":" & NumChange
    Debug.Print "Book time:" & Format(Now() - t, "hh:mm:ss")
NoCaller:
Next c

Escape:
'Cleanup code goes here
With Application
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
    .EnableEvents = True
End With
Debug.Print "MergeC"
TotalTime = Format(Now() - t, "hh:mm:ss")
Debug.Print "Total:" & TotalTime
MsgBox "Caller data import complete" & vbLf & "Run time: " & TotalTime, vbOKOnly, "Data merge complete"
   On Error GoTo 0
   Exit Sub

MergeC_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MergeC of Module Merge in VBAProject" & vbCrLf & vbCrLf & "When you click OK the tool will continue to operate but proceed with caution in case something has become corrupt. Or you could close the file and reopen to try again." & vbCrLf & "This is probably our fault, so please accept our apologies in advance." & vbCrLf & vbCrLf & "Please send an email to keith.manning@kmaone.com and mention the error codes above.", vbCritical, "VBAProject Error"
'Error reset/resume code goes here

End Sub