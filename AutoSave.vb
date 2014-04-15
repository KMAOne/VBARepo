Public vartimer As Variant
Const TimeOut = 30 'in minutes
 
Sub Salva()
Dim MasterBook As Workbook
Dim TestBook As Workbook
Dim CallerData As Workbook
Dim DataFileName As String
Dim DataFileIsOpen As Boolean
Set MasterBook = ThisWorkbook
    
    CurrentPath = ThisWorkbook.Path
    Caller = Sheets("LoginX").Range("C1")
    UserClass = Sheets("LoginX").Range("D1")
    Select Case UserClass
        Case "PM"
            Exit Sub 'PM logged in no need for autosave
        Case "Caller"
            DataFileName = Replace(Caller, " ", "") & "Data.xlsx" 'Set Name of master file to a string
            DataFileIsOpen = False 'set boolean value to not open
            For Each TestBook In Application.Workbooks 'loop through open books to see if master file is open
                If UCase(TestBook.Name) = UCase(DataFileName) Then
                    DataFileIsOpen = True
                    Exit For
                End If
            Next TestBook
            Select Case DataFileIsOpen
                Case False 'Since DataFile is not open, open it as MasterData
                    Set CallerData = Workbooks.Open(CurrentPath & "\" & Caller & "Data.xlsx") 'Define CallerData and open DataFile as CallerData
                Case Else ' Since DataFile is already open, just define it as CallerData
                    Set CallerData = Workbooks(DataFileName)
            End Select
            CallerData.Save
    End Select
    Call Tempo
End Sub
 
Sub Tempo()
    vartimer = Format(Now + TimeSerial(0, TimeOut, 0), "hh:mm:ss")
    If vartimer = "" Then Exit Sub
    Application.OnTime TimeValue(vartimer), "Salva"
End Sub
 
Sub Limpa()
    Call Salva
    On Error Resume Next
    Application.OnTime earliesttime:=vartimer, _
    procedure:="Salva", schedule:=False
    On Error GoTo 0
End Sub