Private Sub Workbook_Open()
    Call Tempo
Sheets("Login").Activate
End Sub
 
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call Limpa
End Sub