'On file open, checks to see if login data was previously entered, if not, it runs setup and opens the login form
Private Sub Workbook_Open()
Call Master.CheckIfSetup
End Sub
