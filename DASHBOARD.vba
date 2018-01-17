'Code for the DASHBOARD sheet, must have command button for logout and manual refresh

Private Sub LogoutBtn_Click()
Call Master.Logout
End Sub

Private Sub RefreshBtn_Click()
Call Master.SaveWebData(Master.XLSURL, Master.SourceDataFileName, Master.DataSheetName)
Call Master.UpdateTime
End Sub


