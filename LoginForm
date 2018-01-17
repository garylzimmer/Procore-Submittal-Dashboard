'Build GUI Form with Textfields for Email, Password, Project Home URL, and a command button for login

Public Sub LoginBtn_Click()

'Save Login Form user variables to Master module
Master.Email = LoginForm.UserTxt.Value
Master.password = LoginForm.PWTxt.Value
Master.HomeURL = LoginForm.URLTxt.Value
Master.key = "session[email]=" & Email & "&session[password]=" & password
Master.XLSURL = Replace(HomeURL, "home", "submittal_logs/list.xlsx")

'Write the user variables to DEBUG sheet for storage
Sheets("DEBUG").Range("C3") = Master.Email
Sheets("DEBUG").Range("C4") = Master.password
Sheets("DEBUG").Range("C5") = Master.HomeURL
Sheets("DEBUG").Range("C6") = Master.key
Sheets("DEBUG").Range("C7") = Master.XLSURL

'Read the sitewide constants from DEBUG
Master.SourceDataFileName = Sheets("DEBUG").Range("C9")
Master.DataSheetName = Sheets("DEBUG").Range("C10")

'Begin Login Procedure
Call Master.LoginProcore("https://app.procore.com/sessions", Master.key)
Call Master.SaveWebData(Master.XLSURL, Master.SourceDataFileName, Master.DataSheetName)
Call Master.UpdateTime
SETUP = True

Unload Me
End Sub
