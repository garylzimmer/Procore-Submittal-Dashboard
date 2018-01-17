'Master Module file containing all common subs, functions, and variables

'Declaring User Defined and Computed Variables
Public Email As String
Public password As String
Public HomeURL As String
Public key As String
Public SourceDataFileName As String
Public DataSheetName As String
Public XLSURL As String


'Constants for Submittal Table
'SourceDataFileName = "list.xlsx"
'DataSheetName = "RAW DATA"

'Function SaveWebData takes source data xls/xlsx file url, that file name and ext, and the target data sheet's name
Public Function SaveWebData(URL As String, sourcefilename As String, targetsheet As String)

'hourly calls for refresh
Application.OnTime Now + TimeValue("01:00"), "SaveWebData(XLSURL, SourceDataFileName, DataSheetName)"
Application.OnTime Now + TimeValue("01:00"), "UpdateTime"

Dim WHTTP As Object
Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5.1")


'Save current window name
DashboardWindow = ActiveWorkbook.Name
'clear the target data sheet
Sheets(targetsheet).Cells.ClearContents

'Previous attempt to get login page auth cookie and save it found unnecessary for now
'WHTTP.Open "GET", URL, False
'WHTTP.SetRequestHeader "Cookie", cookie
'WHTTP.Send

'Open the Submittal list.xlsx link in excel directly
Workbooks.Open (URL)

'activate new file window
Windows(sourcefilename).Activate

'Select all used data ranges
ActiveSheet.UsedRange.Select

'Copy that selection to clipboard
Selection.Copy

'activate original dashboard excel window
Windows(DashboardWindow).Activate

'Paste from clipboard to target sheet at A1
Sheets(targetsheet).Paste Destination:=Worksheets(targetsheet).Range("A1")

'Clear Clipboard
Application.CutCopyMode = False

'Close the list.xlsx excel window
Windows(sourcefilename).Close

End Function


'Function to Login takes the url for the "proper" login page (seems to always be app.procore.com/sessions, not app.procore.com/login)
'and the previously built "key" string containing username and password in proper syntax
Public Function LoginProcore(loginurl As String, key As String)

  Set XMLHttpRequest = CreateObject("Msxml2.XMLHTTP.6.0")

  'example key = "method=login&session[email]=EMAIL@WEB.COM&pass=PASSWORD123"
   
  With XMLHttpRequest
   .Open "POST", "https://app.procore.com/sessions", False
   .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
   .Send "method=login&" & key
  End With

End Function

'Sub for changing the "last updated" field in the dashboard sheet
Public Sub UpdateTime()
Worksheets("DASHBOARD").Range("C2").Value = Now()
End Sub
Public Sub CheckIfSetup()
If Worksheets("DEBUG").Range("C12").Value = False Then
    MsgBox "Need to run a setup"
    LoginForm.Show
    Sheets("DEBUG").Range("C12").Value = True
End If
End Sub

'Sub for logging out of the dashboard and wiping all user defined variables (username, password, homeurl, and all computed variables)
'Currently keeps the downloaded data
'TODO: Stop autorefreshing data if logged out!

Public Sub Logout()
'declare relog question
Dim relog As Integer

'Clear Debug Sheet
Worksheets("DEBUG").Range("C3:C7").Value = ""
Email = ""
password = ""
HomeURL = ""
 key = ""
 XLSURL = ""
 
 'Setup boolean tracks if login form is filled out or not
 SETUP = False
 Worksheets("DEBUG").Range("C12").Value = False
 
 relog = MsgBox("Would you like to log in with a different account/project?", vbYesNo + vbQuestion, "Relog?")
 
 If relog = vbYes Then
    LoginForm.Show
ElseIf relog = vbNo Then
    MsgBox ("FYI, If you close and reopen, you will be asked for a new login...")
End If

End Sub
