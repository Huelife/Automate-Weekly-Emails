Option Explicit

'Declare variables
Dim xlApp, xlBook1, xlBook2, path1, path2
path1 = "C:\Users\To\File\Location\email_logs_macro.xlsm"

'Create Excel object and open Macro WB1
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
Set xlBook1 = xlApp.Workbooks.Open(path1, 0 True)

'Open/close WB2, Executing Macro
path2 = "C:\Users\To\File\Location\logs.xlsm"
Set xlBook2 = xlApp.Workbooks.Open(path2, 0, True)
xlApp.Run "'" & path1 & "'!email_logs_macro.email_logs_macro"
xlBook2.Close

'Close WB1 and Excel
xlBook1.Close
xlApp.Quit

Set xlApp = Nothing
Set xlBook1 = Nothing
Set xlBook2 = Nothing

WScript.Quit
