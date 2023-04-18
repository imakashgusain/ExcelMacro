Option Explicit

Dim xlApp, xlBook

Set xlApp = CreateObject("Excel.Application")
'~~> Change Path here
Set xlBook = xlApp.Workbooks.Open("D:\Book1.xlsm", 0, False)
xlApp.Run "AddCorrespondingCells"
xlBook.Save
xlBook.Close
xlApp.Quit

Set xlBook = Nothing
Set xlApp = Nothing
