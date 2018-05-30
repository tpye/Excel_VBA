Attribute VB_Name = "CanAmCMS_UserAudit_Validation"
Public reportFile As String
Public reportNameString As String
Option Base 1
Public Sub CreateCMSuserReport()
Dim fileSelect As Office.FileDialog
Dim myRangeRow As Range
Dim myRangeColumn As Range
Dim finalrow As Integer
Dim finalColumn As Integer
Dim wsReport As Worksheet
Dim wbProfile As Workbook
Dim wsProfile As Worksheet
Dim txtFileName As String
Dim userProfile As String
Dim permission As String
Dim CurrentRow As Integer
Dim CurrentColumn As Integer

Dim testRange As Range

 
'On Error GoTo ErrorHandler

Application.ScreenUpdating = False

Set fileSelect = Application.FileDialog(msoFileDialogFilePicker)
' select file
With fileSelect

      .AllowMultiSelect = False

      ' Set the title of the dialog box.
      .Title = "Please select the file."

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "All Files", "*.txt"
      .Filters.Add "All Files", "*.csv"
      .Filters.Add "Excel 2010", "*.xlsx"
      .Filters.Add "Excel 2010", "*.xlsm"
      .InitialFileName = "P:\CSG\BusApps\common\Trevorp\Reports"

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
        txtFileName = .SelectedItems(1)

      End If
End With


With ActiveSheet.QueryTables.Add(Connection:="TEXT;" + txtFileName, Destination:=Range("A1"))
        .Name = fileName
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileTabDelimiter = True
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
End With

reportFile = Split(txtFileName, ".", 2)(0) & " " & MonthName(Month(Date)) & " " & Year(Date) & ".xlsx"

With ActiveWorkbook
    .SaveAs fileName:=reportFile, FileFormat:=51
End With


Application.ScreenUpdating = False

reportNameString = Right(reportFile, Len(reportFile) - 87)

'define and open needed workbooks and worksheets
Set wb = Workbooks(reportNameString)
Set wsReport = wb.Worksheets(1)
Set wsErrorReport = wb.Worksheets.Add
wsErrorReport.Name = "Error Report"
Set wbProfile = Workbooks.Open("P:\CSG\BusApps\CanAm\CanAm CMS User Management\CanAm CMS User Profile Details.xlsx")
Set wsProfile = wbProfile.Worksheets("User Profiles")

wsReport.Name = "CMS User Report"

' Count number of rows and columns in report to determin location of bottom right cell address
finalrow = wsReport.Cells(Rows.Count, 1).End(xlUp).Row
finalColumn = wsReport.Cells(1, Columns.Count).End(xlToLeft).column

' setting up error report headers
With wsErrorReport
        .Range("A1").Value = "User Name"
        .Range("B1").Value = "Name"
        .Range("C1").Value = "User Profile"
        .Range("D1").Value = "Permission"
        .Range("E1").Value = "Error Note"
        .Range("A1:E1").Font.Bold = True
End With

Set reportTestRange = wsReport.Range(wsReport.Cells(2, 4), wsReport.Cells(finalrow, finalColumn))
Set reportMycell = wsReport.Range("D2")


For Each reportMycell In reportTestRange            ' loop through the permission range of the CMS user report and compare with expected user profiles
    CurrentColumn = reportMycell.column
    CurrentRow = reportMycell.Row
    permission = wsReport.Cells(1, CurrentColumn)   ' set the current permission to be tested
    userProfile = wsReport.Cells(CurrentRow, 3)     ' set the user profile for permission test
    userName = wsReport.Cells(CurrentRow, 1)        ' hold CMS user name in memory for use in report if test fails
    Name = wsReport.Cells(CurrentRow, 2)            ' hold CMS user full name in memory for use in report if test fails

        Set permissionRange = wsProfile.Range("b1", wsProfile.Range("xfd1").End(xlToRight))
        Set profileRange = wsProfile.Range("a2", wsProfile.Range("A1048576").End(xlUp))
    
    Set found1 = permissionRange.Find(what:=permission, LookAt:=xlWhole)
            If found1 Is Nothing Then
                            With wsErrorReport.Range("A1048576").End(xlUp) ' permission not found in user profile matrix
                                .Offset(1, 0).Value = userName
                                .Offset(0, 1).Value = Name
                                .Offset(0, 2).Value = userProfile
                                .Offset(0, 3).Value = permission
                                .Offset(0, 4).Value = "Permission not found in user profile matrix"
                            End With
                            GoTo SkipIfStatement
                        Else ' Nothing
            End If
        
    testColumn = found1.column
    Set found = profileRange.Find(what:=userProfile, LookAt:=xlWhole)
            If found Is Nothing Then
                            With wsErrorReport.Range("A1048576").End(xlUp) ' user_profile not found in user profile matrix
                                .Offset(1, 0).Value = userName
                                .Offset(0, 1).Value = Name
                                .Offset(0, 2).Value = userProfile
                                .Offset(0, 3).Value = permission
                                .Offset(0, 4).Value = "user_profile not found in user profile matrix"
                            End With
                            GoTo SkipIfStatement
                        Else
            End If
        
    testRow = found.Row
    Set testCell = wsProfile.Cells(testRow, testColumn)
        
   testcellAddress = testCell.Address
   
      If reportMycell.Value = testCell.Value Then
                       'nothing
                Else
                    
                With wsErrorReport.Range("A1048576").End(xlUp) ' reportMycell carries a value of 1 and with are expected test cell to have an interior colour formatting
                        .Offset(1, 0).Value = userName
                        .Offset(0, 1).Value = Name
                        .Offset(0, 2).Value = userProfile
                        .Offset(0, 3).Value = permission
                        .Offset(0, 4).Value = "Users permission (" & permission & ") is not consistent with their current profile of " & userProfile
                End With
        End If
    
SkipIfStatement:
Next reportMycell

With wsErrorReport
    .Columns("A").AutoFit
    .Columns("B").AutoFit
    .Columns("C").AutoFit
    .Columns("D").AutoFit
    .Columns("E").AutoFit
End With



Dim myArray() As Variant
Dim myCount As Integer

RowCount = wsProfile.Range("a2", wsProfile.Range("A1048576").End(xlUp)).Rows.Count

ReDim Preserve myArray(1 To RowCount)

' load data into myArray
With wsProfile
    For myCount = 1 To RowCount
        myArray(myCount) = .Range("A1").Offset(myCount)
    Next myCount
End With
    
' Create and name worksheets for user profile reports
With wb
    For myCount = 1 To RowCount
        If InStr(myArray(myCount), "/") Then
            .Sheets.Add.Name = Replace(myArray(myCount), "/", " ")
        Else
            .Sheets.Add.Name = myArray(myCount)
        End If
    Next myCount
End With

' add headers profile worksheets
With wb
    For myCount = 1 To RowCount
        If InStr(myArray(myCount), "/") <> 0 Then
            With .Worksheets(Replace(myArray(myCount), "/", " "))
                .Range("A1").Value = "User"
                .Range("B1").Value = "Full Name"
                .Range("C1").Value = "User Profile"
                .Range("A1:C1").Font.Bold = True
            End With
        Else
            With .Worksheets(myArray(myCount))
                .Range("A1").Value = "User"
                .Range("B1").Value = "Full Name"
                .Range("C1").Value = "User Profile"
                .Range("A1:C1").Font.Bold = True
            End With
        End If
    Next myCount
End With

' add users to user profile worksheets
Set myCell = wsReport.Range("A2")
Set myRange = wsReport.Range("A2", wsReport.Range("a1048576").End(xlUp))

For Each myCell In myRange
If myCell.Offset(0, 2).Value = "" Then
    GoTo skip
Else
    If InStr(myCell.Offset(0, 2).Value, "/") <> 0 Then
        With wb.Worksheets(Replace(myCell.Offset(0, 2).Value, "/", " ")).Range("A1048576").End(xlUp)
            .Offset(1, 0).Value = myCell.Value
            .Offset(1, 1).Value = myCell.Offset(0, 1).Value
            .Offset(1, 2).Value = myCell.Offset(0, 2).Value
        End With
'        Worksheets(Replace(myCell.Value, "/", " ")).Range("A:C").Columns.AutoFit
    Else
        With wb.Worksheets(myCell.Offset(0, 2).Value).Range("A1048576").End(xlUp)
            .Offset(1, 0).Value = myCell.Value
            .Offset(1, 1).Value = myCell.Offset(0, 1).Value
            .Offset(1, 2).Value = myCell.Offset(0, 2).Value
        End With
'        Worksheets(myCell.Value).Range("A:C").Columns.AutoFit
    End If
End If
skip:
Next myCell

Workbooks(reportNameString).Save
wbProfile.Close

Application.ScreenUpdating = True


MsgBox "Procedure Complete"

Exit Sub


'ErrorHandler:
'MsgBox "Procedure has failed" & vbNewLine & Err.Description

End Sub


Public Sub ACM_UserValidation()

Dim firstnameColumn As Integer
Dim surnameColumn As Integer
Dim fullName As String
Dim wbACM As Workbook
Dim wbUserAudit As Workbook
Dim wsACM As Worksheet
Dim finalrow As Integer


Set wbACM = Workbooks("ACM_PeopleDirectoryExport_" & Year(Now()) & Month(Now()) & Day(Now()) & ".csv")
Set wbUserAudit = Workbooks(reportNameString)
Set wsACM = wbUserAudit.Worksheets("ACM User")
Set wsDirectory = wbACM.Sheets(1)
Set directoryRange = wsDirectory.UsedRange

wbACM.Open
rowNum = 2

With wsDirectory
    
    firstname = .Cells(rowNum, firstnameColumn).Value
    lastname = .Cells(rowNum, surnameColumn).Value
    fullName = firstname & " " & lastname
End With

End Sub


