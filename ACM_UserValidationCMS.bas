Attribute VB_Name = "ACM_UserValidationCMS"
Public Sub ACM_CMSuserValidationReport()

Dim CMS As Workbook
Dim Charlie As Workbook
Dim Report As Workbook
Dim wsCMS As Worksheet
Dim wsCharlie As Worksheet
Dim wsCMSReport As Worksheet
Dim wsCharlieReport As Worksheet
Dim found As Range
Dim myCell As Range

'Application.ScreenUpdating = False

' create CMS user delta report workbook
Set Report = Workbooks.Add
Set wsCMSReport = Report.Worksheets("Sheet1")
Set wsCharlieReport = Report.Worksheets.Add
wsCMSReport.Name = "Non ACM Staff"
wsCMSReport.Range("A1").Value = "Users in CanAmCMS that are not listed as ACM staff on Charlie"
wsCharlieReport.Name = "Non CMS User"
wsCharlieReport.Range("A1").Value = "Users Listed as ACM employees on Charlie but do not have an Active CanAmCMS account."

If Month(Date) < 10 Then
    reportMonth = 0 & Month(Date)
Else
    reportMonth = Month(Date)
End If

reportYear = Year(Date)

If Day(Date) < 10 Then
    reportDay = 0 & Day(Date)
Else
    reportDay = Day(Date)
End If

'set workbook for comaprison
Set CMS = Workbooks("ACM User_" & reportYear & reportMonth & ".xlsx")
Set Charlie = Workbooks("ACM_PeopleDirectoryExport_" & reportYear & reportMonth & reportDay & ".csv")


Set wsCMS = CMS.Worksheets("ACM User")
Set wsCharlie = Charlie.Worksheets("ACM_PeopleDirectoryExport_" & reportYear)



Set myCell = wsCMS.Range("C7")

' right the report for the ACM CMS users who are not in People Directory
Do Until myCell.Value = ""
    fullName = Trim(myCell.Value)
    Set found = wsCharlie.Range("D1", wsCharlie.Range("D1048576").End(xlUp)).Find(what:=fullName, LookAt:=xlWhole)
        If found Is Nothing Then
        wsCMSReport.Range("A1048576").End(xlUp).Offset(1, 0).Value = fullName
        End If
    Set myCell = myCell.Offset(1, 0)
Loop

Set myCell = wsCharlie.Range("d2")

' right the report of people in the people Directory who do not have CMS access
Do Until myCell.Value = ""
    fullName = Trim(myCell.Value)
    Set found = wsCMS.Range("C7", wsCMS.Range("C1048576").End(xlUp)).Find(what:=fullName, LookAt:=xlWhole)
        If found Is Nothing Then
        wsCharlieReport.Range("A1048576").End(xlUp).Offset(1, 0).Value = fullName
        End If
    Set myCell = myCell.Offset(1, 0)
Loop

Report.SaveAs fileName:="P:\CSG\BusApps\Common\Trevorp\Reports\CanAM_Reports\CanAmCMS_UserAudit\ACM CMS User Report" & reportMonth & reportYear

'Application.ScreenUpdating = True
MsgBox ("Procedure is Complete")

End Sub
