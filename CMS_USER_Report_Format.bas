Attribute VB_Name = "CMS_USER_Report_Format"
Sub CMS_UserReportFormat()
Attribute CMS_UserReportFormat.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CMS_UserReportFormat Macro
'

'
Set wb = Workbooks("CanAmCMS_UserAudit April 2018.xlsx")

For i = 1 To 12
    With wb.Worksheets(i)
        wsName = .Name
            If wsName = "Error Report" Or wsName = "CMS User Report" Then
                GoTo skipProcess
                Else
                'nothing
            End If
        With .Range("A1:C1")
            .Interior.Color = RGB(100, 100, 100)
            .Font.Color = RGB(255, 255, 255)
        End With
        For j = 1 To 5
            .Rows(1).EntireRow.Insert
        Next j
        .Columns(1).EntireColumn.Insert
        With .Cells(2, 2)
            .Value = "CanAmCMS " & wsName & " User Report"
            .Font.Size = 11
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        With .Cells(3, 2)
            .Value = "CanAm Insurance"
            .HorizontalAlignment = xlCenter
        End With
        With .Cells(4, 2)
            .Value = Format(Now(), "yyyy-mm-dd")
            .HorizontalAlignment = xlCenter
        End With
            .Range("b2:d2").Merge
            .Range("b3:d3").Merge
            .Range("b4:d4").Merge
        For n = 2 To 4
            .Columns(n).AutoFit
        Next n
        .Copy '<-- Move
    End With

If Month(Now()) < 10 Then
    monthNum = 0 & Month(Now())
    Else
    monthNum = Month(Now())
End If
 
 
ActiveWorkbook.SaveAs fileName:="P:\CSG\BusApps\Common\Trevorp\Reports\CanAM_Reports\CanAmCMS_UserAudit\CanAmCMS_201804\" & wsName & "_" & Year(Now()) & monthNum
    
skipProcess:
Next i

End Sub
