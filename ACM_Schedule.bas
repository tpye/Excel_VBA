Attribute VB_Name = "ACM_Schedule"

Public Sub QuestionCleanUp()
Dim myRange As Range
Dim myCell As Range

Set myRange = ActiveSheet.Range("i2", ActiveSheet.Range("i1048576").End(xlUp))
Set myCell = ActiveSheet.Range("i2")

For Each myCell In myRange
    Select Case myCell.Value
        Case "A) to work a stretch of the same shift during consecutive work days in order to maintain a more consistent schedule."
            myCell.Value = "A) Same Shift"
        Case "B) not to have the same shift during a stretch of consecutive work days, if the fluctuating schedule can result in working fewer of my less desired shifts."
            myCell.Value = "B)Varied Schedule"
        Case "B) varied schedule"
            myCell.Value = "B)Varied Schedule"
        Case Else
            'nothing
        End Select
Next myCell

MsgBox "Complete"

End Sub

Public Sub Chritmas_List()
Dim wsHoliday As Worksheet
Dim ws As Worksheet
Dim myRange As Range
Dim myCell As Range
Dim MyCell2 As Range

Set wsHoliday = Workbooks("2016 Holiday Work Preference Survey_ CSC (Responses).xlsx").Worksheets("Form Responses 1")
Set ws = Workbooks("HolidayLists.xlsx").Worksheets("Christmas")
'Set myCell2 = ws.Range("A1048576").End(xlUp)
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [1st preference (most preferred day to work)]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Sunday, December 25th (Christmas Day)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "1st"
            .Offset(0, 1).Value = myCell.Offset(0, -7) & " " & myCell.Offset(0, -6)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -5)
            .Offset(0, 3).Value = myCell.Offset(0, -4).Value
            .Offset(0, 4).Value = myCell.Offset(0, 8).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 10).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 12).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 14).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 16).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell

' SECOND PREFERENCE
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [2nd preference]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found on second preference"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Sunday, December 25th (Christmas Day)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "2nd"
            .Offset(0, 1).Value = myCell.Offset(0, -8) & " " & myCell.Offset(0, -7)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -6)
            .Offset(0, 3).Value = myCell.Offset(0, -5).Value
            .Offset(0, 4).Value = myCell.Offset(0, 7).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 9).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 11).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 13).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 15).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell

' THIRD PREFERENCE
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [3rd preference]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found on 3rd preference"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Sunday, December 25th (Christmas Day)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "3rd"
            .Offset(0, 1).Value = myCell.Offset(0, -9) & " " & myCell.Offset(0, -8)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -7)
            .Offset(0, 3).Value = myCell.Offset(0, -6).Value
            .Offset(0, 4).Value = myCell.Offset(0, 6).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 8).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 10).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 12).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 14).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell

' FOURTH PREFERENCE
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [4th preference]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found on 4th preference"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Sunday, December 25th (Christmas Day)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "4th"
            .Offset(0, 1).Value = myCell.Offset(0, -10) & " " & myCell.Offset(0, -9)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -8)
            .Offset(0, 3).Value = myCell.Offset(0, -7).Value
            .Offset(0, 4).Value = myCell.Offset(0, 5).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 7).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 9).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 11).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 13).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell

' FIFTH PREFERENCE
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [5th preference]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found on 5th preference"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Sunday, December 25th (Christmas Day)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "5th"
            .Offset(0, 1).Value = myCell.Offset(0, -11) & " " & myCell.Offset(0, -10)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -9)
            .Offset(0, 3).Value = myCell.Offset(0, -8).Value
            .Offset(0, 4).Value = myCell.Offset(0, 4).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 6).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 8).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 10).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 12).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell


' SIXTH PREFERENCE
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [6th preference]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found on 6th preference"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Sunday, December 25th (Christmas Day)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "6th"
            .Offset(0, 1).Value = myCell.Offset(0, -12) & " " & myCell.Offset(0, -11)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -10)
            .Offset(0, 3).Value = myCell.Offset(0, -9).Value
            .Offset(0, 4).Value = myCell.Offset(0, 3).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 5).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 7).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 9).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 11).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell


' SEVENTH PREFERENCE
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [7th preference (least preferred day to work)]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found on 7th preference"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Sunday, December 25th (Christmas Day)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "7th"
            .Offset(0, 1).Value = myCell.Offset(0, -13) & " " & myCell.Offset(0, -12)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -11)
            .Offset(0, 3).Value = myCell.Offset(0, -10).Value
            .Offset(0, 4).Value = myCell.Offset(0, 2).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 4).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 6).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 8).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 10).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell


Set myRange = ws.Range("C2", ws.Range("C104875").End(xlUp))
myRange.NumberFormat = "General"



End Sub


Public Sub News_Eve()
Dim wsHoliday As Worksheet
Dim ws As Worksheet
Dim myRange As Range
Dim myCell As Range
Dim MyCell2 As Range

Set wsHoliday = Workbooks("2016 Holiday Work Preference Survey_ CSC (Responses).xlsx").Worksheets("Form Responses 1")
Set ws = Workbooks("HolidayLists.xlsx").Worksheets("New Years Eve")
'Set myCell2 = ws.Range("A1048576").End(xlUp)
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [1st preference (most preferred day to work)]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Saturday, December 31st (New Year's Eve)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "1st"
            .Offset(0, 1).Value = myCell.Offset(0, -7) & " " & myCell.Offset(0, -6)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -5)
            .Offset(0, 3).Value = myCell.Offset(0, -4).Value
            .Offset(0, 4).Value = myCell.Offset(0, 8).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 10).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 12).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 14).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 16).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell

' SECOND PREFERENCE
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [2nd preference]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found on second preference"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Saturday, December 31st (New Year's Eve)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "2nd"
            .Offset(0, 1).Value = myCell.Offset(0, -8) & " " & myCell.Offset(0, -7)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -6)
            .Offset(0, 3).Value = myCell.Offset(0, -5).Value
            .Offset(0, 4).Value = myCell.Offset(0, 7).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 9).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 11).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 13).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 15).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell

' THIRD PREFERENCE
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [3rd preference]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found on 3rd preference"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Saturday, December 31st (New Year's Eve)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "3rd"
            .Offset(0, 1).Value = myCell.Offset(0, -9) & " " & myCell.Offset(0, -8)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -7)
            .Offset(0, 3).Value = myCell.Offset(0, -6).Value
            .Offset(0, 4).Value = myCell.Offset(0, 6).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 8).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 10).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 12).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 14).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell

' FOURTH PREFERENCE
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [4th preference]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found on 4th preference"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Saturday, December 31st (New Year's Eve)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "4th"
            .Offset(0, 1).Value = myCell.Offset(0, -10) & " " & myCell.Offset(0, -9)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -8)
            .Offset(0, 3).Value = myCell.Offset(0, -7).Value
            .Offset(0, 4).Value = myCell.Offset(0, 5).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 7).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 9).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 11).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 13).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell

' FIFTH PREFERENCE
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [5th preference]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found on 5th preference"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Saturday, December 31st (New Year's Eve)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "5th"
            .Offset(0, 1).Value = myCell.Offset(0, -11) & " " & myCell.Offset(0, -10)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -9)
            .Offset(0, 3).Value = myCell.Offset(0, -8).Value
            .Offset(0, 4).Value = myCell.Offset(0, 4).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 6).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 8).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 10).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 12).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell


' SIXTH PREFERENCE
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [6th preference]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found on 6th preference"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Saturday, December 31st (New Year's Eve)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "6th"
            .Offset(0, 1).Value = myCell.Offset(0, -12) & " " & myCell.Offset(0, -11)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -10)
            .Offset(0, 3).Value = myCell.Offset(0, -9).Value
            .Offset(0, 4).Value = myCell.Offset(0, 3).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 5).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 7).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 9).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 11).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell


' SEVENTH PREFERENCE
Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [7th preference (least preferred day to work)]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found on 7th preference"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))


For Each myCell In myRange

    If myCell.Value = "Saturday, December 31st (New Year's Eve)" Then
        Set MyCell2 = ws.Range("A1048576").End(xlUp).Offset(1, 0)
        With MyCell2
            .Value = "7th"
            .Offset(0, 1).Value = myCell.Offset(0, -13) & " " & myCell.Offset(0, -12)
            .Offset(0, 2).Value = Date - myCell.Offset(0, -11)
            .Offset(0, 3).Value = myCell.Offset(0, -10).Value
            .Offset(0, 4).Value = myCell.Offset(0, 2).Value     ' <-- First preference shift
            .Offset(0, 5).Value = myCell.Offset(0, 4).Value    ' <-- Second preference shift
            .Offset(0, 6).Value = myCell.Offset(0, 6).Value    ' <-- Third preference shift
            .Offset(0, 7).Value = myCell.Offset(0, 8).Value    ' <-- Fouth preference shift
            .Offset(0, 8).Value = myCell.Offset(0, 10).Value    ' <-- Last preference shift
        End With
    Else
            ' Nothing
    End If
Next myCell


Set myRange = ws.Range("C2", ws.Range("C104875").End(xlUp))
myRange.NumberFormat = "General"



End Sub



Public Sub columnNumberFind_test()
Dim wsHoliday As Worksheet
Dim ws As Worksheet
Dim myRange As Range
Dim myCell As Range
Dim MyCell2 As Range
Dim column As Integer
Dim Column2 As Integer


Set wsHoliday = Workbooks("2016 Holiday Work Preference Survey_ CSC (Responses).xlsx").Worksheets("Form Responses 1")
Set ws = Workbooks("HolidayLists.xlsx").Worksheets("NewYears Eve")

Set myCell = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [1st preference (most preferred day to work)]", LookIn:=xlValues, LookAt:=xlWhole)
    If myCell Is Nothing Then
            MsgBox "Column Not Found"
            Exit Sub
        Else
            ' nothing
    End If
Set myRange = wsHoliday.Range(myCell, myCell.End(xlDown))

column = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="1-20161231", LookIn:=xlValues, LookAt:=xlWhole).column
Column2 = wsHoliday.Range("A1", wsHoliday.Range("XFD1").End(xlToLeft)).Find(what:="Work Preferences [1st preference (most preferred day to work)]", LookIn:=xlValues, LookAt:=xlWhole).column

MsgBox column
MsgBox Column2





End Sub
