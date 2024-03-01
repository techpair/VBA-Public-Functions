Function GenerateCustomDate(inputDate As Date, excludeWeekends As Boolean, Optional dateFormat As String = "yyyy-mm-dd") As String
    Dim adjustedDate As String

    ' Check if user wants to exclude weekends
    If excludeWeekends Then
        ' Exclude Saturdays (vbSaturday = 7) and Sundays (vbSunday = 1)
        While Weekday(inputDate) = vbSaturday Or Weekday(inputDate) = vbSunday
            inputDate = inputDate + IIf(Weekday(inputDate) = vbSaturday, 2, 1)
        Wend
    End If

    ' Format the date using the specified date format
    adjustedDate = Format(inputDate, dateFormat)

    ' Return the formatted date string
    GenerateCustomDate = adjustedDate
End Function
