' Copyright (c) Microsoft. All rights reserved.
' Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.

Option Explicit

Const OlFolderCalendar = 9
Const OlAppointment = 1
Const OlFree = 0

Dim appOlk
Dim nmSession
Dim fldCalendar

'Year 2021
Dim arrNormal21 : arrNormal21 = Array("New Year's Day,1,1","National Foundation Day,2,11","Emperor's Birthday,2,23","Shōwa Day,4,29","Constitution Day,5,3","Greenery Day,5,4","Children's Day,5,5","Marine Day,7,19,7,22","Sports Day,10,11,7,23","Mountain Day,8,11,8,8","Culture Day,11,3","Labor Thanksgiving Day,11,23","New Year's Eve,12,31")
'Year 2022
Dim arrNormal22 : arrNormal22 = Array("New Year's Day,1,1","National Foundation Day,2,11","Emperor's Birthday,2,23","Shōwa Day,4,29","Constitution Day,5,3","Greenery Day,5,4","Children's Day,5,5","Mountain Day,8,11","Culture Day,11,3","Labor Thanksgiving Day,11,23","New Year's Eve,12,31")

Dim arrHappyMon2021 : arrHappyMon2021 = Array("Coming of Age Day,1,2","Respect for the Aged Day,9,3")
Dim arrHappyMon2022: arrHappyMon2022 = Array("Coming of Age Day,1,2","Marine Day,7,3","Respect for the Aged Day,9,3","Sports Day,10,2")

'Vernal Equinox Day
Dim arrAEquinox : arrAEquinox = Array(20,21,21,20,20,20,21,20,20,20)
'Autumnal Equinox Day
Dim arrVEquinox : arrVEquinox = Array(23,23,23,22,23,23,23,22,23,23)
'Observed
Dim arrSubHoliday : arrSubHoliday = Array("Mountain Day,2021/8/9","New Year's Day,2023/1/2","National Foundation Day,2024/2/12","Children's Day,2024/5/6","Mountain Day,2024/8/12","Autumnal Equinox Day,2024/9/23","Emperor's Birthday,2025/2/24","Greenery Day,2025/5/6","Labor Thanksgiving Day,2025/11/24","Children's Day,2026/5/6","Vernal Equinox Day,2027/3/22","National Foundation Day,2029/2/12","Shōwa Day,2029/4/30","Autumnal Equinox Day,2029/9/24","Children's Day,2030/5/6","Mountain Day,2030/8/12","Culture Day,2030/11/4")

Dim rslt
rslt = MsgBox("Add holiday?", 68)

If rslt = 6 then
    Set appOlk = CreateObject("Outlook.Application")
    Set nmSession = appOlk.GetNamespace("MAPI")
    Set fldCalendar = nmSession.GetDefaultFolder(OlFolderCalendar)

    Dim iYear

    For iYear = 2019 to 2028
        Dim itmOld
        ' Delete old Emperor's Birthday first
        Set itmOld = fldCalendar.Items.Find( _
            "[CATEGORIES] = 'Holiday' AND [LOCATION] = 'Japan' AND [SUBJECT] = 'Emperor''s Birthday' AND [START] >= '" & iYear & "/12/01' AND [END] <= '" & iYear & "/12/31'")

        If Not itmOld Is Nothing Then
            itmOld.Delete
        End If
    Next

    For iYear = 2021 to 2030
        Dim i
        Dim strName
        Dim arrRec

        ' Regular Holiday
        If iYear = 2021 then
            AddNormalHolidays iYear, arrNormal21
        Else
            AddNormalHolidays iYear, arrNormal22
        End If

        ' Happy Monday
        If iYear = 2021 then
            AddHappyMondays iYear, arrHappyMon2021
        Else
            AddHappyMondays iYear, arrHappyMon2022
        End If

        ' Irregular Holiday
        AddOneHoliday "Vernal Equinox Day", iYear & "/3/" & arrAEquinox(iYear - 2021)
        AddOneHoliday "Autumnal Equinox Day", iYear & "/9/" & arrVEquinox(iYear - 2021)

    Next
    
    ' Observed
    For i = LBound(arrSubHoliday) To UBound(arrSubHoliday)
        arrRec = Split(arrSubHoliday(i),",")
        AddOneHoliday "" & arrRec(0) & " (Observed)", arrRec(1)
    Next

    ' Sandwiched between Respect for the Aged Day and Autumnal Equinox Day
    AddOneHoliday "People's Day", "2026/9/22"
    MsgBox "Finished adding Holidays."
End If

' Add regular holiday
Sub AddNormalHolidays( iYear, arrHolidays )
    Dim i
    Dim arrRec

    For i = LBound(arrHolidays) To UBound(arrHolidays)
        arrRec = Split(arrHolidays(i),",")

        If UBound(arrRec) = 4 Then
            MoveHoliday arrRec(0), iYear & "/" & arrRec(1) & "/" & arrRec(2), iYear & "/" & arrRec(3) & "/" & arrRec(4)
        Else
            AddOneHoliday arrRec(0), iYear & "/" & arrRec(1) & "/" & arrRec(2)
        End If
    Next
End Sub

' Add Happy Monday
Sub AddHappyMondays( iYear, arrHappyMon )
    Dim i
    Dim arrRec
    Dim iWeek
    Dim iDay

    For i = LBound(arrHappyMon) To UBound(arrHappyMon)
        arrRec = Split(arrHappyMon(i),",")
        iWeek = Weekday(iYear & "/" & arrRec(1) & "/1", vbTuesday) - 1
        iDay = 7 * CInt(arrRec(2)) - iWeek
        AddOneHoliday arrRec(0), iYear & "/" & arrRec(1) & "/" & iDay
    Next
End Sub

' Add holiday
Sub AddOneHoliday( strName, strHoliday )
    Dim itmOneHoliday
    
    Set itmOneHoliday = fldCalendar.Items.Find( _
        "[CATEGORIES] = 'Holiday' AND [LOCATION] = 'Japan' AND [START] >= '" & _
        strHoliday & " 00:00 AM' AND [END] <= '" & _
        DateAdd("d", CDate(strHoliday), 1) & "'")
        
    ' WScript.Echo "[CATEGORIES] = 'Holiday' AND [LOCATION] = 'Japan' AND [START] >= '" & _
    '    strHoliday & " 00:00 AM' AND [END] <= '" & _
    '    DateAdd("d", CDate(strHoliday), 1) & "'"
        
    If itmOneHoliday Is Nothing Then
        Set itmOneHoliday = appOlk.CreateItem(OlAppointment)
        itmOneHoliday.Subject = strName
        itmOneHoliday.Start = strHoliday
        itmOneHoliday.AllDayEvent = True
        itmOneHoliday.BusyStatus = OlFree
        itmOneHoliday.ReminderSet = False
        itmOneHoliday.Location = "Japan"
        itmOneHoliday.Categories = "Holiday"
        itmOneHoliday.Save
        ' WScript.Echo "added: " & strHoliday & itmOneHoliday.Subject
    Else
        ' WScript.Echo "found: " & strHoliday & itmOneHoliday.Subject
    End If
End Sub

' Move holiday
Sub MoveHoliday( strName, strOldDay, strNewDay )
    Dim itmOneHoliday
    
    Set itmOneHoliday = fldCalendar.Items.Find( _
        "[CATEGORIES] = 'Holiday' AND [LOCATION] = 'Japan' AND [START] >= '" & _
        strOldDay & " 00:00 AM' AND [END] <= '" & _
        DateAdd("d", CDate(strOldDay), 1) & "'")

    If itmOneHoliday Is Nothing Then
        Set itmOneHoliday = appOlk.CreateItem(OlAppointment)
    End If

    itmOneHoliday.Subject = strName
    itmOneHoliday.Start = strNewDay
    itmOneHoliday.AllDayEvent = True
    itmOneHoliday.BusyStatus = OlFree
    itmOneHoliday.ReminderSet = False
    itmOneHoliday.Location = "Japan"
    itmOneHoliday.Categories = "Holiday"
    itmOneHoliday.Save
End Sub
