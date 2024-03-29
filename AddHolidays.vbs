﻿' Copyright (c) Microsoft. All rights reserved.
' Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.

Option Explicit

Const OlFolderCalendar = 9
Const OlAppointment = 1
Const OlFree = 0

Dim appOlk
Dim nmSession
Dim fldCalendar

' 2021 年以降用
Dim arrNormal21 : arrNormal21 = Array("元日,1,1","建国記念の日,2,11","天皇誕生日,2,23","昭和の日,4,29","憲法記念日,5,3","みどりの日,5,4","こどもの日,5,5","海の日,7,19,7,22","スポーツの日,10,11,7,23","山の日,8,11,8,8","文化の日,11,3","勤労感謝の日,11,23","大みそか,12,31")
' 2022 年以降用
Dim arrNormal22 : arrNormal22 = Array("元日,1,1","建国記念の日,2,11","天皇誕生日,2,23","昭和の日,4,29","憲法記念日,5,3","みどりの日,5,4","こどもの日,5,5","山の日,8,11","文化の日,11,3","勤労感謝の日,11,23","大みそか,12,31")

Dim arrHappyMon2021 : arrHappyMon2021 = Array("成人の日,1,2","敬老の日,9,3")
Dim arrHappyMon2022: arrHappyMon2022 = Array("成人の日,1,2","海の日,7,3","敬老の日,9,3","スポーツの日,10,2")

'春分の日
Dim arrAEquinox : arrAEquinox = Array(20,21,21,20,20,20,21,20,20,20)
'秋分の日
Dim arrVEquinox : arrVEquinox = Array(23,23,23,22,23,23,23,22,23,23)
'振替休日
Dim arrSubHoliday : arrSubHoliday = Array("山の日,2021/8/9","元旦,2023/1/2","建国記念の日,2024/2/12","こどもの日,2024/5/6","山の日,2024/8/12","秋分の日,2024/9/23","天皇誕生日,2025/2/24","みどりの日,2025/5/6","勤労感謝の日,2025/11/24","こどもの日,2026/5/6","春分の日,2027/3/22","建国記念の日,2029/2/12","昭和の日,2029/4/30","秋分の日,2029/9/24","こどもの日,2030/5/6","山の日,2030/8/12","文化の日,2030/11/4")

Dim rslt
rslt = MsgBox("祝日を追加しますか?", 68)

If rslt = 6 then
    Set appOlk = CreateObject("Outlook.Application")
    Set nmSession = appOlk.GetNamespace("MAPI")
    Set fldCalendar = nmSession.GetDefaultFolder(OlFolderCalendar)

    Dim iYear

    For iYear = 2019 to 2028
        Dim itmOld
        ' 古い天皇誕生日を先に削除
        Set itmOld = fldCalendar.Items.Find( _
            "[分類項目] = '祝日' AND [場所] = '日本' AND [件名] = '天皇誕生日' AND [開始日] >= '" & iYear & "/12/01' AND [終了日] <= '" & iYear & "/12/31'")

        If Not itmOld Is Nothing Then
            itmOld.Delete
        End If
    Next

    For iYear = 2021 to 2030
        Dim i
        Dim strName
        Dim arrRec

        ' 通常の祝日
        If iYear = 2021 then
            AddNormalHolidays iYear, arrNormal21
        Else
            AddNormalHolidays iYear, arrNormal22
        End If

        ' ハッピーマンデー
        If iYear = 2021 then
            AddHappyMondays iYear, arrHappyMon2021
        Else
            AddHappyMondays iYear, arrHappyMon2022
        End If

        ' 不定期の祝日
        AddOneHoliday "春分の日", iYear & "/3/" & arrAEquinox(iYear - 2021)
        AddOneHoliday "秋分の日", iYear & "/9/" & arrVEquinox(iYear - 2021)

    Next
    
    ' 振替休日
    For i = LBound(arrSubHoliday) To UBound(arrSubHoliday)
        arrRec = Split(arrSubHoliday(i),",")
        AddOneHoliday "振替休日 (" & arrRec(0) & ")", arrRec(1)
    Next

    ' 敬老の日と秋分の日に挟まれるため
    AddOneHoliday "国民の祝日", "2026/9/22"
    MsgBox("祝日を追加しました。")
End If

' 配列から通常の祝日を追加
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

' 配列からハッピーマンデーを追加
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

' 祝日がなければ追加
Sub AddOneHoliday( strName, strHoliday )
    Dim itmOneHoliday
    
    Set itmOneHoliday = fldCalendar.Items.Find( _
        "[分類項目] = '祝日' AND [場所] = '日本' AND [開始日] >= '" & _
        strHoliday & " 00:00 AM' AND [終了日] <= '" & _
        DateAdd("d", CDate(strHoliday), 1) & "'")
        
    ' WScript.Echo "[分類項目] = '祝日' AND [場所] = '日本' AND [開始日] >= '" & _
    '    strHoliday & " 00:00 AM' AND [終了日] <= '" & _
    '    DateAdd("d", CDate(strHoliday), 1) & "'"
        
    If itmOneHoliday Is Nothing Then
        Set itmOneHoliday = appOlk.CreateItem(OlAppointment)
        itmOneHoliday.Subject = strName
        itmOneHoliday.Start = strHoliday
        itmOneHoliday.AllDayEvent = True
        itmOneHoliday.BusyStatus = OlFree
        itmOneHoliday.ReminderSet = False
        itmOneHoliday.Location = "日本"
        itmOneHoliday.Categories = "祝日"
        itmOneHoliday.Save
        ' WScript.Echo "added: " & strHoliday & itmOneHoliday.Subject
    Else
        ' WScript.Echo "found: " & strHoliday & itmOneHoliday.Subject
    End If
End Sub

' 祝日を移動
Sub MoveHoliday( strName, strOldDay, strNewDay )
    Dim itmOneHoliday
    
    Set itmOneHoliday = fldCalendar.Items.Find( _
        "[分類項目] = '祝日' AND [場所] = '日本' AND [開始日] >= '" & _
        strOldDay & " 00:00 AM' AND [終了日] <= '" & _
        DateAdd("d", CDate(strOldDay), 1) & "'")

    If itmOneHoliday Is Nothing Then
        Set itmOneHoliday = appOlk.CreateItem(OlAppointment)
    End If

    itmOneHoliday.Subject = strName
    itmOneHoliday.Start = strNewDay
    itmOneHoliday.AllDayEvent = True
    itmOneHoliday.BusyStatus = OlFree
    itmOneHoliday.ReminderSet = False
    itmOneHoliday.Location = "日本"
    itmOneHoliday.Categories = "祝日"
    itmOneHoliday.Save
End Sub
