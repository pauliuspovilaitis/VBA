Function holidays(date1 As Date, date2 As Date) As Double

Dim sventes(0 To 12) As Date
sventes(0) = DateSerial(Year(Now()), 1, 1)
sventes(1) = DateSerial(Year(Now()), 2, 16)
sventes(2) = DateSerial(Year(Now()), 3, 11)
sventes(3) = DateSerial(Year(Now()), 4, 20)
sventes(4) = DateSerial(Year(Now()), 4, 21)
sventes(5) = DateSerial(Year(Now()), 5, 1)
sventes(6) = DateSerial(Year(Now()), 6, 24)
sventes(7) = DateSerial(Year(Now()), 7, 6)
sventes(8) = DateSerial(Year(Now()), 8, 15)
sventes(9) = DateSerial(Year(Now()), 11, 1)
sventes(10) = DateSerial(Year(Now()), 12, 24)
sventes(11) = DateSerial(Year(Now()), 12, 25)
sventes(12) = DateSerial(Year(Now()), 12, 26)

Dim yra As Double
yra = 0
Dim DateLooper As Date

For DateLooper = date1 To date2

         Dim m As Variant
         For m = LBound(sventes) To UBound(sventes)
            If sventes(m) = DateLooper _
                And Weekday(DateLooper) <> vbSaturday And _
                Weekday(DateLooper) <> vbSunday Then
                
                yra = yra + 1
            End If
         Next m
Next DateLooper

holidays = yra

End Function
