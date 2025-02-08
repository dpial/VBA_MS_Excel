Attribute VB_Name = "my_windowHanning"
'Окно Хэннинга(Ханна). Формула "0.5*( 1 - Cos(2*Pi*n/(N-1)) )".
'  Получает количество точек (одно число).
'  Возвращает: указанную точку окна или массив всех точек окна.
Public Function windowHanning( _
        totalPoints As Long, _
        Optional numberPoint As Variant = Null _
    ) As Variant
    
    Const twoPi = 6.28318530717959
    Dim isOdd As Boolean
    Dim endFirstPart As Long
    Dim beginSecondPart As Long
    Dim i As Long
    Dim totalPointsMinusOne As Long
    'Массив для результата.
    Dim wHanning() As Double
    
    'Проверка входных данных и расчет указанной точки окна или для точек меньше 3-х.
    '  #Н/Д, если точек меньше 1.
    If totalPoints < 1 Then
        windowHanning = CVErr(xlErrNA)
        Exit Function
    End If
    
    totalPointsMinusOne = totalPoints - 1
    
    '  Расчет указанной точки окна.
    If Not (IsNull(numberPoint)) Then
        numberPoint = CLng(numberPoint)
        If (numberPoint >= 0) And (numberPoint <= totalPointsMinusOne) Then
            If totalPointsMinusOne = 0 Then
                windowHanning = 0
            Else
                windowHanning = 0.5 * (1 - Cos(twoPi * numberPoint / totalPointsMinusOne))
            End If
        Else
            windowHanning = CVErr(xlErrNA)
        End If
        Exit Function
    End If
    
    
    ReDim wHanning(0 To totalPointsMinusOne, 0 To 0)
    
    '  Расчет, если точек меньше 3-х.
    If totalPoints < 3 Then
        For i = 0 To totalPointsMinusOne
            wHanning(i, 0) = 0
        Next
        windowHanning = wHanning
        Exit Function
    End If
    
    'Основной расчет (когда точек больше 3-х и нет указанной точки).
    isOdd = totalPoints Mod 2 = 1
    
    '  Если количество точек нечетное число, то центральная точка равна 1.
    If isOdd Then
        endFirstPart = totalPointsMinusOne \ 2 - 1
        wHanning(endFirstPart + 1, 0) = 1
        beginSecondPart = endFirstPart + 2
    Else
        endFirstPart = totalPoints \ 2 - 1
        beginSecondPart = endFirstPart + 1
    End If
    
    '  Первая половина данных, расчитывается по формуле. Остальные зеркально заполняются.
    '    расчет первой половины.
    For i = 0 To endFirstPart
        wHanning(i, 0) = 0.5 * (1 - Cos(twoPi * i / totalPointsMinusOne))
    Next
    
    '    зеркальное заполнение остальных.
    For i = beginSecondPart To totalPointsMinusOne
        wHanning(i, 0) = wHanning(endFirstPart - (i - beginSecondPart), 0)
    Next
    
    windowHanning = wHanning
End Function
