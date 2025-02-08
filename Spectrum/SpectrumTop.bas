Attribute VB_Name = "my_SpectrumTop"
'Ќайти более точно вершину спектра по 3-м точкам (массивы: Freq - частота, Amp - амплитуда)
Public Function SpectrumTop(Freq As Variant, Amp As Variant) As Variant
    'Const maxRow = 3
    Dim FreqUp, FreqDown As Boolean
    Dim arrResult(1 To 1, 1 To 2) As Double
    Dim i As Byte
    Dim maxAmp As Double
    Dim k1, k2, kTop, bTop, kRes, bRes As Double
    'точка в стороне (через которую проводим зеркальную пр€мую)
    Dim Dot3X, Dot3Y As Double
    'точка вершины равнобедренного треугольника (полученного из пр€мой и ее зеркальной)
    Dim DotTopX, DotTopY As Double
    
    'ѕровер€ем входные данные (количество и чтобы были числами).
    SpectrumTop = CVErr(xlErrNA)
    
    If data_err(Freq) Or data_err(Amp) Then
        Exit Function
    End If
    
    'ѕровер€ем входные данные на соответствие задачи.
    '  ѕолное равенство амплитуд недопустимо (запрет расчета).
    If (Amp(2, 1) = Amp(1, 1)) And (Amp(2, 1) = Amp(3, 1)) Then
        Exit Function
    End If
    
    '  „астота должна быть упор€дочена (расчет можно продолжать).
    FreqUp = (Freq(2, 1) > Freq(1, 1)) And (Freq(2, 1) < Freq(3, 1))
    FreqDown = (Freq(2, 1) < Freq(1, 1)) And (Freq(2, 1) > Freq(3, 1))
    If Not (FreqUp Or FreqDown) Then
        Exit Function
    End If
    
    
    '–ешаем задачу.
    'ѕримечание: уравнение пр€мых y=k*x+b, (x-x1)/(x2-x1)=(y-y1)/(y2-y1)
    'x - это Freq, y - это Amp
    '  ќпредел€ем круто понимающуюс€/спускающуюс€ пр€мую и точку в стороне.
    k1 = (Amp(2, 1) - Amp(1, 1)) / (Freq(2, 1) - Freq(1, 1))
    k2 = (Amp(3, 1) - Amp(2, 1)) / (Freq(3, 1) - Freq(2, 1))
    If Abs(k1) > Abs(k2) Then
        kTop = k1
        Dot3X = Freq(3, 1)
        Dot3Y = Amp(3, 1)
    Else
        kTop = k2
        Dot3X = Freq(1, 1)
        Dot3Y = Amp(1, 1)
    End If
    bTop = Amp(2, 1) - kTop * Freq(2, 1)
    
    '  «екрально отражаем пр€мую. ѕроводим зеркальную через точку в стороне
    '  и находим точку пересечени€ пр€мых (вершина равнобедренного треугольника).
    DotTopX = (kTop * Dot3X + Dot3Y - bTop) / 2 / kTop
    DotTopY = kTop * DotTopX + bTop
    
    '  Y вершины меньше максимального Y данных (запрет расчета).
    maxAmp = 0
    For i = 1 To 3
        If maxAmp < Amp(i, 1) Then maxAmp = Amp(i, 1)
    Next
    
    If DotTopY < maxAmp Then
        Exit Function
    End If
    
    '  Y вершины совпадает с максимальным Y-ком данных,
    '  то вершина это ответ задачи.
    If DotTopY = maxAmp Then
        arrResult(1, 1) = DotTopX
        arrResult(1, 2) = DotTopY
    Else
        '  ќпредел€ем пр€мую проход€щую через точку в стороне
        '  и точку с координатами (X центральной точки, Y вершины треугольника).
        kRes = (DotTopY - Dot3Y) / (Freq(2, 1) - Dot3X)
        bRes = Dot3Y - kRes * Dot3X
        
        '  Ќаходим точку на пересечении найденой пр€мой
        '  и вертикали проход€щей через вершину треугольника (ответ).
        arrResult(1, 1) = DotTopX
        arrResult(1, 2) = kRes * DotTopX + bRes
    End If
    
    SpectrumTop = arrResult
End Function

Private Function data_err(x As Variant) As Boolean
    'ѕроверка данных с листа книги.
    '3 строки, 1 столбец, числа - это нормально.
    Const maxRow = 3, col = 1
    Dim row As Long
    
    data_err = True
    
    If Not (IsArray(x)) Then
        Exit Function
    Else
        x = x.Value
    End If
    
    If (UBound(x, 1) <> maxRow) Or (UBound(x, 2) <> col) Then
        Exit Function
    End If
    
    For row = 1 To maxRow
        If Not (IsNumeric(x(row, col))) Then
            Exit Function
        End If
    Next
    
    data_err = False
End Function
