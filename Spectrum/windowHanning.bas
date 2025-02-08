Attribute VB_Name = "my_windowHanning"
'���� ��������(�����). ������� "0.5*( 1 - Cos(2*Pi*n/(N-1)) )".
'  �������� ���������� ����� (���� �����).
'  ����������: ��������� ����� ���� ��� ������ ���� ����� ����.
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
    '������ ��� ����������.
    Dim wHanning() As Double
    
    '�������� ������� ������ � ������ ��������� ����� ���� ��� ��� ����� ������ 3-�.
    '  #�/�, ���� ����� ������ 1.
    If totalPoints < 1 Then
        windowHanning = CVErr(xlErrNA)
        Exit Function
    End If
    
    totalPointsMinusOne = totalPoints - 1
    
    '  ������ ��������� ����� ����.
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
    
    '  ������, ���� ����� ������ 3-�.
    If totalPoints < 3 Then
        For i = 0 To totalPointsMinusOne
            wHanning(i, 0) = 0
        Next
        windowHanning = wHanning
        Exit Function
    End If
    
    '�������� ������ (����� ����� ������ 3-� � ��� ��������� �����).
    isOdd = totalPoints Mod 2 = 1
    
    '  ���� ���������� ����� �������� �����, �� ����������� ����� ����� 1.
    If isOdd Then
        endFirstPart = totalPointsMinusOne \ 2 - 1
        wHanning(endFirstPart + 1, 0) = 1
        beginSecondPart = endFirstPart + 2
    Else
        endFirstPart = totalPoints \ 2 - 1
        beginSecondPart = endFirstPart + 1
    End If
    
    '  ������ �������� ������, ������������� �� �������. ��������� ��������� �����������.
    '    ������ ������ ��������.
    For i = 0 To endFirstPart
        wHanning(i, 0) = 0.5 * (1 - Cos(twoPi * i / totalPointsMinusOne))
    Next
    
    '    ���������� ���������� ���������.
    For i = beginSecondPart To totalPointsMinusOne
        wHanning(i, 0) = wHanning(endFirstPart - (i - beginSecondPart), 0)
    Next
    
    windowHanning = wHanning
End Function
