Attribute VB_Name = "my_windowHamming"
'���� ��������. ������� "0.53836 - 0.46164*Cos(2*Pi*n/(N-1))".
'  �������� ���������� ����� (���� �����).
'  ����������: ��������� ����� ���� ��� ������ ���� ����� ����.
Public Function windowHamming( _
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
    Dim wHamming() As Double
    
    '�������� ������� ������ � ������ ��������� ����� ���� ��� ��� ����� ������ 3-�.
    '  #�/�, ���� ����� ������ 1.
    If totalPoints < 1 Then
        windowHamming = CVErr(xlErrNA)
        Exit Function
    End If
    
    totalPointsMinusOne = totalPoints - 1
    
    '  ������ ��������� ����� ���� (���� ��� �������).
    If Not (IsNull(numberPoint)) Then
        numberPoint = CLng(numberPoint)
        If (numberPoint >= 0) And (numberPoint <= totalPointsMinusOne) Then
            If totalPointsMinusOne = 0 Then
                windowHamming = 0.53836 - 0.46164
            Else
                windowHamming = 0.53836 - 0.46164 * Cos(twoPi * numberPoint / totalPointsMinusOne)
            End If
        Else
            windowHamming = CVErr(xlErrNA)
        End If
        Exit Function
    End If
    
    
    ReDim wHamming(0 To totalPointsMinusOne, 0 To 0)
    
    '  ������, ���� ����� ������ 3-�.
    If totalPoints < 3 Then
        For i = 0 To totalPointsMinusOne
            wHamming(i, 0) = 0.53836 - 0.46164
        Next
        windowHamming = wHamming
        Exit Function
    End If
    
    '�������� ������ (����� ����� ������ 3-� � ��� ��������� �����).
    isOdd = totalPoints Mod 2 = 1
    
    '  ���� ���������� ����� �������� �����, �� ����������� ����� ����� 1.
    If isOdd Then
        endFirstPart = totalPointsMinusOne \ 2 - 1
        wHamming(endFirstPart + 1, 0) = 1
        beginSecondPart = endFirstPart + 2
    Else
        endFirstPart = totalPoints \ 2 - 1
        beginSecondPart = endFirstPart + 1
    End If
    
    '  ������ �������� ������, ������������� �� �������. ��������� ��������� �����������.
    '    ������ ������ ��������.
    For i = 0 To endFirstPart
        wHamming(i, 0) = 0.53836 - 0.46164 * Cos(twoPi * i / totalPointsMinusOne)
    Next
    
    '    ���������� ���������� ���������.
    For i = beginSecondPart To totalPointsMinusOne
        wHamming(i, 0) = wHamming(endFirstPart - (i - beginSecondPart), 0)
    Next
    
    windowHamming = wHamming
End Function
