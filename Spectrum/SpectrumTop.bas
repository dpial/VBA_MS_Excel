Attribute VB_Name = "my_SpectrumTop"
'����� ����� ����� ������� ������� �� 3-� ������ (�������: Freq - �������, Amp - ���������)
Public Function SpectrumTop(Freq As Variant, Amp As Variant) As Variant
    'Const maxRow = 3
    Dim FreqUp, FreqDown As Boolean
    Dim arrResult(1 To 1, 1 To 2) As Double
    Dim i As Byte
    Dim maxAmp As Double
    Dim k1, k2, kTop, bTop, kRes, bRes As Double
    '����� � ������� (����� ������� �������� ���������� ������)
    Dim Dot3X, Dot3Y As Double
    '����� ������� ��������������� ������������ (����������� �� ������ � �� ����������)
    Dim DotTopX, DotTopY As Double
    
    '��������� ������� ������ (���������� � ����� ���� �������).
    SpectrumTop = CVErr(xlErrNA)
    
    If data_err(Freq) Or data_err(Amp) Then
        Exit Function
    End If
    
    '��������� ������� ������ �� ������������ ������.
    '  ������ ��������� �������� ����������� (������ �������).
    If (Amp(2, 1) = Amp(1, 1)) And (Amp(2, 1) = Amp(3, 1)) Then
        Exit Function
    End If
    
    '  ������� ������ ���� ����������� (������ ����� ����������).
    FreqUp = (Freq(2, 1) > Freq(1, 1)) And (Freq(2, 1) < Freq(3, 1))
    FreqDown = (Freq(2, 1) < Freq(1, 1)) And (Freq(2, 1) > Freq(3, 1))
    If Not (FreqUp Or FreqDown) Then
        Exit Function
    End If
    
    
    '������ ������.
    '����������: ��������� ������ y=k*x+b, (x-x1)/(x2-x1)=(y-y1)/(y2-y1)
    'x - ��� Freq, y - ��� Amp
    '  ���������� ����� ������������/������������ ������ � ����� � �������.
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
    
    '  ��������� �������� ������. �������� ���������� ����� ����� � �������
    '  � ������� ����� ����������� ������ (������� ��������������� ������������).
    DotTopX = (kTop * Dot3X + Dot3Y - bTop) / 2 / kTop
    DotTopY = kTop * DotTopX + bTop
    
    '  Y ������� ������ ������������� Y ������ (������ �������).
    maxAmp = 0
    For i = 1 To 3
        If maxAmp < Amp(i, 1) Then maxAmp = Amp(i, 1)
    Next
    
    If DotTopY < maxAmp Then
        Exit Function
    End If
    
    '  Y ������� ��������� � ������������ Y-��� ������,
    '  �� ������� ��� ����� ������.
    If DotTopY = maxAmp Then
        arrResult(1, 1) = DotTopX
        arrResult(1, 2) = DotTopY
    Else
        '  ���������� ������ ���������� ����� ����� � �������
        '  � ����� � ������������ (X ����������� �����, Y ������� ������������).
        kRes = (DotTopY - Dot3Y) / (Freq(2, 1) - Dot3X)
        bRes = Dot3Y - kRes * Dot3X
        
        '  ������� ����� �� ����������� �������� ������
        '  � ��������� ���������� ����� ������� ������������ (�����).
        arrResult(1, 1) = DotTopX
        arrResult(1, 2) = kRes * DotTopX + bRes
    End If
    
    SpectrumTop = arrResult
End Function

Private Function data_err(x As Variant) As Boolean
    '�������� ������ � ����� �����.
    '3 ������, 1 �������, ����� - ��� ���������.
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
