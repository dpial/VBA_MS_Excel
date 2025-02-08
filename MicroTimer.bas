Attribute VB_Name = "MicroTimer"

'������ �������� ����������
'  ��� ������ �������� ���������� ����

'  ��������� PtrSafe ��� 64 ������� excel "Declare PtrSafe Function"
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Function MicroTimer() As Double
    Dim cyTicks1 As Currency
    Static cyFrequency As Currency
    '��������� ������������
    MicroTimer = 0
    '��������� �������, ������ ���������� ������ (������)
    '�������, ����������� � ���������, � �������
    If cyFrequency = 0 Then getFrequency cyFrequency
    '��������� ���������� ������ (������) ������� � ��������
    '�� �������� �������
    getTickCount cyTicks1
    '���������� ���������� ������, ��������� � ��������, ���
    '���������� ������ = ���������� ������ / �������
    If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency
End Function

'������ �������������
'Public Function aaa(a As Variant) As Variant
'    Dim StartTime As Double, ElapsedTime As Double
'
'    StartTime = MicroTimer
'
'    '����� ����� ���, �������� �������� ��������
'
'    ElapsedTime = MicroTimer - StartTime
'
'    '����� ����������
'    Debug.Print CDec(ElapsedTime)
'End Function

