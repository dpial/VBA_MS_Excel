Attribute VB_Name = "my_isVolatile"
'isVolatile ������� ���������� ��������� ����� ����.
'  ��������� ������ � isVolatile ������� ������ ���������.
'  isVolatile ������ ���� �������.
'  ����������: ���� ������� ������� ��������������� ������ ���,
'    ���� ���� ������ �� ���������� (Volatile functions).
Public Function isVolatile(x As Variant) As Integer
    Const maxCount = 32767
    Static count As Integer
    
    If count >= maxCount Then count = 0
    count = count + 1
    isVolatile = count
    
    '������ �������� � isVolatile, ����� ������������.
    '����� �������� �� �����.
    x = x.Value
End Function
