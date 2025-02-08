Attribute VB_Name = "MicroTimer"

'Таймер высокого разрешения
'  для замера скорости выполнения кода

'  поставить PtrSafe для 64 битного excel "Declare PtrSafe Function"
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Function MicroTimer() As Double
    Dim cyTicks1 As Currency
    Static cyFrequency As Currency
    'Обнуление МикроТаймера
    MicroTimer = 0
    'Получение частоты, равной количеству тактов («тиков»)
    'таймера, встроенного в процессор, в секунду
    If cyFrequency = 0 Then getFrequency cyFrequency
    'Получение количества тактов («тиков») таймера с полуночи
    'до текущего времени
    getTickCount cyTicks1
    'Вычисление количества секунд, прошедших с полуночи, где
    'количество секунд = количество тактов / частота
    If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency
End Function

'Пример использования
'Public Function aaa(a As Variant) As Variant
'    Dim StartTime As Double, ElapsedTime As Double
'
'    StartTime = MicroTimer
'
'    'здесь пишем код, скорость которого измеряем
'
'    ElapsedTime = MicroTimer - StartTime
'
'    'вывод результата
'    Debug.Print CDec(ElapsedTime)
'End Function

