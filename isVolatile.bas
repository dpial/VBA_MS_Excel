Attribute VB_Name = "my_isVolatile"
'isVolatile считает количество пересчета самой себя.
'  Отправьте данные в isVolatile которые хотите проверить.
'  isVolatile вернет свой счетчик.
'  Примечание: есть функции которые пересчитываются каждый раз,
'    даже если данные не изменились (Volatile functions).
Public Function isVolatile(x As Variant) As Integer
    Const maxCount = 32767
    Static count As Integer
    
    If count >= maxCount Then count = 0
    count = count + 1
    isVolatile = count
    
    'Данные вошедшие в isVolatile, нужно использовать.
    'Иначе работать не будет.
    x = x.Value
End Function
