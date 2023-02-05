Sub МакросДляФормированияСпискаЗаказов()

SourceFileName = ActiveWorkbook.Name ' Имя файла с обедами

ChDir "S:\Общая информация\Обеды\Обед\SelfAdhesivePapersTemplate"
Workbooks.Open Filename:= _
    "S:\Общая информация\Обеды\Обед\SelfAdhesivePapersTemplate\template.xlsm"
Cells.Select
Range("E2").Activate
Selection.ClearContents

' Здесь запоминается число, которое берется как имя листа в исходном файле с обедами
Windows(SourceFileName).Activate
CURRENT_DATE_TIME = ActiveSheet.Name ' Текущее число месяца

Const MacrosFileName As String = "template.xlsm" ' Имя файла с макросом
' Константы

' Пользователи
Const UsersColumnNumber As Byte = 1     ' Номер столбца, в котором располагаются имена пользователей
Const FirstLineOfUserNames As Byte = 5  ' Номер строки, с которой начинаются имена пользователей
Const LastLineOfUserNames As Byte = 51  ' Номер строки, в которой заканчиваются имена пользователей
' Столбцы заказа конкретных блюд
Const FirstColumnNumberOfOrders = 2     ' Номер первого столбца, с которого начинается область, где пользователи могут вписывать цифры заказа
Const LastColumnNumberOfOrders = 39     ' Номер последнего столбца области, где пользователи могут вписывать цифры заказа
' Раздел меню
    Const FoodNumberColumn = 41         ' Столбец, где указывается номер блюда
    Const FoodNameColumn = 42           ' Столбец, где указывается название блюда
    Const FoodStartLine = 6             ' Номер строки, с которой начинаются номера и названия блюд (одно значение для двух стоблцов меню)
    Const FoodEndLine = 42              ' Номер строки, в которой заканчиваются номера и названия блюд
' Раздел заказов блюд
    Const FoodNumberLine = 2            ' Номер строки, в которой указываются кодовые номер блюд из меню
    Const FoodPriceLine = 3             ' Номер строки, в которой указана стоимость каждого блюда из меню
' Выходной файл - лист с самоклеящимися бумажками (65 шт на одном листе)
Const OutputColumnCount = 5             ' Количество столбцов на листе с бумажками
Const MAX_FOOD_NAME_LENGTH = 30         ' Максимальное количество символов в имени блюда, которое будет скопировано на выходную бумажку

currentUsersThatMakeOrderCount = 0      ' Количество пользователей, сделавших заказ
currentPaperNumber = 0                  ' Текущая заполняемая бумажка на листе с самоклейками

    ' Цикл, проходящий по всем пользователям
    For i = FirstLineOfUserNames To LastLineOfUserNames
        ' Запоминаем имя пользователя, оно пригодится для того, чтобы потом вставлять строчки с его заказами
        ' Перейти в книгу с обедами, выделить ячейку с именем пользователя, скопировать ее
        Windows(SourceFileName).Activate
        userFullName = ActiveSheet.Cells(i, UsersColumnNumber).Value
            
        ' Теперь нужно пройти циклом по строке пользователя и выяснить, заказывал ли он обед?
        isUserMakeOrder = False
        currentUserTotalOrdersCount = 0 ' Количество позиций, заказанных текущим пользователем
        currentNumberOfFood = 0 ' Число столбцов, на которое мы отошли от начала строки (требуется для распознавания названия блюда по порядковому номеру)
        For j = FirstColumnNumberOfOrders To LastColumnNumberOfOrders
            
            Windows(SourceFileName).Activate
            ActiveSheet.Cells(i, j).Select
            If Len(ActiveCell) Then ' В этой ячейке что-то есть
                isUserMakeOrder = True ' Помечаем, что этот пользователь что-то заказал
                foodName = ActiveSheet.Cells(FoodStartLine + currentNumberOfFood, FoodNameColumn).Value ' Сначала запомним название блюда
                
                ' Теперь нужно узнать, сколько штук товара заказал пользователь в этой ячейке
                ' Для этого нужно знать количество денег, которое указал пользователь и цену за 1 единицу товара
                priceForOneCount = ActiveSheet.Cells(FoodPriceLine, j).Value ' Узнаем цену за 1 единицу товара
                userPriceString = ActiveSheet.Cells(i, j).Value ' Узнаем, какую сумму вписал пользователь, реализована защита от ввода нечислового значения
                If IsNumeric(userPriceString) Then
                userMoneyAmount = userPriceString
                End If
                ' Количество единиц товара, которое заказал пользователь
                currentUserTotalOrdersCount = Application.RoundUp((userMoneyAmount / priceForOneCount), 0)
                If (currentUserTotalOrdersCount > 0) Then
                    ' Если пользователь заказал что-то, нужно вписать строчку в выходной файл.
                    ' На каждый заказанный товар нужно печатать отдельную строчку
                    For k = 1 To currentUserTotalOrdersCount
                        Windows(MacrosFileName).Activate
                        currentColumn = currentPaperNumber Mod OutputColumnCount
                        currentLine = Application.RoundDown(currentPaperNumber / OutputColumnCount, 0)
                        foodNameLength = WorksheetFunction.Min(Len(foodName), MAX_FOOD_NAME_LENGTH)
                        ' Формирование выходной ячейки
                        ActiveSheet.Cells(currentLine + 1, currentColumn + 1).Value = userFullName + " " + CURRENT_DATE_TIME + " # " + Mid(foodName, 1, foodNameLength) ' Вставить имя пользователя, на второй строке - имя блюда
                        currentPaperNumber = currentPaperNumber + 1
                    Next k
                End If
            End If
            currentNumberOfFood = currentNumberOfFood + 1
        Next j ' Цикл по столбцам, в которых пользователи отмечают то, что они заказали
    Next i ' Цикл по пользователям
    
'###############################################################################
' Закрытие файла с обедами
Windows(SourceFileName).Activate
    
End Sub



