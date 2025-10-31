Sub GetContactsForMultipleOGRN()
    Dim http As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ogrn As String
    Dim url As String
    Dim response As String
    Dim contactInfo As String
    
    ' Настройки
    Set ws = ThisWorkbook.ActiveSheet
    lastRow = 1179 ' Последняя заполенная строка (ОГРН)
    
    ' Создаем HTTP объект
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Проходим по всем строкам
    For i = 3 To lastRow ' Начальная строка
        ' Берем ОГРН из столбца D
        ogrn = Trim(ws.Cells(i, 4).Value)
        
        ' Пропускаем пустые ячейки
        If ogrn = "" Then
            ws.Cells(i, 14).Value = "ОГРН не указан"
            GoTo NextIteration
        End If
        
        ' Формируем URL запроса
        url = "https://api.checko.ru/v2/company?key=ВАШ_API_КЛЮЧ&ogrn=" & ogrn ' Вставтье ваш ключ
        
        ' Выполняем запрос
        On Error Resume Next
        With http
            .Open "GET", url, False
            .send
            If .Status = 200 Then
                response = .responseText
                contactInfo = ExtractContactsFromJSON(response)
            Else
                contactInfo = "Ошибка HTTP: " & .Status
            End If
        End With
        On Error GoTo 0
        
        ' Записываем результат в столбец N
        ws.Cells(i, 14).Value = contactInfo
        
        ' Статус в столбце O
        ws.Cells(i, 15).Value = "Обработано: " & Format(Now, "hh:mm:ss")
        
        ' Пауза между запросами
        Application.Wait (Now + TimeValue("00:00:01"))
        
NextIteration:
        ' Обновляем прогресс
        Application.StatusBar = "Обработано: " & i - 2 & " из " & lastRow - 2
        DoEvents
    Next i
    
    Application.StatusBar = False
    Set http = Nothing
    Set ws = Nothing
    
    MsgBox "Обработка завершена!", vbInformation
End Sub

' Функция для извлечения контактов через поиск в тексте JSON
Function ExtractContactsFromJSON(jsonText As String) As String
    Dim phones As String
    Dim emails As String
    Dim website As String
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    ' Ищем телефоны в JSON
    phones = ExtractPhones(jsonText)
    
    ' Ищем email в JSON
    emails = ExtractEmails(jsonText)
    
    ' Ищем сайт в JSON
    website = ExtractWebsite(jsonText)
    
    ' Формируем результат
    result = ""
    If phones <> "" Then result = result & "Тел: " & phones & vbCrLf
    If emails <> "" Then result = result & "Email: " & emails & vbCrLf
    If website <> "" Then result = result & "Сайт: " & website & vbCrLf
    
    If result = "" Then result = "Контакты не найдены"
    
    ExtractContactsFromJSON = result
    Exit Function
    
ErrorHandler:
    ExtractContactsFromJSON = "Ошибка извлечения контактов"
End Function

' Функция для извлечения телефонов
Function ExtractPhones(jsonText As String) As String
    Dim phones As String
    Dim startPos As Long
    Dim endPos As Long
    Dim tempText As String
    
    ' Ищем раздел с телефонами
    startPos = InStr(1, jsonText, """Тел"":", vbTextCompare)
    If startPos > 0 Then
        ' Ищем начало массива
        startPos = InStr(startPos, jsonText, "[")
        If startPos > 0 Then
            ' Ищем конец массива
            endPos = InStr(startPos, jsonText, "]")
            If endPos > 0 Then
                ' Извлекаем текст между [ и ]
                tempText = Mid(jsonText, startPos, endPos - startPos + 1)
                
                ' Извлекаем номера телефонов
                phones = ExtractValuesFromArray(tempText)
            End If
        End If
    End If
    
    ExtractPhones = phones
End Function

' Функция для извлечения email
Function ExtractEmails(jsonText As String) As String
    Dim emails As String
    Dim startPos As Long
    Dim endPos As Long
    Dim tempText As String
    
    ' Ищем раздел с email
    startPos = InStr(1, jsonText, """Емэйл"":", vbTextCompare)
    If startPos > 0 Then
        startPos = InStr(startPos, jsonText, "[")
        If startPos > 0 Then
            endPos = InStr(startPos, jsonText, "]")
            If endPos > 0 Then
                tempText = Mid(jsonText, startPos, endPos - startPos + 1)
                emails = ExtractValuesFromArray(tempText)
            End If
        End If
    End If
    
    ExtractEmails = emails
End Function

' Функция для извлечения сайта
Function ExtractWebsite(jsonText As String) As String
    Dim website As String
    Dim startPos As Long
    Dim endPos As Long
    Dim tempText As String
    
    ' Ищем раздел с сайтом
    startPos = InStr(1, jsonText, """ВебСайт"":", vbTextCompare)
    If startPos > 0 Then
        ' Ищем значение после двоеточия
        startPos = InStr(startPos, jsonText, ":") + 1
        If startPos > 0 Then
            ' Ищем конец значения
            endPos = InStr(startPos, jsonText, ",")
            If endPos = 0 Then endPos = InStr(startPos, jsonText, "}")
            If endPos > 0 Then
                tempText = Trim(Mid(jsonText, startPos, endPos - startPos))
                ' Убираем кавычки
                website = Replace(tempText, """", "")
                If website = "null" Then website = ""
            End If
        End If
    End If
    
    ExtractWebsite = website
End Function

' Функция для извлечения значений из массива JSON
Function ExtractValuesFromArray(arrayText As String) As String
    Dim result As String
    Dim values() As String
    Dim i As Long
    
    ' Убираем квадратные скобки и разбиваем по запятым
    arrayText = Replace(arrayText, "[", "")
    arrayText = Replace(arrayText, "]", "")
    arrayText = Replace(arrayText, """", "")
    
    If Trim(arrayText) = "" Then
        ExtractValuesFromArray = ""
        Exit Function
    End If
    
    values = Split(arrayText, ",")
    result = ""
    
    For i = 0 To UBound(values)
        If Trim(values(i)) <> "" Then
            result = result & Trim(values(i)) & "; "
        End If
    Next i
    
    ' Убираем последнюю точку с запятой
    If Len(result) > 0 Then
        result = Left(result, Len(result) - 2)
    End If
    
    ExtractValuesFromArray = result
End Function
