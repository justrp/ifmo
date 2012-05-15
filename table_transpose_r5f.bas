Attribute VB_Name = "table_transpose"
'Макрос, транспонирующий таблицу, в которой находится курсор. Исходная таблица остается без изменений; транспонированная-
'появляется в конце документа, после разрыва строки.

'Чиковани А. 3645
'23.04.2012


'Структура, хранящая информацию об объединениях
Public Type cellInfo
    rs As Integer   'rowspan
    cs As Integer   'colspan
    w As Integer    'width
    shiftH As Integer 'смещение по горизонтали - нужно для корректной транспонировки
    shiftV As Integer 'аналогично - по вертикали
End Type

'Структура, хранящая индексы ячеек
Public Type indexes
    i As Integer    'индекс по вертикали
    j As Integer    'индекс по горизонтали
End Type


Sub getinfo(ByRef tinfo() As cellInfo, ByRef rinfo() As Integer, t As Table)
'tinfo - массив, хранящий данные об объединениях ячейки
'rinfo - массив, содержащий количество ячеек в строках
't - указатель на исходную таблицу
'Метод исследует таблицу, сохраняя данные об объединениях ее ячеек
    For Each el In rinfo  'заполняем нулями (будет заполняться позже)
        el = 0
    Next el
    
    For Each aCell In t.Range.cells   'проходим всю таблицу, собирая информацию о ее клетках
        aCell.Select
        i = aCell.rowIndex
        j = aCell.ColumnIndex
        tinfo(i, j).rs = (Selection.Information(wdEndOfRangeRowNumber) - Selection.Information(wdStartOfRangeRowNumber))  'количество слитых строк -1
        tinfo(i, j).w = aCell.width         ' ширина ячейки (требуется для последующего определения атрибута cs
        tinfo(i, j).cs = 0                  'количество слитых столбцов (будет рассчитано позже)
        rinfo(i) = rinfo(i) + 1             'количество ячеек в строке (это другой массив)
        tinfo(i, j).shiftH = 0              'смещение по горизонтали - нужно для корректной транспонировки, будет заполняться позже
        tinfo(i, j).shiftV = 0              'аналогично - по вертикали
    Next aCell
    
    Call mergingResearch(tinfo, rinfo, t)   'заполняем массив данными об объединениях - cs, shiftH и shiftV
    
End Sub

Function getPerfIndex(ByRef rinfo() As Integer, maxColsAm As Integer)
'rinfo - массив, содержащий количество ячеек в строках
'maxColsAm - количество столбцов в таблице = количество ячеек в "идеальной" строке
'Функция возвращает индекс "идеальной" строки
    getPerfIndex = 1
    For i = 1 To UBound(rinfo)          'просматриваем все строки таблицы
        If rinfo(i) = maxColsAm Then    '-если количество ячеек в строке равно количеству столбцов в таблице,
            getPerfIndex = i            'значит это - "идеальная" строка, возвращаем ее индекс
            Exit For
        End If
    Next i
End Function

Sub mergingResearch(ByRef tinfo() As cellInfo, ByRef rinfo() As Integer, t As Table)
'tinfo - массив, хранящий данные об объединениях ячейки
'rinfo - массив, содержащий количество ячеек в строках
't - указатель на исходную таблицу
'Метод заполняет массив tinfo данными об объединениях - cs, shiftH и shiftV

    Dim i, j, perfIndex, w, k, shift As Integer
    'i,j,k - указатели
    'perfIndex - индекс "идеальной" строки
    'w - служит для суммирования параметра width ячеек идеальной строки -- для сравнения с шириной рассматриваемой ячейки
    'shift - смещение MSWord-индекса ячейки относительно логического, применяется в случае наличия объединений

    perfIndex = getPerfIndex(rinfo, t.Columns.count)    'ищем индекс "идеальной" строки

    For Each c In t.Range.cells                 'проходим все ячейки в таблице
        shift = 0           'обнуляем смещение
        If (c.rowIndex <> perfIndex) Then           '"идеальную" строку не рассматриваем - в ней нет объединений
            i = c.rowIndex                          'запоминаем индексы
            j = c.ColumnIndex
            k = j
            w = tinfo(perfIndex, k).w               'находим ширину ячейки, аналогичной рассматриваемой, находящейся в "идеальной" строке
            While (tinfo(i, j).w - w > 5)           'пока ширина не совпадает (5 - допустимая погрешность суммы ширин границ)
                w = w + tinfo(perfIndex, k).w       'увеличиваем параметр cs (colspan) горизонтального объединения
                tinfo(i, j).cs = tinfo(i, j).cs + 1 'при этом увеличивая параметр сравнения ширины w на ширину следующей ячейки "идеальной" строки
                k = k + 1
                shift = shift + 1
            Wend

            If tinfo(i, j).cs <> 0 Then                     'в случае, если ячейки были объединены, следует увеличить параметр гор.смещения
                For z = i To (i + tinfo(i, j).rs)           'для всех ячеек, чьи индексы будут зависеть от этого объединения
                    For k = (j + 1) To UBound(tinfo, 2)
                        tinfo(z, k).shiftH = tinfo(z, k).shiftH + shift
                    Next k
                Next z
            End If
        End If
    Next c
    

    
    For i = 1 To UBound(tinfo, 1)           'после того, как все объединения выявлены, следует запомнить информацию о вертикальном сдвиге
            For j = 1 To UBound(tinfo, 2)   'индексов ячеек - они потребуются для корректного выполнения транспонировки. Их удобнее хранить,
                If tinfo(i, j).rs <> 0 Then 'нежели постоянно высчитывать
                    For z = i + tinfo(i, j).rs + 1 To UBound(tinfo, 1)  'если ячейка имеет вертикальное объединение, следует увеличить параметр верт.смещения
                        For k = j + tinfo(i, j).shiftH To j + tinfo(i, j).cs    'для всех ячеек, чьи индексы будут зависеть от этого объединения
                            tinfo(z, k).shiftV = tinfo(z, k).shiftV + tinfo(i, j).rs
                        Next k
                    Next z
                End If
            Next j
    Next i

End Sub

Sub applyMerges(ByRef tinfo() As cellInfo, newT As Table)
'tinfo - массив, хранящий данные об объединениях ячейки
'newT - указатель на новую таблицу
'Метод объединяет ячейки новой таблицы на основании данных, хранящихся в tinfo, на лету транспонируя таблицу
    For i = 1 To UBound(tinfo, 2)                       'проходим массив, хранящий информацию об объединениях
            For j = 1 To UBound(tinfo, 1)
                If tinfo(j, i).w <> 0 Then              'и, в случае, если рассматриваемый элемент хранит информацию о существующей ячейке
                    If (tinfo(j, i).cs + tinfo(j, i).rs) <> 0 Then  'содержащей объединения, объединяет транспонированные ячейки новой таблицы
                        newT.Cell(i + tinfo(j, i).shiftH, j - tinfo(j, i).shiftV).Merge newT.Cell(i + tinfo(j, i).shiftH + tinfo(j, i).cs, j + tinfo(j, i).rs)
                    End If
                End If
            Next j
    Next i
End Sub

Sub fillTable(t As Table, newT As Table, ByRef tinfo() As cellInfo)
't - указатель на исходную таблицу
'newT - указатель на новую таблицу
'tinfo - массив, хранящий данные об объединениях ячейки
'Метод заполняет новую таблицу данными из исходной таблицы
    Dim i, j, m, n, k As Integer             'указатели
    Dim cinfo() As indexes         'массив, хранящий данные об индексах ячеек - нужен для обратного цикла типа foreach
    ReDim cinfo(1 To t.Rows.count * t.Columns.count)
    
    k = 1
    For Each c In t.Range.cells             'проходим все ячейки исходной таблицы, запоминая их индексы
        cinfo(k).i = c.rowIndex
        cinfo(k).j = c.ColumnIndex
        k = k + 1                           'и их количество
    Next c
        
    While k <> 1            'после чего - проходим по исходной таблице "обратным foreach'ем - от правой нижней ячейки к левой верхней
        k = k - 1
        
        i = cinfo(k).i                      'запоминаем индексы, на основании которых передаем содержание ячейки исходной таблицы
        j = cinfo(k).j                      'в ячейку новой таблицы, не забывая сменить индексы на транспонированные
        m = j + tinfo(i, j).shiftH          ' m и n - смещенные индексы, указывающие на ячейки в транспонированной таблице
        n = i - tinfo(i, j).shiftV
        
        'копируем значения из исходной таблицы - и передаем в транспонируемую:
        t.Cell(i, j).Range.Copy
        newT.Cell(m, n).Shading.BackgroundPatternColorIndex = t.Cell(i, j).Shading.BackgroundPatternColorIndex
        If t.Cell(i, j).Shading.BackgroundPatternColor <> -1 Then
            newT.Cell(m, n).Shading.BackgroundPatternColor = t.Cell(i, j).Shading.BackgroundPatternColor
        End If
        newT.Cell(m, n).Borders = t.Cell(i, j).Borders
        newT.Cell(m, n).Borders.OutsideLineStyle = t.Cell(i, j).Borders.OutsideLineStyle
        newT.Cell(m, n).Borders.OutsideColorIndex = t.Cell(i, j).Borders.OutsideColorIndex
        newT.Cell(m, n).Borders.OutsideColor = t.Cell(i, j).Borders.OutsideColor
        newT.Cell(m, n).Borders.Enable = t.Cell(i, j).Borders.Enable
        newT.Cell(m, n).Borders.DistanceFromTop = t.Cell(i, j).Borders.DistanceFromTop
        newT.Cell(m, n).Borders.DistanceFromRight = t.Cell(i, j).Borders.DistanceFromRight
        newT.Cell(m, n).Borders.DistanceFromLeft = t.Cell(i, j).Borders.DistanceFromLeft
        newT.Cell(m, n).Borders.DistanceFromBottom = t.Cell(i, j).Borders.DistanceFromBottom
        newT.Cell(m, n).Range.PasteAndFormat (wdPasteDefault)
    Wend
End Sub

Sub createTransponedTable(ByRef tinfo() As cellInfo, t As Table)
'tinfo - массив, хранящий данные об объединениях ячейки
't - указатель на исходную таблицу
'Метод создает новую таблицу, объединяя ее ячейки и заполняя данными из исходной таблицы, транспонируя ее

    Selection.EndKey Unit:=wdStory  'новая таблица добавляется в конец документа,
    Selection.TypeParagraph         'после разрыва строки

    'создаем новую таблицу, используя данные старой
    ActiveDocument.tables.Add Range:=Selection.Range, NumRows:=t.Columns.count, NumColumns:= _
        t.Rows.count, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    Dim newT As Table   'указатель на новую таблицу
    Set newT = Selection.tables(1)
    With newT
        .Style = t.Style '"Сетка таблицы"
        .ApplyStyleHeadingRows = t.ApplyStyleFirstColumn
        .ApplyStyleLastRow = t.ApplyStyleLastColumn
        .ApplyStyleFirstColumn = t.ApplyStyleHeadingRows
        .ApplyStyleLastColumn = t.ApplyStyleLastRow
        .Borders.OutsideLineStyle = t.Borders.OutsideLineStyle
        .Borders.InsideLineStyle = t.Borders.InsideLineStyle
    End With

   Call applyMerges(tinfo, newT)    'объединяем ячейки
   Call fillTable(t, newT, tinfo)   'заполняем таблицу содержимым (текст, стили)
End Sub

Sub table_transpose()
'Макрос, транспонирующий таблицу, в которой находится курсор. Исходная таблица остается без изменений; транспонированная-
'появляется в конце документа, после разрыва строки.
    Dim t As Table                  'транспонируем таблицу, в которой
    Set t = Selection.tables(1)     'находится курсор

    Dim tinfo() As cellInfo         'массив, хранящий данные об объединениях ячейки
    ReDim tinfo(1 To t.Rows.count, 1 To t.Columns.count)
    Dim rinfo() As Integer          'массив, содержащий количество ячеек в строках
    ReDim rinfo(1 To t.Rows.count)
    Call getinfo(tinfo, rinfo, t)   'собираем информацию об объединениях ячеек исходной таблицы
    Call createTransponedTable(tinfo, t)    'и, на основании полученной информации, создаем новую таблицу, заполняя ее
                                            'информацией из исходной таблицы

End Sub

