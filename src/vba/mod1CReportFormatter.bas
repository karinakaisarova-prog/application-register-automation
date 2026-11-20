Attribute VB_Name = "mod1CReportFormatter"
Option Explicit

Public Sub StartForColleague()
    Dim fd As FileDialog
    Dim wbTarget As Workbook

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Выберите выгрузку из 1С"
    fd.Filters.Clear
    fd.Filters.Add "Файлы Excel", "*.xls; *.xlsx; *.xlsm; *.csv"
    fd.AllowMultiSelect = False

    If fd.Show = -1 Then
        Set wbTarget = Workbooks.Open(fd.SelectedItems(1))
        wbTarget.Sheets(1).Activate
        Call Format1CReportExact
    End If
End Sub

Public Sub Format1CReportExact()
    Dim ws As Worksheet
    Dim lastCol As Long, lastRow As Long
    Dim col As Long, row As Long
    Dim headerRow As Long, qtyCol As Long, dateCol As Long
    Dim startTime As Double
    Dim cell As Range
    Dim isColEmpty As Boolean
    Dim cleanText As String, paramText As String
    Dim colQtyLet As String, colSumNoVatLet As String, colSumVatLet As String
    Dim maxTableCol As Long
    Dim needInsert As Boolean
    Dim titleRow As Long, formulaRow As Long, dataStartRow As Long
    Dim dateRange As Range
    Dim sampleEndRow As Long

    startTime = Timer
    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 1. Поиск строки с шапкой основной таблицы
    headerRow = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 15 Then
        On Error Resume Next
        lastRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        On Error GoTo 0
    End If

    For row = 1 To lastRow
        For col = 1 To 5
            If InStr(1, Trim(CStr(ws.Cells(row, col).Value)), "Поступление (акт", vbTextCompare) = 1 Then
                headerRow = row
                Exit For
            End If
        Next col
        If headerRow > 0 Then Exit For
    Next row

    If headerRow = 0 Then
        MsgBox "В выбранном файле не найдена таблица 'Поступление (акт, накладная, УПД)'!", vbExclamation
        GoTo Cleanup
    End If

    ' 2. Разъединение объединенных ячеек на листе
    ws.Cells.UnMerge

    ' 3. Перенос реквизитов выгрузки в столбец А
    For row = 4 To headerRow - 1
        paramText = ""
        For col = 2 To 10
            cleanText = Trim(Replace(Replace(Replace(CStr(ws.Cells(row, col).Value), Chr(160), ""), vbCr, " "), vbLf, " "))
            If Len(cleanText) > 0 Then
                paramText = cleanText
                ws.Cells(row, col).ClearContents
            End If
        Next col

        If Len(paramText) > 0 Then
            If InStr(1, paramText, "Имя объекта:", vbTextCompare) = 1 Then
                ws.Cells(row, 1).ClearContents
            Else
                ws.Cells(row, 1).Value = paramText
                ws.Cells(row, 1).WrapText = False
            End If
        End If
    Next row

    ' 4. Удаление пустых столбцов внутри таблицы
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    On Error Resume Next
    maxTableCol = ws.Range(ws.Rows(headerRow), ws.Rows(lastRow)).Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    On Error GoTo 0
    If maxTableCol > lastCol Then lastCol = maxTableCol

    For col = lastCol To 1 Step -1
        isColEmpty = True
        For Each cell In ws.Range(ws.Cells(headerRow, col), ws.Cells(lastRow, col))
            If Not IsEmpty(cell.Value) Then
                cleanText = Trim(Replace(Replace(Replace(CStr(cell.Value), Chr(160), ""), vbCr, ""), vbLf, ""))
                If Len(cleanText) > 0 Then
                    isColEmpty = False
                    Exit For
                End If
            End If
        Next cell

        If isColEmpty Then
            ws.Range(ws.Cells(headerRow, col), ws.Cells(lastRow, col)).Delete Shift:=xlToLeft
        End If
    Next col

    ' 5. Поиск столбцов с датой и количеством
    qtyCol = 0
    dateCol = 0
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For col = 1 To lastCol
        If Trim(CStr(ws.Cells(headerRow, col).Value)) = "Количество" Then qtyCol = col
        If Trim(CStr(ws.Cells(headerRow, col).Value)) = "Дата" Then dateCol = col
    Next col

    If qtyCol = 0 Then
        MsgBox "Не найден столбец 'Количество'!", vbExclamation
        GoTo Cleanup
    End If

    ' Проверка свободного места над столбцами итогов
    needInsert = False
    If headerRow <= 3 Then
        needInsert = True
    Else
        For Each cell In ws.Range(ws.Cells(headerRow - 3, qtyCol), ws.Cells(headerRow - 1, qtyCol + 4))
            If Len(Trim(Replace(CStr(cell.Value), Chr(160), ""))) > 0 Then
                needInsert = True
                Exit For
            End If
        Next cell
    End If

    If needInsert Then
        ws.Rows(headerRow & ":" & headerRow + 2).Insert Shift:=xlDown
        headerRow = headerRow + 3
        lastRow = lastRow + 3
    End If

    titleRow = headerRow - 3
    formulaRow = headerRow - 2
    dataStartRow = headerRow + 1

    ' 6. Установка формата даты ДД.ММ.ГГГГ
    If dateCol > 0 And lastRow >= dataStartRow Then
        Set dateRange = ws.Range(ws.Cells(dataStartRow, dateCol), ws.Cells(lastRow, dateCol))
        On Error Resume Next
        dateRange.TextToColumns Destination:=dateRange.Cells(1, 1), DataType:=xlDelimited, FieldInfo:=Array(1, xlDMYFormat)
        On Error GoTo 0
        dateRange.NumberFormat = "dd.mm.yyyy"
    End If

    ' 7. Добавление заголовков и формул итогов
    colQtyLet = Split(ws.Cells(1, qtyCol).Address, "$")(1)
    colSumNoVatLet = Split(ws.Cells(1, qtyCol + 2).Address, "$")(1)
    colSumVatLet = Split(ws.Cells(1, qtyCol + 4).Address, "$")(1)

    ws.Cells(titleRow, qtyCol).Value = "Промеж. итоги по кол-ву"
    ws.Cells(titleRow, qtyCol + 1).Value = "Средневзвешенная стоимость"
    ws.Cells(titleRow, qtyCol + 2).Value = "Промеж. итоги сумма без НДС"
    ws.Cells(titleRow, qtyCol + 3).Value = "Средневзвешенная стоимость с НДС"
    ws.Cells(titleRow, qtyCol + 4).Value = "Промеж. итоги сумма с НДС"

    With ws.Range(ws.Cells(titleRow, qtyCol), ws.Cells(titleRow, qtyCol + 4))
        .WrapText = True
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With

    ws.Cells(formulaRow, qtyCol).Formula = "=SUBTOTAL(9, " & colQtyLet & dataStartRow & ":" & colQtyLet & lastRow & ")"
    ws.Cells(formulaRow, qtyCol + 1).Formula = "=IFERROR(" & colSumNoVatLet & formulaRow & "/" & colQtyLet & formulaRow & ", 0)"
    ws.Cells(formulaRow, qtyCol + 2).Formula = "=SUBTOTAL(9, " & colSumNoVatLet & dataStartRow & ":" & colSumNoVatLet & lastRow & ")"
    ws.Cells(formulaRow, qtyCol + 3).Formula = "=IFERROR(" & colSumVatLet & formulaRow & "/" & colQtyLet & formulaRow & ", 0)"
    ws.Cells(formulaRow, qtyCol + 4).Formula = "=SUBTOTAL(9, " & colSumVatLet & dataStartRow & ":" & colSumVatLet & lastRow & ")"
    ws.Range(ws.Cells(formulaRow, qtyCol), ws.Cells(formulaRow, qtyCol + 4)).Font.Bold = True

    ' 8. Числовое форматирование и шрифт
    ws.Cells(formulaRow, qtyCol).NumberFormat = "#,##0"
    ws.Range(ws.Cells(formulaRow, qtyCol + 1), ws.Cells(formulaRow, qtyCol + 4)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(dataStartRow, qtyCol), ws.Cells(lastRow, qtyCol)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(dataStartRow, qtyCol + 1), ws.Cells(lastRow, qtyCol + 4)).NumberFormat = "#,##0.00"

    ws.Cells.Font.Name = "Times New Roman"
    ws.Cells(2, 1).Font.Size = 14
    ws.Cells(2, 1).Font.Bold = True

    ' 9. Настройка фильтра и ширины столбцов
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, qtyCol + 4)).AutoFilter

    sampleEndRow = WorksheetFunction.Min(lastRow, headerRow + 100)
    ws.Range(ws.Cells(titleRow, 1), ws.Cells(sampleEndRow, qtyCol + 4)).Columns.AutoFit

    For col = 1 To qtyCol + 4
        If ws.Columns(col).ColumnWidth > 35 Then ws.Columns(col).ColumnWidth = 35
        If ws.Columns(col).ColumnWidth < 12 Then ws.Columns(col).ColumnWidth = 12
    Next col

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Таблица успешно обработана за " & Format(Timer - startTime, "0.00") & " сек!", vbInformation
End Sub
