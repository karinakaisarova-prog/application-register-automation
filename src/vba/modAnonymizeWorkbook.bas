Attribute VB_Name = "modAnonymizeWorkbook"
Option Explicit

Public Sub AnonymizeTextData()
    Dim wbSource As Workbook
    Dim wbCopy As Workbook
    Dim ws As Worksheet
    Dim foundCell As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim firstUsedRow As Long
    Dim headerRow As Long
    Dim col As Long
    Dim rowNum As Long
    Dim headerValue As Variant
    Dim cellValue As Variant
    Dim dict As Object
    Dim formulaColumns() As Boolean
    Dim savePath As Variant
    Dim suggestedPath As String
    Dim extension As String
    Dim baseName As String
    Dim dotPosition As Long
    Dim replacedCount As Long
    Dim skippedFormulaColumns As Long
    Dim skippedDateColumns As Long
    Dim answer As VbMsgBoxResult
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim copyCreated As Boolean

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    On Error GoTo ErrorHandler

    Set wbSource = ActiveWorkbook

    If wbSource Is Nothing Then Exit Sub

    If wbSource.Worksheets.Count <> 1 Then
        MsgBox "В книге должен быть ровно один лист. Обработка отменена.", vbExclamation
        Exit Sub
    End If

    Set ws = wbSource.Worksheets(1)

    If wbSource.Path = "" Then
        MsgBox "Сначала сохрани исходную книгу на компьютере, затем запусти макрос ещё раз.", vbExclamation
        Exit Sub
    End If

    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

    If foundCell Is Nothing Then
        MsgBox "На листе нет данных.", vbInformation
        Exit Sub
    End If

    firstUsedRow = foundCell.Row

    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    lastRow = foundCell.Row

    Set foundCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    lastCol = foundCell.Column

    headerRow = Application.InputBox("Укажи номер строки с заголовками столбцов.", "Обезличивание данных", firstUsedRow, Type:=1)

    If headerRow < 1 Or headerRow >= lastRow Then Exit Sub

    ReDim formulaColumns(1 To lastCol)

    For col = 1 To lastCol
        For rowNum = headerRow + 1 To lastRow
            If ws.Cells(rowNum, col).HasFormula Then
                formulaColumns(col) = True
                skippedFormulaColumns = skippedFormulaColumns + 1
                Exit For
            End If
        Next rowNum
    Next col

    dotPosition = InStrRev(wbSource.Name, ".")
    If dotPosition > 0 Then
        baseName = Left$(wbSource.Name, dotPosition - 1)
        extension = Mid$(wbSource.Name, dotPosition)
    Else
        baseName = wbSource.Name
        extension = ".xlsx"
    End If

    suggestedPath = wbSource.Path & Application.PathSeparator & baseName & "_обезличенная" & extension
    savePath = Application.GetSaveAsFilename(InitialFileName:=suggestedPath, Title:="Сохранить обезличенную копию")

    If VarType(savePath) = vbBoolean Then Exit Sub

    savePath = CStr(savePath)

    If LCase$(Right$(savePath, Len(extension))) <> LCase$(extension) Then
        savePath = savePath & extension
    End If

    If Len(Dir$(savePath)) > 0 Then
        MsgBox "Файл с таким именем уже существует. Выбери другое имя, чтобы не перезаписать его.", vbExclamation
        Exit Sub
    End If

    answer = MsgBox("Будет создана отдельная копия книги." & vbCrLf & _
                    "Исходный файл останется без изменений." & vbCrLf & vbCrLf & _
                    "Текстовые значения в столбцах без формул будут заменены на заголовок столбца и номер." & vbCrLf & _
                    "Даты, числа и формулы останутся без изменений." & vbCrLf & vbCrLf & _
                    "Продолжить?", vbQuestion + vbYesNo + vbDefaultButton2, "Подтверждение обезличивания")

    If answer <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    wbSource.SaveCopyAs CStr(savePath)
    copyCreated = True

    Set wbCopy = Workbooks.Open(Filename:=CStr(savePath), UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    Set ws = wbCopy.Worksheets(1)

    For col = 1 To lastCol
        If Not formulaColumns(col) Then
            headerValue = ws.Cells(headerRow, col).Value2

            If Not IsError(headerValue) Then
                If Len(Trim$(CStr(headerValue))) > 0 Then
                    If HeaderContainsDate(CStr(headerValue)) Then
                        skippedDateColumns = skippedDateColumns + 1
                    Else
                        Set dict = CreateObject("Scripting.Dictionary")
                        dict.CompareMode = vbBinaryCompare

                        For rowNum = headerRow + 1 To lastRow
                            With ws.Cells(rowNum, col)
                                If Not .HasFormula Then
                                    cellValue = .Value2

                                    If VarType(cellValue) = vbString Then
                                        If Len(CStr(cellValue)) > 0 Then
                                            If Not LooksLikeDate(CStr(cellValue)) Then
                                                If Not dict.Exists(CStr(cellValue)) Then
                                                    dict.Add CStr(cellValue), CStr(headerValue) & " " & CStr(dict.Count + 1)
                                                End If

                                                If .Hyperlinks.Count > 0 Then .Hyperlinks.Delete
                                                .Value2 = dict(CStr(cellValue))
                                                replacedCount = replacedCount + 1
                                            End If
                                        End If
                                    End If
                                End If
                            End With
                        Next rowNum
                    End If
                End If
            End If
        End If
    Next col

    wbCopy.Save
    wbCopy.Close SaveChanges:=False
    Set wbCopy = Nothing

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation

    Debug.Print "Обезличивание завершено. Заменено ячеек: " & replacedCount
    Debug.Print "Пропущено столбцов с формулами: " & skippedFormulaColumns
    Debug.Print "Пропущено столбцов с датами по заголовку: " & skippedDateColumns
    Debug.Print "Файл: " & CStr(savePath)

    MsgBox "Готово. Создана обезличенная копия:" & vbCrLf & CStr(savePath) & vbCrLf & vbCrLf & _
           "Заменено текстовых ячеек: " & replacedCount & vbCrLf & _
           "Пропущено столбцов с формулами: " & skippedFormulaColumns & vbCrLf & _
           "Пропущено столбцов с датами: " & skippedDateColumns, vbInformation

    Exit Sub

ErrorHandler:
    On Error Resume Next

    If Not wbCopy Is Nothing Then wbCopy.Close SaveChanges:=False

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation

    If copyCreated And Len(CStr(savePath)) > 0 Then
        If Len(Dir$(CStr(savePath))) > 0 Then Kill CStr(savePath)
    End If

    Debug.Print "Ошибка обезличивания: " & Err.Number & " — " & Err.Description
    MsgBox "Не удалось завершить обезличивание." & vbCrLf & Err.Number & " — " & Err.Description & vbCrLf & _
           "Исходная книга не изменена.", vbCritical
End Sub

Private Function HeaderContainsDate(ByVal headerText As String) As Boolean
    Dim normalizedText As String

    normalizedText = LCase$(Trim$(headerText))

    HeaderContainsDate = _
        (InStr(1, normalizedText, "дата", vbTextCompare) > 0) Or _
        (InStr(1, normalizedText, "date", vbTextCompare) > 0) Or _
        (InStr(1, normalizedText, "период", vbTextCompare) > 0)
End Function

Private Function LooksLikeDate(ByVal cellText As String) As Boolean
    Dim normalizedText As String

    normalizedText = Trim$(cellText)

    If Not IsDate(normalizedText) Then Exit Function

    LooksLikeDate = _
        (InStr(1, normalizedText, ".", vbBinaryCompare) > 0) Or _
        (InStr(1, normalizedText, "/", vbBinaryCompare) > 0) Or _
        (InStr(1, normalizedText, "-", vbBinaryCompare) > 0) Or _
        (InStr(1, normalizedText, ",", vbBinaryCompare) > 0) Or _
        (InStr(1, normalizedText, " ", vbBinaryCompare) > 0)
End Function
