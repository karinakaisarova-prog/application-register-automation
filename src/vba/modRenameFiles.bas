Attribute VB_Name = "modRenameFiles"
Option Explicit

' Safe file renaming from the existing register sheet.
' Columns: B=Object, C=Application number, D=Title, G=File hyperlink.
' Preview + explicit confirmation; never overwrites an existing destination.

Private Const FIRST_DATA_ROW As Long = 3
Private Const COL_OBJECT As Long = 2
Private Const COL_APP_NUMBER As Long = 3
Private Const COL_TITLE As Long = 4
Private Const COL_LINK As Long = 7
Private Const LOG_SHEET As String = "Журнал переименования"

Public Sub RenameFilesFromRegister()
    Dim ws As Worksheet, logWs As Worksheet
    Dim lastRow As Long, r As Long, logRow As Long
    Dim oldPath As String, newPath As String, newName As String, ext As String
    Dim objectName As String, appNo As String, titleText As String
    Dim plan As Collection, rec As Object, seen As Object, item As Variant
    Dim preview As String, problems As String, failure As String, batchId As String
    Dim renamed As Long, failed As Long
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim oldStatusBar As Variant
    Dim oldCancelKey As XlEnableCancelKey
    Dim renamedItems As Collection
    Dim rollbackItem As Variant
    Dim rollbackError As String

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    oldCancelKey = Application.EnableCancelKey
    Set renamedItems = New Collection
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, COL_LINK).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then
        MsgBox "На активном листе нет строк для переименования.", vbInformation
        GoTo CleanExit
    End If

    Set plan = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For r = FIRST_DATA_ROW To lastRow
        DoEvents
        oldPath = RegisterFilePath(ws.Cells(r, COL_LINK))
        If Len(oldPath) = 0 Then GoTo ContinueRow
        If Len(Dir$(oldPath)) = 0 Then
            problems = problems & "Строка " & r & ": исходный файл не найден — " & oldPath & vbCrLf
            GoTo ContinueRow
        End If

        objectName = Trim$(CStr(ws.Cells(r, COL_OBJECT).Value))
        appNo = Trim$(CStr(ws.Cells(r, COL_APP_NUMBER).Value))
        titleText = Trim$(CStr(ws.Cells(r, COL_TITLE).Value))
        If Len(objectName) = 0 Or Len(appNo) = 0 Or Len(titleText) = 0 Then
            problems = problems & "Строка " & r & ": заполните Объект, № Прил. и Наименование." & vbCrLf
            GoTo ContinueRow
        End If

        ext = FileExtensionFromPath(oldPath)
        newName = CleanFileName(objectName & "_Приложение № " & appNo & "_" & titleText) & ext
        newPath = Left$(oldPath, InStrRev(oldPath, Application.PathSeparator)) & newName
        If StrComp(oldPath, newPath, vbTextCompare) = 0 Then GoTo ContinueRow
        If seen.Exists(newPath) Then
            problems = problems & "Строка " & r & ": дублируется целевое имя — " & newName & vbCrLf
            GoTo ContinueRow
        End If
        seen(newPath) = True
        If Len(Dir$(newPath)) > 0 Then
            problems = problems & "Строка " & r & ": целевой файл уже существует — " & newPath & vbCrLf
            GoTo ContinueRow
        End If

        Set rec = CreateObject("Scripting.Dictionary")
        rec("Row") = r
        rec("OldPath") = oldPath
        rec("NewPath") = newPath
        rec("OldName") = Mid$(oldPath, InStrRev(oldPath, Application.PathSeparator) + 1)
        rec("NewName") = newName
        plan.Add rec
ContinueRow:
    Next r

    If plan.Count = 0 Then
        MsgBox "Нет безопасных переименований для выполнения." & IIf(Len(problems) > 0, vbCrLf & problems, ""), vbInformation
        GoTo CleanExit
    End If

    preview = "Будут переименованы файлы: " & plan.Count & vbCrLf & vbCrLf
    For Each item In plan
        preview = preview & CStr(item("OldName")) & vbCrLf & "  → " & CStr(item("NewName")) & vbCrLf & vbCrLf
        If Len(preview) > 5500 Then
            preview = preview & "... Остальные файлы будут записаны в журнал." & vbCrLf
            Exit For
        End If
    Next item
    If Len(problems) > 0 Then preview = preview & "Пропущенные строки:" & vbCrLf & problems
    If MsgBox(preview & vbCrLf & "Продолжить? Исходные файлы будут переименованы.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Подтверждение переименования") <> vbYes Then GoTo CleanExit

    Set logWs = GetOrCreateRenameLog(LOG_SHEET)
    If Len(CStr(logWs.Cells(1, 1).Value)) = 0 Then
        logWs.Range("A1:G1").Value = Array("Дата", "Лист перечня", "Строка перечня", "Старый путь", "Новое имя", "Новый путь", "Статус/ошибка")
    End If
    If Len(CStr(logWs.Cells(1, 8).Value)) = 0 Then logWs.Cells(1, 8).Value = "ID пакета"
    logRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 1
    batchId = Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(logRow, "000000")

    For Each item In plan
        DoEvents
        failure = TryRenameFile(CStr(item("OldPath")), CStr(item("NewPath")))
        logWs.Cells(logRow, 1).Value = Now
        logWs.Cells(logRow, 2).Value = ws.Name
        logWs.Cells(logRow, 3).Value = item("Row")
        logWs.Cells(logRow, 4).Value = item("OldPath")
        logWs.Cells(logRow, 5).Value = item("NewName")
        logWs.Cells(logRow, 6).Value = item("NewPath")
        If Len(failure) = 0 Then
            renamedItems.Add item
            On Error Resume Next
            ws.Cells(CLng(item("Row")), COL_LINK).Hyperlinks.Delete
            ws.Hyperlinks.Add Anchor:=ws.Cells(CLng(item("Row")), COL_LINK), _
                Address:=CStr(item("NewPath")), TextToDisplay:="Открыть файл"
            On Error GoTo Failed
            logWs.Cells(logRow, 7).Value = "Переименован"
            logWs.Cells(logRow, 8).Value = batchId
            renamed = renamed + 1
        Else
            logWs.Cells(logRow, 7).Value = "Ошибка: " & failure
            logWs.Cells(logRow, 8).Value = batchId
            failed = failed + 1
        End If
        logRow = logRow + 1
        Application.StatusBar = "Переименование: " & (renamed + failed) & " из " & plan.Count
    Next item

    FormatRenameLog logWs
    MsgBox "Переименование завершено. Успешно: " & renamed & "; ошибок: " & failed & _
           IIf(Len(problems) > 0, vbCrLf & vbCrLf & "Пропуски:" & vbCrLf & problems, ""), vbInformation
    GoTo CleanExit

Failed:
    failure = Err.Description
    If Err.Number = 18 Then failure = "Операция отменена пользователем (Ctrl+Break)."
    ' If a batch was interrupted, reverse every rename completed in this run.
    On Error Resume Next
    If Not renamedItems Is Nothing Then
        For Each rollbackItem In renamedItems
            rollbackError = TryRenameFile(CStr(rollbackItem("NewPath")), CStr(rollbackItem("OldPath")))
            If Len(rollbackError) = 0 Then
                If Not ws Is Nothing Then
                    ws.Cells(CLng(rollbackItem("Row")), COL_LINK).Hyperlinks.Delete
                    ws.Hyperlinks.Add Anchor:=ws.Cells(CLng(rollbackItem("Row")), COL_LINK), _
                        Address:=CStr(rollbackItem("OldPath")), TextToDisplay:="Открыть файл"
                End If
                ' Keep the journal consistent with the actual file state after rollback.
                If Not logWs Is Nothing And Len(batchId) > 0 Then
                    For r = 2 To logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row
                        If Trim$(CStr(logWs.Cells(r, 8).Value)) = batchId And _
                           StrComp(CStr(logWs.Cells(r, 4).Value), CStr(rollbackItem("OldPath")), vbTextCompare) = 0 Then
                            logWs.Cells(r, 7).Value = "Автоматически отменён после прерывания"
                            Exit For
                        End If
                    Next r
                End If
            Else
                problems = problems & "Не удалось автоматически отменить: " & CStr(rollbackItem("NewPath")) & _
                           " (" & rollbackError & ")" & vbCrLf
                If Not logWs Is Nothing And Len(batchId) > 0 Then
                    For r = 2 To logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row
                        If Trim$(CStr(logWs.Cells(r, 8).Value)) = batchId And _
                           StrComp(CStr(logWs.Cells(r, 4).Value), CStr(rollbackItem("OldPath")), vbTextCompare) = 0 Then
                            logWs.Cells(r, 7).Value = "Ошибка автоматического отката: " & rollbackError
                            Exit For
                        End If
                    Next r
                End If
            End If
        Next rollbackItem
    End If
    On Error GoTo 0
    MsgBox failure & vbCrLf & _
           IIf(Len(problems) > 0, "Проблемы и пропуски:" & vbCrLf & problems & vbCrLf, "") & _
           "Выполненные в этом запуске переименования автоматически отменены, где это удалось. " & _
           "Проверьте журнал и пути файлов.", vbExclamation
CleanExit:
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
End Sub

Public Sub RollbackLastRenameBatch()
    Dim logWs As Worksheet, registerWs As Worksheet
    Dim lastRow As Long, r As Long, registerRow As Long, logRow As Long
    Dim oldPath As String, newPath As String, statusText As String
    Dim restored As Long, skipped As Long, errText As String, sheetName As String, batchId As String
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim oldStatusBar As Variant, oldCancelKey As XlEnableCancelKey
    Dim reversedItems As Collection, item As Object, doneItem As Object
    Dim failure As String, recoveryError As String, itemPath As Variant

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    oldCancelKey = Application.EnableCancelKey
    Set reversedItems = New Collection
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set logWs = ThisWorkbook.Worksheets(LOG_SHEET)
    lastRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Журнал переименования пуст.", vbInformation
        GoTo CleanExit
    End If

    batchId = ""
    ' Find the newest batch that still has successful, not-yet-rolled-back renames.
    For r = lastRow To 2 Step -1
        If Trim$(CStr(logWs.Cells(r, 7).Value)) = "Переименован" And _
           Len(Trim$(CStr(logWs.Cells(r, 8).Value))) > 0 Then
            batchId = Trim$(CStr(logWs.Cells(r, 8).Value))
            Exit For
        End If
    Next r
    If Len(batchId) = 0 Then
        MsgBox "В журнале нет переименованных файлов, доступных для отката. Старые записи без ID пакета автоматически не откатываются.", vbExclamation
        GoTo CleanExit
    End If
    If MsgBox("Отменить последнее переименование (пакет " & batchId & ")? " & _
              "Если исходное имя уже занято, файл будет пропущен. При прерывании макрос попытается вернуть уже откатанные файлы к новым именам.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Откат переименования") <> vbYes Then GoTo CleanExit

    For r = lastRow To 2 Step -1
        DoEvents
        Application.StatusBar = "Откат переименования: запись " & r & " из " & lastRow
        statusText = CStr(logWs.Cells(r, 7).Value)
        If statusText = "Переименован" And Trim$(CStr(logWs.Cells(r, 8).Value)) = batchId Then
            oldPath = CStr(logWs.Cells(r, 4).Value)
            newPath = CStr(logWs.Cells(r, 6).Value)
            If Len(Dir$(newPath)) = 0 Or Len(Dir$(oldPath)) > 0 Then
                logWs.Cells(r, 7).Value = "Откат пропущен: новый файл отсутствует или исходное имя занято"
                skipped = skipped + 1
            Else
                errText = TryRenameFile(newPath, oldPath)
                If Len(errText) = 0 Then
                    ' Track immediately after the filesystem change so interruption recovery can find it.
                    Set doneItem = CreateObject("Scripting.Dictionary")
                    doneItem("LogRow") = r
                    doneItem("OldPath") = oldPath
                    doneItem("NewPath") = newPath
                    doneItem("SheetName") = CStr(logWs.Cells(r, 2).Value)
                    doneItem("RegisterRow") = CLng(Val(logWs.Cells(r, 3).Value))
                    reversedItems.Add doneItem

                    logWs.Cells(r, 7).Value = "Откатан"
                    sheetName = CStr(doneItem("SheetName"))
                    registerRow = CLng(doneItem("RegisterRow"))
                    Set registerWs = Nothing
                    On Error Resume Next
                    Set registerWs = ThisWorkbook.Worksheets(sheetName)
                    If Not registerWs Is Nothing And registerRow > 0 Then
                        registerWs.Cells(registerRow, COL_LINK).Hyperlinks.Delete
                        registerWs.Hyperlinks.Add Anchor:=registerWs.Cells(registerRow, COL_LINK), _
                            Address:=oldPath, TextToDisplay:="Открыть файл"
                    End If
                    On Error GoTo Failed
                    restored = restored + 1
                Else
                    logWs.Cells(r, 7).Value = "Ошибка отката: " & errText
                    skipped = skipped + 1
                End If
            End If
        End If
    Next r
    FormatRenameLog logWs
    MsgBox "Откат завершён. Восстановлено: " & restored & "; пропущено: " & skipped, vbInformation
    GoTo CleanExit

Failed:
    failure = Err.Description
    If Err.Number = 18 Then failure = "Операция отменена пользователем (Ctrl+Break)."
    ' Best-effort recovery: restore files already renamed back to their post-rename paths.
    On Error Resume Next
    Application.EnableCancelKey = xlDisabled
    For Each item In reversedItems
        recoveryError = TryRenameFile(CStr(item("OldPath")), CStr(item("NewPath")))
        logRow = CLng(item("LogRow"))
        If Len(recoveryError) = 0 Then
            logWs.Cells(logRow, 7).Value = "Переименован"
            Set registerWs = Nothing
            Set registerWs = ThisWorkbook.Worksheets(CStr(item("SheetName")))
            If Not registerWs Is Nothing And CLng(item("RegisterRow")) > 0 Then
                registerWs.Cells(CLng(item("RegisterRow")), COL_LINK).Hyperlinks.Delete
                registerWs.Hyperlinks.Add Anchor:=registerWs.Cells(CLng(item("RegisterRow")), COL_LINK), _
                    Address:=CStr(item("NewPath")), TextToDisplay:="Открыть файл"
            End If
        Else
            logWs.Cells(logRow, 7).Value = "Ошибка восстановления после прерывания: " & recoveryError
            failure = failure & vbCrLf & "Не удалось вернуть файл к новому имени: " & CStr(item("OldPath")) & _
                      " (" & recoveryError & ")"
        End If
    Next item
    On Error GoTo 0
    MsgBox "Откат прерван. Макрос попытался вернуть уже обработанные файлы к состоянию до отката." & vbCrLf & _
           failure & vbCrLf & "Проверьте журнал переименования и фактические имена файлов.", vbExclamation
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
    On Error GoTo 0
End Sub

Private Function TryRenameFile(ByVal oldPath As String, ByVal newPath As String) As String
    On Error GoTo RenameError
    If Len(Dir$(oldPath)) = 0 Then
        TryRenameFile = "Исходный файл не найден."
        Exit Function
    End If
    If Len(Dir$(newPath)) > 0 Then
        TryRenameFile = "Целевой файл уже существует."
        Exit Function
    End If
    Name oldPath As newPath
    Exit Function
RenameError:
    TryRenameFile = Err.Description
    Err.Clear
End Function

Private Function RegisterFilePath(ByVal c As Range) As String
    On Error Resume Next
    If c.Hyperlinks.Count > 0 Then
        RegisterFilePath = CStr(c.Hyperlinks(1).Address)
        If Len(RegisterFilePath) = 0 Then RegisterFilePath = CStr(c.Hyperlinks(1).SubAddress)
    Else
        RegisterFilePath = CStr(c.Value)
    End If
    On Error GoTo 0
    RegisterFilePath = NormalizeRegisterPath(RegisterFilePath)
End Function

Private Function NormalizeRegisterPath(ByVal value As String) As String
    value = Trim$(value)
    If LCase$(Left$(value, 8)) = "file:///" Then value = Mid$(value, 9)
    value = Replace(value, "%20", " ")
    value = Replace(value, "/", "\")
    NormalizeRegisterPath = value
End Function

Private Function FileExtensionFromPath(ByVal filePath As String) As String
    Dim slashPos As Long, dotPos As Long
    slashPos = InStrRev(filePath, Application.PathSeparator)
    dotPos = InStrRev(filePath, ".")
    If dotPos > slashPos Then FileExtensionFromPath = Mid$(filePath, dotPos)
End Function

Private Function CleanFileName(ByVal value As String) As String
    Dim invalidChars As Variant, i As Long
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(invalidChars) To UBound(invalidChars)
        value = Replace(value, CStr(invalidChars(i)), "_")
    Next i
    Do While InStr(value, "__") > 0
        value = Replace(value, "__", "_")
    Loop
    CleanFileName = Trim$(value)
End Function

Private Function GetOrCreateRenameLog(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateRenameLog = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateRenameLog Is Nothing Then
        Set GetOrCreateRenameLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateRenameLog.Name = sheetName
    End If
End Function

Private Sub FormatRenameLog(ByVal ws As Worksheet)
    With ws.UsedRange
        .Font.Name = "Times New Roman"
        .Borders.LineStyle = xlContinuous
    End With
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(198, 224, 180)
    ws.Columns("A:H").AutoFit
End Sub
