Attribute VB_Name = "modRenameFiles"
Option Explicit

' Safe file renaming from the existing register sheet.
' Expected columns: B=Object, C=Application number, D=Title, G=File hyperlink.
' The macro previews every planned rename and asks for confirmation before changing files.
' Existing destination files are never overwritten.

Private Const FIRST_DATA_ROW As Long = 3
Private Const COL_OBJECT As Long = 2
Private Const COL_APP_NUMBER As Long = 3
Private Const COL_TITLE As Long = 4
Private Const COL_LINK As Long = 7
Private Const LOG_SHEET As String = "Журнал переименования"

Public Sub RenameFilesFromRegister()
    Dim ws As Worksheet
    Dim logWs As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim oldPath As String
    Dim newPath As String
    Dim oldName As String
    Dim newName As String
    Dim ext As String
    Dim objectName As String
    Dim appNo As String
    Dim titleText As String
    Dim plan As Collection
    Dim rec As Object
    Dim seen As Object
    Dim preview As String
    Dim problems As String
    Dim item As Variant
    Dim i As Long
    Dim logRow As Long
    Dim renamed As Long
    Dim failed As Long
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim errText As String

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
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

        ext = Mid$(oldPath, InStrRev(oldPath, "."))
        newName = CleanFileName(objectName & "_Приложение № " & appNo & "_" & titleText) & ext
        newPath = Left$(oldPath, InStrRev(oldPath, Application.PathSeparator)) & newName
        If StrComp(oldPath, newPath, vbTextCompare) = 0 Then GoTo ContinueRow

        If seen.Exists(newPath) Then
            problems = problems & "Строка " & r & ": одинаковое целевое имя уже запланировано — " & newName & vbCrLf
            GoTo ContinueRow
        End If
        seen(newPath) = True
        If Len(Dir$(newPath)) > 0 Then
            problems = problems & "Строка " & r & ": файл с целевым именем уже существует — " & newPath & vbCrLf
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
        MsgBox "Нет безопасных переименований для выполнения." & IIf(Len(problems) > 0, vbCrLf & vbCrLf & problems, ""), vbInformation
        GoTo CleanExit
    End If

    preview = "Будут переименованы файлы: " & plan.Count & vbCrLf & vbCrLf
    For Each item In plan
        preview = preview & CStr(item("OldName")) & vbCrLf & "  → " & CStr(item("NewName")) & vbCrLf & vbCrLf
        If Len(preview) > 6000 Then
            preview = preview & "... (остальные строки есть в журнале после подтверждения)" & vbCrLf
            Exit For
        End If
    Next item
    If Len(problems) > 0 Then preview = preview & "Пропущенные строки:" & vbCrLf & problems
    If MsgBox(preview & vbCrLf & "Продолжить? Исходные файлы будут переименованы.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Подтверждение переименования") <> vbYes Then
        GoTo CleanExit
    End If

    Set logWs = GetOrCreateRenameLog(LOG_SHEET)
    If Len(CStr(logWs.Cells(1, 1).Value)) = 0 Then
        logWs.Range("A1:G1").Value = Array("Дата", "Строка перечня", "Старый путь", "Новое имя", "Новый путь", "Статус", "Ошибка")
    End If
    logRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 1
    For Each item In plan
        DoEvents
        On Error GoTo RenameFailed
        Name CStr(item("OldPath")) As CStr(item("NewPath"))
        ws.Hyperlinks.Add Anchor:=ws.Cells(CLng(item("Row")), COL_LINK), _
            Address:=CStr(item("NewPath")), TextToDisplay:="Открыть файл"
        logWs.Cells(logRow, 1).Value = Now
        logWs.Cells(logRow, 2).Value = item("Row")
        logWs.Cells(logRow, 3).Value = item("OldPath")
        logWs.Cells(logRow, 4).Value = item("NewName")
        logWs.Cells(logRow, 5).Value = item("NewPath")
        logWs.Cells(logRow, 6).Value = "Переименован"
        logWs.Cells(logRow, 7).Value = ""
        logRow = logRow + 1
        renamed = renamed + 1
        On Error GoTo Failed
    Next item

    FormatRenameLog logWs
    MsgBox "Переименование завершено. Успешно: " & renamed & "; ошибок: " & failed & _
           IIf(Len(problems) > 0, vbCrLf & vbCrLf & "Пропуски:" & vbCrLf & problems, ""), vbInformation
    GoTo CleanExit

RenameFailed:
    errText = Err.Description
    logWs.Cells(logRow, 1).Value = Now
    logWs.Cells(logRow, 2).Value = item("Row")
    logWs.Cells(logRow, 3).Value = item("OldPath")
    logWs.Cells(logRow, 4).Value = item("NewName")
    logWs.Cells(logRow, 5).Value = item("NewPath")
    logWs.Cells(logRow, 6).Value = "Ошибка"
    logWs.Cells(logRow, 7).Value = errText
    logRow = logRow + 1
    failed = failed + 1
    Err.Clear
    On Error GoTo Failed
    Resume NextItem
NextItem:
    On Error GoTo Failed
    Resume Next
Failed:
    errText = Err.Description
    MsgBox "Операция переименования прервана: " & errText, vbExclamation
CleanExit:
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
End Sub

Public Sub RollbackLastRenameBatch()
    Dim logWs As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim oldPath As String
    Dim newPath As String
    Dim statusText As String
    Dim restored As Long
    Dim skipped As Long
    Dim errText As String

    On Error GoTo Failed
    Set logWs = ThisWorkbook.Worksheets(LOG_SHEET)
    lastRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Журнал переименования пуст.", vbInformation
        Exit Sub
    End If
    If MsgBox("Попытаться отменить все успешные переименования из журнала в обратном порядке? " & _
              "Если исходное имя уже занято, файл будет пропущен.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Откат переименования") <> vbYes Then Exit Sub

    For r = lastRow To 2 Step -1
        statusText = CStr(logWs.Cells(r, 6).Value)
        If statusText = "Переименован" Then
            oldPath = CStr(logWs.Cells(r, 3).Value)
            newPath = CStr(logWs.Cells(r, 5).Value)
            If Len(Dir$(newPath)) = 0 Or Len(Dir$(oldPath)) > 0 Then
                logWs.Cells(r, 6).Value = "Откат пропущен"
                logWs.Cells(r, 7).Value = "Новый файл отсутствует или исходное имя уже занято."
                skipped = skipped + 1
            Else
                On Error Resume Next
                Name newPath As oldPath
                If Err.Number = 0 Then
                    logWs.Cells(r, 6).Value = "Откатан"
                    logWs.Cells(r, 7).Value = ""
                    restored = restored + 1
                Else
                    logWs.Cells(r, 6).Value = "Ошибка отката"
                    logWs.Cells(r, 7).Value = Err.Description
                    skipped = skipped + 1
                    Err.Clear
                End If
                On Error GoTo Failed
            End If
        End If
    Next r
    MsgBox "Откат завершён. Восстановлено: " & restored & "; пропущено: " & skipped, vbInformation
    Exit Sub
Failed:
    errText = Err.Description
    MsgBox "Не удалось выполнить откат: " & errText, vbExclamation
End Sub

Private Function RegisterFilePath(ByVal c As Range) As String
    On Error Resume Next
    If c.Hyperlinks.Count > 0 Then
        RegisterFilePath = CStr(c.Hyperlinks(1).Address)
    Else
        RegisterFilePath = CStr(c.Value)
    End If
    On Error GoTo 0
End Function

Private Function CleanFileName(ByVal value As String) As String
    Dim invalidChars As Variant
    Dim i As Long
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
    ws.Columns("A:G").AutoFit
End Sub
