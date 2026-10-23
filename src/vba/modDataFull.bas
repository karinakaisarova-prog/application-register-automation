Attribute VB_Name = "modDataFull"
Option Explicit

' Builds a unique list of Section + Responsible Person pairs from the active register.
' Source columns: F = Responsible Person, H = Section; headers are on row 2, data starts on row 3.
' Output sheet: data_full. Blank values are ignored.

Private Const SOURCE_FIRST_ROW As Long = 3
Private Const COL_RESPONSIBLE As Long = 6
Private Const COL_SECTION As Long = 8
Private Const OUTPUT_SHEET As String = "data_full"

Public Sub BuildDataFull()
    Dim sourceWs As Worksheet
    Dim outputWs As Worksheet
    Dim backupWs As Worksheet
    Dim damagedWs As Worksheet
    Dim recoverySheetName As String
    Dim outputWasExisting As Boolean
    Dim operationCompleted As Boolean
    Dim lastRow As Long
    Dim r As Long
    Dim sectionText As String
    Dim personText As String
    Dim key As String
    Dim pairs As Object
    Dim item As Variant
    Dim pairValue As Variant
    Dim outRow As Long
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim oldStatusBar As Variant
    Dim oldCancelKey As XlEnableCancelKey

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    oldCancelKey = Application.EnableCancelKey
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set sourceWs = GetActiveRegisterSheet()
    If sourceWs Is Nothing Then
        MsgBox "Не найден лист перечня в книге результата. Сначала сформируйте перечень.", vbExclamation
        GoTo CleanExit
    End If
    If LCase$(sourceWs.Name) = LCase$(OUTPUT_SHEET) Or LCase$(sourceWs.Name) = "data" Then
        MsgBox "Сначала откройте лист перечня приложений и запустите макрос ещё раз.", vbInformation
        GoTo CleanExit
    End If

    If Trim$(CStr(sourceWs.Cells(2, COL_SECTION).Value)) <> "Раздел" Or _
       Trim$(CStr(sourceWs.Cells(2, COL_RESPONSIBLE).Value)) <> "Ответственное лицо" Then
        MsgBox "Активный лист не похож на перечень приложений: ожидаются заголовки 'Ответственное лицо' в F2 и 'Раздел' в H2.", vbExclamation
        GoTo CleanExit
    End If

    lastRow = Application.Max(sourceWs.Cells(sourceWs.Rows.Count, COL_SECTION).End(xlUp).Row, _
                               sourceWs.Cells(sourceWs.Rows.Count, COL_RESPONSIBLE).End(xlUp).Row)
    Set pairs = CreateObject("Scripting.Dictionary")
    pairs.CompareMode = vbTextCompare

    For r = SOURCE_FIRST_ROW To lastRow
        DoEvents
        sectionText = Trim$(CStr(sourceWs.Cells(r, COL_SECTION).Value))
        personText = Trim$(CStr(sourceWs.Cells(r, COL_RESPONSIBLE).Value))
        If Len(sectionText) > 0 And Len(personText) > 0 Then
            key = sectionText & Chr$(30) & personText
            If Not pairs.Exists(key) Then pairs.Add key, Array(sectionText, personText)
        End If
        Application.StatusBar = "Формирование data_full: строка " & r & " из " & lastRow
    Next r

    On Error Resume Next
    Set outputWs = GetRegisterWorkbook().Worksheets(OUTPUT_SHEET)
    On Error GoTo Failed
    outputWasExisting = Not outputWs Is Nothing

    ' Keep a full worksheet copy until the rebuild has completed successfully.
    If outputWasExisting Then
        outputWs.Copy After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count)
        Set backupWs = ActiveSheet
        backupWs.Name = GetUniqueDataFullBackupName()
    Else
        Set outputWs = GetOrCreateDataFullSheet()
    End If

    outputWs.Cells.Clear
    outputWs.Range("A1:B1").Value = Array("Раздел", "Ответственное лицо")
    outRow = 2
    For Each item In pairs.Keys
        pairValue = pairs(item)
        outputWs.Cells(outRow, 1).Value = pairValue(0)
        outputWs.Cells(outRow, 2).Value = pairValue(1)
        outRow = outRow + 1
    Next item

    If outRow > 2 Then
        With outputWs.Sort
            .SortFields.Clear
            .SortFields.Add Key:=outputWs.Range("A2:A" & outRow - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=outputWs.Range("B2:B" & outRow - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange outputWs.Range("A1:B" & outRow - 1)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If

    With outputWs.Range("A1:B" & Application.Max(1, outRow - 1))
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With
    With outputWs.Range("A1:B1")
        .Interior.Color = RGB(198, 224, 180)
        .Font.Bold = True
        .Font.Italic = True
        .WrapText = True
    End With
    outputWs.Columns("A:B").AutoFit
    outputWs.Activate
    operationCompleted = True
    If Not backupWs Is Nothing Then
        Application.DisplayAlerts = False
        backupWs.Delete
        Application.DisplayAlerts = oldAlerts
        Set backupWs = Nothing
    End If
    MsgBox "Лист data_full обновлён. Уникальных пар: " & pairs.Count, vbInformation
    GoTo CleanExit

Failed:
    Dim failureText As String
    Dim wasCancelled As Boolean
    failureText = Err.Description
    wasCancelled = (Err.Number = 18)

    ' Restore the previous output if an error or Ctrl+Break occurs mid-rebuild.
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not backupWs Is Nothing Then
        Set damagedWs = outputWs
        If damagedWs Is Nothing Then
            Err.Clear
            backupWs.Name = OUTPUT_SHEET
            If Err.Number = 0 Then
                Set outputWs = backupWs
                Set backupWs = Nothing
            Else
                Err.Clear
                failureText = failureText & vbCrLf & "Не удалось вернуть резервной копии исходное имя. Резервная копия сохранена под именем " & backupWs.Name & "."
            End If
        Else
            recoverySheetName = GetUniqueDataFullRecoveryName()
            Err.Clear
            damagedWs.Name = recoverySheetName
            If Err.Number = 0 Then
                Err.Clear
                backupWs.Name = OUTPUT_SHEET
                If Err.Number = 0 Then
                    Set outputWs = backupWs
                    Set backupWs = Nothing
                    Err.Clear
                    damagedWs.Delete
                    If Err.Number <> 0 Then
                        Err.Clear
                        failureText = failureText & vbCrLf & "Предыдущий лист восстановлен; неполный лист сохранён под именем " & recoverySheetName & "."
                    End If
                Else
                    Err.Clear
                    damagedWs.Name = OUTPUT_SHEET
                    failureText = failureText & vbCrLf & "Не удалось переименовать резервную копию. Она сохранена под именем " & backupWs.Name & "."
                End If
            Else
                Err.Clear
                failureText = failureText & vbCrLf & "Не удалось безопасно переименовать неполный лист. Резервная копия сохранена под именем " & backupWs.Name & "."
            End If
        End If
    ElseIf Not outputWasExisting And Not outputWs Is Nothing And Not operationCompleted Then
        outputWs.Delete
    End If
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0

    If wasCancelled Then
        MsgBox "Формирование data_full отменено пользователем. Предыдущий лист восстановлен, если он существовал." & _
               IIf(InStr(1, failureText, "сохранён", vbTextCompare) > 0 Or InStr(1, failureText, "не удалось", vbTextCompare) > 0, vbCrLf & failureText, ""), vbInformation
    Else
        MsgBox "Не удалось сформировать data_full: " & failureText & _
               IIf(outputWasExisting, vbCrLf & "Предыдущий лист восстановлен, если операция успела изменить его.", ""), vbExclamation
    End If

CleanExit:
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
End Sub

Private Function GetOrCreateDataFullSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateDataFullSheet = GetRegisterWorkbook().Worksheets(OUTPUT_SHEET)
    On Error GoTo 0
    If GetOrCreateDataFullSheet Is Nothing Then
        Set GetOrCreateDataFullSheet = GetRegisterWorkbook().Worksheets.Add(After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count))
        GetOrCreateDataFullSheet.Name = OUTPUT_SHEET
    End If
End Function

Private Function GetUniqueDataFullRecoveryName() As String
    Dim candidate As String
    Dim suffix As Long
    candidate = "_failed_data_full"
    suffix = 1
    Do While DataFullWorksheetExists(candidate)
        candidate = "_failed_df_" & CStr(suffix)
        suffix = suffix + 1
    Loop
    GetUniqueDataFullRecoveryName = candidate
End Function

Private Function DataFullWorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetRegisterWorkbook().Worksheets(sheetName)
    DataFullWorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function GetUniqueDataFullBackupName() As String
    Dim baseName As String
    Dim candidate As String
    Dim suffix As Long
    Dim ws As Worksheet
    Dim exists As Boolean

    baseName = "df_bak_" & Format$(Now, "yymmddhhnnss")
    candidate = baseName
    suffix = 1
    Do
        exists = False
        For Each ws In GetRegisterWorkbook().Worksheets
            If StrComp(ws.Name, candidate, vbTextCompare) = 0 Then
                exists = True
                Exit For
            End If
        Next ws
        If Not exists Then Exit Do
        candidate = baseName & "_" & CStr(suffix)
        suffix = suffix + 1
    Loop
    GetUniqueDataFullBackupName = candidate
End Function
