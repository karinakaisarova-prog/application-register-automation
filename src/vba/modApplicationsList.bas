Attribute VB_Name = "modApplicationsList"
Option Explicit

' Application Register Automation
' Main list builder for desktop Excel on Windows.
' Requires no additional VBA references (FileSystemObject and Word are late-bound).
'
' Entry point: CreateMultiFolderApplicationsListCorrected
' Folder depth: up to MAX_FOLDER_DEPTH below each selected root.
' Manual values in column H ("Раздел") are preserved by full file path.
' No FreezePanes are used.

Private Const HEADER_ROW As Long = 2
Private Const FIRST_DATA_ROW As Long = 3
Private Const MAX_FOLDER_DEPTH As Long = 3
Private Const TEMP_SORT_COL As Long = 11
Private Const DATA_SHEET_NAME As String = "data"

Private mFiles As Collection
Private mSections As Object
Private mLogText As String
Private mRootPath As String

Public Sub CreateMultiFolderApplicationsListCorrected()
    Dim dlg As FileDialog
    Dim roots As Collection
    Dim rootItem As Variant
    Dim ws As Worksheet
    Dim sheetName As String
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim cancelOrError As Boolean
    Dim errText As String

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    mLogText = "Формирование перечня приложений" & vbCrLf

    Set roots = New Collection
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    With dlg
        .Title = "Выберите одну или несколько исходных папок"
        .AllowMultiSelect = True
        If .Show <> -1 Then GoTo CleanExit
        For Each rootItem In .SelectedItems
            roots.Add CStr(rootItem)
        Next rootItem
    End With

    Set mSections = CreateObject("Scripting.Dictionary")
    mSections.CompareMode = vbTextCompare
    SaveExistingSections ActiveSheet, mSections

    Set mFiles = New Collection
    For Each rootItem In roots
        mRootPath = CStr(rootItem)
        LogLine "Поиск файлов: " & mRootPath
        CollectFiles mRootPath, mRootPath, 0
    Next rootItem

    If mFiles.Count = 0 Then
        MsgBox "Подходящие файлы не найдены." & vbCrLf & mLogText, vbInformation, "Перечень приложений"
        GoTo CleanExit
    End If

    sheetName = SafeSheetName(GetFolderName(CStr(roots(1))))
    Set ws = GetOrCreateSheet(sheetName)
    BuildOutputSheet ws, mFiles, mSections
    ApplyLookups ws
    ws.Columns(TEMP_SORT_COL).Hidden = True
    ws.Activate
    Application.StatusBar = False
    MsgBox "Перечень сформирован. Найдено файлов: " & mFiles.Count & vbCrLf & vbCrLf & mLogText, _
           vbInformation, "Перечень приложений"
    GoTo CleanExit

Failed:
    cancelOrError = (Err.Number = 18)
    errText = Err.Description
    If cancelOrError Then
        LogLine "Операция отменена пользователем."
    Else
        LogLine "Ошибка " & Err.Number & ": " & errText
    End If
    On Error Resume Next
    Application.StatusBar = False
    MsgBox IIf(cancelOrError, "Операция отменена.", "Не удалось завершить формирование перечня: " & errText) & _
           vbCrLf & vbCrLf & mLogText, vbExclamation, "Перечень приложений"

CleanExit:
    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
End Sub

Private Sub CollectFiles(ByVal rootPath As String, ByVal currentPath As String, ByVal depth As Long)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim fileItem As Object
    Dim ext As String
    Dim info As Object
    Dim objectName As String

    On Error GoTo FolderError
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(currentPath)

    For Each fileItem In folder.Files
        ext = LCase$(fso.GetExtensionName(fileItem.Name))
        If IsSupportedExtension(ext) Then
            Set info = ParseFileName(CStr(fileItem.Name))
            objectName = GetObjectName(rootPath, currentPath)
            Dim rec As Object
            Set rec = CreateObject("Scripting.Dictionary")
            rec("Path") = CStr(fileItem.Path)
            rec("FileName") = CStr(fileItem.Name)
            rec("Extension") = ext
            rec("Object") = objectName
            rec("AppNumber") = CStr(info("AppNumber"))
            rec("Title") = CStr(info("Title"))
            rec("IsReport") = IsMainReport(CStr(fileItem.Name))
            rec("Folder") = CStr(currentPath)
            mFiles.Add rec
        End If
    Next fileItem

    If depth < MAX_FOLDER_DEPTH Then
        For Each subFolder In folder.SubFolders
            DoEvents
            Application.StatusBar = "Обработка папки: " & CStr(subFolder.Path)
            CollectFiles rootPath, CStr(subFolder.Path), depth + 1
        Next subFolder
    End If
    Exit Sub

FolderError:
    LogLine "Не удалось прочитать папку: " & currentPath & " (" & Err.Description & ")"
    Err.Clear
End Sub

Private Function IsSupportedExtension(ByVal ext As String) As Boolean
    IsSupportedExtension = (ext = "doc" Or ext = "docx" Or ext = "pdf" Or _
                            ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "xlsb")
End Function

Private Function IsMainReport(ByVal fileName As String) As Boolean
    IsMainReport = (LCase$(Left$(Trim$(fileName), 6)) = "отчет_")
End Function

Private Function ParseFileName(ByVal fileName As String) As Object
    Dim fso As Object
    Dim baseName As String
    Dim result As Object
    Dim rx As Object
    Dim matches As Object
    Dim appNumber As String
    Dim titleText As String
    Dim p As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    baseName = fso.GetBaseName(fileName)
    Set result = CreateObject("Scripting.Dictionary")
    appNumber = ""
    titleText = baseName

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = "приложени[ея]\s*№?\s*([0-9]+[А-ЯA-Z]?)"
    If rx.Test(baseName) Then
        Set matches = rx.Execute(baseName)
        appNumber = matches(0).SubMatches(0)
        p = InStr(1, baseName, matches(0).Value, vbTextCompare)
        titleText = Trim$(Mid$(baseName, p + Len(matches(0).Value)))
        titleText = Replace(titleText, "_", " ")
        titleText = Trim$(titleText)
        If Left$(titleText, 1) = "_" Then titleText = Mid$(titleText, 2)
        If Right$(titleText, 1) = "_" Then titleText = Left$(titleText, Len(titleText) - 1)
    ElseIf LCase$(Left$(baseName, 6)) = "отчет_" Then
        titleText = Mid$(baseName, 7)
        appNumber = ""
    Else
        titleText = Replace(baseName, "_", " ")
    End If

    If Len(titleText) = 0 Then titleText = baseName
    result("AppNumber") = appNumber
    result("Title") = titleText
    Set ParseFileName = result
End Function

Private Function GetObjectName(ByVal rootPath As String, ByVal currentPath As String) As String
    Dim fso As Object
    Dim relativePath As String
    Dim parts() As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    rootPath = RemoveTrailingSlash(rootPath)
    currentPath = RemoveTrailingSlash(currentPath)

    If StrComp(rootPath, currentPath, vbTextCompare) = 0 Then
        GetObjectName = GetFolderName(rootPath)
        Exit Function
    End If

    relativePath = Mid$(currentPath, Len(rootPath) + 2)
    parts = Split(relativePath, "\")
    If UBound(parts) >= 0 And Len(parts(0)) > 0 Then
        GetObjectName = parts(0)
    Else
        GetObjectName = GetFolderName(currentPath)
    End If
End Function

Private Sub SaveExistingSections(ByVal ws As Worksheet, ByVal target As Object)
    Dim lastRow As Long
    Dim r As Long
    Dim filePath As String
    Dim sectionValue As String

    On Error GoTo Done
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then Exit Sub

    For r = FIRST_DATA_ROW To lastRow
        filePath = GetCellLinkPath(ws.Cells(r, 7))
        sectionValue = CStr(ws.Cells(r, 8).Value)
        If Len(filePath) > 0 And Len(sectionValue) > 0 Then target(filePath) = sectionValue
    Next r
Done:
End Sub

Private Function GetCellLinkPath(ByVal c As Range) As String
    On Error Resume Next
    If c.Hyperlinks.Count > 0 Then
        GetCellLinkPath = CStr(c.Hyperlinks(1).Address)
    Else
        GetCellLinkPath = CStr(c.Value)
    End If
    On Error GoTo 0
End Function

Private Sub BuildOutputSheet(ByVal ws As Worksheet, ByVal files As Collection, ByVal sections As Object)
    Dim headers As Variant
    Dim i As Long
    Dim r As Long
    Dim rec As Object
    Dim sectionValue As String
    Dim sortKey As String

    headers = Array("№ п/п", "Объект", "№ Прил.", "Наименование", "Отдел", _
                    "Ответственное лицо", "Ссылка на файл", "Раздел", "Стр. отчёта", "Кол-во страниц")

    ws.Cells.Clear
    On Error Resume Next
    ws.Hyperlinks.Delete
    On Error GoTo 0

    With ws.Range("A1:J1")
        .Merge
        .Value = "Перечень приложений"
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 30

    For i = LBound(headers) To UBound(headers)
        ws.Cells(HEADER_ROW, i + 1).Value = headers(i)
    Next i
    ws.Cells(HEADER_ROW, TEMP_SORT_COL).Value = "Временный ключ"

    r = FIRST_DATA_ROW
    For i = 1 To files.Count
        Set rec = files(i)
        ws.Cells(r, 2).Value = rec("Object")
        ws.Cells(r, 3).Value = rec("AppNumber")
        ws.Cells(r, 4).Value = rec("Title")
        ws.Cells(r, 7).Value = "Открыть файл"
        ws.Hyperlinks.Add Anchor:=ws.Cells(r, 7), Address:=CStr(rec("Path")), TextToDisplay:="Открыть файл"

        sectionValue = ""
        If sections.Exists(CStr(rec("Path"))) Then sectionValue = CStr(sections(CStr(rec("Path"))))
        ws.Cells(r, 8).Value = sectionValue

        If IsSpecialDepartmentFile(CStr(rec("FileName")), CStr(rec("Title"))) Then
            ws.Cells(r, 5).Value = "—"
        End If

        sortKey = UCase$(CStr(rec("Object"))) & "|" & IIf(CBool(rec("IsReport")), "0", "1") & "|" & _
                  Format$(Val(CStr(rec("AppNumber"))), "000000") & "|" & UCase$(CStr(rec("Title")))
        ws.Cells(r, TEMP_SORT_COL).Value = sortKey
        r = r + 1
    Next i

    If r > FIRST_DATA_ROW Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range(ws.Cells(FIRST_DATA_ROW, TEMP_SORT_COL), ws.Cells(r - 1, TEMP_SORT_COL)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(r - 1, TEMP_SORT_COL))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
        For i = FIRST_DATA_ROW To r - 1
            ws.Cells(i, 1).Value = i - FIRST_DATA_ROW + 1
        Next i
    End If

    ApplyTableFormatting ws, r - 1
End Sub

Private Sub ApplyLookups(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim appName As String

    If Not WorksheetExists(DATA_SHEET_NAME) Then
        LogLine "Предупреждение: лист 'data' не найден; столбцы 'Отдел' и 'Ответственное лицо' не заполнены автоматически."
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    For r = FIRST_DATA_ROW To lastRow
        appName = Trim$(CStr(ws.Cells(r, 4).Value))
        If IsSpecialDepartmentFile(CStr(ws.Cells(r, 3).Value), appName) Then
            ws.Cells(r, 5).Value = "—"
        Else
            ws.Cells(r, 5).FormulaLocal = "=ЕСЛИОШИБКА(ИНДЕКС(data!$B:$B;ПОИСКПОЗ($D" & r & ";data!$A:$A;0));"""")"
            ws.Cells(r, 6).FormulaLocal = "=ЕСЛИОШИБКА(ИНДЕКС(data!$C:$F;ПОИСКПОЗ($D" & r & ";data!$A:$A;0);ПОИСКПОЗ($B" & r & ";data!$C$1:$F$1;0));"""")"
        End If
    Next r
End Sub

Private Function IsSpecialDepartmentFile(ByVal fileNameOrNumber As String, ByVal titleText As String) As Boolean
    Dim s As String
    s = LCase$(fileNameOrNumber & " " & titleText)
    IsSpecialDepartmentFile = (InStr(1, s, "отчет_2 кв 2026", vbTextCompare) > 0 Or _
                               InStr(1, s, "отчет 2 кв 2026", vbTextCompare) > 0 Or _
                               InStr(1, s, "2 кв 2026", vbTextCompare) > 0 Or _
                               InStr(1, s, "реестр рисков", vbTextCompare) > 0)
End Function

Private Sub ApplyTableFormatting(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim fullRange As Range
    Dim widths As Variant
    Dim i As Long

    If lastRow < HEADER_ROW Then lastRow = HEADER_ROW
    Set fullRange = ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow, TEMP_SORT_COL))
    With fullRange
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, 10))
        .Interior.Color = RGB(198, 224, 180)
        .Font.Color = RGB(31, 78, 45)
        .Font.Italic = True
        .Font.Bold = False
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .RowHeight = 32
    End With

    widths = Array(6, 12, 8, 11, 16, 22, 16, 11, 10, 10)
    For i = LBound(widths) To UBound(widths)
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i
    ws.Columns(TEMP_SORT_COL).Hidden = True
    ws.Columns(1).HorizontalAlignment = xlCenter
    ws.Columns(3).HorizontalAlignment = xlCenter
    ws.Columns(8).HorizontalAlignment = xlCenter
    ws.Columns(9).HorizontalAlignment = xlCenter
    ws.Columns(10).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(lastRow, 10)).RowHeight = 20
    ws.Columns(4).WrapText = True
    ws.Columns(6).WrapText = True
    ws.Columns(7).WrapText = True
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function SafeSheetName(ByVal value As String) As String
    Dim invalidChars As Variant
    Dim i As Long
    invalidChars = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(invalidChars) To UBound(invalidChars)
        value = Replace(value, CStr(invalidChars(i)), "_")
    Next i
    value = Trim$(value)
    If Len(value) = 0 Then value = "Перечень приложений"
    If Len(value) > 31 Then value = Left$(value, 31)
    SafeSheetName = value
End Function

Private Function GetFolderName(ByVal folderPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFolderName = fso.GetFolder(RemoveTrailingSlash(folderPath)).Name
End Function

Private Function RemoveTrailingSlash(ByVal value As String) As String
    Do While Len(value) > 3 And (Right$(value, 1) = "\" Or Right$(value, 1) = "/")
        value = Left$(value, Len(value) - 1)
    Loop
    RemoveTrailingSlash = value
End Function

Private Sub LogLine(ByVal message As String)
    mLogText = mLogText & Format$(Now, "hh:nn:ss") & "  " & message & vbCrLf
    Application.StatusBar = message
    DoEvents
End Sub
