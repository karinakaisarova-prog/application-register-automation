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

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set sourceWs = ActiveSheet
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

    Set outputWs = GetOrCreateDataFullSheet()
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
    MsgBox "Лист data_full обновлён. Уникальных пар: " & pairs.Count, vbInformation
    GoTo CleanExit

Failed:
    If Err.Number = 18 Then
        MsgBox "Формирование data_full отменено пользователем.", vbInformation
    Else
        MsgBox "Не удалось сформировать data_full: " & Err.Description, vbExclamation
    End If

CleanExit:
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
End Sub

Private Function GetOrCreateDataFullSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateDataFullSheet = ThisWorkbook.Worksheets(OUTPUT_SHEET)
    On Error GoTo 0
    If GetOrCreateDataFullSheet Is Nothing Then
        Set GetOrCreateDataFullSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateDataFullSheet.Name = OUTPUT_SHEET
    End If
End Function
