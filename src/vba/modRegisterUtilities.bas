Attribute VB_Name = "modRegisterUtilities"
Option Explicit

' Builds a unique list of pairs from columns H (Раздел) and F (Ответственное лицо).
' Source register is not changed. Output sheet is named "data_full".

Public Sub BuildDataFull()
    Const FIRST_DATA_ROW As Long = 3
    Const SECTION_COL As Long = 8
    Const PERSON_COL As Long = 6
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim sectionValue As String
    Dim personValue As String
    Dim key As String
    Dim pairs As Object
    Dim k As Variant
    Dim outRow As Long

    On Error GoTo Failed
    Set sourceWs = ActiveSheet
    If LCase$(sourceWs.Name) = "data" Or LCase$(sourceWs.Name) = "data_full" Then
        MsgBox "Перейдите на лист сформированного перечня и повторите запуск.", vbInformation
        Exit Sub
    End If

    lastRow = sourceWs.Cells(sourceWs.Rows.Count, SECTION_COL).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then
        MsgBox "В перечне нет строк данных.", vbInformation
        Exit Sub
    End If

    Set pairs = CreateObject("Scripting.Dictionary")
    pairs.CompareMode = vbTextCompare
    For r = FIRST_DATA_ROW To lastRow
        sectionValue = Trim$(CStr(sourceWs.Cells(r, SECTION_COL).Value))
        personValue = Trim$(CStr(sourceWs.Cells(r, PERSON_COL).Value))
        If Len(sectionValue) > 0 And Len(personValue) > 0 Then
            key = sectionValue & Chr$(30) & personValue
            If Not pairs.Exists(key) Then pairs.Add key, Array(sectionValue, personValue)
        End If
    Next r

    Set targetWs = GetOrCreateDataFullSheet("data_full")
    targetWs.Cells.Clear
    targetWs.Range("A1:B1").Value = Array("Раздел", "Ответственное лицо")
    outRow = 2
    For Each k In pairs.Keys
        targetWs.Cells(outRow, 1).Value = pairs(k)(0)
        targetWs.Cells(outRow, 2).Value = pairs(k)(1)
        outRow = outRow + 1
    Next k

    With targetWs.Range("A1:B" & Application.Max(1, outRow - 1))
        .Font.Name = "Times New Roman"
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With
    With targetWs.Range("A1:B1")
        .Interior.Color = RGB(198, 224, 180)
        .Font.Bold = True
    End With
    targetWs.Columns("A:B").AutoFit
    MsgBox "Лист data_full обновлён. Уникальных пар: " & pairs.Count, vbInformation
    Exit Sub

Failed:
    MsgBox "Не удалось сформировать data_full: " & Err.Description, vbExclamation
End Sub

Private Function GetOrCreateDataFullSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateDataFullSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateDataFullSheet Is Nothing Then
        Set GetOrCreateDataFullSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateDataFullSheet.Name = sheetName
    End If
End Function
