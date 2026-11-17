Attribute VB_Name = "modRegisterPanel"
Option Explicit

' Панель управления макросами на листе перечня.
' Кнопки размещаются справа от таблицы, начиная с колонки L,
' чтобы не сдвигать строки перечня и не менять формулы.

Private Const PANEL_PREFIX As String = "regPanel_"
Private Const PANEL_LEFT_COLUMN As Long = 12
Private Const BUTTON_WIDTH As Double = 190
Private Const BUTTON_HEIGHT As Double = 30
Private Const BUTTON_GAP As Double = 6

Public Sub AddRegisterControlPanel()
    Dim ws As Worksheet
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo WorkbookNotFound
    Set ws = GetActiveRegisterSheet()
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Не найден лист перечня в книге результата. Сначала сформируйте перечень и откройте его лист.", _
               vbExclamation, "Панель макросов"
        Exit Sub
    End If

    On Error GoTo PanelFailed
    CreateRegisterControlPanel ws
    On Error GoTo 0

    MsgBox "Панель макросов добавлена справа от таблицы. Если её не видно, прокрутите лист вправо.", _
           vbInformation, "Панель макросов"
    Exit Sub

WorkbookNotFound:
    If Err.Number = 18 Then
        On Error GoTo 0
        Err.Raise 18
    End If
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error GoTo 0
    MsgBox "Не удалось определить книгу с листом «data»." & vbCrLf & _
           "Ошибка " & failureNumber & ": " & failureDescription, _
           vbExclamation, "Панель макросов"
    Exit Sub

PanelFailed:
    If Err.Number = 18 Then
        On Error GoTo 0
        Err.Raise 18
    End If
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error GoTo 0
    MsgBox "Не удалось полностью обновить панель макросов. Новые кнопки могли быть добавлены, а прежние — остаться." & vbCrLf & _
           "Ошибка " & failureNumber & ": " & failureDescription, _
           vbExclamation, "Панель макросов"
End Sub

Public Sub CreateRegisterControlPanel(ByVal ws As Worksheet)
    Dim buttonCaptions As Variant
    Dim macroNames As Variant
    Dim i As Long
    Dim shp As Shape
    Dim leftPos As Double
    Dim topPos As Double
    Dim actionName As String
    Dim workbookRef As String
    Dim batchPrefix As String
    Dim replacementStarted As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String

    If ws Is Nothing Then Exit Sub

    On Error GoTo CreationFailed

    buttonCaptions = Array( _
        "Повторно создать перечень", _
        "Анализ упоминаний в отчёте", _
        "Подсчитать страницы", _
        "Создать data_full", _
        "Переименовать файлы", _
        "Откатить переименование", _
        "Проверить нумерацию таблиц Word", _
        "Проверить маркированные списки Word", _
        "Открыть памятку", _
        "Обновить панель")

    macroNames = Array( _
        "CreateMultiFolderApplicationsListCorrected", _
        "ScanWordReportMentions", _
        "CountPagesForRegister", _
        "BuildDataFull", _
        "RenameFilesFromRegister", _
        "RollbackLastRenameBatch", _
        "CheckAndNumberWordTables", _
        "CheckAndFixWordBulletLists", _
        "OpenMacroLaunchGuide", _
        "AddRegisterControlPanel")

    batchPrefix = GetUniquePanelPrefix(ws)
    workbookRef = Replace(ThisWorkbook.Name, "'", "''") & "'!"
    leftPos = ws.Columns(PANEL_LEFT_COLUMN).Left
    topPos = ws.Rows(1).Top

    For i = LBound(buttonCaptions) To UBound(buttonCaptions)
        Set shp = ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            leftPos, _
            topPos + i * (BUTTON_HEIGHT + BUTTON_GAP), _
            BUTTON_WIDTH, _
            BUTTON_HEIGHT)

        shp.Name = batchPrefix & CStr(i + 1)
        shp.TextFrame2.TextRange.Text = CStr(buttonCaptions(i))
        shp.TextFrame2.TextRange.Font.Name = "Arial"
        shp.TextFrame2.TextRange.Font.Size = 9
        shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.Placement = xlFreeFloating
        shp.OnAction = "'" & workbookRef & CStr(macroNames(i))
    Next i

    replacementStarted = True

    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(PANEL_PREFIX)) = PANEL_PREFIX Then
            If Left$(ws.Shapes(i).Name, Len(batchPrefix)) <> batchPrefix Then
                ws.Shapes(i).Delete
            End If
        End If
    Next i

    On Error GoTo 0
    Exit Sub

CreationFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description

    If Not replacementStarted And Len(batchPrefix) > 0 Then
        On Error Resume Next
        For i = ws.Shapes.Count To 1 Step -1
            If Left$(ws.Shapes(i).Name, Len(batchPrefix)) = batchPrefix Then
                ws.Shapes(i).Delete
            End If
        Next i
        On Error GoTo 0
    End If

    If failureNumber = 18 Then Err.Raise 18
    Err.Raise failureNumber, "CreateRegisterControlPanel", failureDescription
End Sub

Private Function GetUniquePanelPrefix(ByVal ws As Worksheet) As String
    Dim basePrefix As String
    Dim candidate As String
    Dim suffix As Long
    Dim i As Long
    Dim exists As Boolean

    basePrefix = PANEL_PREFIX & "new_" & Format$(Now, "yyyymmddhhnnss")
    candidate = basePrefix & "_"

    Do
        exists = False
        For i = 1 To ws.Shapes.Count
            If Left$(ws.Shapes(i).Name, Len(candidate)) = candidate Then
                exists = True
                Exit For
            End If
        Next i

        If exists Then
            suffix = suffix + 1
            candidate = basePrefix & "_" & CStr(suffix) & "_"
        End If
    Loop While exists

    GetUniquePanelPrefix = candidate
End Function

Public Sub DeleteRegisterControlPanel(ByVal ws As Worksheet)
    Dim i As Long
    If ws Is Nothing Then Exit Sub

    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(PANEL_PREFIX)) = PANEL_PREFIX Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub

Public Sub OpenMacroLaunchGuide()
    On Error GoTo OpenError
    ThisWorkbook.FollowHyperlink Address:="https://github.com/karinakaisarova-prog/application-register-automation/blob/main/docs/macros-launch-order.md"
    Exit Sub
OpenError:
    MsgBox "Не удалось открыть памятку автоматически." & vbCrLf & _
           "Откройте ссылку вручную:" & vbCrLf & _
           "https://github.com/karinakaisarova-prog/application-register-automation/blob/main/docs/macros-launch-order.md", _
           vbExclamation, "Памятка по запуску"
End Sub
