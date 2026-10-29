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

    Set ws = GetActiveRegisterSheet()
    If ws Is Nothing Then
        MsgBox "Не найден лист перечня в книге результата. Сначала сформируйте перечень и откройте его лист.", _
               vbExclamation, "Панель макросов"
        Exit Sub
    End If

    CreateRegisterControlPanel ws
    MsgBox "Панель макросов добавлена справа от таблицы. Если её не видно, прокрутите лист вправо.", _
           vbInformation, "Панель макросов"
End Sub

Public Sub CreateRegisterControlPanel(ByVal ws As Worksheet)
    Dim buttonNames As Variant
    Dim buttonCaptions As Variant
    Dim macroNames As Variant
    Dim i As Long
    Dim shp As Shape
    Dim leftPos As Double
    Dim topPos As Double
    Dim actionName As String
    Dim workbookRef As String

    If ws Is Nothing Then Exit Sub

    DeleteRegisterControlPanel ws

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

        shp.Name = PANEL_PREFIX & CStr(i + 1)
        shp.TextFrame2.TextRange.Text = CStr(buttonCaptions(i))
        shp.TextFrame2.TextRange.Font.Name = "Arial"
        shp.TextFrame2.TextRange.Font.Size = 9
        shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.Placement = xlFreeFloating
        shp.OnAction = "'" & workbookRef & CStr(macroNames(i))
    Next i
End Sub

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
