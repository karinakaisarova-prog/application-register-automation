$ErrorActionPreference = "Stop"

# Builds a macro-enabled Excel launcher from the repository's .bas modules.
# Requires Windows, desktop Excel, and Excel's "Trust access to the VBA project object model".
$repoRoot = Split-Path -Parent $PSScriptRoot
$vbaDir = Join-Path $repoRoot "src\vba"
$moduleFiles = @(
    "modWorkbookContext.bas",
    "modRegisterPanel.bas",
    "modApplicationsList.bas",
    "modReportAnalysis.bas",
    "modTableNumbering.bas",
    "modBulletListAudit.bas",
    "modRenameFiles.bas",
    "modDataFull.bas",
    "modSafetySupport.bas"
)

foreach ($name in $moduleFiles) {
    $fullPath = Join-Path $vbaDir $name
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        throw ("Не найден модуль: " + $fullPath + [Environment]::NewLine + "Убедитесь, что архив репозитория полностью распакован.")
    }
}

Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.SaveFileDialog
try {
    $dialog.Title = "Сохранить книгу запуска автоматизации"
    $dialog.Filter = "Книга Excel с поддержкой макросов (*.xlsm)|*.xlsm"
    $dialog.FileName = "ApplicationRegisterAutomation.xlsm"
    $dialog.OverwritePrompt = $true
    $dialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Установка отменена. Ничего не сохранено." -ForegroundColor Yellow
        exit 0
    }
    $outputPath = $dialog.FileName
    if ([IO.Path]::GetExtension($outputPath) -ne ".xlsm") {
        $outputPath += ".xlsm"
    }
}
finally {
    $dialog.Dispose()
}

$excel = $null
$workbook = $null
$success = $false
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()

    try {
        $vbProject = $workbook.VBProject
        $null = $vbProject.VBComponents.Count
    }
    catch {
        throw ("Excel заблокировал автоматическое добавление VBA-модулей. В Excel откройте Файл → Параметры → Центр управления безопасностью → Параметры центра управления безопасностью → Параметры макросов и включите «Доверять доступ к объектной модели проектов VBA». Затем закройте Excel и запустите установщик снова. Включайте эту настройку только если понимаете риск и доверяете исходникам проекта. Подробности: " + $_.Exception.Message)
    }

    foreach ($name in $moduleFiles) {
        $fullPath = Join-Path $vbaDir $name
        $null = $vbProject.VBComponents.Import($fullPath)
        Write-Host "Добавлен модуль: $name"
    }

    # 52 = xlOpenXMLWorkbookMacroEnabled
    $workbook.SaveAs($outputPath, 52)
    $success = $true
}
finally {
    if ($workbook -ne $null) {
        if (-not $success) {
            try { $workbook.Close($false) } catch {}
        } else {
            try { $workbook.Close($true) } catch {}
        }
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
    }
    if ($workbook -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
    if ($excel -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($success) {
    Write-Host ""
    Write-Host ("Готово: " + $outputPath) -ForegroundColor Green
    Write-Host "Откройте эту .xlsm-книгу в настольном Excel и сначала выполните Debug → Compile VBAProject (Alt+F11)."
    Write-Host "Не запускайте макросы до успешной компиляции. Лист data и реальные справочные данные подготовьте отдельно."
    Read-Host "Нажмите Enter для завершения"
}
