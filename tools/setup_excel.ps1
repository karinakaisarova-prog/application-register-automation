$ErrorActionPreference = "Stop"

# Creates a macro-enabled COPY of the user's existing workbook, preserving all sheets,
# imports the project's VBA modules, and leaves the original workbook untouched.
# Requires Windows, desktop Excel, and trusted access to the VBA project object model.
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

$sourceDialog = New-Object System.Windows.Forms.OpenFileDialog
try {
    $sourceDialog.Title = "Выберите ваш Excel-файл с листом data"
    $sourceDialog.Filter = "Книги Excel (*.xlsx;*.xlsm;*.xlsb;*.xls)|*.xlsx;*.xlsm;*.xlsb;*.xls"
    $sourceDialog.Multiselect = $false
    $sourceDialog.CheckFileExists = $true
    if ($sourceDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Установка отменена. Ничего не создано." -ForegroundColor Yellow
        exit 0
    }
    $sourcePath = [IO.Path]::GetFullPath($sourceDialog.FileName)
}
finally {
    $sourceDialog.Dispose()
}

$outputDialog = New-Object System.Windows.Forms.SaveFileDialog
try {
    $outputDialog.Title = "Сохранить копию этой книги с макросами"
    $outputDialog.Filter = "Книга Excel с поддержкой макросов (*.xlsm)|*.xlsm"
    $outputDialog.FileName = ([IO.Path]::GetFileNameWithoutExtension($sourcePath) + "_с_макросами.xlsm")
    $outputDialog.OverwritePrompt = $true
    $outputDialog.InitialDirectory = [IO.Path]::GetDirectoryName($sourcePath)
    if ($outputDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Установка отменена. Ничего не создано." -ForegroundColor Yellow
        exit 0
    }
    $outputPath = [IO.Path]::GetFullPath($outputDialog.FileName)
    if ([IO.Path]::GetExtension($outputPath) -ne ".xlsm") {
        $outputPath += ".xlsm"
    }
    if ([string]::Equals($sourcePath, $outputPath, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Нельзя сохранять поверх исходного файла. Выберите другое имя: оригинал останется без изменений."
    }
}
finally {
    $outputDialog.Dispose()
}

$excel = $null
$workbook = $null
$success = $false
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbook = $excel.Workbooks.Open($sourcePath, 0, $true)

    $sourceData = $null
    foreach ($sheet in $workbook.Worksheets) {
        if ([string]::Equals([string]$sheet.Name, "data", [StringComparison]::Ordinal)) {
            $sourceData = $sheet
            break
        }
    }
    if ($sourceData -eq $null) {
        throw "В выбранной книге не найден лист с точным именем data. Исходная книга не изменена. Выберите правильный файл."
    }

    try {
        $vbProject = $workbook.VBProject
        $null = $vbProject.VBComponents.Count
    }
    catch {
        throw ("Excel заблокировал автоматическое добавление VBA-модулей. В Excel откройте Файл → Параметры → Центр управления безопасностью → Параметры центра управления безопасностью → Параметры макросов и включите «Доверять доступ к объектной модели проектов VBA». Затем закройте Excel и запустите установщик снова. Включайте эту настройку только если понимаете риск и доверяете исходникам проекта. Подробности: " + $_.Exception.Message)
    }

    # Convert/save a separate macro-enabled copy before importing any code.
    # The selected original is opened read-only and will be closed without saving.
    $workbook.SaveAs($outputPath, 52)

    # Replace only modules with the same names as the project modules to avoid duplicate components.
    foreach ($name in $moduleFiles) {
        $moduleName = [IO.Path]::GetFileNameWithoutExtension($name)
        for ($i = $vbProject.VBComponents.Count; $i -ge 1; $i--) {
            $component = $vbProject.VBComponents.Item($i)
            if ([string]::Equals([string]$component.Name, $moduleName, [StringComparison]::OrdinalIgnoreCase)) {
                $vbProject.VBComponents.Remove($component)
            }
        }
        $fullPath = Join-Path $vbaDir $name
        $null = $vbProject.VBComponents.Import($fullPath)
        Write-Host ("Добавлен модуль: " + $name)
    }

    $workbook.Save()
    $success = $true
}
finally {
    if ($workbook -ne $null) {
        try { $workbook.Close($success) } catch {}
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
    }
    if ($workbook -ne $null) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {} }
    if ($excel -ne $null) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($success) {
    Write-Host ""
    Write-Host ("Готово: " + $outputPath) -ForegroundColor Green
    Write-Host "Это отдельная копия выбранной книги: все листы, включая data, сохранены в ней."
    Write-Host "Перечень будет формироваться в этой же копии, а не в отдельной книге результата."
    Write-Host "Исходный файл не изменён. Откройте новую .xlsm-книгу и выполните Alt+F11 → Debug → Compile VBAProject."
    Write-Host "Не запускайте макросы до успешной компиляции."
    Read-Host "Нажмите Enter для завершения"
}
