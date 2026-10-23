$ErrorActionPreference = "Stop"

# Builds a macro-enabled workbook, imports the project's VBA modules,
# and copies the existing "data" worksheet from a workbook chosen by the user.
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
    $sourceDialog.Title = "Выберите существующую Excel-книгу, в которой уже есть лист data"
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
    $outputDialog.Title = "Сохранить новую книгу с макросами и листом data"
    $outputDialog.Filter = "Книга Excel с поддержкой макросов (*.xlsm)|*.xlsm"
    $outputDialog.FileName = "ApplicationRegisterAutomation.xlsm"
    $outputDialog.OverwritePrompt = $true
    $outputDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
    if ($outputDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Установка отменена. Ничего не создано." -ForegroundColor Yellow
        exit 0
    }
    $outputPath = [IO.Path]::GetFullPath($outputDialog.FileName)
    if ([IO.Path]::GetExtension($outputPath) -ne ".xlsm") {
        $outputPath += ".xlsm"
    }
    if ([string]::Equals($sourcePath, $outputPath, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Нельзя сохранять новую книгу поверх исходного файла со справочником data. Выберите другое имя или папку."
    }
}
finally {
    $outputDialog.Dispose()
}

$excel = $null
$workbook = $null
$sourceWorkbook = $null
$success = $false
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $sourceWorkbook = $excel.Workbooks.Open($sourcePath, 0, $true)
    $sourceData = $null
    foreach ($sheet in $sourceWorkbook.Worksheets) {
        if ([string]::Equals([string]$sheet.Name, "data", [StringComparison]::Ordinal)) {
            $sourceData = $sheet
            break
        }
    }
    if ($sourceData -eq $null) {
        throw "В выбранной книге не найден лист с точным именем data. Исходная книга не изменена. Выберите правильный файл."
    }

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
        Write-Host ("Добавлен модуль: " + $name)
    }

    # Copy the user's existing lookup sheet; copying does not edit the source workbook.
    $sourceData.Copy($workbook.Worksheets.Item(1))
    $dataSheet = $null
    foreach ($sheet in $workbook.Worksheets) {
        if ([string]::Equals([string]$sheet.Name, "data", [StringComparison]::Ordinal)) {
            $dataSheet = $sheet
            break
        }
    }
    if ($dataSheet -eq $null) {
        throw "Не удалось скопировать лист data. Новая книга не будет сохранена."
    }

    # Keep only the copied data sheet in the launcher workbook.
    for ($i = $workbook.Worksheets.Count; $i -ge 1; $i--) {
        $sheet = $workbook.Worksheets.Item($i)
        if (-not [string]::Equals([string]$sheet.Name, "data", [StringComparison]::Ordinal)) {
            $sheet.Delete()
        }
    }

    # 52 = xlOpenXMLWorkbookMacroEnabled
    $workbook.SaveAs($outputPath, 52)
    $success = $true
}
finally {
    if ($sourceWorkbook -ne $null) {
        try { $sourceWorkbook.Close($false) } catch {}
    }
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
    foreach ($obj in @($sourceWorkbook, $workbook, $excel)) {
        if ($obj -ne $null) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {}
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($success) {
    Write-Host ""
    Write-Host ("Готово: " + $outputPath) -ForegroundColor Green
    Write-Host "Лист data скопирован из выбранной книги; исходный файл не изменялся."
    Write-Host "Откройте новую .xlsm-книгу в настольном Excel и сначала выполните Debug → Compile VBAProject (Alt+F11)."
    Write-Host "Не запускайте макросы до успешной компиляции."
    Read-Host "Нажмите Enter для завершения"
}
