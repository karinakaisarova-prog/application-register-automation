$ErrorActionPreference = "Stop"

# Installs project VBA modules directly into the user's already-created working copy.
# If the chosen file is not .xlsm, Excel converts it to .xlsm because other formats cannot store VBA.
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
$dialog = New-Object System.Windows.Forms.OpenFileDialog
try {
    $dialog.Title = "Выберите свою рабочую копию Excel-книги с листом data"
    $dialog.Filter = "Книги Excel (*.xlsm;*.xlsx;*.xlsb;*.xls)|*.xlsm;*.xlsx;*.xlsb;*.xls"
    $dialog.Multiselect = $false
    $dialog.CheckFileExists = $true
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Установка отменена. Ничего не изменено." -ForegroundColor Yellow
        exit 0
    }
    $workbookPath = [IO.Path]::GetFullPath($dialog.FileName)
}
finally {
    $dialog.Dispose()
}

$excel = $null
$workbook = $null
$success = $false
$finalPath = $workbookPath
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    # Open the already-existing working copy for editing, not read-only.
    $workbook = $excel.Workbooks.Open($workbookPath, 0, $false)

    $dataSheet = $null
    foreach ($sheet in $workbook.Worksheets) {
        if ([string]::Equals([string]$sheet.Name, "data", [StringComparison]::Ordinal)) {
            $dataSheet = $sheet
            break
        }
    }
    if ($dataSheet -eq $null) {
        throw "В выбранной книге не найден лист с точным именем data. Закройте установщик, выберите правильную рабочую копию и повторите."
    }

    # .xlsx/.xlsb/.xls cannot store VBA projects. Convert the selected working copy
    # to a macro-enabled workbook in the same folder; existing .xlsm is edited in place.
    if (-not [string]::Equals([IO.Path]::GetExtension($workbookPath), ".xlsm", [StringComparison]::OrdinalIgnoreCase)) {
        $finalPath = [IO.Path]::ChangeExtension($workbookPath, ".xlsm")
        if (Test-Path -LiteralPath $finalPath) {
            throw ("Рядом уже существует файл .xlsm: " + $finalPath + [Environment]::NewLine + "Чтобы не перезаписать его случайно, переименуйте его или выберите рабочую книгу с другим именем.")
        }
        $workbook.SaveAs($finalPath, 52)
        Write-Host ("Рабочая книга преобразована в формат .xlsm: " + $finalPath)
    }

    try {
        $vbProject = $workbook.VBProject
        $null = $vbProject.VBComponents.Count
    }
    catch {
        throw ("Excel заблокировал автоматическое добавление VBA-модулей. В Excel откройте Файл → Параметры → Центр управления безопасностью → Параметры центра управления безопасностью → Параметры макросов и включите «Доверять доступ к объектной модели проектов VBA». Затем закройте Excel и запустите установщик снова. Включайте эту настройку только если понимаете риск и доверяете исходникам проекта. Подробности: " + $_.Exception.Message)
    }

    # Replace same-named project modules so repeated installation does not create duplicates.
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
    Write-Host ("Готово. Макросы добавлены в рабочую книгу: " + $finalPath) -ForegroundColor Green
    Write-Host "Лист data и остальные листы остались в этой же книге. Перечень будет создаваться в ней же."
    Write-Host "Откройте эту .xlsm-книгу и выполните Alt+F11 → Debug → Compile VBAProject."
    Write-Host "Не запускайте макросы до успешной компиляции."
    Read-Host "Нажмите Enter для завершения"
}
