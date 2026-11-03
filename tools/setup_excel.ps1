$ErrorActionPreference = "Stop"

# Устанавливает модули VBA проекта непосредственно в уже созданную рабочую копию книги.
# If the chosen file is not .xlsm, Excel converts it to .xlsm because other formats cannot store VBA.
# Требуются Windows, настольный Excel и разрешённый программный доступ к объектной модели проекта VBA.
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
$backupPath = $null
$convertedOutputCreated = $false

# Сохраняет выбранную книгу до автоматизации Excel или преобразования формата.
try {
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupName = [IO.Path]::GetFileNameWithoutExtension($workbookPath) + ".backup_before_macro_install_" + $stamp + [IO.Path]::GetExtension($workbookPath)
    $backupPath = Join-Path ([IO.Path]::GetDirectoryName($workbookPath)) $backupName
    Copy-Item -LiteralPath $workbookPath -Destination $backupPath -ErrorAction Stop
    Write-Host ("Резервная копия исходной книги: " + $backupPath)
}
catch {
    throw ("Не удалось создать резервную копию исходной книги. Установка остановлена без изменений. Подробности: " + $_.Exception.Message)
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    # Открывает существующую рабочую копию для редактирования, а не только для чтения.
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

    # Проверяет программный доступ к VBA до изменения формата файла или содержимого книги.
    try {
        $vbProject = $workbook.VBProject
        $null = $vbProject.VBComponents.Count
    }
    catch {
        throw ("Excel заблокировал автоматическое добавление VBA-модулей. В Excel откройте Файл → Параметры → Центр управления безопасностью → Параметры центра управления безопасностью → Параметры макросов и включите «Доверять доступ к объектной модели проектов VBA». Затем закройте Excel и запустите установщик снова. Включайте эту настройку только если понимаете риск и доверяете исходникам проекта. Подробности: " + $_.Exception.Message)
    }

    # Форматы .xlsx, .xlsb и .xls не поддерживают хранение проекта VBA. Преобразует выбранную рабочую копию
    # в книгу с поддержкой макросов в той же папке; существующий файл .xlsm редактируется на месте.
    if (-not [string]::Equals([IO.Path]::GetExtension($workbookPath), ".xlsm", [StringComparison]::OrdinalIgnoreCase)) {
        $finalPath = [IO.Path]::ChangeExtension($workbookPath, ".xlsm")
        if (Test-Path -LiteralPath $finalPath) {
            throw ("Рядом уже существует файл .xlsm: " + $finalPath + [Environment]::NewLine + "Чтобы не перезаписать его случайно, переименуйте его или выберите рабочую книгу с другим именем.")
        }
        # Целевого файла ранее не существовало. Помечает его для удаления перед SaveAs
        # чтобы удалить неполный результат, если SaveAs или импорт модуля завершится ошибкой.
        $convertedOutputCreated = $true
        $workbook.SaveAs($finalPath, 52)
        Write-Host ("Рабочая книга преобразована в формат .xlsm: " + $finalPath)
    }

    # Заменяет одноимённые модули проекта, чтобы повторная установка не создавала дубликаты.
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
    if (-not $success -and $convertedOutputCreated -and $finalPath -ne $workbookPath) {
        # Удаляет только новый файл .xlsm, но никогда не удаляет выбранный исходный файл.
        # Исходная книга и её резервная копия, созданная до установки, сохраняются.
        try {
            if (Test-Path -LiteralPath $finalPath -PathType Leaf) {
                Remove-Item -LiteralPath $finalPath -Force -ErrorAction Stop
                Write-Host ("Удалена неполная .xlsm-копия после неудачной установки: " + $finalPath)
            }
        }
        catch {
            Write-Warning ("Не удалось удалить неполную .xlsm-копию. Проверьте файл вручную: " + $finalPath + ". " + $_.Exception.Message)
        }
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
    Write-Host ("Резервная копия исходной книги сохранена здесь: " + $backupPath)
    Read-Host "Нажмите Enter для завершения"
}
