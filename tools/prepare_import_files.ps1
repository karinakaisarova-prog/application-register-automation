# Готовит копии модулей src/vba для ручного импорта через редактор VBA
# (File → Import File…). Редактор VBA читает .bas только в кодировке Windows
# (на русской Windows — 1251), а в репозитории модули хранятся в UTF-8:
# без перекодировки вся кириллица в коде превратилась бы в «кракозябры».
#
# Запуск из папки репозитория:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\prepare_import_files.ps1
# Копии появятся в папке import-cp1251 рядом с src. Импортируйте их, а не файлы из src/vba.

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$sourceDir = Join-Path $repoRoot 'src\vba'
$targetDir = Join-Path $repoRoot 'import-cp1251'
New-Item -ItemType Directory -Force -Path $targetDir | Out-Null

$cp1251 = [Text.Encoding]::GetEncoding(1251)
foreach ($file in Get-ChildItem -LiteralPath $sourceDir -Filter '*.bas') {
    $text = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)
    $text = $text -replace "`r?`n", "`r`n"
    # символы, которых нет в кодировке Windows-1251, VBA всё равно не покажет — предупреждаем
    $roundTrip = $cp1251.GetString($cp1251.GetBytes($text))
    if ($roundTrip -ne $text) {
        Write-Warning ("В файле " + $file.Name + " есть символы вне кодировки Windows-1251: они станут «?».")
    }
    [IO.File]::WriteAllText((Join-Path $targetDir $file.Name), $text, $cp1251)
    Write-Host ("Готово: " + $file.Name)
}
Write-Host ("Копии для импорта: " + $targetDir)
