from pathlib import Path
import re
import sys

root = Path(__file__).resolve().parents[1]
modules = sorted((root / 'src' / 'vba').glob('*.bas'))
errors = []
seen = {}
required = {
    'StartRegisterAutomation',
    'CreateMultiFolderApplicationsListCorrected',
    'ScanWordReportMentions',
    'CountPagesForRegister',
    'BuildDataFull',
    'RenameFilesFromRegister',
    'RollbackLastRenameBatch',
    'CheckAndNumberWordTables',
    'CheckAndFixWordBulletLists',
}

for path in modules:
    lines = path.read_text(encoding='utf-8-sig').splitlines()
    stack = []

    for number, line in enumerate(lines, 1):
        if not line.lstrip().startswith("'") and 'FreezePanes' in line:
            errors.append(f'{path.relative_to(root)}:{number}: unexpected FreezePanes reference')

        match = re.match(
            r'^\s*(?:(?:Public|Private|Friend)\s+)?(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\b',
            line, re.I
        )
        end = re.match(r'^\s*End\s+(Sub|Function)\s*$', line, re.I)

        if match:
            kind, name = match.group(1).lower(), match.group(2)
            seen.setdefault(name.lower(), []).append((path, number, name))
            stack.append((kind, name, number))
        elif end:
            kind = end.group(1).lower()
            if not stack:
                errors.append(f'{path.relative_to(root)}:{number}: unmatched End {kind}')
            else:
                started, name, start_line = stack.pop()
                if started != kind:
                    errors.append(
                        f'{path.relative_to(root)}:{number}: End {kind} mismatches '
                        f'{started} {name} at {start_line}'
                    )

    for kind, name, number in stack:
        errors.append(f'{path.relative_to(root)}:{number}: missing End {kind} for {name}')

    # Check GoTo targets within each procedure. This catches misspelled labels
    # that a simple count of Sub/Function endings cannot detect.
    procedure_start = None
    procedure_lines = []
    procedure_name = ''

    for number, line in enumerate(lines, 1):
        if line.lstrip().startswith("'"):
            continue

        start_match = re.match(
            r'^\s*(?:(?:Public|Private|Friend)\s+)?(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\b',
            line, re.I
        )
        if start_match:
            procedure_start = number
            procedure_name = start_match.group(2)
            procedure_lines = []
            continue

        if procedure_start is None:
            continue

        if re.match(r'^\s*End\s+(?:Sub|Function)\s*$', line, re.I):
            labels = {
                label_match.group(1).lower()
                for item in procedure_lines
                if (label_match := re.match(
                    r"^\s*([A-Za-z_][A-Za-z0-9_]*)\s*:\s*(?:'.*)?$", item
                ))
            }
            for offset, item in enumerate(procedure_lines, procedure_start + 1):
                goto_match = re.search(r'\bGoTo\s+([A-Za-z_][A-Za-z0-9_]*)', item, re.I)
                if goto_match and goto_match.group(1).lower() not in labels:
                    errors.append(
                        f'{path.relative_to(root)}:{offset}: GoTo target '
                        f'"{goto_match.group(1)}" not found in procedure {procedure_name}'
                    )

            procedure_start = None
            procedure_lines = []
            procedure_name = ''
            continue

        procedure_lines.append(line)

for name, entries in seen.items():
    if len(entries) > 1:
        errors.append(
            'Duplicate procedure/function name ' + entries[0][2] + ': ' +
            ', '.join(f'{p.relative_to(root)}:{n}' for p, n, _ in entries)
        )

# Every macro wired to a visible panel button must resolve to a Public Sub.
public_procedures = set()
for path in modules:
    source = path.read_text(encoding='utf-8-sig')
    public_procedures.update(
        name.lower()
        for name in re.findall(r'^\s*Public\s+Sub\s+([A-Za-z_][A-Za-z0-9_]*)\b', source, re.I | re.M)
    )

panel_module = root / 'src' / 'vba' / 'modRegisterPanel.bas'
if panel_module.exists():
    panel_source = panel_module.read_text(encoding='utf-8-sig')
    macro_array = re.search(r'\bmacroNames\s*=\s*Array\((.*?)\)', panel_source, re.I | re.S)
    if not macro_array:
        errors.append('modRegisterPanel.bas: macroNames array not found')
    else:
        for macro_name in re.findall(r'"([^"]+)"', macro_array.group(1)):
            if macro_name.lower() not in public_procedures:
                errors.append('modRegisterPanel.bas: button target is not a Public Sub: ' + macro_name)

# Object detection is limited to the selected folder tree, then the filename.
# A matching parent folder outside the selected root must not override the object.
list_module = root / 'src' / 'vba' / 'modApplicationsList.bas'
if list_module.exists():
    source = list_module.read_text(encoding='utf-8-sig')
    object_match = re.search(
        r'Private Function GetObjectName\b(.*?)^End Function\s*$',
        source, re.I | re.M | re.S
    )
    if not object_match:
        errors.append('modApplicationsList.bas: GetObjectName function not found')
    else:
        object_code = object_match.group(1)
        if 'selectedRootPath' not in object_code:
            errors.append('modApplicationsList.bas: GetObjectName must receive the selected root path')
        if 'If StrComp(folderPath, selectedRootPath, vbTextCompare) = 0 Then Exit Do' not in object_code:
            errors.append('modApplicationsList.bas: object detection must stop at the selected root')
        if 'GetObjectName(currentPath, CStr(fileItem.Name), rootPath)' not in source:
            errors.append('modApplicationsList.bas: GetObjectName must receive the selected root for each file')
        boundary_match = re.search(
            r'Private Function FindObjectInText\b(.*?)^End Function\s*$',
            source, re.I | re.M | re.S
        )
        if not boundary_match or 'Ёё' not in boundary_match.group(1):
            errors.append('modApplicationsList.bas: object-name boundaries must treat Ё/ё as Cyrillic letters')

# Keep folder scans responsive so Ctrl+Break can be handled during large directories.
if list_module.exists():
    scan_source = list_module.read_text(encoding='utf-8-sig')
    file_loop = re.search(r'For Each fileItem In folder\.Files(.*?)Next fileItem', scan_source, re.I | re.S)
    if not file_loop or 'DoEvents' not in file_loop.group(1) or 'Application.StatusBar' not in file_loop.group(1):
        errors.append('modApplicationsList.bas: file scan must yield to Excel and update progress inside the file loop')

# Guard SharePoint path normalization: keep ordinary Windows paths unchanged,
# convert HTTP(S) roots to WebDAV UNC, and use VBA's single-backslash separator.
if list_module.exists():
    sharepoint_source = list_module.read_text(encoding='utf-8-sig')
    sharepoint_guards = (
        'roots.Add NormalizeSelectedRootPath(CStr(rootItem))',
        'Private Function NormalizeSelectedRootPath(ByVal selectedPath As String) As String',
        'urlPath = Replace(urlPath, "/", "\\")',
        'NormalizeSelectedRootPath = "\\\\" & webDavServer & "\\DavWWWRoot"',
        'NormalizeSelectedRootPath = NormalizeSelectedRootPath & "\\" & urlPath',
        'If Len(GetFolderName) = 0 Then',
    )
    for fragment in sharepoint_guards:
        if fragment not in sharepoint_source:
            errors.append('modApplicationsList.bas: SharePoint path handling missing or malformed: ' + fragment)

# File renaming must treat a broken register hyperlink update as a batch failure.
rename_module = root / "src" / "vba" / "modRenameFiles.bas"
if rename_module.exists():
    rename_source = rename_module.read_text(encoding="utf-8-sig")
    if "Не удалось обновить гиперссылку в строке" not in rename_source or "Err.Raise vbObjectError + 2401" not in rename_source:
        errors.append("modRenameFiles.bas: hyperlink-update failure must trigger rollback")

# Rollback must report a failed hyperlink repair even when the filesystem rename itself succeeded.
if rename_module.exists():
    recovery_source = rename_module.read_text(encoding='utf-8-sig')
    for fragment in (
        'Private Function RestoreRegisterHyperlink(',
        'Файл возвращён; ошибка восстановления ссылки:',
        'Файл возвращён, но ссылка в строке ',
    ):
        if fragment not in recovery_source:
            errors.append('modRenameFiles.bas: rollback hyperlink-repair reporting missing: ' + fragment)

# Duplicate application numbers are compared after normalizing leading zeros, while retaining suffixes.
if list_module.exists():
    duplicate_source = list_module.read_text(encoding='utf-8-sig')
    duplicate_fragments = (
        'NormalizeApplicationNumberForDuplicateCheck(appNumber)',
        'NormalizeApplicationNumberForDuplicateCheck = value',
        '01 = 1, 001A = 1A',
    )
    for fragment in duplicate_fragments:
        if fragment not in duplicate_source:
            errors.append(f'modApplicationsList.bas: duplicate-number normalization missing: {fragment}')

# Guard filename parsing against common grammatical forms of "Приложение".
if list_module.exists():
    parser_source = list_module.read_text(encoding='utf-8-sig')
    parser_forms = (
        'приложение|приложения|приложению|приложением|приложении|приложений|прил',
        'matches(0).SubMatches(2)',
    )
    for fragment in parser_forms:
        if fragment not in parser_source:
            errors.append(f'modApplicationsList.bas: filename parser missing expected form/index: {fragment}')

# Duplicate highlighting must use the same normalized key as the counting pass;
# otherwise values such as 01 and 1 are counted together but not all highlighted.
if list_module.exists():
    duplicate_key = 'key = UCase$(objectName) & "|" & NormalizeApplicationNumberForDuplicateCheck(appNumber)'
    if list_module.read_text(encoding='utf-8-sig').count(duplicate_key) < 2:
        errors.append('modApplicationsList.bas: duplicate counting and highlighting must use the same normalized key')

# Special department marker applies only to the two explicitly named base filenames.
if list_module.exists():
    special_source = list_module.read_text(encoding='utf-8-sig')
    special_match = re.search(
        r'Private Function IsSpecialDepartmentFile\b(.*?)^End Function\s*$',
        special_source, re.I | re.M | re.S
    )
    if not special_match:
        errors.append('modApplicationsList.bas: IsSpecialDepartmentFile function not found')
    else:
        special_code = special_match.group(1)
        if 'Replace(baseName, "ё", "е")' not in special_code:
            errors.append('modApplicationsList.bas: special-file matching must preserve underscores for exact base-name checks')
        if 'Replace(baseName, "_", " ")' in special_code:
            errors.append('modApplicationsList.bas: special-file matching must not normalize underscores into spaces')
        for exact_name in ('отчет_2 кв 2026', 'реестр рисков'):
            if exact_name not in special_code:
                errors.append('modApplicationsList.bas: exact special-file name missing: ' + exact_name)

readme = root / 'README.md'
if readme.exists():
    readme_text = readme.read_text(encoding='utf-8-sig').lower()
    contradictory_phrases = (
        'внешние родительские папки выше корня не используются',
        'код из внешней папки выше выбранного корня не используется',
        'only inside the selected root',
        'включая родительские папки выше выбранного корня',
    )
    for phrase in contradictory_phrases:
        if phrase in readme_text:
            errors.append(f'README.md: outdated object-detection rule remains: {phrase}')

# Keep all user-facing documentation and manual test cases aligned with the
# selected-root boundary; otherwise an obsolete test can reintroduce the old rule.
for doc_path in (root / 'docs' / 'requirements.md', root / 'tests' / 'test-checklist.md'):
    if doc_path.exists():
        doc_text = doc_path.read_text(encoding='utf-8-sig').lower()
        obsolete_object_rules = (
            'включая папки выше выбранной исходной папки',
            'включая родителей выше выбранного корня',
            'объект определяется по этой папке',
        )
        for phrase in obsolete_object_rules:
            if phrase in doc_text:
                errors.append(f'{doc_path.relative_to(root)}: outdated object-detection rule remains: {phrase}')

# Guard mention-scanning behavior that is easy to regress:
# common grammatical forms must be recognized and each match must use its own start page.
analysis_module = root / 'src' / 'vba' / 'modReportAnalysis.bas'
if analysis_module.exists():
    analysis_source = analysis_module.read_text(encoding='utf-8-sig')
    required_fragments = (
        'приложение|приложения|приложению|приложением|приложении|приложений|прил',
        'mt.FirstIndex',
        'rng.Information(1)',
        'failed = failed + 1',
        'ws.Cells(r, 10).Value = "Нет ссылки на файл"',
        'oldPageValues = ws.Range("J3:J" & lastRow).Formula',
        'If pageSnapshotTaken Then ws.Range("J3:J" & lastRow).Formula = oldPageValues',
        'Application.EnableCancelKey = oldCancelKey',
    )
    for fragment in required_fragments:
        if fragment not in analysis_source:
            errors.append(f'modReportAnalysis.bas: required mention-scanning logic missing: {fragment}')

# Page-count helper must propagate Ctrl+Break instead of swallowing it under local Resume Next.
if analysis_module.exists():
    workbook_pages_match = re.search(
        r'Private Function CountWorkbookPrintPages\b(.*?)^End Function\s*$',
        analysis_source, re.I | re.M | re.S
    )
    if not workbook_pages_match:
        errors.append('modReportAnalysis.bas: CountWorkbookPrintPages function not found')
    else:
        workbook_pages_code = workbook_pages_match.group(1)
        if 'If pageBreakError = 18 Then' not in workbook_pages_code or 'Err.Raise pageBreakError' not in workbook_pages_code:
            errors.append('modReportAnalysis.bas: workbook page-count helper must propagate Ctrl+Break')
        if 'DoEvents' not in workbook_pages_code:
            errors.append('modReportAnalysis.bas: workbook page-count helper must yield to Excel')

rename_module = root / 'src' / 'vba' / 'modRenameFiles.bas'
if rename_module.exists():
    rename_source = rename_module.read_text(encoding='utf-8-sig')
    rollback_guards = (
        'For rollbackIndex = renamedItems.Count To 1 Step -1',
        'Set rollbackItem = renamedItems(rollbackIndex)',
        'For recoveryIndex = reversedItems.Count To 1 Step -1',
        'Set item = reversedItems(recoveryIndex)',
        'Application.EnableCancelKey = xlDisabled',
    )
    for fragment in rollback_guards:
        if fragment not in rename_source:
            errors.append(f'modRenameFiles.bas: reverse-order recovery logic missing: {fragment}')

    # A no-op first candidate must still reserve its source path, otherwise a
    # duplicate register row can rename the same physical file a second time.
    source_claim = rename_source.find('seenSource(oldPath) = r')
    no_op_check = rename_source.find('If StrComp(oldPath, newPath, vbTextCompare) = 0 Then')
    if source_claim < 0 or no_op_check < 0 or source_claim > no_op_check:
        errors.append('modRenameFiles.bas: source path must be claimed before the no-op rename check')
    if 'имя файла уже соответствует полям перечня; переименование не требуется.' not in rename_source:
        errors.append('modRenameFiles.bas: no-op rename candidate should be reported as a skipped row')

# Guard table-caption recognition: a SEQ field must be in the nearest non-empty
# paragraph and the visible paragraph must start with the expected caption label.
table_module = root / 'src' / 'vba' / 'modTableNumbering.bas'
if table_module.exists():
    table_source = table_module.read_text(encoding='utf-8-sig')
    caption_match = re.search(
        r'Private Function HasTableSequenceField\b(.*?)^End Function\s*$',
        table_source, re.I | re.M | re.S
    )
    if not caption_match:
        errors.append('modTableNumbering.bas: HasTableSequenceField function not found')
    else:
        caption_code = caption_match.group(1)
        for fragment in (
            'Left$(trimmedText, Len("Таблица"))',
            'StrComp(',
            'For Each fld In para.Fields',
            'If Len(trimmedText) > 0 Then',
        ):
            if fragment not in caption_code:
                errors.append('modTableNumbering.bas: caption recognition guard missing: ' + fragment)

# The final Word table audit must be populated on a staging sheet before replacing the preview.
if table_module.exists():
    final_report_guards = (
        'Set finalWs = GetRegisterWorkbook().Worksheets.Add',
        'FormatTableReport finalWs, r - 1',
        'ws.Name = finalOldName',
        'finalWs.Name = TABLE_REPORT_SHEET',
        'FinalReportCommitFailed:',
        'If Not finalWs Is Nothing Then',
        'Deleting the old preview is cleanup, not part of the Word edit transaction.',
        'If Err.Number = 0 Then',
    )
    for fragment in final_report_guards:
        if fragment not in table_source:
            errors.append('modTableNumbering.bas: staged final-report replacement guard missing: ' + fragment)

# Guard against creating duplicate captions when a visible manual caption already exists.
if table_module.exists():
    for fragment in (
        'Private Function IsTableCaptionText(',
        'manualCaptionCount = manualCaptionCount + 1',
        'Not IsTableCaptionText(captionText)',
        'para.Information(WD_WITHIN_TABLE) = False',
    ):
        if fragment not in table_source:
            errors.append('modTableNumbering.bas: manual-caption protection missing: ' + fragment)

# Guard the current single-workbook startup design. The startup macro must
# validate the exact data sheet, assign the selected workbook context, and
# report a handled error without silently swallowing it. Do not assert resultWb
# cleanup: this version intentionally does not create a separate result workbook.
context_module = root / 'src' / 'vba' / 'modWorkbookContext.bas'
if context_module.exists():
    context_source = context_module.read_text(encoding='utf-8-sig')
    startup_guards = (
        'Set sourceWb = ThisWorkbook',
        'Set sourceData = sourceWb.Worksheets("data")',
        'Set gRegisterWorkbook = sourceWb',
        'CreateMultiFolderApplicationsListCorrected',
        'failureNumber = Err.Number',
        'failureDescription = Err.Description',
    )
    for fragment in startup_guards:
        if fragment not in context_source:
            errors.append(f'modWorkbookContext.bas: single-workbook startup guard missing: {fragment}')

# Required launch points must be externally callable macros, not merely private
# procedures or functions with the same names.
required_public_subs = {
    name.lower()
    for name in re.findall(
        r'^\s*Public\s+Sub\s+([A-Za-z_][A-Za-z0-9_]*)\b',
        '\n'.join(path.read_text(encoding='utf-8-sig') for path in modules),
        re.I | re.M
    )
}
missing_public = sorted(name for name in required if name.lower() not in required_public_subs)
if missing_public:
    errors.append('Required entry points must be Public Sub: ' + ', '.join(missing_public))

missing = sorted(x for x in required if x.lower() not in seen)
if missing:
    errors.append('Missing required entry points: ' + ', '.join(missing))


# Guard the installer contract: every VBA source module must be listed exactly
# once in the installer, and the CMD launcher must stay ASCII-only because
# Windows CMD can misread UTF-8 Cyrillic text depending on the active code page.
installer_path = root / 'tools' / 'setup_excel.ps1'
launcher_path = root / 'tools' / 'Установить_книгу_Excel.cmd'
if installer_path.exists():
    installer_source = installer_path.read_text(encoding='utf-8-sig')
    listed_modules = re.findall(r'"(mod[A-Za-z0-9]+\.bas)"', installer_source)
    expected_modules = sorted(path.name for path in modules)
    if sorted(listed_modules) != expected_modules:
        errors.append(
            'setup_excel.ps1: module list does not match src/vba files; '
            f'listed={sorted(listed_modules)}, expected={expected_modules}'
        )
if launcher_path.exists():
    launcher_bytes = launcher_path.read_bytes()
    if any(byte > 127 for byte in launcher_bytes):
        errors.append('Установить_книгу_Excel.cmd: launcher must remain ASCII-only')

# Keep installer recovery behavior covered by static regression checks.
if installer_path.exists():
    for fragment in (
        '$backupPath = Join-Path',
        '$convertedOutputCreated = $true',
        'if (-not $success -and $convertedOutputCreated -and $finalPath -ne $workbookPath)',
        'Remove-Item -LiteralPath $finalPath -Force -ErrorAction Stop',
    ):
        if fragment not in installer_source:
            errors.append('setup_excel.ps1: installer safety guard missing: ' + fragment)



# Mixed leading indentation can indicate nested list levels. The bullet fixer
# must stop before editing such a document instead of flattening those levels.
bullet_module = root / 'src' / 'vba' / 'modBulletListAudit.bas'
if bullet_module.exists():
    bullet_source = bullet_module.read_text(encoding='utf-8-sig')
    for fragment in (
        'GetBulletIndentLevel(para, txt)',
        'If currentIndent <> runIndent Then unsafeGroup = True',
        'If unsafeGroup Then unsafeGroupsFound = unsafeGroupsFound + 1',
        'If unsafeGroupsFound > 0 Then',
        'Автоматическое исправление не выполнено, чтобы не потерять структуру вложенности.',
        'Private Function GetBulletIndentLevel(ByVal para As Object, ByVal paragraphText As String) As Double',
        'GetBulletIndentLevel = CDbl(para.Format.LeftIndent)',
    ):
        if fragment not in bullet_source:
            errors.append('modBulletListAudit.bas: nested-list safety guard missing: ' + fragment)


# Guard Word-document edits: a same-folder backup must be made before changes,
# and a failed/cancelled edit must attempt restoration after closing Word.
for module_name, function_prefix, log_prefix in (
    ('modTableNumbering.bas', 'CreateWordBackupPath', 'LogTableBackup'),
    ('modBulletListAudit.bas', 'CreateBulletBackupPath', 'LogBulletBackup'),
):
    module_path = root / 'src' / 'vba' / module_name
    if module_path.exists():
        module_source = module_path.read_text(encoding='utf-8-sig')
        recovery_fragments = (
            'backupPath = ' + function_prefix + '(reportPath)',
            'FileCopy reportPath, backupPath',
            'changesStarted = True',
            'If changesStarted And Len(backupPath) > 0 Then',
            'If Not wdDoc Is Nothing Then wdDoc.Close False',
            'If Not wdApp Is Nothing Then wdApp.Quit',
            'fsoRestore.CopyFile backupPath, reportPath, True',
            'Автовосстановление не удалось. Резервная копия сохранена:',
            'Application.EnableCancelKey = oldCancelKey',
        )
        for fragment in recovery_fragments:
            if fragment not in module_source:
                errors.append(f'{module_name}: Word backup/recovery guard missing: {fragment}')
        backup_pos = module_source.find('FileCopy reportPath, backupPath')
        change_pos = module_source.find('changesStarted = True', backup_pos)
        first_edit_pos = module_source.find('EnsureTableCaptionLabel wdDoc') if module_name == 'modTableNumbering.bas' else module_source.find('wordRange.Delete')
        if backup_pos < 0 or change_pos < backup_pos or first_edit_pos < 0 or backup_pos > first_edit_pos:
            # Bullet list edit is inside a later loop; backup must precede its first deletion.
            errors.append(f'{module_name}: backup must be created before document edits')

print(f'Modules scanned: {len(modules)}; procedures/functions: {sum(len(x) for x in seen.values())}')
if errors:
    print('\n'.join('ERROR: ' + x for x in errors))
    sys.exit(1)

print('PASS: no duplicate procedure names, unmatched procedure endings, or unresolved GoTo labels; required entry points found.')
print('NOTE: this is static validation, not VBA compilation or Excel/Word runtime testing.')
