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

# Guard a critical project requirement: object detection must walk the full
# folder chain, including parents above the selected root, before checking the filename.
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
        if 'fso.GetParentFolderName(folderPath)' not in object_code:
            errors.append('modApplicationsList.bas: GetObjectName must walk parent folders')
        if re.search(r'StrComp\(\s*folderPath\s*,\s*rootPath', object_code, re.I):
            errors.append('modApplicationsList.bas: GetObjectName must not stop at the selected root')
        # Boundary characters are defined in FindObjectInText, which owns
        # the regular expression; GetObjectName only walks folders and calls it.
        boundary_match = re.search(
            r'Private Function FindObjectInText\b(.*?)^End Function\s*$',
            source, re.I | re.M | re.S
        )
        if not boundary_match or 'Ёё' not in boundary_match.group(1):
            errors.append('modApplicationsList.bas: object-name boundaries must treat Ё/ё as Cyrillic letters')

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
    )
    for phrase in contradictory_phrases:
        if phrase in readme_text:
            errors.append(f'README.md: outdated object-detection rule remains: {phrase}')

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
    )
    for fragment in required_fragments:
        if fragment not in analysis_source:
            errors.append(f'modReportAnalysis.bas: required mention-scanning logic missing: {fragment}')

# Guard transactional recovery for batch file renaming: both failure paths must unwind in reverse order.
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

print(f'Modules scanned: {len(modules)}; procedures/functions: {sum(len(x) for x in seen.values())}')
if errors:
    print('\n'.join('ERROR: ' + x for x in errors))
    sys.exit(1)

print('PASS: no duplicate procedure names, unmatched procedure endings, or unresolved GoTo labels; required entry points found.')
print('NOTE: this is static validation, not VBA compilation or Excel/Word runtime testing.')
