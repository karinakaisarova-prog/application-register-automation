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

# Guard table-caption recognition: a SEQ field must be in the nearest non-empty
# paragraph and the visible paragraph must start with the expected caption label.
table_module = root / 'src' / 'vba' / 'modTableNumbering.bas'
if table_module.exists():
    table_source = table_module.read_text(encoding='utf-8-sig')
    caption_match = re.search(
        r'Private Function HasTableSequenceField\\b(.*?)^End Function\\s*
context_module = root / 'src' / 'vba' / 'modWorkbookContext.bas'
if context_module.exists():
    context_source = context_module.read_text(encoding='utf-8-sig')
    startup_guards = (
        'failureNumber = Err.Number',
        'failureDescription = Err.Description',
        'If Not resultWb Is Nothing Then',
        'resultWb.Close SaveChanges:=False',
        'If resultWb Is gRegisterWorkbook Then Set gRegisterWorkbook = Nothing',
    )
    for fragment in startup_guards:
        if fragment not in context_source:
            errors.append(f'modWorkbookContext.bas: first-run cleanup guard missing: {fragment}')

missing = sorted(x for x in required if x.lower() not in seen)
if missing:
    errors.append('Missing required entry points: ' + ', '.join(missing))

print(f'Modules scanned: {len(modules)}; procedures/functions: {sum(len(x) for x in seen.values())}')
if errors:
    print('\n'.join('ERROR: ' + x for x in errors))
    sys.exit(1)

print('PASS: no duplicate procedure names, unmatched procedure endings, or unresolved GoTo labels; required entry points found.')
print('NOTE: this is static validation, not VBA compilation or Excel/Word runtime testing.')
,
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

# Guard first-run cleanup: a failed setup must not leave a partially initialized result workbook open.
context_module = root / 'src' / 'vba' / 'modWorkbookContext.bas'
if context_module.exists():
    context_source = context_module.read_text(encoding='utf-8-sig')
    startup_guards = (
        'failureNumber = Err.Number',
        'failureDescription = Err.Description',
        'If Not resultWb Is Nothing Then',
        'resultWb.Close SaveChanges:=False',
        'If resultWb Is gRegisterWorkbook Then Set gRegisterWorkbook = Nothing',
    )
    for fragment in startup_guards:
        if fragment not in context_source:
            errors.append(f'modWorkbookContext.bas: first-run cleanup guard missing: {fragment}')

missing = sorted(x for x in required if x.lower() not in seen)
if missing:
    errors.append('Missing required entry points: ' + ', '.join(missing))

print(f'Modules scanned: {len(modules)}; procedures/functions: {sum(len(x) for x in seen.values())}')
if errors:
    print('\n'.join('ERROR: ' + x for x in errors))
    sys.exit(1)

print('PASS: no duplicate procedure names, unmatched procedure endings, or unresolved GoTo labels; required entry points found.')
print('NOTE: this is static validation, not VBA compilation or Excel/Word runtime testing.')
