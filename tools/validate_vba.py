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

missing = sorted(x for x in required if x.lower() not in seen)
if missing:
    errors.append('Missing required entry points: ' + ', '.join(missing))

print(f'Modules scanned: {len(modules)}; procedures/functions: {sum(len(x) for x in seen.values())}')
if errors:
    print('\n'.join('ERROR: ' + x for x in errors))
    sys.exit(1)

print('PASS: no duplicate procedure names, unmatched procedure endings, or unresolved GoTo labels; required entry points found.')
print('NOTE: this is static validation, not VBA compilation or Excel/Word runtime testing.')
