from pathlib import Path
import re
import sys

root = Path(__file__).resolve().parents[1]
modules = sorted((root / 'src' / 'vba').glob('*.bas'))
errors = []
seen = {}
required = {'StartRegisterAutomation', 'CreateMultiFolderApplicationsListCorrected', 'ScanWordReportMentions', 'CountPagesForRegister', 'BuildDataFull', 'RenameFilesFromRegister', 'RollbackLastRenameBatch', 'CheckAndNumberWordTables', 'CheckAndFixWordBulletLists'}
for path in modules:
    lines = path.read_text(encoding='utf-8-sig').splitlines()
    stack = []
    for number, line in enumerate(lines, 1):
        if not line.lstrip().startswith("'") and 'FreezePanes' in line:
            errors.append(f'{path.relative_to(root)}:{number}: unexpected FreezePanes reference')
        match = re.match(r'^\s*(?:(?:Public|Private|Friend)\s+)?(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\b', line, re.I)
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
                    errors.append(f'{path.relative_to(root)}:{number}: End {kind} mismatches {started} {name} at {start_line}')
    for kind, name, number in stack:
        errors.append(f'{path.relative_to(root)}:{number}: missing End {kind} for {name}')

    # Check GoTo targets inside each procedure. This catches misspelled labels
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
        if re.match(r'^\s*End\s+(?:Sub|Function)\s*
    if len(entries) > 1:
        errors.append('Duplicate procedure/function name ' + entries[0][2] + ': ' + ', '.join(f'{p.relative_to(root)}:{n}' for p,n,_ in entries))
missing = sorted(x for x in required if x.lower() not in seen)
if missing:
    errors.append('Missing required entry points: ' + ', '.join(missing))
print(f'Modules scanned: {len(modules)}; procedures/functions: {sum(len(x) for x in seen.values())}')
if errors:
    print('\n'.join('ERROR: ' + x for x in errors))
    sys.exit(1)
print('PASS: no duplicate procedure names or unmatched procedure endings; required entry points found.')
print('NOTE: this is static validation, not VBA compilation or Excel/Word runtime testing.')
, line, re.I):
            labels = {
                m.group(1).lower()
                for item in procedure_lines
                if (m := re.match(r'^\s*([A-Za-z_][A-Za-z0-9_]*)\s*:\s*(?:\'.*)?
    if len(entries) > 1:
        errors.append('Duplicate procedure/function name ' + entries[0][2] + ': ' + ', '.join(f'{p.relative_to(root)}:{n}' for p,n,_ in entries))
missing = sorted(x for x in required if x.lower() not in seen)
if missing:
    errors.append('Missing required entry points: ' + ', '.join(missing))
print(f'Modules scanned: {len(modules)}; procedures/functions: {sum(len(x) for x in seen.values())}')
if errors:
    print('\n'.join('ERROR: ' + x for x in errors))
    sys.exit(1)
print('PASS: no duplicate procedure names or unmatched procedure endings; required entry points found.')
print('NOTE: this is static validation, not VBA compilation or Excel/Word runtime testing.')
, item))
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
        errors.append('Duplicate procedure/function name ' + entries[0][2] + ': ' + ', '.join(f'{p.relative_to(root)}:{n}' for p,n,_ in entries))
missing = sorted(x for x in required if x.lower() not in seen)
if missing:
    errors.append('Missing required entry points: ' + ', '.join(missing))
print(f'Modules scanned: {len(modules)}; procedures/functions: {sum(len(x) for x in seen.values())}')
if errors:
    print('\n'.join('ERROR: ' + x for x in errors))
    sys.exit(1)
print('PASS: no duplicate procedure names or unmatched procedure endings; required entry points found.')
print('NOTE: this is static validation, not VBA compilation or Excel/Word runtime testing.')
