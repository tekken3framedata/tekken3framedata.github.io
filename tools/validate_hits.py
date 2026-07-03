#!/usr/bin/env python3
"""
Validate that frame data column counts match the expected number of hits
derived from Command Full (presses) and Hits Per Move.

Reads sources/<character>_step10.xlsx and reports rows where the number of
space-separated values in a Full column doesn't match the total hit count.

Usage:
    python3 tools/validate_hits.py
"""

import glob
import os
import sys

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from step_10_hits_per_move import HITS_PER_MOVE


def count_space_separated(field):
    if field is None or str(field).strip() == '':
        return 0
    return len(str(field).split())


def expected_hits(hits_per_move):
    """Sum of hits from Hits Per Move column."""
    if not hits_per_move or str(hits_per_move).strip() == '':
        return 0
    return sum(int(x) for x in str(hits_per_move).split() if x.isdigit())


def validate_file(path):
    char = os.path.basename(path).replace('_step10.xlsx', '')
    wb = load_workbook(path)
    ws = wb.active
    headers = [cell.value for cell in ws[2]]

    cmd_full_idx = headers.index('Command Full')
    hpm_idx = headers.index('Hits Per Move')
    move_name_idx = headers.index('Move Name Full')
    from_stance_idx = headers.index('From Stance')
    uuid_idx = headers.index('UUID')

    per_hit_cols = {
        'Block Adv Full': headers.index('Block Adv Full'),
        'Hit Adv Full': headers.index('Hit Adv Full'),
        'Counter Hit Adv Full': headers.index('Counter Hit Adv Full'),
    }

    essential_cols = {
        'Speed': headers.index('Speed'),
        'Damage': headers.index('Damage'),
        'Block Adv': headers.index('Block Adv'),
    }

    mismatches = []

    for row_idx, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
        cmd_full = row[cmd_full_idx]
        if not cmd_full or str(cmd_full).strip() == '' or str(cmd_full) == 'Command Full':
            continue

        hpm = row[hpm_idx]
        total_hits = expected_hits(hpm)

        issues = []

        missing = [name for name, idx in essential_cols.items()
                   if not row[idx] or str(row[idx]).strip() in ('', '-')]
        uuid = row[uuid_idx] or ''
        has_override = uuid in HITS_PER_MOVE
        if len(missing) == len(essential_cols) and not has_override:
            issues.append(f'missing data: {", ".join(missing)}')
        elif total_hits > 0:
            for col_name, col_idx in per_hit_cols.items():
                val = row[col_idx]
                if val is None or str(val).strip() == '':
                    continue
                actual = count_space_separated(val)
                if actual != total_hits:
                    issues.append(f'{col_name}: {actual} values')

        if issues:
            fs = row[from_stance_idx] or ''
            prefix = f'{fs} ' if fs else ''
            move_name = row[move_name_idx] or ''
            uuid = row[uuid_idx] or ''
            mismatches.append((row_idx, prefix, cmd_full, move_name, hpm, total_hits, issues, uuid))

    return char, mismatches


def main():
    characters = sys.argv[1:] if len(sys.argv) > 1 else None

    if characters:
        files = [f'sources/{c}_step10.xlsx' for c in characters]
    else:
        files = sorted(glob.glob('sources/*_step10.xlsx'))

    if not files:
        print("No step10 files found in sources/")
        sys.exit(1)

    total_issues = 0
    for path in files:
        if not os.path.exists(path):
            print(f"File not found: {path}")
            continue
        if os.path.basename(path).startswith('~$'):
            continue
        char, mismatches = validate_file(path)
        print(f'=== {char} ===')
        if mismatches:
            for row_idx, prefix, cmd_full, move_name, hpm, total_hits, issues, uuid in mismatches:
                print(f'  Row {row_idx}: "{move_name}"')
                print(f'    UUID: {uuid}')
                print(f'    Command: {prefix}{cmd_full}')
                if total_hits:
                    print(f'    Expected {total_hits} hits (Hits Per Move: {hpm})')
                for issue in issues:
                    print(f'    - {issue}')
            total_issues += len(mismatches)
        else:
            print('  OK')
        print()

    if total_issues:
        print(f'{total_issues} rows with mismatches.')
        sys.exit(1)
    else:
        print('All rows valid.')


if __name__ == '__main__':
    main()
