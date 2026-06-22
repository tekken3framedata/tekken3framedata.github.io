#!/usr/bin/env python3
"""
List all unique bracket notations [xxx] found in Command columns across all characters.

Reads sources/<character>_step6.xlsx files (the input for step 6b/7).
Reports each unique bracket notation with example characters and commands.

Usage:
    python3 tools/list_bracket_notations.py
"""

import glob
import os
import re
from collections import defaultdict

from openpyxl import load_workbook


BRACKET_PATTERN = re.compile(r'\[[^\]]+\]')


def parse_sections(ws):
    sections = []
    row = 1
    while row <= ws.max_row:
        cell = ws.cell(row=row, column=1)
        if cell.value and cell.font.bold:
            next_cell = ws.cell(row=row + 1, column=1) if row + 1 <= ws.max_row else None
            if next_cell and next_cell.value == 'Command' and next_cell.font.bold:
                heading = cell.value
                row += 1

                header_cells = []
                for col in range(1, ws.max_column + 1):
                    val = ws.cell(row=row, column=col).value
                    if val:
                        header_cells.append(val)
                row += 1

                data_rows = []
                while row <= ws.max_row:
                    first_cell = ws.cell(row=row, column=1)
                    if first_cell.value is None and all(
                        ws.cell(row=row, column=c).value is None
                        for c in range(1, len(header_cells) + 1)
                    ):
                        row += 1
                        break
                    if first_cell.font.bold:
                        break

                    row_dict = {}
                    for col_idx, col_name in enumerate(header_cells):
                        val = ws.cell(row=row, column=col_idx + 1).value
                        row_dict[col_name] = val or ''
                    data_rows.append(row_dict)
                    row += 1

                sections.append((heading, data_rows))
            else:
                row += 1
        else:
            row += 1

    return sections


def find_characters():
    files = glob.glob('sources/*_step6.xlsx')
    return sorted(
        os.path.basename(f).replace('_step6.xlsx', '')
        for f in files
        if not os.path.basename(f).startswith('~$')
    )


def main():
    import sys
    characters = find_characters()
    if not characters:
        print("No step6 files found in sources/")
        return

    if len(sys.argv) > 1:
        characters = [c for c in characters if c in sys.argv[1:]]

    # notation -> list of (character, full_command)
    notations = defaultdict(list)

    for char in characters:
        input_path = f"sources/{char}_step6.xlsx"
        wb = load_workbook(input_path, read_only=True)
        ws = wb.active
        sections = parse_sections(ws)
        wb.close()

        for _section_name, rows in sections:
            for row in rows:
                for col_name, val in row.items():
                    val_str = str(val or '')
                    matches = BRACKET_PATTERN.findall(val_str)
                    for match in matches:
                        notations[match].append((char, col_name, val_str))

    print(f"Found {len(notations)} unique bracket notations across {len(characters)} characters\n")
    print("=" * 80)

    for notation in sorted(notations.keys()):
        examples = notations[notation]
        chars_with = sorted(set(ex[0] for ex in examples))
        cols_with = sorted(set(ex[1] for ex in examples))
        print(f"\n{notation}  ({len(examples)} occurrences, {len(chars_with)} characters)")
        print(f"  Columns: {', '.join(cols_with)}")
        print(f"  Characters: {', '.join(chars_with)}")
        print(f"  Examples:")
        seen = set()
        for char, col, val in examples:
            if len(seen) >= 5:
                print(f"    ...")
                break
            key = (char, col, val)
            if key not in seen:
                seen.add(key)
                print(f"    {char} [{col}]: {val}")


if __name__ == '__main__':
    main()
