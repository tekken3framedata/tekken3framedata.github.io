#!/usr/bin/env python3
"""
Step 6b: Process bracket notations in Command column.

Reads sources/<character>_step6.xlsx. Extracts bracket notations and creates
new columns based on rules:
  - [~5]: Taggable=TRUE, strip from Command. Strip [Tag] from Move Name.

Output: sources/<character>_step6b.xlsx

Usage:
    python3 step_6b_bracket_notations.py
"""

import glob
import os
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font


TAG_PATTERN = re.compile(r'\s*\[~5\]')
NAME_TAG_PATTERN = re.compile(r'\s*\[Tag\]')

STANCE_RECOVERY = {
    'lee': [
        {
            'uuid': '96cf7074-dd8c-4b40-a94e-b20afb923e00',
            'strip_cmd': re.compile(r'\s*\[~3\]'),
            'strip_name': re.compile(r'\s*\[Hit Man Stance\]'),
            'to_stance': '(~3)HMS',
        },
        {
            'uuid': '5b2fca48-dee4-430a-9737-31438e2ffb94',
            'strip_cmd': re.compile(r'\s*\[~3\]'),
            'strip_name': re.compile(r'\s*\[Hit Man Stance\]'),
            'to_stance': '(~3)HMS',
        },
        {
            'uuid': '80993965-6214-47a8-a5aa-f15fddcf516d',
            'strip_cmd': re.compile(r'\s*\[~3\]'),
            'strip_name': re.compile(r'\s*\[Hit Man Stance\]'),
            'to_stance': '(~3)HMS',
        },
    ],
}


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


def process_sections(sections, char):
    tag_count = 0
    stance_count = 0

    stance_rules = {}
    for rule in STANCE_RECOVERY.get(char, []):
        stance_rules[rule['uuid']] = rule

    for _, rows in sections:
        for row in rows:
            cmd = str(row.get('Command', ''))

            if TAG_PATTERN.search(cmd):
                tag_count += 1
                row['Taggable'] = 'TRUE'
                row['Command'] = TAG_PATTERN.sub('', cmd).strip()
                alt_cmds = str(row.get('Alt Commands', ''))
                if alt_cmds:
                    row['Alt Commands'] = TAG_PATTERN.sub('', alt_cmds).strip()
                move_name = str(row.get('Move Name', ''))
                row['Move Name'] = NAME_TAG_PATTERN.sub('', move_name).strip()
            else:
                row['Taggable'] = ''

            uuid = str(row.get('UUID', ''))
            if uuid in stance_rules:
                rule = stance_rules[uuid]
                stance_count += 1
                row['Command'] = rule['strip_cmd'].sub('', str(row.get('Command', ''))).strip()
                row['Move Name'] = rule['strip_name'].sub('', str(row.get('Move Name', ''))).strip()
                existing = str(row.get('To Stance', ''))
                if existing:
                    row['To Stance'] = existing + ';' + rule['to_stance']
                else:
                    row['To Stance'] = rule['to_stance']

    return tag_count, stance_count


OUTPUT_COLUMNS = [
    'Command', 'Alt Commands', 'Move Name', 'To Stance',
    'Speed', 'Block Adv', 'Hit Adv', 'Counter Hit Adv',
    'Damage', 'Damage Sum', 'Hit Range',
    'Throw Type', 'Throw Escape', 'Properties', 'Notes',
    'Taggable',
    'UUID', 'ML UUID', 'Parent UUID', 'Unmatched',
]


HIDDEN_COLUMNS = set()


def autofit_columns(ws, header_row_numbers):
    for col in ws.columns:
        col_letter = col[0].column_letter
        header_value = None
        for cell in col:
            if cell.row in header_row_numbers and cell.value:
                header_value = cell.value
                break
        if header_value in HIDDEN_COLUMNS:
            ws.column_dimensions[col_letter].width = 0
            ws.column_dimensions[col_letter].hidden = True
            continue
        max_len = 0
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_len + 2


def find_characters():
    files = glob.glob('sources/*_step6.xlsx')
    return sorted(
        os.path.basename(f).replace('_step6.xlsx', '')
        for f in files
        if not os.path.basename(f).startswith('~$')
    )


def main():
    characters = find_characters()
    if not characters:
        print("No step6 files found in sources/")
        return
    print(f"Found characters: {characters}")

    for char in characters:
        input_path = f"sources/{char}_step6.xlsx"
        output_path = f"sources/{char}_step6b.xlsx"

        print(f"\n=== {char.capitalize()} ===")
        wb_in = load_workbook(input_path)
        ws_in = wb_in.active

        sections = parse_sections(ws_in)
        tag_count, stance_count = process_sections(sections, char)

        print(f"  {tag_count} rows tagged, {stance_count} stance recoveries")

        wb_out = Workbook()
        ws = wb_out.active
        ws.title = "Merged"
        row_num = 1
        header_row_numbers = set()

        for section_name, rows in sections:
            cell = ws.cell(row=row_num, column=1, value=section_name)
            cell.font = Font(bold=True)
            row_num += 1

            for col, h in enumerate(OUTPUT_COLUMNS, 1):
                cell = ws.cell(row=row_num, column=col, value=h)
                cell.font = Font(bold=True)
            header_row_numbers.add(row_num)
            row_num += 1

            for row_dict in rows:
                for col, col_name in enumerate(OUTPUT_COLUMNS, 1):
                    val = row_dict.get(col_name, '')
                    if val:
                        ws.cell(row=row_num, column=col, value=val)
                row_num += 1

            row_num += 1

        autofit_columns(ws, header_row_numbers)
        wb_out.save(output_path)
        print(f"  Written to {output_path}")

    print(f"\nDone. Processed {len(characters)} characters.")


if __name__ == '__main__':
    main()
