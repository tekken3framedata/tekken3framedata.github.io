#!/usr/bin/env python3
"""
Step 9: Apply manual corrections to the final data.

Reads sources/<character>_step8c.xlsx and applies hand-specified corrections:
  - Modify cell values in existing rows (matched by UUID)
  - Insert new rows (after a specified UUID)
  - Delete rows (matched by UUID)

New columns introduced by corrections are added to all sections.

Output: sources/<character>_step9.xlsx

Usage:
    python3 step_9_manual_corrections.py
"""

import glob
import os

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font


CORRECTIONS = {
    'jin': {
        'modify': [],
        'insert_after': [],
        'delete': [],
    },
    'julia': {
        'modify': [],
        'insert_after': [],
        'delete': [],
    },
    'ling': {
        'modify': [
            {
                'uuid': '085e1996-b16b-4a63-a3d0-ae8152d065ea',
                'set': {
                    'Properties': 'RC GB OC OCb',
                    'T3 Speed': '17 x',
                    'T3 Block Adv': 'x -6',
                    'T3 Hit Adv': 'x -15',
                    'T3 Counter Hit Adv': 'x -15',
                },
            },
            {
                'uuid': 'b929cc95-3142-4aba-8191-88bca5e84bf1',
                'set': {
                    'Speed Full': '15 x',
                    'Block Adv Full': 'x -6',
                    'Hit Adv Full': 'x +4',
                    'Counter Hit Adv Full': 'x +4',
                    'To Stance': '',
                    'Properties': 'GB OC OCb',
                    'T3 Speed': '17 x',
                    'T3 Block Adv': '-5~6',
                    'T3 Hit Adv': '+2',
                    'T3 Counter Hit Adv': '+2',
                },
            },
            {
                'uuid': '0a724735-c333-4d7c-9b41-4bda114587c0',
                'set': {'Hit Range Full': 'm', 'Hit Range': 'm'},
            },
            {
                'uuid': '087533d3-7b32-4028-9903-882f53af94af',
                'set': {'Hit Range Full': 'm', 'Hit Range': 'm'},
            },
        ],
        'insert_after': [],
        'delete': [],
    },
}


def parse_sections(ws):
    """Parse a worksheet into sections."""
    sections = []
    row = 1
    while row <= ws.max_row:
        cell = ws.cell(row=row, column=1)
        if cell.value and cell.font.bold:
            next_cell = ws.cell(row=row + 1, column=1) if row + 1 <= ws.max_row else None
            if next_cell and next_cell.value == 'From Stance' and next_cell.font.bold:
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

                sections.append((heading, header_cells, data_rows))
            else:
                row += 1
        else:
            row += 1

    return sections


def apply_corrections(sections, corrections):
    """Apply modify/insert/delete corrections to parsed sections."""
    uuid_col = 'UUID'

    new_columns = []

    for mod in corrections.get('modify', []):
        target_uuid = mod['uuid']
        found = False
        for _, header_cells, rows in sections:
            for row in rows:
                if row.get(uuid_col) == target_uuid:
                    for col_name, value in mod['set'].items():
                        if col_name not in header_cells:
                            new_columns.append(col_name)
                        row[col_name] = value
                    found = True
                    break
            if found:
                break
        if not found:
            print(f"  WARNING: UUID {target_uuid} not found for modify")

    for col_name in new_columns:
        for _, header_cells, _ in sections:
            if col_name not in header_cells:
                header_cells.append(col_name)

    for ins in corrections.get('insert_after', []):
        after_uuid = ins['after_uuid']
        new_row = ins['row']
        found = False
        for _, header_cells, rows in sections:
            for i, row in enumerate(rows):
                if row.get(uuid_col) == after_uuid:
                    row_dict = {col: '' for col in header_cells}
                    row_dict.update(new_row)
                    rows.insert(i + 1, row_dict)
                    found = True
                    break
            if found:
                break
        if not found:
            print(f"  WARNING: UUID {after_uuid} not found for insert_after")

    for delete_uuid in corrections.get('delete', []):
        found = False
        for _, _, rows in sections:
            for i, row in enumerate(rows):
                if row.get(uuid_col) == delete_uuid:
                    rows.pop(i)
                    found = True
                    break
            if found:
                break
        if not found:
            print(f"  WARNING: UUID {delete_uuid} not found for delete")

    return sections


HIDDEN_COLUMNS = {'Command', 'Move Name', 'Damage Sum', 'Hit Range',
                  'Block Adv', 'Hit Adv', 'Counter Hit Adv', 'Speed'}


def autofit_columns(ws, header_row_numbers):
    """Auto-fit column widths. Hidden columns get width 0."""
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
    """Find all characters that have step8c xlsx files."""
    files = glob.glob('sources/*_step8c.xlsx')
    return sorted(
        os.path.basename(f).replace('_step8c.xlsx', '')
        for f in files
        if not os.path.basename(f).startswith('~$')
    )


def main():
    characters = find_characters()
    if not characters:
        print("No step8c files found in sources/")
        return
    print(f"Found characters: {characters}")

    for char in characters:
        input_path = f"sources/{char}_step8c.xlsx"
        output_path = f"sources/{char}_step9.xlsx"

        print(f"\n=== {char.capitalize()} ===")
        wb_in = load_workbook(input_path)
        ws_in = wb_in.active

        sections = parse_sections(ws_in)

        corrections = CORRECTIONS.get(char, {'modify': [], 'insert_after': [], 'delete': []})
        sections = apply_corrections(sections, corrections)

        wb_out = Workbook()
        ws = wb_out.active
        ws.title = "Merged"
        row_num = 1
        header_row_numbers = set()

        for section_name, header_cells, rows in sections:
            cell = ws.cell(row=row_num, column=1, value=section_name)
            cell.font = Font(bold=True)
            row_num += 1

            for col, h in enumerate(header_cells, 1):
                cell = ws.cell(row=row_num, column=col, value=h)
                cell.font = Font(bold=True)
            header_row_numbers.add(row_num)
            row_num += 1

            for row_dict in rows:
                for col, col_name in enumerate(header_cells, 1):
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
