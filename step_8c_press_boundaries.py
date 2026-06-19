#!/usr/bin/env python3
"""
Step 8c: Mark press boundaries in Command Full with spaces.

Reads sources/<character>_step8b.xlsx. Inserts spaces before press-separating
characters (, < ~) in Command Full to mark where each new button press begins.

Rules:
  - < always separates presses (delay into new button)
  - ~ separates if followed by a token containing a button (1-4)
  - , separates if previous token has a button and next token has a button
    (with exceptions for direction-only contexts)

Example: "1<2,4~1+4,2" → "1 <2 ,4 ~1+4 ,2"

Output: sources/<character>_step8c.xlsx

Usage:
    python3 step_8c_press_boundaries.py
"""

import glob
import os
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font


DIRECTIONS = {'f', 'b', 'd', 'u', 'df', 'db', 'uf', 'ub', 'N', 'F', 'B', 'D', 'U',
              'DF', 'DB', 'UF', 'UB', 'FC'}


def token_has_button(tok):
    return bool(re.search(r'[1-4]', tok))


def add_spaces(cmd):
    """Insert spaces between presses in a command string."""
    trailing = ''
    m = re.search(r'(\s+-\s*(\[.*?\])?\s*)$', cmd)
    if m:
        trailing = m.group(0)
        cmd_clean = cmd[:m.start()]
    else:
        cmd_clean = cmd

    result = []
    i = 0
    while i < len(cmd_clean):
        ch = cmd_clean[i]
        if ch == ':':
            result.append(' ')
            result.append(ch)
        elif ch == '<':
            result.append(' ')
            result.append(ch)
        elif ch == '~':
            rest = cmd_clean[i+1:]
            next_sep = re.search(r'[,<~:]', rest)
            next_part = rest[:next_sep.start()] if next_sep else rest
            before = ''.join(result).split(' ')[-1] if result else ''
            if token_has_button(next_part) and token_has_button(before):
                result.append(' ')
            result.append(ch)
        elif ch == ',':
            before = ''.join(result).split(' ')[-1] if result else ''
            rest = cmd_clean[i+1:]
            next_sep = re.search(r'[,<~:]', rest)
            next_part = rest[:next_sep.start()] if next_sep else rest
            prev_has_btn = token_has_button(before)
            prev_ends_dir = False
            if '~' not in before:
                parts = before.split('<')
                last = parts[-1]
                if last in DIRECTIONS:
                    prev_ends_dir = True
            curr_is_dir_btn = bool(re.match(r'^[fbudFBUDN][fbudFBUDN]*\+[1-4]', next_part))
            curr_is_motion = '~' in next_part
            curr_has_btn = token_has_button(next_part)
            prev_is_pure_dir = not token_has_button(before)
            curr_is_pure_dir = not curr_has_btn
            if prev_is_pure_dir and curr_is_pure_dir:
                result.append(ch)
            elif prev_is_pure_dir and curr_is_dir_btn:
                result.append(ch)
            elif prev_is_pure_dir and curr_is_motion:
                result.append(ch)
            elif prev_is_pure_dir and curr_has_btn and not curr_is_dir_btn:
                result.append(ch)
            elif prev_has_btn and prev_ends_dir:
                result.append(ch)
            elif prev_has_btn and not prev_ends_dir and curr_has_btn:
                result.append(' ')
                result.append(ch)
            elif prev_has_btn and curr_is_pure_dir:
                result.append(ch)
            else:
                result.append(ch)
        else:
            result.append(ch)
        i += 1

    final = ''.join(result).strip() + trailing
    final = re.sub(r'  +', ' ', final)
    return final


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


def process_sections(sections):
    """Apply add_spaces to Command Full in all rows. Returns count of modified rows."""
    modified = 0
    for _, _, rows in sections:
        for row in rows:
            cmd_full = str(row.get('Command Full', '') or '')
            if cmd_full:
                spaced = add_spaces(cmd_full)
                if spaced != cmd_full:
                    row['Command Full'] = spaced
                    modified += 1
    return modified


def autofit_columns(ws, header_row_numbers, hidden_columns):
    """Auto-fit column widths. Hidden columns get width 0."""
    for col in ws.columns:
        col_letter = col[0].column_letter
        header_value = None
        for cell in col:
            if cell.row in header_row_numbers and cell.value:
                header_value = cell.value
                break
        if header_value in hidden_columns:
            ws.column_dimensions[col_letter].width = 0
            ws.column_dimensions[col_letter].hidden = True
            continue
        max_len = 0
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_len + 2


OUTPUT_COLUMNS = [
    'From Stance',
    'Command', 'Command Full', 'Alt Commands', 'Move Name', 'Move Name Full', 'To Stance',
    'Speed', 'Speed Full', 'Block Adv', 'Block Adv Full', 'Hit Adv', 'Hit Adv Full',
    'Counter Hit Adv', 'Counter Hit Adv Full', 'Damage', 'Damage Sum', 'Damage Full',
    'Hit Range', 'Hit Range Full', 'Throw Type', 'Throw Escape', 'Properties', 'Notes',
    'Taggable', 'UUID', 'ML UUID', 'Parent UUID', 'Unmatched',
]

HIDDEN_COLUMNS = {'Command', 'Move Name', 'Damage Sum', 'Hit Range',
                  'Block Adv', 'Hit Adv', 'Counter Hit Adv', 'Speed'}


def find_characters():
    """Find all characters that have step8b xlsx files."""
    files = glob.glob('sources/*_step8b.xlsx')
    return sorted(
        os.path.basename(f).replace('_step8b.xlsx', '')
        for f in files
        if not os.path.basename(f).startswith('~$')
    )


def main():
    characters = find_characters()
    if not characters:
        print("No step8b files found in sources/")
        return
    print(f"Found characters: {characters}")

    for char in characters:
        input_path = f"sources/{char}_step8b.xlsx"
        output_path = f"sources/{char}_step8c.xlsx"

        print(f"\n=== {char.capitalize()} ===")
        wb_in = load_workbook(input_path)
        ws_in = wb_in.active

        sections = parse_sections(ws_in)
        modified = process_sections(sections)

        print(f"  {modified} Command Full values spaced")

        wb_out = Workbook()
        ws = wb_out.active
        ws.title = "Merged"
        row_num = 1
        header_row_numbers = set()

        for section_name, header_cells, rows in sections:
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

        autofit_columns(ws, header_row_numbers, HIDDEN_COLUMNS)
        wb_out.save(output_path)
        print(f"  Written to {output_path}")

    print(f"\nDone. Processed {len(characters)} characters.")


if __name__ == '__main__':
    main()
