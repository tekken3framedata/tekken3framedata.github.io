#!/usr/bin/env python3
"""
Step 4: Split (A_B) alternative notation into primary command and alternatives.

Reads sources/<character>_step3.xlsx:
  - Resolves (A_B) groups to a primary command
  - Generates Alt Commands column with remaining combinations
  - Follow-up moves (Parent UUID set) get resolved but no Alt Commands

Output: sources/<character>_step4.xlsx

Usage:
    python3 step_4_alternatives.py
"""

import glob
import os
import re
from itertools import product

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font


PREFERRED_COMMAND = {
    'julia': {
        'WR+1': 'd,d/f+1',
    },
}

OUTPUT_COLUMNS = [
    'Command', 'Alt Commands', 'Move Name', 'To Stance', 'Speed', 'Block Adv', 'Hit Adv',
    'Counter Hit Adv', 'Damage', 'Hit Range', 'Throw Type', 'Throw Escape', 'Properties', 'Notes',
    'UUID', 'ML UUID', 'Parent UUID', 'Unmatched',
]

ALT_PATTERN = re.compile(r'\(([^)]*_[^)]*)\)')


def split_alternatives(row_dicts, character):
    """Split (A_B) alternative notation into primary command and alternatives."""
    char_prefs = PREFERRED_COMMAND.get(character.lower(), {})

    for row in row_dicts:
        cmd = row.get('Command', '')
        is_followup = bool(row.get('Parent UUID'))

        alt_groups = ALT_PATTERN.findall(cmd)
        if not alt_groups:
            row['Alt Commands'] = ''
            continue

        parts = ALT_PATTERN.split(cmd)
        groups = [g.split('_') for g in alt_groups]

        chosen = [g[0] for g in groups]

        primary = parts[0]
        for i, opt in enumerate(chosen):
            primary += opt + parts[2 * i + 2]
        row['Command'] = primary

        all_commands = []
        for combo in product(*groups):
            built = parts[0]
            for i, opt in enumerate(combo):
                built += opt + parts[2 * i + 2]
            all_commands.append(built)

        if char_prefs and primary in char_prefs:
            preferred = char_prefs[primary]
            if preferred in all_commands:
                primary = preferred
                row['Command'] = primary

        if all(len(g) == 1 for g in groups):
            row['Alt Commands'] = ''
            continue

        alts = [c for c in all_commands if c != primary]
        if is_followup:
            prefix = re.match(r'^–+ ', primary)
            if prefix:
                alts = [a[len(prefix.group()):] if a.startswith(prefix.group()) else a for a in alts]
        row['Alt Commands'] = '; '.join(alts)


def parse_sections(ws):
    """Parse a merged worksheet into sections."""
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


def autofit_columns(ws):
    """Auto-fit all column widths based on cell content."""
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_len + 2


def find_characters():
    """Find all characters that have step3 xlsx files."""
    files = glob.glob('sources/*_step3.xlsx')
    return sorted(
        os.path.basename(f).replace('_step3.xlsx', '')
        for f in files
        if not os.path.basename(f).startswith('~$')
    )


def main():
    characters = find_characters()
    if not characters:
        print("No step3 files found in sources/")
        return
    print(f"Found characters: {characters}")

    for char in characters:
        input_path = f"sources/{char}_step3.xlsx"
        output_path = f"sources/{char}_step4.xlsx"

        print(f"\n=== {char.capitalize()} ===")
        wb_in = load_workbook(input_path)
        ws_in = wb_in.active

        sections = parse_sections(ws_in)

        alt_count = 0
        for _, rows in sections:
            split_alternatives(rows, char)
            alt_count += sum(1 for r in rows if r.get('Alt Commands'))

        print(f"  {alt_count} rows with alternatives")

        wb_out = Workbook()
        ws = wb_out.active
        ws.title = "Merged"
        row_num = 1

        for section_name, rows in sections:
            cell = ws.cell(row=row_num, column=1, value=section_name)
            cell.font = Font(bold=True)
            row_num += 1

            for col, h in enumerate(OUTPUT_COLUMNS, 1):
                cell = ws.cell(row=row_num, column=col, value=h)
                cell.font = Font(bold=True)
            row_num += 1

            for row_dict in rows:
                for col, col_name in enumerate(OUTPUT_COLUMNS, 1):
                    val = row_dict.get(col_name, '')
                    if val:
                        ws.cell(row=row_num, column=col, value=val)
                row_num += 1

            row_num += 1

        autofit_columns(ws)
        wb_out.save(output_path)
        print(f"  Written to {output_path}")

    print(f"\nDone. Processed {len(characters)} characters.")


if __name__ == '__main__':
    main()
