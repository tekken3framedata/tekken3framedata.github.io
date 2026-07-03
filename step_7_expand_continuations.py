#!/usr/bin/env python3
"""
Step 7: Expand continuation moves into full command and name forms.

Reads sources/<character>_step6b.xlsx. Adds seven new columns:
  - Command Full: parent command + separator + own command for the full chain
  - Move Name Full: parent name – own name chain
  - Damage Full: parent Damage Sum + own Damage Sum chain (comma-separated)
  - Hit Range Full: parent hit range + own hit range chain
  - Block Adv Full: parent block adv + own block adv (space-separated)
  - Hit Adv Full: parent hit adv + own hit adv (space-separated)
  - Counter Hit Adv Full: parent counter hit adv + own (space-separated)

Separator logic for Command Full:
  - Empty if suffix starts with ~ or <
  - Comma otherwise

Move Name Full uses " – " (space-endash-space) as separator.

Output: sources/<character>_step7.xlsx

Usage:
    python3 step_7_expand_continuations.py
"""

import glob
import os
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font


def strip_continuation_prefix(value):
    """Remove leading – markers and space from a value."""
    if not value:
        return ''
    return re.sub(r'^–+ ', '', str(value))


def parse_sections(ws):
    """Parse a worksheet into sections."""
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


def get_continuation_level(command):
    """Determine continuation level from leading – markers."""
    if not command:
        return 0
    match = re.match(r'^(–+) ', str(command))
    if not match:
        return 0
    return len(match.group(1))


def expand_section(rows):
    """Expand continuations within a section using Parent UUID lookups."""
    rows_by_uuid = {}
    for row in rows:
        uuid = row.get('UUID', '')
        if uuid:
            rows_by_uuid[uuid] = row
        ml_uuid = row.get('ML UUID', '')
        if ml_uuid:
            rows_by_uuid[ml_uuid] = row

    for row in rows:
        cmd = row.get('Command', '')
        level = get_continuation_level(cmd)

        if level == 0:
            row['Command Full'] = cmd
            row['Alt Commands Full'] = row.get('Alt Commands', '')
            row['Move Name Full'] = row.get('Move Name', '')
            row['Speed Full'] = row.get('Speed', '')
            row['Damage Full'] = row.get('Damage', '')
            row['Hit Range Full'] = row.get('Hit Range', '')
            row['Block Adv Full'] = row.get('Block Adv', '')
            row['Hit Adv Full'] = row.get('Hit Adv', '')
            row['Counter Hit Adv Full'] = row.get('Counter Hit Adv', '')
            continue

        own_cmd = strip_continuation_prefix(cmd)
        own_name = strip_continuation_prefix(row.get('Move Name', ''))
        own_damage = row.get('Damage', '')
        own_range = row.get('Hit Range', '')
        own_block = row.get('Block Adv', '')
        own_hit = row.get('Hit Adv', '')
        own_ch = row.get('Counter Hit Adv', '')

        parent_uuid = row.get('Parent UUID', '')
        parent_row = rows_by_uuid.get(parent_uuid)

        if parent_row:
            parent_cmd = parent_row.get('Command Full', parent_row.get('Command', ''))
            parent_name = parent_row.get('Move Name Full', parent_row.get('Move Name', ''))
            parent_speed = parent_row.get('Speed Full', parent_row.get('Speed', ''))
            parent_damage = parent_row.get('Damage Full', parent_row.get('Damage', ''))
            parent_range = parent_row.get('Hit Range Full', parent_row.get('Hit Range', ''))
            parent_block = parent_row.get('Block Adv Full', parent_row.get('Block Adv', ''))
            parent_hit = parent_row.get('Hit Adv Full', parent_row.get('Hit Adv', ''))
            parent_ch = parent_row.get('Counter Hit Adv Full', parent_row.get('Counter Hit Adv', ''))
        else:
            parent_cmd = ''
            parent_name = ''
            parent_speed = ''
            parent_damage = ''
            parent_range = ''
            parent_block = ''
            parent_hit = ''
            parent_ch = ''

        if parent_cmd:
            separator = '' if own_cmd.startswith(('~', '<')) else ','
            full_cmd = parent_cmd + separator + own_cmd
        else:
            full_cmd = own_cmd

        if parent_name and own_name:
            full_name = parent_name + ' – ' + own_name
        elif own_name:
            full_name = own_name
        else:
            full_name = parent_name

        if parent_damage and own_damage:
            full_damage = str(parent_damage) + ',' + str(own_damage)
        elif own_damage:
            full_damage = own_damage
        else:
            full_damage = parent_damage

        if parent_range and own_range:
            full_range = parent_range + own_range
        elif own_range:
            full_range = own_range
        else:
            full_range = parent_range

        def join_space_separated(parent_val, own_val):
            p = str(parent_val).strip() if parent_val else ''
            o = str(own_val).strip() if own_val else ''
            if p and o:
                return p + ' ' + o
            return o or p

        row['Command Full'] = full_cmd
        alt_cmds = row.get('Alt Commands', '')
        if alt_cmds and parent_cmd:
            alt_full = []
            for alt in alt_cmds.split('; '):
                sep = '' if alt.startswith(('~', '<')) else ','
                alt_full.append(parent_cmd + sep + alt)
            row['Alt Commands Full'] = '; '.join(alt_full)
        else:
            row['Alt Commands Full'] = alt_cmds
        row['Move Name Full'] = full_name
        row['Speed Full'] = parent_speed
        row['Damage Full'] = full_damage
        row['Hit Range Full'] = full_range
        row['Block Adv Full'] = join_space_separated(parent_block, own_block)
        row['Hit Adv Full'] = join_space_separated(parent_hit, own_hit)
        row['Counter Hit Adv Full'] = join_space_separated(parent_ch, own_ch)


OUTPUT_COLUMNS = [
    'Command', 'Command Full', 'Alt Commands', 'Alt Commands Full', 'Move Name', 'Move Name Full', 'To Stance',
    'Speed', 'Speed Full', 'Block Adv', 'Block Adv Full', 'Hit Adv', 'Hit Adv Full',
    'Counter Hit Adv', 'Counter Hit Adv Full', 'Damage', 'Damage Sum', 'Damage Full',
    'Hit Range', 'Hit Range Full', 'Throw Type', 'Throw Escape', 'Properties', 'Notes',
    'Taggable', 'UUID', 'ML UUID', 'Parent UUID', 'Unmatched',
]


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
    """Find all characters that have step6 xlsx files."""
    files = glob.glob('sources/*_step6b.xlsx')
    return sorted(
        os.path.basename(f).replace('_step6b.xlsx', '')
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
        input_path = f"sources/{char}_step6b.xlsx"
        output_path = f"sources/{char}_step7.xlsx"

        print(f"\n=== {char.capitalize()} ===")
        wb_in = load_workbook(input_path)
        ws_in = wb_in.active

        sections = parse_sections(ws_in)

        continuation_count = 0
        for _, rows in sections:
            expand_section(rows)
            continuation_count += sum(1 for r in rows if get_continuation_level(r.get('Command', '')) > 0)

        print(f"  {continuation_count} continuations expanded")

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
