#!/usr/bin/env python3
"""
Step 8: Split multi-hit rows into individual rows per hit.

Reads sources/<character>_step7.xlsx. A row is multi-hit if the Block Adv column
contains space-separated values. Hit count is determined by the Block Adv column.

Each hit becomes its own row with:
  - UUID: {original_uuid}-1, {original_uuid}-2, ...
  - Command: "{original_command} (First)", "(Second)", etc.
  - Damage, Block Adv, Hit Adv, Counter Hit Adv, Hit Range split accordingly

If splitting would only produce sub-rows whose commands already exist in the
table, the split is skipped. If a sub-row's command exists but data differs,
a note is written to the Processor Notes column.

Output: sources/<character>_step8.xlsx

Usage:
    python3 step_8_split_hits.py
"""

import glob
import os
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font


ORDINALS = ['First', 'Second', 'Third', 'Fourth', 'Fifth', 'Sixth',
            'Seventh', 'Eighth', 'Ninth', 'Tenth', 'Eleventh', 'Twelfth']


def split_space_separated(value, count):
    """Split space-separated values, padding with '' if needed."""
    if not value:
        return [''] * count
    parts = str(value).split()
    while len(parts) < count:
        parts.append('')
    return parts[:count]


def split_space_separated_skip_x(value, count):
    """Split space-separated values, skipping 'x' entries."""
    if not value:
        return [''] * count
    parts = [p for p in str(value).split() if p != 'x']
    while len(parts) < count:
        parts.append('')
    return parts[:count]


def split_damage(value, count):
    """Split comma-separated damage values, padding with '' if needed."""
    if not value:
        return [''] * count
    parts = str(value).split(',')
    while len(parts) < count:
        parts.append('')
    return [p.strip() for p in parts[:count]]


def split_hit_range(value, count):
    """Split Hit Range into individual characters."""
    if not value:
        return [''] * count
    chars = list(str(value))
    while len(chars) < count:
        chars.append('')
    return chars[:count]


def build_sub_move_name(original_name, index):
    """Build move name for a split row."""
    ordinal = ORDINALS[index] if index < len(ORDINALS) else str(index + 1)
    return f"{original_name} ({ordinal})"


def parse_progressive_commands(cmd_full, hit_count):
    """Parse Command Full into progressive sub-commands, one per hit.

    Takes the last hit_count hits from the command, building progressive
    commands from the start up to each hit point.

    E.g. '1<2,4~1+4,2' with 2 hits -> ['1<2,4~1+4', '1<2,4~1+4,2']
    """
    clean = re.sub(r'\s+-\s*\[.*?\]\s*$', '', cmd_full)
    groups = re.split(r'(?=[,~<])', clean)
    groups = [g for g in groups if g]

    hit_indices = [i for i, g in enumerate(groups) if re.search(r'[1-4]', g)]

    if len(hit_indices) < hit_count:
        return None

    relevant_hits = hit_indices[-hit_count:]
    progressive = []
    for h_idx in relevant_hits:
        progressive.append(''.join(groups[:h_idx + 1]))
    return progressive


def rows_data_matches(row_a, row_b, fields):
    """Check if two rows have matching data for given fields."""
    for f in fields:
        a = str(row_a.get(f, '') or '')
        b = str(row_b.get(f, '') or '')
        if a != b:
            return False
    return True


def get_hit_count(block_adv):
    """Determine hit count from Block Adv column. Returns 1 if single value.

    'x' values are ignored (not a blockable hit).
    """
    if not block_adv:
        return 1
    parts = [p for p in str(block_adv).split() if p != 'x']
    return len(parts)


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


def process_sections(sections):
    """Split multi-hit rows across all sections."""
    all_cmds = {}
    for _, rows in sections:
        for row in rows:
            cmd_full = row.get('Command Full', '') or row.get('Command', '')
            if cmd_full:
                all_cmds.setdefault(cmd_full, []).append(row)

    split_count = 0
    skip_count = 0

    for section_idx, (heading, rows) in enumerate(sections):
        new_rows = []
        for row in rows:
            hit_count = get_hit_count(row.get('Block Adv', ''))
            if hit_count <= 1:
                new_rows.append(row)
                continue

            original_cmd_full = row.get('Command Full', '') or row.get('Command', '')
            original_name_full = row.get('Move Name Full', '') or row.get('Move Name', '')
            original_uuid = row.get('UUID', '')

            progressive = parse_progressive_commands(original_cmd_full, hit_count)
            if progressive is None:
                new_rows.append(row)
                skip_count += 1
                continue

            all_exist = True
            for sub_cmd in progressive:
                if sub_cmd not in all_cmds:
                    all_exist = False
                    break

            if all_exist:
                skip_count += 1
                new_rows.append(row)
                continue

            blocks = split_space_separated_skip_x(row.get('Block Adv', ''), hit_count)
            hits_adv = split_space_separated_skip_x(row.get('Hit Adv', ''), hit_count)
            counters = split_space_separated_skip_x(row.get('Counter Hit Adv', ''), hit_count)
            damages = split_damage(row.get('Damage', ''), hit_count)
            ranges = split_hit_range(row.get('Hit Range', ''), hit_count)

            for i in range(hit_count):
                sub_row = dict(row)
                sub_row['Command Full'] = progressive[i]
                sub_row['Move Name Full'] = build_sub_move_name(original_name_full, i)
                sub_row['Block Adv'] = blocks[i]
                sub_row['Hit Adv'] = hits_adv[i]
                sub_row['Counter Hit Adv'] = counters[i]
                sub_row['Damage'] = damages[i]
                sub_row['Hit Range'] = ranges[i]
                if original_uuid:
                    sub_row['UUID'] = f"{original_uuid}-{i+1}"
                else:
                    sub_row['UUID'] = ''

                existing = all_cmds.get(sub_row['Command Full'], [])
                if existing:
                    compare_fields = ['Damage', 'Block Adv', 'Hit Adv', 'Counter Hit Adv', 'Hit Range']
                    if not rows_data_matches(sub_row, existing[0], compare_fields):
                        sub_row['Processor Notes'] = f"Conflicts with existing {sub_row['Command Full']}"

                new_rows.append(sub_row)

            for sub_row in new_rows[-hit_count:]:
                cmd_full = sub_row['Command Full']
                all_cmds.setdefault(cmd_full, []).append(sub_row)

            split_count += 1

        sections[section_idx] = (heading, new_rows)

    return split_count, skip_count


OUTPUT_COLUMNS = [
    'Command', 'Command Full', 'Alt Commands', 'Move Name', 'Move Name Full', 'Stance',
    'Hit', 'Block Adv', 'Hit Adv', 'Counter Hit Adv', 'Damage', 'Damage Full',
    'Hit Range', 'Hit Range Full', 'Throw Type', 'Throw Escape', 'Properties', 'Notes',
    'Processor Notes', 'UUID', 'ML UUID', 'Parent UUID', 'Unmatched',
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
    """Find all characters that have step7 xlsx files."""
    files = glob.glob('sources/*_step7.xlsx')
    return sorted(
        os.path.basename(f).replace('_step7.xlsx', '')
        for f in files
        if not os.path.basename(f).startswith('~$')
    )


def main():
    characters = find_characters()
    if not characters:
        print("No step7 files found in sources/")
        return
    print(f"Found characters: {characters}")

    for char in characters:
        input_path = f"sources/{char}_step7.xlsx"
        output_path = f"sources/{char}_step8.xlsx"

        print(f"\n=== {char.capitalize()} ===")
        wb_in = load_workbook(input_path)
        ws_in = wb_in.active

        sections = parse_sections(ws_in)
        split_count, skip_count = process_sections(sections)

        print(f"  {split_count} rows split, {skip_count} skipped (sub-commands already exist)")

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
