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
    'lee': {
        'modify': [],
        'insert_after': [],
        'delete': [
            # Fang Rush – Hit Man Stance (b+1,1,3+4): alternative stance
            # recovery, to stance info goes on parent row instead
            '3ca6233d-6896-42c6-a780-d5112744e6d8',
            # Alternate Fang Rush – Hit Man Stance (b+1,N+1,3+4)
            'a9627d79-0bb4-4abf-9fca-4e78a13b0bb0',
        ],
    },
    'hwoarang': {
        'modify': [
            # Cheap Shot – Retreat (f+3~b): alternative stance recovery,
            # not a cancel — single hit with BT recovery
            {
                'uuid': '86748328-3934-4013-bb61-5f21c8a678d1',
                'set': {
                    'To Stance': 'BT',
                    'Block Adv': '-16',
                    'Block Adv Full': '-16',
                    'Hit Adv': '-5',
                    'Hit Adv Full': '-5',
                    'Counter Hit Adv': '-5',
                    'Counter Hit Adv Full': '-5',
                    'Damage': '25',
                    'Damage Full': '25',
                    'Damage Sum': '25',
                    'Hit Range': 'h',
                    'Hit Range Full': 'h',
                },
            },
        ],
        'insert_after': [],
        'delete': [],
    },
    'jin': {
        'modify': [],
        'insert_after': [],
        'delete': [],
    },
    'julia': {
        'modify': [
            # Arm Whip – Back Push (b+2,1+2): source had (1+2_5) meaning
            # two separate follow-ups, not alternatives. Remove cancel
            # signature — 1+2 is a real hit.
            {
                'uuid': 'a2a00b1a-1750-4988-9b6c-606239b6a0c6',
                'set': {
                    'Alt Commands': '',
                    'Notes': '',
                },
            },
        ],
        'insert_after': [

            # Arm Whip – Tag (b+2,5): the _5 follow-up lost by step 4
            {
                'after_uuid': 'a2a00b1a-1750-4988-9b6c-606239b6a0c6',
                'row': {
                    'Command': '– 5',
                    'Command Full': 'b+2 ,5',
                    'Move Name': '– Back Push into Michelle\'s Lariat',
                    'Move Name Full': 'Arm Whip – Back Push into Michelle\'s Lariat',
                    'Damage': '24',
                    'Damage Sum': '24',
                    'Damage Full': '12,24',
                    'Hit Range': '!',
                    'Hit Range Full': 'h!',
                    'Speed Full': '10',
                    'Taggable': 'TRUE',
                    'Notes': 'Only with Michelle.',
                    'UUID': 'b10b068c-ac9e-49e7-81b0-86920d0c34af',
                    'Parent UUID': 'a72aa904-1124-409f-a75b-cb6b7a95fa1f',
                },
            },
        ],
        'delete': [],
        'delete_by_ml_uuid': [
            # Arm Whip (b+2) in Grappling Arts: this is a hit throw,
            # not a standalone throw — belongs with Special Arts b+2 rows
            '1b569aeb-dd63-4db9-bfe2-caa3a868a11d',
        ],
    },
    'michelle': {
        'modify': [
            {
                'uuid': '8045d9f7-3789-4f1b-b9c6-3335bf7f87f4',
                'set': {
                    'Notes': '',
                },
            },
        ],
        'insert_after': [
            # Arm Whip – Rear Suplex: move from Grappling Arts, fix Parent UUID
            {
                'after_uuid': '8045d9f7-3789-4f1b-b9c6-3335bf7f87f4',
                'row': {
                    'Command': '– d+1+2',
                    'Command Full': 'b+2 ,d+1+2',
                    'Move Name': '– Rear Suplex',
                    'Move Name Full': 'Arm Whip – Rear Suplex',
                    'Speed Full': '10',
                    'Damage': '45',
                    'Damage Sum': '45',
                    'Damage Full': '12,45',
                    'Hit Range': '!',
                    'Hit Range Full': 'h!',
                    'UUID': 'f1c6a2cd-8f86-45e2-8020-1482f8fb6c52',
                    'ML UUID': 'f1c6a2cd-8f86-45e2-8020-1482f8fb6c52',
                    'Parent UUID': '65aaecf4-ea4f-4d97-b0ed-4df6be392720',
                },
            },
            # Arm Whip – Tag Throw (b+2,5): the _5 follow-up lost by step 4
            {
                'after_uuid': 'f1c6a2cd-8f86-45e2-8020-1482f8fb6c52',
                'row': {
                    'Command': '– 5',
                    'Command Full': 'b+2 ,5',
                    'Move Name': '– Back Push into Julia\'s Running Bulldog',
                    'Move Name Full': 'Arm Whip – Back Push into Julia\'s Running Bulldog',
                    'Damage': '24',
                    'Damage Sum': '24',
                    'Damage Full': '12,24',
                    'Hit Range': '!',
                    'Hit Range Full': 'h!',
                    'Speed Full': '10',
                    'Taggable': 'TRUE',
                    'Notes': 'Only with Julia.',
                    'UUID': '1929ca31-425b-4548-b294-f29693b8fe20',
                    'Parent UUID': '65aaecf4-ea4f-4d97-b0ed-4df6be392720',
                },
            },
        ],
        'delete': [
            # Rear Suplex from Grappling Arts (re-inserted in Special Arts)
            'f1c6a2cd-8f86-45e2-8020-1482f8fb6c52',
        ],
        'delete_by_ml_uuid': [
            # Arm Whip (b+2) in Grappling Arts: hit throw, not standalone
            '151f3bbe-77de-4044-9738-f8670c4da698',
        ],
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
            # Shady Lotus to Rain Dance (FC DF+2) — source lumps two moves'
            # frame values together; keep only the single-hit RDS version
            {
                'uuid': 'c86a461d-9853-409e-8dbf-c140f78957bc',
                'set': {
                    'Block Adv': '-4',
                    'Block Adv Full': '-4',
                    'Hit Adv': '+6',
                    'Hit Adv Full': '+6',
                    'Counter Hit Adv': '+6',
                    'Counter Hit Adv Full': '+6',
                    'Hit Range': 's',
                    'Hit Range Full': 's',
                    'Move Name': 'Shady Lotus to Rain Dance',
                    'Move Name Full': 'Shady Lotus to Rain Dance',
                },
            },
            # Shady Lotus to FC (FC df+2,1) — fix move name and hit range
            {
                'uuid': 'a2d0c94a-e4f1-4b84-867d-a8aacd632dde',
                'set': {
                    'Move Name': 'Shady Lotus to FC',
                    'Move Name Full': 'Shady Lotus to FC',
                    'Hit Range': 'ss',
                    'Hit Range Full': 'ss',
                },
            },
        ],
        'insert_after': [],
        'delete': [],
    },
    'jack2': {
        'modify': [
            {
                'uuid': '3be14565-e783-4db3-9168-bf98e290f21d',
                'set': {
                    'To Stance': '(~B)Roll back; (~F)Roll forward',
                },
            },
            {
                'uuid': '674f6841-5fbb-4b15-865f-9ab351c1a0f0',
                'set': {
                    'Block Adv': '-20 -22 -20 -22 -20 -22',
                    'Block Adv Full': '-20 -22 -20 -22 -20 -22',
                    'Hit Adv': '-9 -11 -9 -11 -9 -11',
                    'Hit Adv Full': '-9 -11 -9 -11 -9 -11',
                    'Counter Hit Adv': '-9 -11 -9 -11 -9 -11',
                    'Counter Hit Adv Full': '-9 -11 -9 -11 -9 -11',
                },
            },
            {
                'uuid': 'eb9bbfe4-68c3-4e58-baf7-821687138e91',
                'set': {
                    'Properties': 'JGc RC',
                    'Notes': 'Juggles on the last hit if the first hit was a counter hit.',
                },
            },
            {
                'uuid': 'eb0c424e-d4b0-438e-bfb7-164bbca1894a',
                'set': {
                    'To Stance': '(~B)Roll back; (~F)Roll forward',
                },
            },
            {
                'uuid': '1e575213-1523-40b8-9492-0a92235d1799',
                'set': {
                    'To Stance': '(~B)Roll back; (~F)Roll forward',
                },
            },
            {
                'uuid': '421087d9-77e4-4344-95d3-0ffee0325022',
                'set': {
                    'To Stance': '(~B)Roll back; (~F)Roll forward',
                },
            },
            {
                'uuid': '3f1586e3-7d70-40ce-a201-943ab9917e0a',
                'set': {
                    'To Stance': '(~B)Roll back; (~F)Roll forward; (~U)Hop up',
                },
            },
        ],
        'insert_after': [],
        'delete': [
            # Sit Down – Roll Back/Forward: to-stance data, not moves
            'a96a391b-b0cf-493d-a205-6511f1e34903',
            'df8e44d6-d84c-4ee6-8117-15cb0b65ceb6',
            # Quick Upper Rush: duplicate of Medium Hammer Rush continuations
            '7507a693-7948-45eb-a05e-3707d923e16d',
        ],
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

    for delete_ml_uuid in corrections.get('delete_by_ml_uuid', []):
        found = False
        for _, _, rows in sections:
            for i, row in enumerate(rows):
                if row.get('ML UUID') == delete_ml_uuid:
                    rows.pop(i)
                    found = True
                    break
            if found:
                break
        if not found:
            print(f"  WARNING: ML UUID {delete_ml_uuid} not found for delete_by_ml_uuid")

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
