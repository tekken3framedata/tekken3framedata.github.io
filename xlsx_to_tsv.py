#!/usr/bin/env python3
"""Convert all .xlsx files in the current directory to .tsv files."""

import glob
import os

from openpyxl import load_workbook


def xlsx_to_tsv(xlsx_path):
    tsv_path = os.path.splitext(xlsx_path)[0] + '.tsv'
    wb = load_workbook(xlsx_path, read_only=True)
    ws = wb.active

    header_written = False
    with open(tsv_path, 'w', encoding='utf-8') as f:
        for row in ws.iter_rows(values_only=True):
            uuid_cell = str(row[0]) if row[0] else ''
            if uuid_cell == 'UUID':
                if not header_written:
                    cells = [str(cell) if cell is not None else '' for cell in row]
                    f.write('\t'.join(cells) + '\n')
                    header_written = True
                continue
            # Skip rows without a valid UUID (section headers, blank rows)
            if len(uuid_cell) != 36 or uuid_cell.count('-') != 4:
                continue
            cells = [str(cell) if cell is not None else '' for cell in row]
            f.write('\t'.join(cells) + '\n')

    wb.close()
    print(f"{xlsx_path} -> {tsv_path}")


def main():
    files = sorted(f for f in glob.glob('*.xlsx') if not os.path.basename(f).startswith('~$'))
    if not files:
        print("No .xlsx files found in current directory.")
        return
    for path in files:
        xlsx_to_tsv(path)


if __name__ == '__main__':
    main()
