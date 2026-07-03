#!/usr/bin/env python3
"""
Step 1: Convert HTML source files to XLSX with minimal transformation.

Reads sources/<character>_framedata.html and sources/<character>_movelist.html,
parses their tables section by section, and writes them to an XLSX file.

The only data transformation is expanding #N footnote references in the
movelist Properties column into a Notes column.

Output: sources/<character>_step1.xlsx

Usage:
    python3 step_1_html_to_xlsx.py
"""

import re
import glob
import os

from openpyxl import Workbook
from openpyxl.styles import Font


def read_html(path):
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()


def find_all_sections(content):
    """Find all <h2> section headings that have a table after them."""
    sections = []
    body_start = content.find('<h2>')
    if body_start == -1:
        return sections
    for match in re.finditer(r'<h2>(.*?)</h2>', content[body_start:]):
        heading = re.sub(r'<[^>]+>', '', match.group(1)).strip()
        pos = body_start + match.start()
        table_start = content.find('<table', pos)
        next_h2 = content.find('<h2>', pos + 1)
        if table_start != -1 and (next_h2 == -1 or table_start < next_h2):
            sections.append(heading)
    return sections


def get_table_html(content, heading):
    """Extract the <table>...</table> HTML for a given section heading."""
    start = content.find(f'<h2>{heading}</h2>')
    if start == -1:
        return ""
    table_start = content.find('<table', start)
    table_end = content.find('</table>', table_start)
    return content[table_start:table_end]


def get_footnotes(content, heading):
    """Get footnotes between a section's </table> and the next <h2>.

    Returns dict: '#N' -> footnote text.
    """
    start = content.find(f'<h2>{heading}</h2>')
    if start == -1:
        return {}
    table_end = content.find('</table>', start)
    next_h2 = content.find('<h2>', table_end)
    if next_h2 == -1:
        section = content[table_end:]
    else:
        section = content[table_end:next_h2]

    footnotes = {}
    for match in re.finditer(r'(#\d+)\s+(.*?)(?:<br|[\n\r]|</fieldset)', section):
        key = match.group(1)
        text = match.group(2).strip().rstrip('/')
        footnotes[key] = text
    return footnotes


def replace_continuation_marker(cell_text):
    """Replace continuation syntax with level-appropriate arrows.

    '&nbsp;=&nbsp;X' (1 nbsp before =, level 1) -> '– X'
    '&nbsp;&nbsp;&nbsp;=&nbsp;X' (3 nbsp before =, level 2) -> '–– X'

    Non-continuation cells are stripped normally.
    """
    match = re.match(r'^(&nbsp;)*=(?=~|&nbsp;)(&nbsp;)?', cell_text)
    if not match:
        cell_text = cell_text.replace('&nbsp;', ' ')
        return cell_text.strip()
    leading_nbsps = cell_text.split('=')[0].count('&nbsp;')
    level = (leading_nbsps // 2) + 1
    rest = cell_text[match.end():]
    rest = rest.replace('&nbsp;', ' ').strip()
    return '–' * level + ' ' + rest


def parse_table_rows(table_html):
    """Parse HTML table into list of rows. Each row is (uuid, [cell_texts]).

    First row (header) is included.
    """
    rows = []
    tr_pattern = re.compile(r'<tr([^>]*)>(.*?)</tr>', re.DOTALL)
    td_pattern = re.compile(r'<td[^>]*>(.*?)</td>', re.DOTALL)
    uuid_pattern = re.compile(r'data-uuid="([^"]+)"')

    for tr_match in tr_pattern.finditer(table_html):
        tr_attrs = tr_match.group(1)
        tr_content = tr_match.group(2)

        uuid_match = uuid_pattern.search(tr_attrs)
        row_uuid = uuid_match.group(1) if uuid_match else ''

        cells = []
        for col_idx, td_match in enumerate(td_pattern.finditer(tr_content)):
            cell_text = td_match.group(1)
            cell_text = cell_text.replace('&lt;', '<')
            cell_text = cell_text.replace('&gt;', '>')
            cell_text = cell_text.replace('&amp;', '&')
            cell_text = cell_text.replace('&quot;', '"')
            cell_text = re.sub(r'<[^>]+>', '', cell_text)
            cell_text = replace_continuation_marker(cell_text)
            cells.append(cell_text)
        if cells:
            rows.append((row_uuid, cells))
    return rows


def expand_footnotes(properties, footnotes):
    """Replace #N references in properties with footnote text.

    Returns (clean_properties, notes_text).
    """
    if not properties or not footnotes:
        return properties, ''
    notes = []
    remaining = properties
    for key in sorted(footnotes.keys(), key=lambda k: int(k[1:])):
        if key in remaining:
            notes.append(footnotes[key])
            remaining = remaining.replace(key, '').strip()
    return remaining, '; '.join(notes)


def write_section(ws, row_num, heading, rows, footnotes=None, add_notes_col=False):
    """Write a section to the worksheet.

    Returns the next row number.
    """
    # Section heading
    cell = ws.cell(row=row_num, column=1, value=heading)
    cell.font = Font(bold=True)
    row_num += 1

    if not rows:
        return row_num + 1

    # Header row
    header_uuid, header_cells = rows[0]
    headers = list(header_cells)
    if add_notes_col and 'Notes' not in headers:
        headers.append('Notes')
    headers.append('UUID')

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row_num, column=col, value=h)
        cell.font = Font(bold=True)
    row_num += 1

    # Find Properties column index for footnote expansion
    props_col = None
    if footnotes and add_notes_col:
        try:
            props_col = header_cells.index('Properties')
        except ValueError:
            pass

    notes_col_idx = headers.index('Notes') if 'Notes' in headers else None
    uuid_col_idx = headers.index('UUID')

    # Data rows
    for uuid, cells in rows[1:]:
        notes_text = ''
        row_cells = list(cells)

        if props_col is not None and props_col < len(row_cells):
            clean_props, notes_text = expand_footnotes(row_cells[props_col], footnotes)
            row_cells[props_col] = clean_props

        # Pad to match header length (minus UUID and Notes which we handle)
        while len(row_cells) < len(header_cells):
            row_cells.append('')

        for col, val in enumerate(row_cells, 1):
            ws.cell(row=row_num, column=col, value=val)

        if notes_col_idx is not None:
            ws.cell(row=row_num, column=notes_col_idx + 1, value=notes_text)

        ws.cell(row=row_num, column=uuid_col_idx + 1, value=uuid)
        row_num += 1

    return row_num + 1  # blank row after section


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
    """Find all characters that have both framedata and movelist HTML sources."""
    fd_files = glob.glob('sources/*_framedata.html')
    characters = []
    for fd_path in sorted(fd_files):
        char = os.path.basename(fd_path).replace('_framedata.html', '')
        if char.startswith('~$'):
            continue
        ml_path = f'sources/{char}_movelist.html'
        if os.path.exists(ml_path):
            characters.append(char)
    return characters


def process_character(character):
    """Process a single character's HTML sources into an XLSX file."""
    fd_path = f"sources/{character}_framedata.html"
    ml_path = f"sources/{character}_movelist.html"
    output_path = f"sources/{character}_step1.xlsx"

    print(f"\n=== {character.capitalize()} ===")
    print(f"Reading {fd_path}...")
    fd_html = read_html(fd_path)
    fd_sections = find_all_sections(fd_html)
    print(f"  Sections: {fd_sections}")

    print(f"Reading {ml_path}...")
    ml_html = read_html(ml_path)
    ml_sections = find_all_sections(ml_html)
    print(f"  Sections: {ml_sections}")

    wb = Workbook()

    # Frame Data sheet
    ws_fd = wb.active
    ws_fd.title = "Frame Data"
    row_num = 1
    for section in fd_sections:
        table_html = get_table_html(fd_html, section)
        rows = parse_table_rows(table_html)
        footnotes = get_footnotes(fd_html, section)
        row_num = write_section(ws_fd, row_num, section, rows,
                                footnotes=footnotes, add_notes_col=bool(footnotes))
        print(f"  FD {section}: {len(rows) - 1} rows")

    # Movelist sheet
    ws_ml = wb.create_sheet("Movelist")
    row_num = 1
    for section in ml_sections:
        table_html = get_table_html(ml_html, section)
        rows = parse_table_rows(table_html)
        footnotes = get_footnotes(ml_html, section)
        row_num = write_section(ws_ml, row_num, section, rows,
                                footnotes=footnotes, add_notes_col=bool(footnotes))
        fn_count = len(footnotes)
        print(f"  ML {section}: {len(rows) - 1} rows" +
              (f" ({fn_count} footnotes)" if fn_count else ""))

    for ws in wb.worksheets:
        autofit_columns(ws)

    wb.save(output_path)
    print(f"Written to {output_path}")


def main():
    characters = find_characters()
    if not characters:
        print("No character sources found in sources/")
        return
    print(f"Found characters: {characters}")
    for char in characters:
        process_character(char)
    print(f"\nDone. Processed {len(characters)} characters.")


if __name__ == '__main__':
    main()
