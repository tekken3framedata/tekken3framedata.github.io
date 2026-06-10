#!/usr/bin/env python3
"""
Scrape Tekken frame data and movelist from Tekken Zaibatsu (via Wayback Machine)
and merge them into an XLSX file.

Usage:
    python3 scrape_framedata.py <framedata_url> <movelist_url> <output.xlsx>

Example:
    python3 scrape_framedata.py \
        "https://web.archive.org/web/20201206043428/http://www.tekkenzaibatsu.com/tekkentag/framedata.php?id=julia" \
        "https://web.archive.org/web/20201206042940/http://www.tekkenzaibatsu.com/tekken3/movelist.php?id=julia" \
        julia_framedata_v2.xlsx
"""

import re
import sys
import subprocess
import uuid

from openpyxl import Workbook
from openpyxl.styles import Font, numbers


def fetch_url(url):
    """Fetch URL content using curl (WebFetch doesn't work with web.archive.org)."""
    result = subprocess.run(
        ['curl', '-s', '-L', url],
        capture_output=True, text=True, timeout=60
    )
    return result.stdout


def parse_table(html_content):
    """Parse HTML table rows into list of lists.

    Preserves leading &nbsp; count in the first cell as spaces,
    which encodes continuation depth (1 space = level 1, 3 spaces = level 2).
    """
    rows = []
    tr_pattern = re.compile(r'<tr[^>]*>(.*?)</tr>', re.DOTALL)
    td_pattern = re.compile(r'<td[^>]*>(.*?)</td>', re.DOTALL)

    for tr_match in tr_pattern.finditer(html_content):
        tr_content = tr_match.group(1)
        cells = []
        for col_idx, td_match in enumerate(td_pattern.finditer(tr_content)):
            cell_text = td_match.group(1)
            if col_idx == 0:
                leading_nbsps = len(re.findall(r'&nbsp;', cell_text.split('=')[0])) if '=' in cell_text else 0
                cell_text = cell_text.replace('&nbsp;', ' ')
                cell_text = cell_text.replace('&lt;', '<')
                cell_text = cell_text.replace('&gt;', '>')
                cell_text = cell_text.replace('&amp;', '&')
                cell_text = cell_text.replace('&quot;', '"')
                cell_text = re.sub(r'<[^>]+>', '', cell_text)
                cell_text = cell_text.strip()
                if leading_nbsps > 0 and cell_text.startswith('='):
                    cell_text = (' ' * leading_nbsps) + cell_text
            else:
                cell_text = cell_text.replace('&nbsp;', ' ')
                cell_text = cell_text.replace('&lt;', '<')
                cell_text = cell_text.replace('&gt;', '>')
                cell_text = cell_text.replace('&amp;', '&')
                cell_text = cell_text.replace('&quot;', '"')
                cell_text = re.sub(r'<[^>]+>', '', cell_text)
                cell_text = cell_text.strip()
            cells.append(cell_text)
        if cells:
            rows.append(cells)
    return rows


def find_all_sections(content):
    """Find all <h2> section headings in the page content (excluding Wayback Machine chrome)."""
    sections = []
    # Only look at content after the last wayback toolbar div
    body_start = content.find('<h2>')
    if body_start == -1:
        return sections
    # Find sections that have a table after them
    for match in re.finditer(r'<h2>(.*?)</h2>', content[body_start:]):
        heading = match.group(1).strip()
        heading = re.sub(r'<[^>]+>', '', heading)
        pos = body_start + match.start()
        # Check if there's a table following this heading
        table_start = content.find('<table', pos)
        next_h2 = content.find('<h2>', pos + 1)
        if table_start != -1 and (next_h2 == -1 or table_start < next_h2):
            sections.append(heading)
    return sections


def get_table_section(content, heading):
    """Extract HTML between a <h2>heading</h2> and its closing </table>."""
    start = content.find(f'<h2>{heading}</h2>')
    if start == -1:
        return ""
    table_start = content.find('<table', start)
    table_end = content.find('</table>', table_start)
    return content[table_start:table_end]


def get_footnotes_after_section(content, heading):
    """Get footnotes between a section's </table> and the next <h2>.

    Each section can have its own footnote div with independent #N numbering.
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


def expand_properties(props, footnotes):
    """Replace #N references with footnote text. Returns (clean_props, notes)."""
    if not props:
        return props, ''
    notes = []
    remaining = props
    for key in sorted(footnotes.keys(), key=lambda k: int(k[1:])):
        if key in remaining:
            notes.append(footnotes[key])
            remaining = remaining.replace(key, '').strip()
    return remaining, '; '.join(notes)


def normalize_cmd(cmd):
    """Normalize command for matching.

    - Strip leading whitespace (indentation for continuation depth)
    - Strip [~5] tag (with optional leading ' - ')
    - Replace < with , (TTT uses < for delays, T3 uses ,)
    """
    cmd = cmd.lstrip(' ')
    cmd = re.sub(r'\s*-?\s*\[~5\]', '', cmd).strip()
    cmd = cmd.replace('<', ',')
    return cmd


TEXT_COLUMNS = {'UUID', 'Character', 'Command', 'Alt Commands', 'Move Name', 'Stance', 'Damage',
                'Hit Range', 'Properties', 'Block Adv', 'Hit Adv', 'CH Adv', 'Notes'}


def expand_continuations(row_dicts):
    """Expand '= X' continuation commands and move names into full forms.

    Handles indentation levels:
    - ' = X' (1 leading space) = level 1, continues from parent (level 0)
    - '   = X' (3 leading spaces) = level 2, continues from preceding level 1

    Commands: '(WS+2_3~2)' → '(WS+2_3~2),4' → '(WS+2_3~2),4,4'
    Names: 'Tequila Sunrise' → 'Tequila Sunrise > Razor Sweep' → 'Tequila Sunrise > Razor Sweep > High Kick'
    """
    commands_by_level = {}
    names_by_level = {}
    for row in row_dicts:
        cmd = row['Command']
        name = row.get('Move Name', '')
        stripped = cmd.lstrip(' ')
        leading_spaces = len(cmd) - len(stripped)
        if stripped.startswith('= '):
            suffix = stripped[2:]
            level = 1 if leading_spaces <= 1 else 2
            parent_level = level - 1
            parent_cmd = commands_by_level.get(parent_level, '')
            separator = '' if suffix.startswith(('~', '<')) else ','
            full_cmd = parent_cmd + separator + suffix
            row['Command'] = full_cmd
            row['_is_followup'] = True
            commands_by_level[level] = full_cmd

            own_name = name.lstrip('= ') if name.startswith('= ') else name
            parent_name = names_by_level.get(parent_level, '')
            if parent_name and own_name:
                full_name = parent_name + ' > ' + own_name
            elif own_name:
                full_name = own_name
            else:
                full_name = parent_name
            row['Move Name'] = full_name
            names_by_level[level] = full_name
        else:
            commands_by_level = {0: cmd}
            names_by_level = {0: name}
            row['Command'] = cmd
    for row in row_dicts:
        row['Command'] = re.sub(r'\s*-?\s*\[~5\]', '', row['Command']).strip()
    split_alternatives(row_dicts)
    return row_dicts


PREFERRED_COMMAND = {
    'julia': {
        'WR+1': 'd,df+1',
    },
}

# Commands that are duplicates of each other. The value is the preferred primary;
# the other becomes an Alt Command and its row is removed.
MERGE_DUPLICATES = {
    'julia': [
        (['d+1', 'FC+1'], 'd+1'),
        (['d+2', 'FC+2'], 'd+2'),
        (['d+3', 'FC+3'], 'd+3'),
    ],
}


def split_alternatives(row_dicts):
    """Split (A_B) alternative notation into primary command and alternatives.

    E.g. '(WS+2_3~2),4' → Command='WS+2,4', Alt Commands='3~2,4'
    - Button 5 (TTT tag button) is removed from alternatives.
    - Follow-up moves don't get Alt Commands (only parent moves do).
    - PREFERRED_COMMAND overrides which alternative becomes the primary.
    """
    from itertools import product

    preferred_opts = {}
    for row in row_dicts:
        cmd = row['Command']
        is_followup = row.get('_is_followup', False)
        alt_groups = re.findall(r'\(([^)]*_[^)]*)\)', cmd)
        if not alt_groups:
            row['Alt Commands'] = ''
            continue

        parts = re.split(r'\([^)]*_[^)]*\)', cmd)
        groups = [[opt for opt in g.split('_') if opt != '5'] for g in alt_groups]

        chosen = []
        for group in groups:
            group_key = '_'.join(group)
            if group_key in preferred_opts:
                chosen.append(preferred_opts[group_key])
            else:
                chosen.append(group[0])

        primary = parts[0]
        for i, opt in enumerate(chosen):
            primary += opt + parts[i + 1]
        row['Command'] = primary

        if is_followup or all(len(g) == 1 for g in groups):
            row['Alt Commands'] = ''
            continue

        alts = []
        for combo in product(*groups):
            if list(combo) == chosen:
                continue
            alt = parts[0]
            for i, opt in enumerate(combo):
                alt += opt + parts[i + 1]
            alts.append(alt)

        char_prefs = PREFERRED_COMMAND.get(row.get('Character', '').lower(), {})
        if primary in char_prefs:
            preferred = char_prefs[primary]
            if preferred in alts:
                alts.remove(preferred)
                alts.insert(0, primary)
                primary = preferred
                row['Command'] = primary
                for i, group in enumerate(groups):
                    group_key = '_'.join(group)
                    for opt in group:
                        if opt in preferred and opt != chosen[i]:
                            preferred_opts[group_key] = opt
                            break
        else:
            for i, group in enumerate(groups):
                group_key = '_'.join(group)
                if group_key not in preferred_opts:
                    preferred_opts[group_key] = chosen[i]

        row['Alt Commands'] = '; '.join(alts)


def group_moves(rows):
    """Group moves into parent + follow-ups.

    Follow-up rows start with '=' (possibly with leading spaces for depth).
    Each parent starts a new group; follow-ups attach to the preceding parent.
    """
    groups = []
    for row in rows:
        cmd = row[0].lstrip(' ')
        if cmd.startswith('='):
            if groups:
                groups[-1].append(row)
        else:
            groups.append([row])
    return groups


def lcs(seq_a, seq_b):
    """LCS backtrack returning list of (i, j) matched index pairs."""
    n, m = len(seq_a), len(seq_b)
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            if seq_a[i - 1] == seq_b[j - 1]:
                dp[i][j] = dp[i - 1][j - 1] + 1
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])

    matches = []
    i, j = n, m
    while i > 0 and j > 0:
        if seq_a[i - 1] == seq_b[j - 1]:
            matches.append((i - 1, j - 1))
            i -= 1
            j -= 1
        elif dp[i - 1][j] >= dp[i][j - 1]:
            i -= 1
        else:
            j -= 1
    matches.reverse()
    return matches


def group_match(fd_rows, ml_rows):
    """Match frame data rows to movelist rows using group-based LCS.

    1. Group rows by parent move (non-'=' rows start groups).
    2. Match parent commands between FD and ML using LCS.
    3. Within each matched group, match follow-ups using LCS.

    This prevents generic follow-up commands like '= 4' from matching
    across unrelated parent moves.

    Returns dict: fd_row_index -> ml_row_index.
    """
    fd_groups = group_moves(fd_rows)
    ml_groups = group_moves(ml_rows)

    fd_parent_cmds = [normalize_cmd(g[0][0]) for g in fd_groups]
    ml_parent_cmds = [normalize_cmd(g[0][0]) for g in ml_groups]

    # LCS on parent commands
    parent_matches = lcs(fd_parent_cmds, ml_parent_cmds)
    group_map = dict(parent_matches)

    # Calculate row offsets for each group
    fd_offsets = []
    offset = 0
    for g in fd_groups:
        fd_offsets.append(offset)
        offset += len(g)

    ml_offsets = []
    offset = 0
    for g in ml_groups:
        ml_offsets.append(offset)
        offset += len(g)

    # Build row-level map
    row_map = {}
    for fd_gi, ml_gi in group_map.items():
        fd_g = fd_groups[fd_gi]
        ml_g = ml_groups[ml_gi]
        fd_base = fd_offsets[fd_gi]
        ml_base = ml_offsets[ml_gi]

        # Match parent row
        row_map[fd_base] = ml_base

        # Match follow-ups within the group
        fd_followups = [normalize_cmd(fd_g[i][0]) for i in range(1, len(fd_g))]
        ml_followups = [normalize_cmd(ml_g[i][0]) for i in range(1, len(ml_g))]

        if fd_followups and ml_followups:
            followup_matches = lcs(fd_followups, ml_followups)
            for fd_fi, ml_fi in followup_matches:
                row_map[fd_base + 1 + fd_fi] = ml_base + 1 + ml_fi

    return row_map


def merge_special_arts(fd_rows, ml_rows, footnotes):
    """Merge frame data and movelist Special Arts rows.

    Returns list of dicts with merged data. Unmatched rows get [UNMATCHED] in notes.
    """
    fd_to_ml = group_match(fd_rows, ml_rows)
    matched_count = len(fd_to_ml)

    merged = []
    for fd_idx, fd_row in enumerate(fd_rows):
        if fd_idx in fd_to_ml:
            ml_row = ml_rows[fd_to_ml[fd_idx]]
            raw_props = ml_row[5] if len(ml_row) > 5 else ''
            props, notes = expand_properties(raw_props, footnotes)
            merged.append({
                'command': fd_row[0],
                'move_name': ml_row[1] if len(ml_row) > 1 else '',
                'damage': ml_row[3] if len(ml_row) > 3 else '',
                'hit_range': ml_row[4] if len(ml_row) > 4 else '',
                'properties': props,
                'notes': notes,
                'speed': fd_row[1] if len(fd_row) > 1 else '',
                'block_adv': fd_row[2] if len(fd_row) > 2 else '',
                'hit_adv': fd_row[3] if len(fd_row) > 3 else '',
                'ch_adv': fd_row[4] if len(fd_row) > 4 else '',
            })
        else:
            merged.append({
                'command': fd_row[0],
                'move_name': '',
                'damage': '',
                'hit_range': '',
                'properties': '',
                'notes': '',
                'unmatched': True,
                'speed': fd_row[1] if len(fd_row) > 1 else '',
                'block_adv': fd_row[2] if len(fd_row) > 2 else '',
                'hit_adv': fd_row[3] if len(fd_row) > 3 else '',
                'ch_adv': fd_row[4] if len(fd_row) > 4 else '',
            })

    return merged, matched_count


UNIFIED_HEADERS = ['UUID', 'Character', 'Stance', 'Command', 'Move Name', 'Damage',
                    'Hit Range', 'Properties', 'Speed', 'Block Adv', 'Hit Adv', 'CH Adv',
                    'Alt Commands', 'Notes', 'Unmatched']


def merge_duplicate_commands(row_dicts, character):
    """Merge rows listed in MERGE_DUPLICATES: keep the preferred, add others as Alt Commands."""
    merges = MERGE_DUPLICATES.get(character.lower(), [])
    if not merges:
        return row_dicts
    for commands, primary in merges:
        alts = [c for c in commands if c != primary]
        primary_row = None
        alt_indices = []
        for i, row in enumerate(row_dicts):
            if row['Command'] == primary:
                primary_row = i
            elif row['Command'] in alts:
                alt_indices.append(i)
        if primary_row is not None and alt_indices:
            existing_alt = row_dicts[primary_row].get('Alt Commands', '')
            new_alts = '; '.join(row_dicts[i]['Command'] for i in alt_indices)
            row_dicts[primary_row]['Alt Commands'] = '; '.join(filter(None, [existing_alt, new_alts]))
            for i in sorted(alt_indices, reverse=True):
                row_dicts.pop(i)
    return row_dicts


def write_section_header(ws, row_num, heading):
    """Write a section heading row in bold."""
    cell = ws.cell(row=row_num, column=1, value=heading)
    cell.font = Font(bold=True)
    return row_num + 1


def write_column_headers(ws, row_num):
    """Write unified column headers."""
    for col, header in enumerate(UNIFIED_HEADERS, 1):
        cell = ws.cell(row=row_num, column=col, value=header)
        cell.font = Font(bold=True)
    return row_num + 1


def write_unified_row_xlsx(ws, row_num, row_dict):
    """Write a single data row. Text columns get explicit text format and data_type='s'."""
    row_dict['UUID'] = str(uuid.uuid4())
    for key in ('Command', 'Alt Commands'):
        if row_dict.get(key):
            row_dict[key] = row_dict[key].replace('/', '')
    for col, h in enumerate(UNIFIED_HEADERS, 1):
        value = row_dict.get(h, '')
        cell = ws.cell(row=row_num, column=col, value=value)
        if h in TEXT_COLUMNS:
            cell.data_type = 's'
            cell.number_format = numbers.FORMAT_TEXT
    return row_num + 1


def write_fd_only_section_xlsx(ws, row_num, heading, rows, stance='Default', character=''):
    """Write a frame-data-only section.

    FD columns are: Command, Hit (=Speed), Block Adv, Hit Adv, CH Adv.
    """
    row_num = write_section_header(ws, row_num, heading)
    if not rows:
        return row_num + 1
    row_num = write_column_headers(ws, row_num)
    row_dicts = []
    for row in rows[1:]:  # skip source header
        row_dicts.append({
            'Character': character,
            'Command': row[0] if len(row) > 0 else '',
            'Stance': stance,
            'Speed': row[1] if len(row) > 1 else '',
            'Block Adv': row[2] if len(row) > 2 else '',
            'Hit Adv': row[3] if len(row) > 3 else '',
            'CH Adv': row[4] if len(row) > 4 else '',
        })
    expand_continuations(row_dicts)
    row_dicts = merge_duplicate_commands(row_dicts, character)
    for row_dict in row_dicts:
        row_num = write_unified_row_xlsx(ws, row_num, row_dict)
    return row_num + 1  # blank row


def main():
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)

    framedata_url = sys.argv[1]
    movelist_url = sys.argv[2]
    output_path = sys.argv[3]

    # Extract character name from URL id parameter
    id_match = re.search(r'[?&]id=(\w+)', framedata_url)
    character_name = id_match.group(1).capitalize() if id_match else 'Unknown'

    print("Fetching frame data page...")
    framedata_content = fetch_url(framedata_url)

    print("Fetching movelist page...")
    movelist_content = fetch_url(movelist_url)

    # Discover sections in both sources
    fd_sections = find_all_sections(framedata_content)
    ml_sections = find_all_sections(movelist_content)
    print(f"Frame data sections: {fd_sections}")
    print(f"Movelist sections: {ml_sections}")

    # Parse frame data tables
    fd_tables = {}
    for section in fd_sections:
        rows = parse_table(get_table_section(framedata_content, section))
        fd_tables[section] = rows

    # Parse movelist tables
    ml_tables = {}
    for section in ml_sections:
        rows = parse_table(get_table_section(movelist_content, section))
        ml_tables[section] = rows

    # Parse footnotes per movelist section
    ml_footnotes = {}
    for section in ml_sections:
        ml_footnotes[section] = get_footnotes_after_section(movelist_content, section)

    total_fn = sum(len(v) for v in ml_footnotes.values())
    print(f"Found {total_fn} footnotes total")

    # --- Determine stance names from section names ---
    STANCE_MAP = {
        'Devil Jin Possession Arts': 'Devil Jin Possession',
    }

    # --- Write XLSX ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Frame Data"
    row_num = 1

    # 1. Basic Arts (FD only, no merge needed)
    if 'Basic Arts' in fd_tables:
        row_num = write_fd_only_section_xlsx(ws, row_num, 'BASIC ARTS', fd_tables['Basic Arts'],
                                             character=character_name)

    # 2. Special Arts (merged FD + ML)
    if 'Special Arts' in fd_tables:
        fd_special = fd_tables['Special Arts'][1:]
        ml_special = ml_tables.get('Special Arts', [[]])[1:] if 'Special Arts' in ml_tables else []
        fn_special = ml_footnotes.get('Special Arts', {})

        merged, matched_count = merge_special_arts(fd_special, ml_special, fn_special)
        print(f"Special Arts: matched {matched_count}/{len(fd_special)}")

        unmatched = [r['command'] for r in merged if r.get('unmatched')]
        if unmatched:
            print(f"  Unmatched: {unmatched}")

        row_num = write_section_header(ws, row_num, 'SPECIAL ARTS')
        row_num = write_column_headers(ws, row_num)
        special_rows = []
        for r in merged:
            special_rows.append({
                'Character': character_name,
                'Command': r['command'],
                'Move Name': r['move_name'],
                'Stance': 'Default',
                'Damage': r['damage'],
                'Hit Range': r['hit_range'],
                'Properties': r['properties'],
                'Speed': r['speed'],
                'Block Adv': r['block_adv'],
                'Hit Adv': r['hit_adv'],
                'CH Adv': r['ch_adv'],
                'Notes': r['notes'],
                'Unmatched': 'TRUE' if r.get('unmatched') else '',
            })
        expand_continuations(special_rows)
        for row_dict in special_rows:
            row_num = write_unified_row_xlsx(ws, row_num, row_dict)
        row_num += 1  # blank row

    # 3. Any additional FD-only sections (Devil Jin Possession Arts, etc.)
    fd_only_sections = [s for s in fd_sections
                       if s not in ('Basic Arts', 'Special Arts', 'Grappling Arts', 'Unblockable Arts')]
    for section in fd_only_sections:
        stance = STANCE_MAP.get(section, section.replace(' Arts', ''))
        row_num = write_fd_only_section_xlsx(ws, row_num, section.upper(), fd_tables[section],
                                             stance=stance, character=character_name)

    # 4. Unblockable Arts
    if 'Unblockable Arts' in fd_tables and 'Unblockable Arts' in ml_tables:
        fd_unblock = fd_tables['Unblockable Arts'][1:]
        ml_unblock = ml_tables['Unblockable Arts'][1:]
        fn_unblock = ml_footnotes.get('Unblockable Arts', {})

        merged_ub, ub_matched = merge_special_arts(fd_unblock, ml_unblock, fn_unblock)
        print(f"Unblockable Arts: matched {ub_matched}/{len(fd_unblock)}")

        row_num = write_section_header(ws, row_num, 'UNBLOCKABLE ARTS')
        row_num = write_column_headers(ws, row_num)
        ub_rows = []
        for r in merged_ub:
            ub_rows.append({
                'Character': character_name,
                'Command': r['command'],
                'Move Name': r['move_name'],
                'Stance': 'Default',
                'Damage': r['damage'],
                'Hit Range': r['hit_range'],
                'Properties': r['properties'],
                'Speed': r['speed'],
                'Block Adv': r['block_adv'],
                'Hit Adv': r['hit_adv'],
                'CH Adv': r['ch_adv'],
                'Notes': r['notes'],
                'Unmatched': 'TRUE' if r.get('unmatched') else '',
            })
        expand_continuations(ub_rows)
        for row_dict in ub_rows:
            row_num = write_unified_row_xlsx(ws, row_num, row_dict)
        row_num += 1
    elif 'Unblockable Arts' in ml_tables:
        fn_unblock = ml_footnotes.get('Unblockable Arts', {})
        row_num = write_section_header(ws, row_num, 'UNBLOCKABLE ARTS')
        row_num = write_column_headers(ws, row_num)
        for row in ml_tables['Unblockable Arts'][1:]:
            raw_props = row[5] if len(row) > 5 else ''
            props, notes = expand_properties(raw_props, fn_unblock)
            row_dict = {
                'Character': character_name,
                'Command': row[0] if len(row) > 0 else '',
                'Move Name': row[1] if len(row) > 1 else '',
                'Stance': 'Default',
                'Damage': row[3] if len(row) > 3 else '',
                'Hit Range': row[4] if len(row) > 4 else '',
                'Properties': props,
                'Notes': notes,
            }
            row_num = write_unified_row_xlsx(ws, row_num, row_dict)
        row_num += 1
    elif 'Unblockable Arts' in fd_tables:
        row_num = write_fd_only_section_xlsx(ws, row_num, 'UNBLOCKABLE ARTS', fd_tables['Unblockable Arts'],
                                             character=character_name)

    # 5. Grappling Arts (ML is primary source, merge FD speed if available)
    if 'Grappling Arts' in ml_tables:
        fn_grappling = ml_footnotes.get('Grappling Arts', {})
        fd_grappling = fd_tables.get('Grappling Arts', [[]])[1:] if 'Grappling Arts' in fd_tables else []

        fd_grappling_speed = {}
        for row in fd_grappling:
            if len(row) >= 2:
                fd_grappling_speed[row[0]] = row[1]

        row_num = write_section_header(ws, row_num, 'GRAPPLING ARTS')
        row_num = write_column_headers(ws, row_num)
        for row in ml_tables['Grappling Arts'][1:]:
            raw_props = row[5] if len(row) > 5 else ''
            props, notes = expand_properties(raw_props, fn_grappling)
            escape_cmd = row[4] if len(row) > 4 else ''
            row_dict = {
                'Character': character_name,
                'Command': row[0] if len(row) > 0 else '',
                'Move Name': row[1] if len(row) > 1 else '',
                'Stance': 'Default',
                'Damage': row[3] if len(row) > 3 else '',
                'Hit Range': row[2] if len(row) > 2 else '',
                'Properties': escape_cmd,
                'Speed': fd_grappling_speed.get(row[0], '') if len(row) > 0 else '',
                'Notes': '; '.join(filter(None, [props, notes])),
            }
            row_num = write_unified_row_xlsx(ws, row_num, row_dict)
        row_num += 1

    # 6. String Hit Arts (ML only)
    if 'String Hit Arts' in ml_tables:
        row_num = write_section_header(ws, row_num, 'STRING HIT ARTS')
        row_num = write_column_headers(ws, row_num)
        for row in ml_tables['String Hit Arts'][1:]:
            row_dict = {
                'Character': character_name,
                'Command': row[0] if len(row) > 0 else '',
                'Stance': 'Default',
                'Damage': row[2] if len(row) > 2 else '',
                'Hit Range': row[3] if len(row) > 3 else '',
                'Notes': f"{row[1]} hits" if len(row) > 1 and row[1] else '',
            }
            row_num = write_unified_row_xlsx(ws, row_num, row_dict)
        row_num += 1

    # Auto-fit all column widths
    for col_idx in range(1, len(UNIFIED_HEADERS) + 1):
        max_len = 0
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx, values_only=True):
            val = row[0]
            if val:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = max_len + 2

    wb.save(output_path)
    print(f"\nWritten to {output_path}")
    for section in fd_sections:
        count = len(fd_tables[section]) - 1
        print(f"  FD {section}: {count} rows")
    for section in ml_sections:
        count = len(ml_tables[section]) - 1
        print(f"  ML {section}: {count} rows")


if __name__ == '__main__':
    main()
