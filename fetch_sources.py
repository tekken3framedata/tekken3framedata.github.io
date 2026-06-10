#!/usr/bin/env python3
"""
Fetch Tekken frame data HTML sources from Wayback Machine and save locally.
Adds a data-uuid attribute to each <tr> element in tables for stable identification.

Usage:
    python3 fetch_sources.py <framedata_url> <movelist_url> <character>

Example:
    python3 fetch_sources.py \
        "https://web.archive.org/web/20201206043428/http://www.tekkenzaibatsu.com/tekkentag/framedata.php?id=julia" \
        "https://web.archive.org/web/20201206042940/http://www.tekkenzaibatsu.com/tekken3/movelist.php?id=julia" \
        julia
"""

import re
import sys
import subprocess
import uuid


def fetch_url(url):
    """Fetch URL content using curl."""
    result = subprocess.run(
        ['curl', '-s', '-L', url],
        capture_output=True, text=True, timeout=60
    )
    return result.stdout


def add_uuids_to_rows(html):
    """Add data-uuid attribute to each <tr> that doesn't already have one."""
    def replace_tr(match):
        tag = match.group(0)
        if 'data-uuid=' in tag:
            return tag
        return tag[:-1] + f' data-uuid="{uuid.uuid4()}">'

    return re.sub(r'<tr[^>]*>', replace_tr, html)


def main():
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)

    framedata_url = sys.argv[1]
    movelist_url = sys.argv[2]
    character = sys.argv[3].lower()

    print(f"Fetching frame data for {character}...")
    fd_html = fetch_url(framedata_url)
    fd_html = add_uuids_to_rows(fd_html)

    fd_path = f"sources/{character}_framedata.html"
    with open(fd_path, 'w', encoding='utf-8') as f:
        f.write(fd_html)
    print(f"  Saved {fd_path}")

    print(f"Fetching movelist for {character}...")
    ml_html = fetch_url(movelist_url)
    ml_html = add_uuids_to_rows(ml_html)

    ml_path = f"sources/{character}_movelist.html"
    with open(ml_path, 'w', encoding='utf-8') as f:
        f.write(ml_html)
    print(f"  Saved {ml_path}")

    print("Done.")


if __name__ == '__main__':
    main()
