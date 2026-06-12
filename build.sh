#!/bin/bash
set -e

python3 fetch_sources.py \
    "https://web.archive.org/web/20201206043428/http://www.tekkenzaibatsu.com/tekkentag/framedata.php?id=ling" \
    "https://web.archive.org/web/20201206043457/http://www.tekkenzaibatsu.com/tekkentag/movelist.php?id=ling" \
    ling

python3 fetch_sources.py \
    "https://web.archive.org/web/20201206043428/http://www.tekkenzaibatsu.com/tekkentag/framedata.php?id=jin" \
    "https://web.archive.org/web/20201206043458/http://www.tekkenzaibatsu.com/tekkentag/movelist.php?id=jin" \
    jin

python3 fetch_sources.py \
    "https://web.archive.org/web/20201206043428/http://www.tekkenzaibatsu.com/tekkentag/framedata.php?id=julia" \
    "https://web.archive.org/web/20201206043457/http://www.tekkenzaibatsu.com/tekkentag/movelist.php?id=julia" \
    julia

#python3 scrape_framedata.py julia julia_framedata_v4.xlsx
#python3 scrape_framedata.py jin jin_framedata_v4.xlsx
#python3 scrape_framedata.py ling ling_framedata_v4.xlsx
#python3 xlsx_to_tsv.py
