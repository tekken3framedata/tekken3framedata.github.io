#!/bin/bash
set -e

python3 step_1_html_to_xlsx.py
python3 step_2_parent_uuid.py
python3 step_3_merge.py
python3 step_4_alternatives.py
python3 step_5_notation.py
python3 step_5b_hit_markers.py
python3 step_6_prepare_for_expansion.py
python3 step_6b_bracket_notations.py
python3 step_7_expand_continuations.py
python3 step_8_extract_tags.py
python3 step_8b_from_stance.py
python3 step_8c_press_boundaries.py
python3 step_9_manual_corrections.py
python3 step_10_hits_per_move.py
