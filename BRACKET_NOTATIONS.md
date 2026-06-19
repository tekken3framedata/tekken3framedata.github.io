# Bracket notations in frame data

Found 21 unique bracket notations across 32 characters.
Script: `python3 tools/list_bracket_notations.py`

## Processing status

### Done

| Notation | Handling | Step | Details |
|----------|----------|------|---------|
| `[~5]` | Taggable=TRUE, strip | step 6b | All 32 characters, 210 occurrences |
| `[~3]` | Strip, To Stance = `(~3)HMS` | step 6b | Lee only (3 rows). UUID-based patch in `STANCE_RECOVERY` dict |
| `[~4]` | Row split into two moves | step 6 | Lee only (1 row). `b+3~3 [~4]` → `b+3~3` + `b+3~3:4` (Mist Trap). UUID-based `replace` in `CORRECTIONS` dict |

### Not yet processed (18 remaining)

## To Stance syntax (decided)

To Stance -sarake käyttää tätä syntaksia:
- **Ehdoton:** `HMS` — liike johtaa aina tähän stanceen
- **Vaihtoehtoinen:** `(~3)HMS` — painamalla `~3` → HMS
- **Useita:** `(~3)HMS;(~D)FC` — puolipisteellä erotettu

## Category 1: Taggable — DONE

| Notation | Occurrences | Characters | Meaning |
|----------|-------------|------------|---------|
| `[~5]` | 210 | 32 (all) | Recovery crouch 5f, taggable |

Handled in step 6b: Taggable=TRUE, stripped from Command. `[Tag]` stripped from Move Name.

## Category 2: Lee-specific — DONE

| Notation | Occurrences | Meaning | Handling |
|----------|-------------|---------|----------|
| `[~3]` | 3 | Alternative stance recovery → HMS | Step 6b: strip, write `(~3)HMS` to To Stance |
| `[~4]` | 1 | Just-frame follow-up (`:4` = Mist Trap) | Step 6: split into own row with Parent UUID |

`[~3]` examples (all have `[Hit Man Stance]` in Move Name, stripped):
- `b+4 [~3] [~5]` → Taggable + To Stance `(~3)HMS`
- `b,b+4 [~3]` → To Stance `(~3)HMS`
- `D+4<4<4<4 [~3]` → To Stance `(~3)HMS`

`[~4]` resolved as:
- Original: `b+3~3 [~4]` / Feint Mist Wolf [Mist Trap] / Speed "14 x" / Block "-8 x" / Damage "18 [33]"
- Row 1: `b+3~3` / Feint Mist Wolf / Speed 14 / Block -8 / Damage 18 / Hits Per Move: `0 1`
- Row 2: `b+3~3:4` / Mist Trap / Damage 33 / Parent UUID → row 1
- NOTE: `b+3~3` framedata is likely wrong (TODO.md)

## Category 3: Stance transition (tilde + direction) — TODO

| Notation | Occurrences | Characters | Meaning |
|----------|-------------|------------|---------|
| `[~B]` | 24 | baek, eddy, lei, ling | Recovers backturned |
| `[~D]` | 10 | eddy, forest, ling | Recovers in crouch/down |
| `[~F]` | 8 | lei | Recovers face-down (uppercase) |
| `[~f]` | 4 | lei | Recovers face-down (lowercase) |
| `[~db]` | 1 | eddy | Recovers down-back |

Examples:
- `baek: b+1+2 [~B]`
- `eddy: f,f+3 [~B] [~D]` — two stance markers on one move
- `eddy: df+3 [~D]`
- `lei: – 1 [~F]`
- `lei: b+1 [~f]`
- `eddy: f+1+2 [~db]`

Proposed handling: strip from Command, write to To Stance with `(~input)STANCE` syntax.

## Category 4: Combined recovery/stance (underscore-separated) — TODO

| Notation | Occurrences | Characters | Meaning |
|----------|-------------|------------|---------|
| `[~U_~D]` | 9 | anna, nina | Recovers up then down |
| `[~B_~3]` | 2 | eddy | BT + recovery 3f |
| `[~B_~D]` | 1 | eddy | BT + crouch |
| `[~D_~B]` | 1 | eddy | Crouch + BT |
| `[~d_~u]` | 2 | lei | Down then up |
| `[~f_~b]` | 5 | baek | Face-down then BT |

Examples:
- `anna: – 1,4,2 [~U_~D]`
- `eddy: 3+4 [~B_~3]`
- `eddy: ––– uf+3+4 [~B_~D]`
- `eddy: u+3+4 [~D_~B]`
- `lei: SS+1 [~f] [~d_~u]`
- `baek: – 3 [~f_~b]`

Underscore likely means "then" — two-part recovery sequence or two stance options.

## Category 5: Stance transition (no tilde, active transition) — TODO

| Notation | Occurrences | Characters | Meaning |
|----------|-------------|------------|---------|
| `[u_d]` | 16 | lei, ling | Up/down transition |
| `[f_b]` | 7 | baek | Forward/back transition |
| `[F]` | 3 | hwoarang | Flamingo stance |
| `[d]` | 6 | lei | Crouch/down |
| `[tag]` | 1 | kunimitsu | Tag move |
| `[b+4]` | 1 | eddy | Specific input transition (hold?) |

Examples:
- `lei: f,N+1 [u_d]`
- `baek: 3+4 [f_b]`
- `hwoarang: 3,3,d+3,4 [F]`
- `lei: 3~4 [d]`
- `kunimitsu: 1+3 [tag]`
- `eddy: db+4,4,4,4,... [b+4]`

No-tilde notations differ from tilde ones — possibly "transitions TO stance" vs tilde "RECOVERS IN stance". Need to determine exact semantics.

## Open questions

1. **`[~F]` vs `[~f]`:** Different stances or case typo? Both are Lei only.
2. **`[u_d]` and `[f_b]`:** "Can continue in either direction" (player choice) or "transitions through both"?
3. **Multiple brackets on one move:** `[~B] [~D]`, `[~f] [~d_~u]` — semicolon-joined in To Stance? Or different semantics?
4. **`[b+4]`:** Eddy `db+4,4,4,4,... [b+4]` — what does this mean mechanically? Hold b+4?
5. **`[tag]`:** Kunimitsu `1+3 [tag]` — is this just Taggable by another name, or different?
6. **Stance abbreviations needed:** What abbreviation for each stance?
   - `[~B]` → BT (backturned)? Already used in data.
   - `[~D]` → FC? Or specific crouch stance (RLX for eddy)?
   - `[~F]` / `[~f]` → FCD? (face-down)
   - `[~db]` → ?
   - `[u_d]` → ?
   - `[f_b]` → ?
   - `[F]` → FLA (flamingo, hwoarang)?
   - `[d]` → ?

## Code changes made

### step_8c_press_boundaries.py
- Added `:` as a press boundary separator (always separates, like `<`)
- Added `:` to `next_sep` regex so tilde/comma logic sees it as token boundary
- Result: `b+3~3:4` → `b+3 ~3 :4` (3 presses)

### step_6b_bracket_notations.py
- Added `STANCE_RECOVERY` dict for UUID-based alternative stance processing
- `process_sections()` now takes `char` parameter and applies stance rules
- Lee UUIDs: strips `[~3]` from Command, `[Hit Man Stance]` from Move Name, writes `(~3)HMS` to To Stance

### step_6_prepare_for_expansion.py
- Added `lee` entry in `CORRECTIONS` with `replace` operation
- UUID `a1226d52` split into two rows: `b+3~3` (Feint Mist Wolf) + `b+3~3:4` (Mist Trap)

### step_10_hits_per_move.py
- Added Lee `a1226d52`: `{0: 0}` — first press of `b+3 ~3` produces 0 hits (feint)

### Command notation
- `:` = just-frame input separator. Always a press boundary. Example: `b+3~3:4` = 3 presses.
