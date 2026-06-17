# Step 9: Split multi-hit rows — Design Notes

## Goal

Split rows that contain multiple hits (multiple Block Adv values) into separate rows, one per press.

## Input

`sources/<character>_step8c.xlsx` — output of step 8c.

## How splitting works

### Determining hit count

Use **Block Adv Full** column. Number of space-separated values = number of hits.
If only 1 value, row is not split.

### Press boundaries in Command Full

Step 8c has already inserted spaces between presses in Command Full.
Example: `1 <2 ,4 ~1+4 ,2` (5 presses, space-separated tokens).

Each space-separated token becomes a progressive command:
- `1`
- `1 <2`
- `1 <2 ,4`
- `1 <2 ,4 ~1+4`
- `1 <2 ,4 ~1+4 ,2`

### Progressive (cumulative) columns

All "Full" columns are progressive — each split row shows the values from the start up to that press:

- **Block Adv Full**: space-separated → cumulative. `+3 -5 -7` → `+3`, `+3 -5`, `+3 -5 -7`
- **Hit Adv Full**: same as above
- **Counter Hit Adv Full**: same as above
- **Damage Full**: comma-separated → cumulative. `16,12,25` → `16`, `16,12`, `16,12,25`
- **Hit Range Full**: one character per hit → cumulative. `hhm` → `h`, `hh`, `hhm`
- **Command Full**: progressive commands (see above)

### Move Name Full splitting

Move Name Full uses two separators:
- **` – `** (endash) — added by step 7 for continuation joins
- **` - `** (hyphen) — original source data separating press names within a row

Both represent press boundaries. Split progressively, replacing all separators with endash:
- `Club Fist – Bow - Arrow` → `Club Fist`, `Club Fist – Bow`, `Club Fist – Bow – Arrow`

When all presses share the same name (e.g., "Spinning Demon" for `uf+4,4,4,4`), add ordinal suffixes:
- `Spinning Demon (First)`, `Spinning Demon (Second)`, `Spinning Demon (Third)`, `Spinning Demon (Fourth)`

### Duplicate handling

A split row's uniqueness is determined by **Command Full + From Stance**.

If a split produces a row that already exists:
- **Same frame data** → silently skip (don't add the row)
- **Different frame data** → keep the row but add a note to Processor Notes

### Rows that should NOT be split

These have multiple Block Adv values but only one press (single button, multi-hit):
- `jin: d+3+4`
- `ling: d+1`, `D+1`, `u+1+2`, `SS 4 [~B]`

### Rows to ignore (invalid data)

- `ling: FC,DF+2`

### Known issues requiring manual resolution (in TODO.md)

1. **Jin `1<2,4~1+4,2`** and continuations — continuation chain where Block Adv Full
   includes parent data but the command structure makes the split ambiguous.
2. **Julia `1~2~1`** — tilde chain with more presses than block values.

### Multi-hit presses (presses < block count)

These are cases where one button press produces multiple hits:
- **`1+4`** (jin, julia): produces 2 hits (first hit is unblockable `x`)
- **`u+1+2`** (ling): produces 2 hits (first hit is `x`)

Affected commands:
- jin: `1+4,2` / `1+4,2,4` / `1+4,2,d+4`
- julia: `1+4,3`
- ling: `u+1+2,2` / `u+1+2,2,1` / `u+1+2,3+4`

In all cases the extra hit is marked `x` in Block Adv. These need the `x` values
handled specially during split — either by grouping with the press that made them
or by creating a sub-row for the `x` hit.

## Observations from step8c data

### Damage Full does not always have one value per press

Damage Full is built at continuation level (step 7), not press level.
Examples:

- `1 <2` (2 presses): damage=`16` — single sum for the whole row
- `1 <2 ,4` (3 presses): damage=`16,12` — 2 values (first is parent's sum)
- `b,f+2 <1 <2` (3 presses): damage=`56` — single sum
- `df+1 ,2` (2 presses): damage=`26` — single sum
- `uf+4 ,4 ,4 ,4` (4 presses): damage=`77` — single sum

When Damage Full has fewer values than presses, the values correspond to
continuation boundaries, not press boundaries. Individual press damage values
exist in the **Damage** column (comma-separated per row's own hits).

### Damage column has per-hit values for the current row

- `1 <2`: Damage=`6,10`, Damage Sum=`16`
- `b,f+2 <1 <2`: Damage=`18,14,24`, Damage Sum=`56`
- `uf+4 ,4 ,4 ,4`: Damage=`10,17,25,25`, Damage Sum=`77`

These per-hit values in the Damage column map 1:1 to presses within that row.

### Strategy for progressive Damage Full

For press N in a split row:
1. Count which presses belong to the current continuation row (vs parent data from Damage Full)
2. Use the Damage column's comma-separated values for the current row's presses
3. Build cumulative: parent damage + press 1 damage, parent damage + press 1 + press 2, etc.

Example `b+1~df ,1 ,3` (Damage Full=`15,-,29,20`, Damage=`20`):
- Press 1 (b+1~df): `15,-`
- Press 2 (,1): `15,-,29`
- Press 3 (,3): `15,-,29,20`

Example `1 <2 ,4` (Damage Full=`16,12`, Damage=`12`):
- Press 1 (1): `6` (from parent `1 <2` Damage col first hit)
- Press 2 (<2): `6,10` = `16` (parent sum)
- Press 3 (,4): `16,12`

This requires looking up parent rows' Damage column for per-press breakdown.

### Hit Range Full matches hit count (not press count)

Hit Range Full always has one character per hit (block adv value), including `x` hits
and `-` markers:
- `hhmshh` = 6 chars for block `-3 -5 -7 x -6 0` (6 values)
- `m-m` = 3 chars for block `-2 -14` (2 values) — the `-` is a stance transition marker

### Move Name Full separators in step8c data

Examples from the data:
- `Double Punch – Knee – White Heron` (endash separators from step 7 continuations)
- `Twin Pistons -` (trailing hyphen = stance transition, NOT a name separator)
- `Parting Wave – Crouch Dash – Thunder Godfist – Mid Kick`
- `Shoot the Works` (single name for 3 presses: `1 <2 ,3`)
- `Spinning Demon` (single name for 4 presses: `uf+4 ,4 ,4 ,4`)

Name parts from step 7 endash splits correspond to continuation boundaries,
not individual presses within a row.

### Stance prefix handling

Commands starting with `WS`, `SS`, `FC`, `RDS` have the prefix as part of the
first press token: `WS 1 ,2 -` → tokens are `WS 1` and `,2 -`.

## Validation

Run `python3 tools/validate_split.py` to check current state.

## Command notation reference

- `,` between presses = sequential button presses (e.g., `1,2` = press 1 then press 2)
- `,` between directions = motion input (e.g., `f,N,d,df` = crouch dash motion)
- `<` = delay (press button with timing)
- `~` before button = fast transition/just-frame
- `~` before direction (e.g., `~df`) = stance transition (NOT a separate press)
- `+` between buttons = simultaneous press (e.g., `1+4` = one press)
- WS/SS/FC/RDS = stance prefix (part of same press, separated by space in original notation)
- Trailing ` -` or ` - [X]` = move transitions to stance
