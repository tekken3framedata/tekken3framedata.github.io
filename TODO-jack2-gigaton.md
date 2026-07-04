# Jack-2 Gigaton / Wind Up -korjaussuunnitelma

## Nykytila

Duplikaattirivejä Special Arts ja Unblockable Arts -sectioissa. Debugger on sekä itsenäisenä liikkeenä että Wind Up -continuationina. Wind Up esiintyy kahdella nimellä (Wind Up / Gigaton Start Up).

## Haluttu lopputulos

| Command | Command Full | Alt Commands | Alt Commands Full | Move Name | Damage | Hit Range | Notes |
|---------|-------------|-------------|------------------|-----------|--------|-----------|-------|
| `b,db,d,DF+1` | `b,db,d,DF+1` | `hcf,DF+1` | `hcf,DF+1` | Debugger | 25 | L | |
| `hcf` | `hcf` | `b,b+1` | `b,b+1` | Wind Up (1 Rotation) | - | - | |
| `– 1` | `hcf,1` | | `b,b+1,1` | – Gigaton Punch (1 Rotation) | 20 | m | |
| `– ccw,db,d` | `ccw,db,d` | | | – Wind Up (2 Rotations) | - | - | With b,b+1 input, full 2nd rotation command is b,b+1,b,db,d,df |
| `–– 1` | `ccw,db,d,1` | | | –– Gigaton Punch (2 Rotations) | 40 | ! | |
| `–– ccw,db,d` | `ccw,ccw,db,d` | | | –– Wind Up (3 Rotations) | - | - | With b,b+1 input, full 3rd rotation command is b,b+1,ccw,b,db,d |
| `––– 1` | `ccw,ccw,db,d,1` | | | ––– Gigaton Punch (3 Rotations) | 60 | ! | |
| `––– ccw,db,d` | `ccw,ccw,ccw,db,d` | | | ––– Wind Up (4 Rotations) | - | - | |
| `–––– 1` | `ccw,ccw,ccw,db,d,1` | | | –––– Gigaton Punch (4 Rotations) | 80 | ! | |
| `–––– ccw,db,d` | `ccw,ccw,ccw,ccw,db,d` | | | –––– Wind Up (5 Rotations) | - | - | |
| `––––– 1` | `ccw,ccw,ccw,ccw,db,d,1` | | | ––––– Gigaton Punch (5 Rotations) | 199 | ! | |
| `––––– ccw,db,d` | `ccw,ccw,ccw,ccw,ccw,db,d` | | | ––––– Wind Up (8 Rotations) | - | - | Completes 8 rotations automatically. No Gigaton Punch follow-up. |

## Notaatio

- `ccw` = `hcf,uf,u,ub,b` (täysi ympyrä b:stä b:hen)
- `hcf` = `b,db,d,df,f` (puoliympyrä eteenpäin)

## Poistettavat rivit

- `hcf,DF+1` Wind Up – Debugger → `DF+1` siirtyy Debuggerin alt commandiksi
- `hcf` Gigaton Start Up → duplikaatti, poistetaan kokonaan

## Toteutus

Korjaukset tehdään step 6:ssa (prepare for expansion) jotta step 7 rakentaa Command Full ja Move Name Full oikein.
