# Cancellable Moves

Moves whose attack can be fully aborted by the player before (or instead of) the hit coming out.

## Detection methods

1. **Move name** — follow-up move with "Cancel" in its name (linked via Parent UUID)
2. **Frame data signature** — Hit Range Full ends with `-` and Damage Full ends with `,-` (the cancelled press produces no hit and no damage)

Data source: `sources/<character>_step10.xlsx`

## Katso läpi

| Character | UUID | From | Command | Move Name | Cancel Input | Cancel To | Source |
|-----------|------|------|---------|-----------|-------------|-----------|--------|
| Eddy | 5b499286 | | FC+1+2 | Crying Needle | ~1 (Needle Cancel) | standing | name |
| Eddy | 27918d65 | | db+3,3 | Weed Whacker – Hot Plate | ~D / ~B | RLX / HSP | name |
| Eddy | f24e245a | | b+3 | Hot Plate | ~D / ~B | RLX / HSP | name |
| Eddy | 9b003da6 | HSP | 3 | Hot Plate | ~D / ~B | RLX / HSP | name |
| Eddy | 23a5f617 | | b+4,3,3 | Leg Whip – Weed Whacker – Hot Plate | ~D / ~B | RLX / HSP | name |
| Eddy | d797e17f | | db+3~4,1+2 | Weed Whacker – Shin Cutter – Crying Needle | ~1 (Needle Cancel) | standing | name |
| Eddy | b64dc694 | | 3~4,4,1+2 | Slippery Kick – Slider – Crying Needle | ~1 (Needle Cancel) | standing | name |
| Eddy | c89588be | | b+4,3~4 | Leg Whip – Shin Cutter | ~1 (Needle Cancel) | standing | name |
| Eddy | 522a8f03 | RLX | 1+2 | Crying Needle | ~1 (Needle Cancel) | standing | name |
| Eddy | 7adca466 | HSP | db+3+4 | Fruit Picker | ,b,b | standing | name |
| Eddy | 6345afef | | 1,2,4 | Double Punch – Black Summy | | | frame data (hh-) |
| Eddy | f36b0997 | | db+3,3,B | Weed Whacker – Hot Plate – Handstand Position | ,B | HSP | frame data (mm-) |
| Eddy | 478ef843 | | b+4,3,3,B | Leg Whip – Weed Whacker – Hot Plate – Handstand Position | ,B | HSP | frame data (hLm-) |
| Eddy | 5198b605 | | b+3,B | Hot Plate – Handstand Position | ,B | HSP | frame data (m-) |
| Eddy | b8a0c0e0 | HSP | 3,B | Hot Plate – Handstand Position | ,B | HSP | frame data (m-) |
| Ganryu | 91ee31e2 | | 3+4~D,u | Sumo Sit – Splits – Splits Side Step | ~D,u | | frame data (---) |
| Ganryu | 91dfeae5 | | uf+3+4,F | Sumo Squash – Sumo Sit – Roll Forward | ,F | | frame data (M-) |
| Ganryu | 9d22baaa | | uf+3+4,B | Sumo Squash – Sumo Sit – Roll Back | ,B | | frame data (M-) |
| Ganryu | 0fb45f9e | | uf+3+4~D | Sumo Squash – Sumo Sit – Splits | ~D | | frame data (M-) |
| Ganryu | cd9f0bca | | uf+3+4~D,u | Sumo Squash – Sumo Sit – Splits – SS | ~D,u | | frame data (M--) |
| Ganryu | 2b41d150 | KND | d+1+2,F | Spring Hammer – Sumo Sit – Roll Forward | ,F | | frame data (m-) |
| Ganryu | 9e1331a8 | KND | d+1+2,B | Spring Hammer – Sumo Sit – Roll Back | ,B | | frame data (m-) |
| Ganryu | b4b24070 | KND | d+1+2~D | Spring Hammer – Sumo Sit – Splits | ~D | | frame data (m-) |
| Ganryu | 5f162acc | KND | d+1+2~D,u | Spring Hammer – Sumo Sit – Splits – SS | ~D,u | | frame data (m--) |
| Ganryu | 2fd36c6a | | b+1+2,D | Sumo Tackle – Splits | ,D | | frame data (!-) |
| Jack-2 | a96a391b | | =~3+4,B | Hop Hip Press – Roll Back | ,B | | frame data (M-) |
| Jack-2 | df8e44d6 | | =~3+4,F | Hop Hip Press – Roll Forward | ,F | | frame data (M-) |
| Jun | b5c67ae7 | | f+2 | Charging Strike | ~d (Strike Cancel) | crouching | name |
| Jun | bf03ba4f | | f+3 | Spinning Roundhouse | ,d (Hit Cancel) | crouching | name |
| Kuma | 90ce6d1b | | 1+2~f | Bear Knuckle – Hunting Bear Stance | ~f | HBS | frame data (m-) |
| Kuma | aae54c5e | | WS+1+2<1+2~f | Claw Upper – Bear Knuckle – HBS | ~f | HBS | frame data (mm-) |
| Kuma | 3080d799 | | b+1+2,3+4 | Deadly Claw – Hunting Bear Stance | ,3+4 | HBS | frame data (!-) |
| Kunimitsu | 8a1ad63b | | f,f+3+4,3+4 | Shark Attack – Ninja Flying Shadow | ,3+4 | | frame data (m-) |
| Lei | 087e74b1 | | b+3+4 | Turn Around | ~f | standing | name |
| Ling | b929cc95 | | u+1+2,3+4 | Double Fan – Ginger Snap | ,3+4 | | frame data (mm-) |
| Michelle | 8045d9f7 | | b+2,1+2 | Arm Whip – Back Push | ,1+2 | | frame data (h-) |
| Nina | 3f312a21 | | db+1+2 | Hunting Swan | ,u,u | standing | name |
| Nina | e95ede68 | | SS+1~F | Snake Shot – Crouch Dash | ~F | | frame data (m-) |
| Nina | d03a4da5 | | SS+1~B | Snake Shot – Sway | ~B | | frame data (m-) |
| Prototype Jack | 3b432f29 | | b+1+2 | Clock Up | ~f | standing | name |
| Roger | fcc6c234 | | b+3+4~B | Lunge Animal Kick – Roll Back | ~B | | frame data (m-) |
| Roger | e13d27e9 | | b+3+4,3,4~B | Lunge Animal Kick – AKR – Roll Back | ~B | | frame data (mmm-) |
| Roger | a2e4233a | | b+3+4,3,4,3,4~B | Lunge Animal Kick – AKR – AKR – Roll Back | ~B | | frame data (mmmmm-) |
| Roger | 08b4077f | | b+3+4,3,4,1 | Lunge Animal Kick – AKR – Rolling Animal | ,1 | | frame data (mmm-) |
| Roger | 7b58ef3a | | b+3+4,1 | Lunge Animal Kick – Rolling Animal | ,1 | | frame data (m-) |
| Roger | 711aeb90 | | b+3+4,3,4,1,3,4~B | Lunge AK – AKR – RA – AKR – Roll Back | ~B | | frame data (mmm-mm-) |
| Roger | 7d6af8c7 | | b+3+4,1,3,4~B | Lunge AK – RA – AKR – Roll Back | ~B | | frame data (m-mm-) |
| Roger | 50a600d0 | | b+3+4,1,3,4,3,4~B | Lunge AK – RA – AKR – AKR – Roll Back | ~B | | frame data (m-mmmm-) |
| Roger | 25704ef5 | | b+3+4,1,3,4,1 | Lunge AK – RA – AKR – Rolling Animal | ,1 | | frame data (m-mm-) |
| Roger | a32c1ee8 | | b+3+4,1,3,4,1,3,4~B | Lunge AK – RA – AKR – RA – AKR – Roll Back | ~B | | frame data (m-mm-mm-) |
| True Ogre | 9eefea67 | | df+3+4~DF | Tail Spin – Back Turned Position | ~DF | | frame data (m-) |
| Wang | 9c67a8d1 | | b,b+1 | Heaven Cannon | ~B | standing | name |
| Yoshimitsu | bf02113c | | b+1,1,1,1,1,1,3+4 | Stone Fists – Evasive Side Spin | ,3+4 | | frame data (hhhhhh-) |
| Yoshimitsu | 7dd15930 | | f,f+2,d+3+4 | Ninja Blade Slice – Indian Sit | ,d+3+4 | | frame data (m-) |
| Yoshimitsu | f0840bfe | | 1+2,d+3+4 | Flea – Indian Sit | ,d+3+4 | | frame data (!-) |
| Yoshimitsu | 7eb42781 | | db+1~N,DB | Sword Slice – Sword Delay | ~N,DB | | frame data (!-) |

## Varmistetut

| Character | UUID | From | Command | Move Name | Cancel Input | Cancel To | Source | Comment |
|-----------|------|------|---------|-----------|-------------|-----------|--------|---------|
| Anna | 48b304a4 | | db+1+2 | Hunting Swan | ,u,u | standing | name | |
| Baek | 89dc78f0 | | db+3+4 | Heel Explosion | ,f/b/u/d | FLA | name | cancel works with any direction (f b u d) |
| Forest | 5a04a751 | | db+1+2 | Dragon Fang | ,u,u | standing | name | |
| Hwoarang | 5c727428 | | db+3+4 | Dynamite Heel | ,b | LFS | name | |
| Hwoarang | f2e6a842 | LFS | 1+4 | Power Blast | ,b,b | standing | name | |
| Lee | ad99b613 | | db+1+2 | Silver Fang | ,u,u | standing | name | |
| Ling | 25f9b418 | | b+1+2<1+2 | Phoenix Strike | ,b,b | RDS | name | ub can be canceled only when 1+2 is inputted at the very beginning of hypnotist walk. if ub is delayed, it can't be canceled. |

## Ignoroidut

| Character | UUID | From | Command | Move Name | Cancel Input | Cancel To | Source |
|-----------|------|------|---------|-----------|-------------|-----------|--------|
| Jin | f465515c | DJP | b+1+2 | Lightning Force | ~2~2~3 (Force Cancel) | DJP | name |
| Jin | 1ac36d3f | | b+1~df | Parting Wave – Crouch Dash | ~df | | frame data (m-) |
| Lee | 3ca6233d | | b+1,1,3+4 | Fang Rush – Hit Man Stance | ,3+4 | HMS | frame data (hm-) |
| Julia | a2a00b1a | | b+2,1+2 | Arm Whip – Back Push | ,1+2 | | frame data (h-) |

## Nämä pitää testata

| Character | UUID | From | Command | Move Name | Cancel Input | Cancel To | Source |
|-----------|------|------|---------|-----------|-------------|-----------|--------|
| Bryan | 401d34d7 | | b+2~f+1,u | Fake Out Doom Knuckle – Side Step | ~f+1,u | | frame data (h-) |
