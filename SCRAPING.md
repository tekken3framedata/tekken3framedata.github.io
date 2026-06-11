# Tekken Frame Data Scraping

## Lähteet

Tekken Zaibatsu -sivuston data haetaan Wayback Machinesta (web.archive.org), koska alkuperäinen sivusto ei ole enää toiminnassa.

Kaksi erillistä sivua per hahmo:

1. **Frame data** (TTT): `tekkentag/framedata.php?id=<character>`
   - Sarakkeet: Command, Hit (=Speed), Block Adv, Hit Adv, Counter Hit Adv
   - Osiot vaihtelevat hahmoittain. Tyypillisiä: Basic Arts, Special Arts, Unblockable Arts
   - Joillakin hahmoilla ylimääräisiä osioita (esim. Jin: "Devil Jin Possession Arts")

2. **Movelist** (Tekken 3): `tekken3/movelist.php?id=<character>`
   - Sarakkeet: Command, Move Name, Stance, Damage, Hit Range, Properties
   - Osiot: Grappling Arts, Special Arts, Unblockable Arts, String Hit Arts
   - Jokaisen osion jälkeen voi olla `<div id="footnote">` jossa #1–#N selitykset
   - Footnote-numerot ovat osiokohtaisia (sama #1 voi tarkoittaa eri asiaa eri osioissa!)

## Scripti

```bash
python3 scrape_framedata.py <framedata_url> <movelist_url> <output.xlsx>
```

Scripti löytää kaikki osiot dynaamisesti molemmista sivuista ja yhdistää ne. Output on XLSX (openpyxl).

### Esimerkki

```bash
python3 scrape_framedata.py \
    "https://web.archive.org/web/20201206043428/http://www.tekkenzaibatsu.com/tekkentag/framedata.php?id=julia" \
    "https://web.archive.org/web/20201206042940/http://www.tekkenzaibatsu.com/tekken3/movelist.php?id=julia" \
    julia_framedata_v3.xlsx
```

## Move Patches

Patchit korjaavat lähdedataa ennen matchausta ja expansionia. Tiedosto: `sources/<character>_patches.tsv`.

### Formaatti

TSV (tab-eroteltu). Tyhjät rivit ja `#`-alkuiset rivit ohitetaan.

### Operaatiot

| Operaatio | Sarakkeet | Toiminto |
|-----------|-----------|----------|
| `replace` | `replace <source> <section> <old_cmd> <new_cmd>` | Korvaa komennon |
| `set` | `set <source> <section> <cmd> <column> <value>` | Asettaa yksittäisen sarakearvon |
| `add_after` | `add_after <source> <section> <after_cmd> <col0> <col1> ...` | Lisää uuden rivin komennon jälkeen |
| `delete` | `delete <source> <section> <cmd>` | Poistaa rivin |

- `<source>`: `fd` (frame data) tai `ml` (movelist)
- `<section>`: osion nimi täsmälleen kuten HTML:ssä (esim. `Special Arts`)
- `<column>`: sarakkeen nimi (`Command`, `Speed`, `Block Adv`, `Hit Adv`, `CH Adv` FD:lle; `Command`, `Move Name`, `Stance`, `Damage`, `Hit Range`, `Properties` ML:lle)

### Suoritusjärjestys

Patchit ajetaan heti taulukoiden parsimisen jälkeen, ennen matchausta, continuation expansionia ja kaikkea muuta prosessointia. Tämä mahdollistaa lähdedatan notaatio-ongelmien korjaamisen puhtaasti.

### Tyypillinen käyttötapaus

FC-prefiksilliset liikkeet joissa lähde käyttää pilkkua (`FC,D/F+2`) mutta logiikka tulkitsee pilkun separaattoriksi. Patch korvaa pilkun välilyönnillä (`FC DF+2`) tai yhdistää muuten notaation oikein. Myös tilanteisiin joissa yksi lähderivi edustaa kahta eri liikettä (esim. Lingin Shady Lotus → FC vs. → Rain Dance).

## Matchaus (yhdistäminen)

Frame data- ja movelist-sivujen taulukot yhdistetään Command-sarakkeen perusteella.

### Ryhmäpohjainen LCS-matchaus

Yksinkertainen flat LCS ei toimi koska geneerisiä follow-up-komentoja (kuten `= 4`, `= 3`) esiintyy useissa eri konteksteissa. Algoritmi:

1. **Ryhmittely**: Rivit jaetaan ryhmiin. Parent-liike (ei ala `=`:lla) aloittaa ryhmän, ja kaikki seuraavat `= ...` -rivit kuuluvat siihen.
2. **Parent-matchaus**: Parent-komentojen välillä tehdään LCS.
3. **Follow-up-matchaus**: Matchatun ryhmäparin sisällä follow-upit matchataan omalla LCS:llä.

### Normalisointi vertailua varten

- `[~5]` ja edeltävä ` - ` poistetaan (TTT tag buffer -merkintä)
- `<` korvataan `,`:llä (TTT käyttää `<` viivästykselle, T3 käyttää `,`)

### Unmatchatut rivit

Rivit jotka eivät matchaa saavat `TRUE` Unmatched-sarakkeeseen. Nämä vaativat manuaalista tarkistusta. Syitä:

- **TTT-only liike**: Liike on lisätty TTT:hen eikä sitä ole T3 movelististä (esim. `SS+2`, `b+3`, `d/b+1`)
- **Notaatioero**: Sama liike on merkitty eri tavalla (esim. `f~N,d~d/f+2` vs `f,N,d~d/f+2`, tai `(d/f+1,2_f+1+2,2_WR+1+2,2)` vs `(d/f+1,2_f+1+2,2)` kun TTT lisää WR-variantin)
- **Yhdistämisero**: T3 movelist yhdistää kaksi komentoa samalle riville (esim. `(f,N,d,d/f,f_WS)+4,4`) mutta TTT frame data listaa ne erikseen

Manuaalinen tarkistus: avaa XLSX, suodata Unmatched-sarake (`TRUE`), vertaa movelisti-sivuun ja päätä onko kyseessä TTT-only liike vai pitääkö täydentää nimi käsin.

## Footnote-viittaukset

Properties-sarakkeessa olevat `#1`, `#2` jne. korvataan sivun Foot Notes -selityksillä. Selitys menee Notes-sarakkeeseen, Properties-sarakkeeseen jää muu sisältö (GB, JG, FSc, OB, CH, RC jne.).

Footnote-divit ovat osiokohtaisia: ne sijaitsevat `</table>` ja seuraavan `<h2>` välissä. Regex: `(#\d+)\s+(.*?)(?:<br|[\n\r]|</fieldset)`.

## Continuation expansion

Follow-up-liikkeet (jotka alkavat `= ...`) laajennetaan täysiksi komennoiksi:
- `(WS+2_3~2)` → `(WS+2_3~2),4` → `(WS+2_3~2),4,4`
- Indentaatiotaso (1 välilyönti = level 1, 3 = level 2) määrittää minkä parentin jatko

Move Name laajennetaan vastaavasti `>`:lla erotettuna:
- `Tequila Sunrise` → `Tequila Sunrise > Razor Sweep`

## Multi-hit split

Liikkeet joiden Block/Hit/CH -sarakkeissa on useita välilyönnillä erotettuja arvoja jaetaan erillisiksi riveiksi. Logiikka:

1. Alkuun `x`-arvot poistetaan (tarkoittaa ettei arvoa voida mitata)
2. Jos jää 1 arvo → ei jakoa
3. Jos jää 2+ arvoa → komento pilkotaan viimeisten top-level-separaattoreiden (`,`/`~`) kohdalta
4. Jos parent-komento on jo taulukossa (esim. `1` Basic Artsissa), sille ei luoda omaa riviä

Esimerkki: `1~1` jossa Block=`0 -15` (ja `1` on jo Basic Artsissa):
- Vain yksi rivi: Command=`1~1`, Block=`-15`, Hit=`-4`, CH=`-4`

Esimerkki: `1+4,3` jossa Block=`x -13 12`:
- `x` poistetaan → [-13, 12], 2 arvoa
- Rivi 1: Command=`1+4`, Speed=`20`, Block=`-13`
- Rivi 2: Command=`1+4,3`, Speed=(tyhjä), Block=`12`

Speed-arvo jää vain ensimmäiselle riville (startup). Sulkujen sisällä olevat separaattorit eivät laukaise jakoa.

### Move Name -jakaminen

Move Namessa ` - ` korvataan aina ` > `:lla (tarkoittaa eri osumien nimiä samassa stringissä).

Kun rivi jaetaan ja nimiä on yhtä monta kuin hittejä, nimet jaetaan 1:1:
- `Club Fist > Flash Uppercut` (2 osaa, 2 hittiä) → `WS+2,1`="Club Fist", `WS+2,1,1`="Flash Uppercut"

Kun nimiä on vähemmän kuin hittejä, lisätään ordinaalit:
- `Club Fist > Bow & Arrow` (2 osaa, 3 hittiä) → `WS+2,1`="Club Fist", `WS+2,1,4`="Bow & Arrow (First)", `WS+2,1,4,3`="Bow & Arrow (Second)"

## Alternative-komennot

`(A_B)` -notaatio puretaan: ensimmäinen vaihtoehto tulee Command-sarakkeeseen, loput Alt Commands -sarakkeeseen.
- Button 5 (TTT tag) poistetaan vaihtoehdoista
- `PREFERRED_COMMAND`-dict ohjaa kumpi vaihtoehto on ensisijainen (hahmokohtainen)
- Follow-up-rivit eivät saa omia Alt Commands -arvoja

## TTT-only-liikkeiden filtteröinti

`TAG_ONLY_MOVES`-dict listaa hahmokohtaisesti liikkeet jotka ovat puhtaasti TTT tag-mekaniikkaa. Nämä poistetaan kokonaan outputista (eivät päädy tiedostoon lainkaan). Myös poistettujen parent-liikkeiden jatkot suodattuvat pois.

## Duplicate-yhdistäminen

`MERGE_DUPLICATES`-dict määrittelee komentoparit jotka ovat sama liike eri notaatiolla (esim. `d+1` / `FC+1`). Ensisijainen jää, toisesta tulee Alt Command.

## `/`-merkin poisto

Command- ja Alt Commands -sarakkeista poistetaan kaikki `/`-merkit kirjoitusvaiheessa (esim. `d/f+1` → `df+1`).

## Tuotettu XLSX-rakenne

Yksi yhtenäinen sarakejärjestys kaikille osioille:

```
Character | Stance | Command | Move Name | Damage | Hit Range | Properties | Speed | Block Adv | Hit Adv | CH Adv | Alt Commands | Notes | Unmatched | UUID
```

Osiot erotetaan tyhjällä rivillä ja otsikkorivillä (bold). Sarakeheaderit toistuvat jokaisen osion alussa.

### Osiot järjestyksessä:

1. **BASIC ARTS** — FD only (Speed, Block/Hit/CH Adv)
2. **SPECIAL ARTS** — yhdistetty FD + ML (kaikki sarakkeet)
3. **[Hahmokohtaiset FD-osiot]** — esim. DEVIL JIN POSSESSION ARTS (Stance-sarake kertoo kontekstin)
4. **UNBLOCKABLE ARTS** — yhdistetty FD + ML
5. **GRAPPLING ARTS** — ML primary, FD:stä Speed. Escape-tieto Properties-sarakkeessa.
6. **STRING HIT ARTS** — ML only (Damage, Hit Range, hits-lukumäärä Notes-sarakkeessa)

## Huomioita

- TTT-spesifiset liikkeet suodatetaan pois `TAG_ONLY_MOVES`-dictin perusteella
- `[~5]` TTT frame datassa tarkoittaa tag-bufferia — poistetaan komennoista continuation expansionissa
- Unblockable Arts voi olla sekä FD:ssä (frame data) että ML:ssä (nimi) — scripti yhdistää nämä
- Grappling Arts FD:ssä on vain Speed-sarake, se liitetään ML:n riveihin
- FD-sivun "Grappling Arts" ei sisällä frame advantage -dataa, vain startupin
- Hahmokohtaiset ylimääräiset FD-osiot (kuten Devil Jin Possession Arts) kirjoitetaan sellaisenaan ilman merge-logiikkaa, Stance-sarake ilmaisee kontekstin
- UUID generoidaan jokaiselle riville (`uuid4`)
- Character-sarake populoidaan automaattisesti URL:n `id`-parametrista

## Wayback Machine URL:t

Käytetty snapshot: `20201206` (joulukuu 2020). URL-muoto:
- Frame data: `https://web.archive.org/web/20201206043428/http://www.tekkenzaibatsu.com/tekkentag/framedata.php?id=<character>`
- Movelist: `https://web.archive.org/web/20201206042940/http://www.tekkenzaibatsu.com/tekken3/movelist.php?id=<character>`

Timestamp-osa URL:ssa voi vaihdella hieman hahmoittain, mutta `20201206` toimii Wayback Machinen redirect-logiikan kanssa.
