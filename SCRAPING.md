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
python3 scrape_framedata.py <framedata_url> <movelist_url> <output.tsv>
```

Scripti löytää kaikki osiot dynaamisesti molemmista sivuista ja yhdistää ne.

### Esimerkki

```bash
python3 scrape_framedata.py \
    "https://web.archive.org/web/20201206043428/http://www.tekkenzaibatsu.com/tekkentag/framedata.php?id=julia" \
    "https://web.archive.org/web/20201206042940/http://www.tekkenzaibatsu.com/tekken3/movelist.php?id=julia" \
    julia_framedata.tsv
```

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

Rivit jotka eivät matchaa saavat `[UNMATCHED]` Notes-sarakkeeseen. Nämä vaativat manuaalista tarkistusta. Syitä:

- **TTT-only liike**: Liike on lisätty TTT:hen eikä sitä ole T3 movelististä (esim. `SS+2`, `b+3`, `d/b+1`)
- **Notaatioero**: Sama liike on merkitty eri tavalla (esim. `f~N,d~d/f+2` vs `f,N,d~d/f+2`, tai `(d/f+1,2_f+1+2,2_WR+1+2,2)` vs `(d/f+1,2_f+1+2,2)` kun TTT lisää WR-variantin)
- **Yhdistämisero**: T3 movelist yhdistää kaksi komentoa samalle riville (esim. `(f,N,d,d/f,f_WS)+4,4`) mutta TTT frame data listaa ne erikseen

Manuaalinen tarkistus: avaa TSV, etsi `[UNMATCHED]`, vertaa movelisti-sivuun ja päätä onko kyseessä TTT-only liike vai pitääkö täydentää nimi käsin.

## Footnote-viittaukset

Properties-sarakkeessa olevat `#1`, `#2` jne. korvataan sivun Foot Notes -selityksillä. Selitys menee Notes-sarakkeeseen, Properties-sarakkeeseen jää muu sisältö (GB, JG, FSc, OB, CH, RC jne.).

Footnote-divit ovat osiokohtaisia: ne sijaitsevat `</table>` ja seuraavan `<h2>` välissä. Regex: `(#\d+)\s+(.*?)(?:<br|[\n\r]|</fieldset)`.

## TSV-escaping Google Sheetsiä varten

Google Sheets tulkitsee solun kaavaksi jos se alkaa `=`-merkillä. Tässä datassa ongelma on:
- Command-sarake: `= Flash Elbow`, `= 4` jne. (follow-up-liikkeet)
- Move Name -sarake: `= Flash Elbow` jne.

**Ratkaisu:** Prefixoi `'` (heittomerkki) solun alkuun. Sheets näyttää tekstin ilman heittomerkkiä.

`+6` ja `-2` toimivat sellaisinaan — Sheets tulkitsee ne numeroiksi, mikä on ok.

## Tuotettu TSV-rakenne

Tiedostossa osiot erotettu tyhjällä rivillä, jokaisen alussa otsikkorivi:

```
BASIC ARTS
Command  Speed  Block Adv  Hit Adv  Counter Hit Adv

SPECIAL ARTS
Command  Move Name  Damage  Hit Range  Properties  Notes  Speed  Block Adv  Hit Adv  Counter Hit Adv

[MAHDOLLINEN YLIMÄÄRÄINEN FD-OSIO, esim. DEVIL JIN POSSESSION ARTS]
Command  Hit  Block Adv  Hit Adv  Counter Hit Adv

UNBLOCKABLE ARTS
Command  Move Name  Damage  Hit Range  Properties  Notes  Speed  Block Adv  Hit Adv  Counter Hit Adv

GRAPPLING ARTS
Command  Throw Name  Type  Damage  Escape  Properties  Notes  Speed

STRING HIT ARTS
Command  Hits  Damage  Hit Range
```

## Huomioita

- TTT-spesifiset liikkeet (joita ei ole T3 movelististä) saavat `[UNMATCHED]` Notes-kenttään
- `[~5]` TTT frame datassa tarkoittaa tag-bufferia (voi tehdä liikkeen jälkeen tag outin painamalla 5)
- Unblockable Arts voi olla sekä FD:ssä (frame data) että ML:ssä (nimi) — scripti yhdistää nämä
- Grappling Arts FD:ssä on vain Speed-sarake, se liitetään ML:n riveihin
- FD-sivun "Grappling Arts" ei sisällä frame advantage -dataa, vain startupin
- Hahmokohtaiset ylimääräiset FD-osiot (kuten Devil Jin Possession Arts) kirjoitetaan sellaisenaan ilman merge-logiikkaa

## Wayback Machine URL:t

Käytetty snapshot: `20201206` (joulukuu 2020). URL-muoto:
- Frame data: `https://web.archive.org/web/20201206043428/http://www.tekkenzaibatsu.com/tekkentag/framedata.php?id=<character>`
- Movelist: `https://web.archive.org/web/20201206042940/http://www.tekkenzaibatsu.com/tekken3/movelist.php?id=<character>`

Timestamp-osa URL:ssa voi vaihdella hieman hahmoittain, mutta `20201206` toimii Wayback Machinen redirect-logiikan kanssa.
