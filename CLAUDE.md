# T3FD — Tekken Tag Tournament Frame Data

Tekken Tag Tournament (PS2) -pelin frame data -pipeline. Lähdedata on Tekken Zaibatsun web-sivuilta (Wayback Machine), ja se käsitellään vaiheittain Python-skripteillä rakenteelliseksi dataksi.

## Arkkitehtuuri

Lähde-HTML:t haetaan kerran (`fetch_sources.py` / `build.sh`) ja tallennetaan `sources/`-kansioon. Jokainen hahmo käsitellään saman step-putken läpi. Stepit ajetaan järjestyksessä (`all_steps.sh`), ja jokainen step lukee edellisen outputin ja kirjoittaa oman xlsx-tiedostonsa:

| Step | Skripti | Tehtävä |
|------|---------|---------|
| 1 | `step_1_html_to_xlsx.py` | HTML → xlsx, kaksi sheettiä (Frame Data + Movelist) |
| 1b | `step_1b_anna_fixes.py` | Anna-spesifiset lähdedatakorjaukset |
| 2 | `step_2_parent_uuid.py` | Headerien normalisointi, Parent UUID continuation-liikkeille (– marker) |
| 3 | `step_3_merge.py` | Frame Data + Movelist yhdistäminen komennon perusteella |
| 4 | `step_4_alternatives.py` | (A_B) vaihtoehtonotaatioiden purkaminen, Alt Commands -sarake |
| 5 | `step_5_notation.py` | Notaation normalisointi (FC/WR/WS/SS prefix, slash removal), Damage Sum |
| 5b | `step_5b_hit_markers.py` | Puuttuvien x-markereiden täydennys Speed-sarakkeeseen |
| 6 | `step_6_prepare_for_expansion.py` | Move Name -korjaukset ennen expansion-vaihetta |
| 6b | `step_6b_bracket_notations.py` | Hakasulkenotaatiot: [~5] → Taggable, stance recovery -merkinnät |
| 7 | `step_7_expand_continuations.py` | Continuation-liikkeiden purkaminen: Command Full, Move Name Full, jne. |
| 8 | `step_8_extract_tags.py` | [~5] ja [Tag] -merkintöjen siivous Command Full / Move Name Full -sarakkeista |
| 8b | `step_8b_from_stance.py` | From Stance -sarake (komentoprefixin tai section-nimen perusteella) |
| 8c | `step_8c_press_boundaries.py` | Painallusrajojen merkintä Command Full -sarakkeeseen (välilyönnit ennen , < ~) |
| 9 | `step_9_manual_corrections.py` | Manuaaliset korjaukset (modify/insert/delete UUID:n perusteella) |
| 10 | `step_10_hits_per_move.py` | Hits Per Move -sarake: montako osumaa kukin painallus tuottaa |

## Tiedostorakenne

- `sources/` — HTML-lähteet ja kaikki välivaiheiden xlsx-tiedostot (`<hahmo>_step<N>.xlsx`)
- `tools/` — Validointi- ja analyysiskriptit
- `index.html` — Web-käyttöliittymä datan selailuun
- `framedata.tsv` / `counters.tsv` — Lopullinen tuotantodata

## Periaatteet

- Kaikki datamuutokset on määritelty scriptien logiikassa tai niiden sisäisissä CORRECTIONS-dictionaryissä — ei manuaalisia xlsx-muokkauksia.
- Jokainen rivi tunnistetaan UUID:llä, joka generoidaan step 1:ssä HTML:n `<tr>`-elementeistä.
- Continuation-liikkeet (liikesarjan jatko-osat) linkitetään Parent UUID:llä vanhempaansa.
- Välivaiheiden xlsx-tiedostot ovat debuggausta varten — lopullinen data tulee viimeisestä stepistä.

## Kieli

Koodi, commitit ja dokumentaatio englanniksi. TODO.md ja keskustelu suomeksi.
