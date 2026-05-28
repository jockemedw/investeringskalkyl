# CLAUDE.md — Investeringskalkyl Lejonfastigheter

Ersättare för LM 371 Investeringskalkyl. Ren xlsx (inga makron), 8 flikar, 20-årig kalkyl.
Beställare: Lejonfastigheter AB (kommunalt fastighetsbolag, Linköping).

## Aktuell baseline (iter 8)

| Mätvärde | Värde | Cell |
|----------|-------|------|
| Bindande kravhyra | 11 631 221,72 kr/år | `Resultat!D14` |
| Faktisk IRR EK | 7,6452 % | `Lönsamhetskontroll!C45` |
| Marginal mot IRR-krav (6,3 %) | +1,35 pp | — |

Testfall: Skola (Nyb) 5 000 kvm × 40 000 kr/kvm = 200 Mkr.

## Arbetsflöde

```bash
# Verifiera baseline mot iter8 cached values:
python tests/regression.py

# Bygga aktuell iter (modulärt — en funktion per §10-uppgift):
python build/iter9.py
python tests/regression.py build/iter9.xlsx

# Visuell designgranskning av en .xlsx (Excel COM → PNG → vision-analys):
# Triggar xlsx-review skill — Claude rapporterar layoutproblem, ändrar ingen kod
/xlsx-review build/iter9.xlsx
```

För ny §10-uppgift: lägg till `round_X_*(wb)` i [build/iter9.py](build/iter9.py), anropa från `main()`. Köra → verifiera → designgranska → committa.

## Kärnkonventioner

- **Kirurgisk redigering** av befintlig xlsx via openpyxl-patches — bygg INTE från grunden (D-02, TECH §7)
- **Kassaflöde stannar vid driftnetto** (D-05) — räntor/avskrivningar tillhör Finansiering-fliken
- **LM 371 är auktoritativ referens** (D-14) — avvik bara med dokumenterat skäl
- **Regression validerar varje iter** (D-15) — siffrorna måste reproducera, inte subjektiv bedömning
- **Cached values nollas av openpyxl-spara** — `recalc()` (Excel COM eller LibreOffice) krävs innan regression

## Arbetsstil med Joakim

- **Svenska, koncist** — hellre kort än långt
- **Plan först, exekvering sen** — föreslå explicit plan, han godkänner i batch
- **Tekniska sub-beslut är Claudes** — flagga bara där Joakim faktiskt behöver välja
- **Direkt pushback vid fel** — rätta utan utdragna ursäkter
- "Vi gör som LM 371" är giltigt argument utan vidare diskussion

## Filer

- `Investeringskalkyl_iter8.xlsx` — baseline (auktoritativ — ändra inte direkt)
- `tools/patches.py` — load/save/diff/rename/replace-utilities
- `tools/recalc.py` — Excel COM via PowerShell eller LibreOffice headless
- `tests/regression.py` — verifierar 11 631 222 / 7,65 % / +1,35 pp
- `build/iter9.py` — pågående iter, en funktion per §10-uppgift
- `build/demo_roundtrip.py` — sanity check för pipelinen
- `HANDOFF.md` — projektkontext, fashistorik, arkitektur
- `DECISIONS.md` — designval D-01 → D-19 (D-XX OMPRÖVAR D-YY när ersatts)
- `TECH_NOTES.md` — openpyxl-mönster, default-värden för iter 8
- `OPEN_QUESTIONS.md` — frågor som behöver Joakim-input

## §10 — öppna uppgifter

Status (senaste rad = överst):

- [ ] **NÄSTA:** (öppen — Joakim väljer)
- [x] Round AE: enhetligt designsystem — eyebrow/H1/H2-harmoni över alla 9 flikar (commit f25bfe7). Nya roller i `tools/theme.py`: `eyebrow`, `h1_display`, `h1_subtitle`, `h2_section`, `h3_sub`, `table_header`. Blå banner-sektioner ersatta med tunn underline-rule.
- [x] Round Z: Resultat — utfall vid bindande kravhyra (NPV/IRR/MV-status, refererar Lönsamhetskontroll)
- [x] Round Y: Indata-beskrivningar sektion 6-9 + räntenivå-renamning (Aktuell / Långsiktig)
- [x] Round X: känslighetstabell restvärdesbedömning (Lönsamhetskontroll rad 73+, opt/bedömt/pess × kravhyror + MV)
- [x] Round W: restvärdeskalibrering (Indata sektion 9 yield-justering, Beräkningslogik-guide, Försättsblad-status)
- [x] Round T-V: Försättsblad + Översikt v2 + Indata polish (commit 515e59b)
- [x] Round H: finansmodell-styling (XLSX-skill-standard, commit 11c9fe1)
- [x] Round G: Översikt om till beslutsdokument (commit 3073d1b)
- [x] Round F: designcleanup Lönsamhetskontroll via outline grouping (commit 59f194b)
- [x] Round C: Indata-fält renamning — branschterminologi
- [x] Round E: "År N" → "år 20" (commit c423d5e)
- [x] Round B: pedagogisk omskrivning Beräkningslogik (commit dab830b)
- [ ] Faktisk IRR EK per scenario i känslighetstabellen — kräver full EK-cashflow-rekonstruktion för opt/pess (icke-trivialt, ej börjat)

Joakim väljer ordning. Inga uppgifter blockerar varandra.
