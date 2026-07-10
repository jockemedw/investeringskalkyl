# CLAUDE.md — Investeringskalkyl Lejonfastigheter

Ersättare för LM 371 Investeringskalkyl. Ren xlsx (inga makron), 10 flikar, 20-årig kalkyl.
Beställare: Lejonfastigheter AB (kommunalt fastighetsbolag, Linköping).

**Aktuell produkt:** `build/oneshot/Investeringskalkyl_v2.xlsx` (v2 ersätter iter9, byggd från spec-replay + rundor). Utskrift: 18 sidor för hela boken, trogen-PDF-verifierad. Ifyllnadsupplevelse (ONESHOT-POLISH): blått input-språk med svenska valideringar, bladskydd utan lösenord, öppningsvyer, ren tom mall (D-23–D-25).

## Aktuell baseline (iter 8 → v2 reproducerar exakt)

| Mätvärde | Värde | Cell |
|----------|-------|------|
| Bindande kravhyra | 11 631 221,72 kr/år | `Resultat!D14` |
| Faktisk IRR EK | 7,6452 % | `Lönsamhetskontroll!C45` |
| Marginal mot IRR-krav (6,3 %) | +1,35 pp | — |

Testfall: Skola (Nyb) 5 000 kvm × 40 000 kr/kvm = 200 Mkr.

## Arbetsflöde

```bash
# Bygga v2 (spec-replay + rundor + recalc + regression i ett):
python build/oneshot/build_v2.py

# Enbart regressionsgate (baslinje + nollfelsscan + IRR-scenariovalidering):
python build/oneshot/regression_v2.py

# Trogen utskriftsexport (Excel COM ExportAsFixedFormat) — enda sanningen
# för print/design; PageSetup.Pages.Count och render_local/xlsx-review ljuger:
python build/oneshot/export_pdf.py   # → build/oneshot/print_preview.pdf + .print/*.pdf
```

För ny §10-uppgift: lägg till `round_*(wb)` i [build/oneshot/v2_rounds.py](build/oneshot/v2_rounds.py), registrera i `apply_all()`. OBS: `round_print_polish` ska ligga sist (sidenav sätter radhöjd 30 på rad 2–11 på alla flikar). Köra → PDF-granska → committa.

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

- `build/oneshot/Investeringskalkyl_v2.xlsx` — **den färdiga produkten** (artefakt — byggs av build_v2.py)
- `build/oneshot/build_v2.py` — bygger v2 från spec_iter9.json + v2_rounds
- `build/oneshot/v2_rounds.py` — en funktion per förbättringsrunda, `apply_all()` är pipelinen
- `build/oneshot/regression_v2.py` — gate: baslinje + nollfelsscan + IRR-scenariomatch
- `build/oneshot/export_pdf.py` — trogen PDF-export, sidantal per flik
- `Investeringskalkyl_iter8.xlsx` — historisk baseline (auktoritativ referens — ändra inte)
- `tools/patches.py` — load/save/diff/rename/replace-utilities
- `tools/recalc.py` — Excel COM via PowerShell eller LibreOffice headless
- `tools/theme.py` / `tools/sidenav.py` — designsystem (roller, palett) + sidnavigator
- `tests/regression.py` — verifierar 11 631 222 / 7,65 % / +1,35 pp
- `build/iter9.py` — historisk iter (ersatt av v2)
- `HANDOFF.md` — projektkontext, fashistorik, arkitektur
- `FINAL.md` — slutrapport ONESHOT-FINAL (2026-07-10)
- `DECISIONS.md` — designval D-01 → D-22 (D-XX OMPRÖVAR D-YY när ersatts)
- `TECH_NOTES.md` — openpyxl-mönster, default-värden för iter 8
- `OPEN_QUESTIONS.md` — frågor som behöver Joakim-input

## §10 — öppna uppgifter

Status (senaste rad = överst):

- [x] **ONESHOT-POLISH klar (2026-07-10):** ifyllnadsupplevelsen på skärm — ett blått input-språk (51 svenska valideringar), bladskydd med Tab-vandring (outline-flikar undantagna, D-24), öppningsvy per flik + ifyllnadsguide/status, tom mall utan felkoder, 2 skärm-granskningsvarv + programmatisk interaktionskontroll. Print fortsatt 18 sidor. Se [POLISH.md](POLISH.md), D-23–D-25. Skärmvyer verifieras med `build/oneshot/screenshot_sheets.py`. Ej pushad.
- [x] **ONESHOT-FINAL klar (2026-07-10):** alla öppna §10-uppgifter stängda — se [FINAL.md](FINAL.md). Ej pushad till origin (Joakim pushar efter granskning).
- [x] **FINAL m3 (2026-07-10):** design-excellens-pass, 2 hela trogen-PDF-granskningsvarv över alla 18 sidor. Indata sektion 5-mallrader fixade (commit b9bdb2a).
- [x] **FINAL m2 (2026-07-10):** Faktisk IRR EK per yield-scenario (Lönsamhetskontroll rad 86, dolt EK-cashflow-block rad 89–91). Bedömt == C45 exakt, valideras i regressionsgaten. + D-20: timing-fix i MV-exit-tabellens scenario-IRR (commit a6ab310).
- [x] **FINAL m1 (2026-07-10):** Utskriftsformat v2 klart — hela boken 31 → 18 sidor (Beräkningslogik 16→3 via outline-kollaps av rådatablocket, Dokumentation 3→2, Kassaflöde 4→2, Resultat 2→1). Inga ########, ingen kapad text (commit 0df6d82). Se D-22.
- [x] **MERGE klar (2026-06-13):** `oneshot-v2` fast-forwardad till `main` (HEAD `5c746ca`). Frikopplad från print. Ej pushad till origin ännu.
- [x] Design-pass alla 9 flikar (NIGHTRUN, autonomt): 7 rena, 2 → Joakim — **båda defekterna lösta i v2/FINAL m1** (Översikt 18-sidorsbuggen försvann med v2:s Översikt-redesign; Beräkningslogik ######## löst via tkr-block + outline-kollaps). Se [NIGHTRUN.md](NIGHTRUN.md). Meta kvarstår: `render_local`/`/xlsx-review` opålitligt på flikar med icke-trivial page setup — verifiera med trogen export.
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
- [x] Faktisk IRR EK per scenario i känslighetstabellen — klar i FINAL m2 (se ovan, D-21)

Inga öppna §10-uppgifter. Nästa: Joakim granskar + pushar main till origin.
