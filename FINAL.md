# FINAL — slutrapport ONESHOT-FINAL (2026-07-10)

Mål: `build/oneshot/Investeringskalkyl_v2.xlsx` produktionsklar och visuellt exceptionell.
**Status: klart.** Regression grön hela vägen. Main är redo att pushas (Joakim pushar efter granskning).

## Definition of done — utfall

| Krav | Utfall |
|------|--------|
| 1. Regression grön | ✅ 11 631 221,72 kr/år / 7,6452 % / +1,35 pp + nollfelsscan, varje milstolpe |
| 2. Utskrift komprimerad | ✅ 31 → **18 sidor** för hela boken; inga `########`, ingen kapad text, inga klippta kolumner |
| 3. IRR EK per scenario | ✅ Lönsamhetskontroll rad 86: opt 7,69 % / bedömt 7,65 % / pess 7,60 %; bedömt == C45 **exakt** (valideras i regressionsgaten) |
| 4. Visuellt exceptionellt | ✅ 2 hela trogen-PDF-granskningsvarv över alla 18 sidor efter sista ändringen; enhetlig eyebrow/H1/H2-hierarki |

## Sidantal per flik (trogen PDF, före → efter)

| Flik | Före | Efter |
|------|------|-------|
| Översikt | 1 | 1 |
| Försättsblad | 2 | 2 |
| Indata | 2 | 2 |
| Kassaflöde | 4 | 2 |
| Finansiering | 2 | 2 |
| Resultat | 2 | 1 |
| Grafer | 1 | 1 |
| Lönsamhetskontroll | 2 | 2 |
| Beräkningslogik | **16** | **3** |
| Dokumentation | 3 | 2 |
| **Totalt** | **~31–35** | **18** |

(“Före” = trogen export vid FINAL-start; per-flik-PDF:er avviker från `PageSetup.Pages.Count`, som rapporterar färre sidor än Excel faktiskt exporterar — PDF:en är sanningen.)

## Vad som gjordes

### m1 — Print-pass (commit 0df6d82)
- **Beräkningslogik 16→3:** tekniska rådatablocket (rad 48–206) outline-kollapsat som default — all data kvar, expanderas med `+`; dolda rader skrivs inte ut. `print_area` B1:K230, restvärdesresonemangets B:R-merges omslagna till B:K.
- **Dokumentation 3→2:** prosan wrappar över hela printbredden (B = 152 enheter), storleksmedveten autofit (9 pt body, leading 1,37/1,45), spacers 2,5 pt.
- **Kassaflöde 4→2 / Finansiering 2:** `fitToHeight=1` + upprepade etikettkolumner B:C (årsremsan 2038–2045 saknade radetiketter); långa captions omslagna till B:C-merge.
- **Resultat 2→1:** `#######` i Högsta-kolumnen (E 13→18), Tolkning-raden kluvits av sidbrytning (höjd 48), `fitToHeight=1`.
- **Grafer:** `print_area` N→V (diagramdata 2038–2045 klipptes).
- **Försättsblad:** print utan sidenav-bandet (slutade abrupt mitt på sida 2), `print_area` B1:I92 (värde-merges H:I klipptes annars), radbrytning vid rad 45.
- **H1-nedstaplar** (g/y) klipptes på Översikt (36 pt) och Försättsblad (42 pt): `tools/sidenav.py` sätter radhöjd 30 på rad 2–11 på *alla* flikar → `round_print_polish` flyttad sist i pipelinen.
- **Lönsamhetskontroll:** B51-caption (merge till H) låg utanför print_area (G).

### m2 — IRR EK per scenario (commit a6ab310, D-20/D-21)
- Restvärdestabellen har nu **Faktisk IRR EK per scenario** (rad 86). EK-cashflow rekonstrueras per scenario i dolt outline-block (rad 89–91): linjärt i hyran via Beräkningslogik block D/E (rad 173/188), exit-delta i år 20-kolumnen (Y) för scenarioyielden, vid respektive scenarios egen bindande kravhyra.
- Bedömt-kolumnen är per konstruktion identisk med rad 43 → **IRR == C45 bit-för-bit**; valideringen ligger permanent i `regression_v2.py`.
- **Bugg fixad (D-20):** MV-exit-tabellens ΔMV adderades i kolumn AD (år 25) i stället för Y (år 20) → scenario-IRR hade justeringen 5 år för sent. Korrekt spann nu 5,83–8,36 % (var 6,38–8,14 %); −40 %-scenariot flaggar korrekt `⚠ IRR<krav`. Bas opåverkad.

### m3 — Design-excellens (commit b9bdb2a)
- Indata sektion 5 (re-investering): tomma mallrader renderade som kantlöst blått block med sex lösryckta nollor → hårlinjer mellan raderna + tomt i stället för 0.
- Två fullständiga granskningsvarv (vision, sida för sida) över alla 18 sidor efter sista ändringen — inga kvarstående defekter.

### m4 — Dokumentation
- `DECISIONS.md`: D-20 (exit-delta-timing), D-21 (IRR per scenario), D-22 (printstrategi).
- `CLAUDE.md`: §10 stängd, arbetsflöde/filer synkade mot v2.
- `STATUS.md`: synkad mot verkligt läge.

## Kvarstående kända begränsningar

1. **Expanderat rådatablock + manuell utskrift:** om användaren expanderar Beräkningslogiks kollapsade block på skärm och skriver ut klipps kolumner bortom K (print_area är pedagogikens bredd). Dokumenterat i fliken (B47). Medvetet val (D-22) — alternativet var 16 sidor.
2. **Dokumentation är tät i print:** 9 pt body vid ~0,74 utskriftsskala ≈ 6,6 pt effektivt. Läsbart men kompakt; digitalt (zoom) opåverkat. Priset för 2-sidorsmålet.
3. **IRR-baslinjens definition:** `C45 = IRR(F43:AD43)` inkluderar drift-kassaflöden år 21–25 *efter* exit i år 20 (dubbelräkning i strikt mening). Detta är iter8-baslinjens definition (regression-pinnad, D-14/D-15) och har INTE ändrats — scenario-IRR:erna använder samma definition och är därmed konsistenta inbördes. Flaggas som kandidat för omprövning om baslinjen någon gång öppnas.
4. **`PageSetup.Pages.Count` ljuger:** rapporterar färre sidor än faktisk export (t.ex. Beräkningslogik 2 vs 3). All print-verifiering ska ske mot `export_pdf.py`:s PDF:er.
5. **Kassaflöde/Finansiering sida 2 upprepar C-kolumnen** (print titles B:C): på Kassaflöde betyder det att "Bindande kravhyra 11 631 222 / NPV" syns även på årsremsan 2038–2045 — harmlöst, snarast informativt.

## Verifiering

- `python build/oneshot/build_v2.py` — bygger + recalc + regressionsgate (grön)
- `python build/oneshot/export_pdf.py` — 18 sidor; granskade renderingar i sessionens scratchpad
- Bedömt-scenariots IRR == C45 valideras vid varje regressionskörning

*Rapport skriven av Claude (ONESHOT-FINAL, autonom körning 2026-07-09 → 2026-07-10).*
