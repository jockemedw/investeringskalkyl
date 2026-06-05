# NIGHTRUN — autonomt design-pass över alla 9 flikar

**Start:** 2026-06-06, natt. Joakim sover. Mandat: "kör så länge du kan", inget tillståndsfrågande.
**Baseline vid start:** grön — Hyra 11 631 221,72 / IRR 7,6452 % / +1,35 pp (commit `d424439`).
**Recalc-motor:** endast Excel COM (ingen LibreOffice-fallback installerad). Om Excel COM dör → loopen stannar naturligt.

## Mål
Ett (1) fullt svep över de 9 synliga flikarna. Fixa **objektiva designdefekter**, logga **subjektiva omdesignsbeslut** för Joakim. Committa per flik. Stanna efter sista fliken — ingen om-granskning.

## Per-flik-cykel (kör detta för varje flik)
1. `python tools/render_local.py build/iter9.xlsx` → läs flikens PNG under `.xlsx-review/iter9/`.
2. Triagera mot designchecklistan (se nedan). Notera fynd i ledgern.
3. **Objektiv defekt?** Fixa via ny/ändrad `round_X(wb)` i `build/iter9.py`, anropad från `main()`. Kirurgiskt mönster.
4. `python build/iter9.py` (bygg → XML-patch → recalc → regression i ett svep).
5. **Regression-gate:** måste vara grön. Röd → `git checkout build/iter9.py` (revert rundan), logga som blockerad, gå vidare.
6. Re-rendera, verifiera att defekten är borta och inget nytt brutits.
7. `git add -A && git commit` → `iter9 round XX: <flik> designfix — <kort>`. Uppdatera ledger till ✅.

## Designchecklista (objektiva defekter — fixas)
- Kapad/överflödande text (####, avhuggna rubriker)
- Krockande eller överlappande headers/celler
- Typografi som avviker från designsystemet (Round AE: eyebrow/h1/h2/h3/table_header, Segoe UI)
- Felaktig kolumnbredd/radhöjd (text får inte plats / enormt tomrum)
- Trasig justering (tal vänsterställda, rubriker felcentrerade)
- Fel sifferformat (saknad tusentalsavgränsare, fel decimaler, % vs kr)
- Brutna print_area / sidnav-knappar

## Subjektivt (loggas, rörs INTE)
Ny layout, färgval, omdisponering av sektioner, "skulle se snyggare ut om". → "Kräver Joakim".
Osäker på om det är defekt → logga, rör inte. Konservativ bias.

## Guardrails
- Siffrorna heliga: regression grön ovillkorligt.
- Per-flik-commit ⇒ allt trivialt återställbart.
- Rensa strö-`EXCEL.EXE` + `~$`-lock mellan iterationer vid behov.

---

## Ledger (status per flik — arbetsboks-ordning)

| # | Flik | Status | Fix / not |
|---|------|--------|-----------|
| 1 | Översikt | ⏭️ Joakim | 18-sidors utskrift (hero-bild). Print_area-bugg fixbar, men sidblåsningen = designbeslut. Se "Kräver Joakim". |
| 2 | Försättsblad | ✅ klar | Ren — sidnav, titel, sektioner sitter snyggt. Ingen åtgärd. |
| 3 | Indata | ⏳ pågår | — |
| 4 | Kassaflöde | ⬜ kö | — |
| 5 | Finansiering | ⬜ kö | — |
| 6 | Resultat | ⬜ kö | — |
| 7 | Lönsamhetskontroll | ⬜ kö | — |
| 8 | Beräkningslogik | ⬜ kö | — |
| 9 | Dokumentation | ⬜ kö | — |

Status: ⬜ kö · ⏳ pågår · ✅ klar · ⏭️ hoppad (blockerad) · 🟡 fix klar men subjektivt kvar

## Kräver Joakim (subjektiva beslut jag lät bli)

### Översikt — skrivs ut på 18 sidor (bör vara ~2)
**Diagnos (verifierad i riktiga Excel, inte bara render_local):**
- Fliken har `fitToWidth=1`, `fitToHeight=2`, `fitToPage=True`, landscape, inga manuella sidbrytningar, inga extremt höga rader (radhöjder summerar ~1229 pt ≈ 2 sidor). Med dessa inställningar *borde* den bli ~2 sidor.
- Ändå exporterar Excel den till **18 sidor** — fit-to-page honoreras inte. Enda kvarvarande orsak: den **flytande hero-bilden** (`Image 1`, ~900×285 pt, ett trätak/interiör-foto) spränger utskriftsutbredningen; Excel skalar inte ned flytande objekt med fit-to-page som det gör med celler.
- **Separat bugg jag hittade:** print_area var `B1:F50` men layouten har en högerspalt med merges till **kol L** (`B17:L17`, `J21:L21`, `J43:L43` …). Print_area klipper alltså bort hela högerspalten — halva beslutsdokumentets data skrivs aldrig ut.
- **Jag testade** att bredda print_area till `B1:L50`: det fixar bortklippningen (regression förblev grön) men höjde sidantalet 18→**27** (mer innehåll i den uppskalade arean). Reverterade — vill inte shippa sämre utskrift.

**Beslut du behöver ta (designval, därför lämnat):**
1. Hero-bilden: behållas, krympas, om-ankras (`oneCellAnchor`), eller tas bort? Den driver sidblåsningen.
2. När hero-frågan är löst: sätt print_area till full bredd (`B1:L50`) så högerspalten kommer med.
Misstänkt rotorsak att börja i: `round_u_oversikt_v2` (build/iter9.py ~rad 1398) — satte `fitToHeight=2` men uppdaterade aldrig print_area när högerspalten lades till, och hero-bildens storlek/ankare.

## Blockerat (regression rött / kunde inte lösa)
_(inget än)_

## Commit-logg
_(en rad per committad flik)_
