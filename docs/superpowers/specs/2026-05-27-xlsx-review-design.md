# xlsx-review — visuell feedback-loop för Excel-output

**Datum:** 2026-05-27
**Status:** Design godkänd, plan pågår
**Beslutsfattare:** Joakim Weimar

## Problem

Claude Code genererar .xlsx-filer via openpyxl utan att kunna verifiera hur de faktiskt ser ut. Det leder till layoutdefekter (kapad text, ojusterade kolumner, krockande headers, inkonsistent typografi) som hade upptäckts på en sekund av en människa men som idag bara fångas när Joakim öppnar filen manuellt mellan iterationer.

## Mål

Återanvändbar mekanism som låter Claude "se" en .xlsx-fil som bilder, granska designen mot en checklist, och rapportera konkreta problem. Pure review i MVP — inga automatiska kodändringar.

## Avgränsning

- **Inkluderar:** Render, vision-analys, strukturerad rapportering.
- **Exkluderar:** Auto-fix av build-kod (övervägs som v2). Innehållskorrekthet (regression-testet ansvarar för det). Designkonventioner i sig (de bor i andra skills/CLAUDE.md).

## Arkitektur

### Plats
Global skill: `~/.claude/skills/xlsx-review/`

```
xlsx-review/
├── SKILL.md
├── scripts/
│   └── render.py
└── checklist.md
```

### Trigger
- **Slash-kommando:** `/xlsx-review [<path>]` — utan arg = senaste `.xlsx` i cwd
- **Natural language:** "granska designen i min senaste excel" — Claude triggar via skill description

### Render-pipeline (`render.py`)

1. Öppna xlsx via Excel COM: `win32com.client.Dispatch("Excel.Application")`, `Visible=False`
2. Om filen redan är öppen i Excel: kopiera till temp och öppna kopian read-only
3. För varje icke-dold flik:
   - Print area = befintlig om satt, annars UsedRange
   - Page setup: `Zoom=False`, `FitToPagesWide=1`, `FitToPagesTall=False`, orientation = landscape om bredd > höjd
   - `ExportAsFixedFormat(Type=0, Filename=<temp.pdf>)`
4. Stäng workbook utan att spara
5. För varje temp-PDF: rasterisera till PNG via PyMuPDF (`fitz`), 150 DPI
6. Spara PNG i `<cwd>/.xlsx-review/<basename>/sheet-NN-<sheetname>.png`
7. Skriv JSON till stdout:
   ```json
   {
     "xlsx": "absolut/sökväg/till/källan.xlsx",
     "sheets": [
       {"name": "Indata", "png": "...sheet-01-Indata.png", "hidden": false},
       {"name": "Beräkningslogik", "png": null, "hidden": true}
     ]
   }
   ```

### Skill-flödet (drivs av SKILL.md)

Claude:
1. Resolverar target path (argument eller senaste `.xlsx` i cwd via Glob sorterat på mtime)
2. Kör `python <skill-root>/scripts/render.py <path>` och parser JSON
3. Read varje PNG (vision)
4. Granskar varje flik mot `checklist.md`
5. Skriver strukturerad rapport per flik:

```markdown
## Flik: Indata

🔴 **Kritiskt** — C7 (Kalkylränta): värdet kapas vid 4 tecken pga kolumnbredd 6
   Åtgärd: bredda C till minst 12 tecken eller ändra cellformat till %

🟡 **Förbättring** — B10–B20 (parameterrubriker): inkonsekvent indrag mot B5
   Åtgärd: vänsterjustera alla rubriker eller indragera konsekvent

⚪ **Nit** — A1: tomt med synlig kant
   Åtgärd: ta bort kanten eller fyll med rubrik
```

6. **Stopp** — ingen kodändring. Joakim beslutar nästa steg.

### Default checklist (`checklist.md`)

Grupperade granskningspunkter:

- **Läsbarhet:** kapad text, för smala/breda kolumner, för låga rader
- **Konsistens:** font-family/size över flikar, talformat (decimaler, %), datumformat
- **Hierarki:** fetstil på rubriker och summarader, indrag, färg som signal
- **Alignment:** headers över rätt kolumner, vertikal justering konsekvent
- **Merged cells:** krockar med innehåll, alignment problem
- **Whitespace:** ovanligt stora tomma områden, dubbla tomma rader
- **Print area:** täcker det relevanta, inga halva tabeller, orientation rimlig
- **Färgkodning:** systematisk (input/output/formel) eller godtycklig

## Dependencies

- `pywin32` — Excel COM (Joakim har det redan, ref `settings.local.json`)
- `PyMuPDF` — pure-Python, `pip install PyMuPDF`

## Edge cases

| Scenario | Beteende |
|----------|----------|
| Fil redan öppen i Excel | Kopiera till temp, öppna kopia read-only |
| Excel inte installerat | Fel-meddelande med installationsvägledning |
| Hidden sheets | Hoppas över, flaggas i JSON med `"hidden": true` |
| Tomma flikar (UsedRange = A1) | Render ändå, rapport noterar "tomt ark" |
| Mycket stora flikar (>1000 rader, >50 kol) | Render hela, men varning i stdout om PNG > 5 MB |
| Excel COM-konflikt (krasch under render) | Försök kill orphan Excel-processer, ny isolerad instans |

## Integration

I projektets CLAUDE.md läggs en kort sektion till:

> Efter `build/iterN.py` + `tests/regression.py`: kör `/xlsx-review build/iterN.xlsx` för designgranskning innan commit.

## Stoppvillkor (MVP)

Ingen iteration. Joakim läser rapporten och beslutar.

## Framtida utbyggnad (ej i scope)

- **v2:** Auto-fix loop — Claude föreslår kod-patches, Joakim approvar i batch, regenerar
- **v3:** Hook efter `Write|Edit` på .xlsx-filer för automatisk granskning
- **v4:** Cache av tidigare rapporter för diff mellan iterationer

## Risker

| Risk | Hantering |
|------|-----------|
| Excel COM hänger på Joakims maskin | Timeout 60s per workbook, kill instans, returnera felkod |
| Vision missar problem som är tydliga för människa | Checklist iteras över tid; eftersom Joakim ändå granskar är detta ok i MVP |
| Token-kostnad för bildanalys | 8 flikar × 1 PNG = 8 bilder. Default Sonnet räcker. Joakim kan välja Opus när det är värt det |
| PyMuPDF licens | AGPL — men endast som dependency, ingen distribution. Ok för intern verktygsanvändning |
