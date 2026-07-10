# POLISH — ifyllnadsupplevelsen på skärm (ONESHOT-POLISH)

Mål: `build/oneshot/Investeringskalkyl_v2.xlsx` ska kunna öppnas av en
fastighetsutvecklare som aldrig sett mallen och fyllas i utan lathund — och
kännas som den mest genomtänkta Excel-fil hen arbetat i. Skärmen är sanningen;
PDF är bivillkor (grind i m5).

Verktyg: `build/oneshot/screenshot_sheets.py` — Excel COM med synligt fönster
(ReadOnly), 100 % zoom, skärmdump per fliks öppningsvy + auto-scrollade vyer.

## m1 — Skärm-audit (2026-07-10)

30 skärmdumpar över alla 10 flikar (öppningsvy + scroll) + programmatisk
inventering (validering, skydd, freeze, selection, fills).

### Fynd — interaktion (programmatiskt verifierat)

| # | Fynd | Allvar |
|---|------|--------|
| I1 | **Inget bladskydd någonstans** — varje formel i hela boken kan skrivas sönder av ett felslag | Kritisk |
| I2 | **3 datavalideringar totalt** (Indata C23, C26:C30, C75:C78) — inga prompts, inga felmeddelanden. År, procent, belopp: helt ovaliderat | Kritisk |
| I3 | **Markören står på A1 i nav-kolumnen på alla flikar** — första intrycket är en markering i mörka sidofältet | Hög |
| I4 | Kassaflöde saknar freeze för årshuvud + radetiketter (B1 = bara nav-kolumnen); Finansiering har D6 (rätt mönster) | Hög |
| I5 | Ingen tab-ordning: utan skydd + olåsta inputceller hoppar Tab bara höger | Hög |

### Fynd — input-språket

| # | Fynd | Allvar |
|---|------|--------|
| S1 | **Tre olika input-språk:** Försättsblad = grå SURFACE + kursiv "Fyll i"; Indata sektion 1–9 = grå boxar; Indata sektion 4–5 = ljusblå FFD6EAF8 | Kritisk |
| S2 | **Hyresobjektstabellen (Indata rad 26–30) — bokens viktigaste inmatning — har ingen inputmarkering alls** (Objektnamn, areor, index, avtalsår, vakans, prod-år, budget: omärkta) | Kritisk |
| S3 | Indata C72 (Investeringsgrad) är en formel men har input-fill — ser ut som något man ska fylla i | Medel |
| S4 | Indata C9/C19/C22 (auto-beräkningar mitt bland inputs) omarkerade — ingen skillnad mot inputs | Medel |
| S5 | Motivering-kolumnen i sektion 9 (E75:E78) är input men tom grå utan ledtext | Låg |

### Fynd — skärmbild per flik

| # | Flik | Fynd | Allvar |
|---|------|------|--------|
| V1 | Resultat | **####### i hero-cellen D14 (Bindande kravhyra) på skärm vid 100 %** — kolumn D för smal | Kritisk |
| V2 | Resultat | Etikett "2. Kravhyra för IRR ≥ avkastningsk…" klipps; Förklaring-texten rad 18 klipps vertikalt | Hög |
| V3 | Alla | Mörka nav-pelaren (kolumn A) slutar abrupt mitt på skärmen (Översikt rad 33, Resultat rad 55, Indata rad 91) | Hög |
| V4 | Indata | Tomma objektrader (27–30) visar nollbrus: `0  0  0,0%  0`; Resultat "(ej ifyllt)"-rader likaså; Kassaflöde INFASNING-nollrader | Medel |
| V5 | Indata | Sektion 5 (Re-investering): blå slab, radgränser knappt synliga, rubrikerna svävar | Medel |
| V6 | Lönsamhetskontroll | Gröna felkontroll-trianglar i scenariotabellen (rad 59–60); etiketter klipps mot värdekolumnen ("…vid bindande kravhyr"); status "IRR<krav" trångt mot F | Medel |
| V7 | Dokumentation | Lösryckta linjefragment i kolumn C–N (rester av unmergade bottom-rules) | Medel |
| V8 | Kassaflöde/Indata | UsedRange långt under innehållet (scrollbar antyder rad 165/137) — tomrums-scroll | Låg |
| V9 | Försättsblad | Auto-hämtade fält (Projektnamn, Kalkylstart…) ser identiska ut med det man ska fylla i | Medel |

Bra bas som behålls: eyebrow/H1-hierarkin, sidenav, ledtext-kursiven i Indata
kolumn G, zebra-tabellerna, KPI-bandet på Översikt, "Fyll i"-kursiven som idé.

### Designplan

- **m2 — Ett input-språk.** BLÅ `FFD6EAF8` + hårlinjekant = "här fyller du i",
  överallt (inkl. hyresobjektstabellen). Grått fill försvinner från inputs;
  beräknat = ingen fyllnad. Datavalidering med svensk prompt (titel = fältnamn,
  text = enhet/intervall/exempel) + svenskt felmeddelande på varje inputfält.
  Bladskydd utan lösenord: `locked=False` på inputs, skydd på alla flikar utom
  Lönsamhetskontroll/Beräkningslogik (bladskydd blockerar outline-expandering
  av de kollapsade blocken — dokumenteras som D-24). Tab vandrar då mellan
  inputfälten. C72 gör om till omarkerad beräkning (S3).
- **m3 — Vägledning & öppningsvy.** Selection på första inputfält per flik
  (Försättsblad C9, Indata C5; övriga hero/B2), Kassaflöde freeze D5,
  nollbrus → tomt/"–", nav-pelaren förlängs (sidenav bg_end_row min 120),
  ifyllnadsstatus på Försättsblad, "tom mall"-test på kopia.
- **m4 — Excellens.** ≥2 hela skärmdump-varv efter sista fixen + programmatisk
  interaktionskontroll (validering/lås/skydd + COM-stickprov: recalc,
  outline-expandering, sidenav-länkar).
- **m5 — Grindar.** Regression grön (varje milstolpe), export_pdf ≤18 sidor,
  dokumentation.

Before-skärmdumpar: `scratchpad/screens_m1/` (session) — nyckelbilder kopieras
till `docs/polish/` i m5.

## m2 — Input-språket (2026-07-10)

Implementerat som `round_input_language` + `round_sheet_protection` i
[build/oneshot/v2_rounds.py](build/oneshot/v2_rounds.py); input-stilen bor i
`tools/theme.py` (`INPUT`, `INPUT_EDGE`, `mark_input()`).

- **Ett input-språk (D-23):** blå `D6EAF8` + hårlinjekant på ALLA inputceller —
  inkl. hyresobjektstabellen (S2 löst), DoU-tabellen och Försättsbladets fält.
  Gråa input-fills ersatta; C72 städad till ren beräkning (S3).
- **51 datavalideringar** med svensk prompt (fältnamn + enhet/exempel) och
  svenskt felmeddelande: år 1990–2100, procent 0–100 %, belopp ≥ 0,
  kalkylperiod 1–25, objektnr 1–5, dropdowns. Verifierat på skärm (prompten
  renderas) och programmatiskt.
- **Bladskydd utan lösenord (D-24):** alla flikar utom Lönsamhetskontroll +
  Beräkningslogik (xlsx-skydd blockerar outline-expandering; deras kollapsade
  block är del av designen). Tab vandrar mellan olåsta inputfält.
  COM-verifierat: låst cell avvisas, input skrivbar, outline expanderbar,
  sidnav-länkar fungerar. Regression grön.
- Semi-inputs: Indata C9 (auto-formel som får överskrivas) olåst utan blå
  fill, med förklarande prompt; underskrifter/bildytor på Försättsblad olåsta
  utan fill.

**Verktygsfynd (viktigt):** Excel målar ALDRIG om rutnätet vid programmatisk
scroll i bakgrundsfönster (ScrollRow/Goto/Select uppdaterar modellen, inte
pixlarna; zoom målar om). `screenshot_sheets.py` löser det genom att baka in
scroll-positionen i en temporär kopia (XML-patch av sheetView/pane
topLeftCell) så öppningsritningen hamnar rätt, och fångar fönstret med
PrintWindow-API:t — fokus stjäls inte och andra fönster kan inte förorena
bilden (CopyFromScreen fångade användarens webbläsare).

Kvar till m3: Motivering-kolumnen E75:E78 för smal för text; nollbrus i tomma
objektrader; öppningsvy/markörposition; nav-pelarens abrupta slut.

## m3 — Vägledning & öppningsvy (2026-07-10)

`round_guidance` + `round_empty_state` i v2_rounds.py; nav-pelaren förlängd i
`tools/sidenav.py` (min 80 rader).

- **Öppningsvy:** markören landar på första inputfältet (Försättsblad C9,
  Indata C5 — DV-prompten visas direkt vid öppning), övriga flikar B2. Filen
  öppnas på Försättsblad. Kassaflöde fryser årshuvud + etikettkolumner (D5).
- **Ifyllnadsguide på Försättsblad** (G16–G21): SÅ FYLLER DU I 1-2-3 med
  hyperlänkar till Indata/Resultat, "Blå fält = inmatning"-förklaring och
  levande nyckelfältsstatus ("Ifyllnad: X av 11 nyckelfält" → "✓ Alla
  nyckelfält ifyllda"; 'Fyll i'-platshållare räknas som tomma).
- **Sektion 9 omlagd:** Motivering fick K:R (var 6 enheter smal i E),
  vägledningstexter F:J.
- **Nollbrus borta:** tomma objektrader i Indata/Resultat/Kassaflöde visar
  tomt i stället för 0/0,0 %; "(ej använd)"-kolumnerna alltid tysta.
- **Tom mall-testet** (kopia, Skola-raden rensad via COM — validerade samtidigt
  att skyddet släpper igenom ifyllnad): 518 #DIV/0! på presentationsytorna →
  0. IFERROR-wrap + Indata!$R$31=0-guards ("–" i stället för "0 kr/år",
  "✗ Nedskrivningsrisk", "MV", "Typ: Befintligt", "Projektnamn 0",
  "· Kalkylstart"). Motorflikarna lämnas medvetet (fel där är normala mitt i
  ifyllnad och försvinner med första objektraden).
- **V1/V2 från m1 löst:** Resultat-hero D14 (#######) → kolumn D 24 enheter;
  B-etiketterna hela (B 36).
- Regression grön efter varje bygge.

## m4 — Design-excellens (2026-07-10)

Granskningsvarv 1 (25 vyer) → fyra fynd → `round_m4_polish` + ignoredErrors:

- Översikt: krav 3-etiketten klipptes mot statuschipen → "MV år 20 ≥ bokfört
  värde". Försättsblad: flödesstegen 2–3 fick länkutseende (understrykning);
  "Marknadsvärde, senaste värdering" → "Senaste marknadsvärde" (klipptes).
- Lönsamhetskontroll: B 34 / F 12 ("…vid bindande kravhyr|", "⚠ IRR<krav"
  klipptes). Gröna felkontrollstrianglar i scenariotabellen släckta via
  `<ignoredErrors>`-XML-patch i `post_save` (openpyxl saknar API;
  numberStoredAsText + formula; Excel bevarar elementet vid recalc-save).
- Dokumentation: lösryckta kantlinjefragment i C:V rensade.

**Verifiering efter sista ändringen:** två hela skärmdumpsvarv (25 vyer × 2).
Varv 1 vision-granskat — inga nya fynd. Varv 2 pixeldiffat mot varv 1: alla
diffar låg i menyfliksområdet (y ≤ 187), arken pixelidentiska; vision-stickprov
rena. Programmatisk interaktionskontroll grön: skyddsmatris, 51 valideringar,
olåsta inputceller, öppningsmarkörer (C9/C5/D5/D6/B2), Kassaflöde-freeze D5,
aktiv flik Försättsblad. Regression grön.

## m5 — Slutrapport ONESHOT-POLISH (2026-07-10)

**Mål uppnått:** filen öppnas på Försättsblad med markören i första inputfältet
och en levande DV-prompt; blå fält + hårlinje är den enda "fyll i här"-signalen
i hela boken; varje inputfält har svensk prompt och svenskt felmeddelande; Tab
vandrar mellan fälten; formler kan inte skrivas sönder; ifyllnadsguiden 1-2-3
med status leder genom flödet; en helt tom mall visar "–" i stället för
518 felkoder.

**Grindar:**
- Regression grön i varje milstolpe (11 631 221,72 / 7,6452 % / +1,35 pp +
  IRR-scenariomatch).
- Print: trogen PDF fortsatt **18 sidor** (== FINAL-baslinjen), inga ####,
  ingen kapad text; blå inputfält skriver ut korrekt (visar ifyllnadsytor
  även på papper). OBS: Pages.Count rapporterar 14 — den ljuger fortsatt,
  PDF:en är sanningen.

**Before/after:** [docs/polish/](docs/polish/) — Indata (omärkt hyresobjekts-
tabell → blått input-språk), Resultat (#######-hero → hel), Försättsblad
(anonym → guide + status), tom mall-Översikt ("–" i stället för felkaskad).

**Kvarstående begränsningar (medvetna):**
- Lönsamhetskontroll/Beräkningslogik är oskyddade (D-24: xlsx-skydd dödar
  outline-expandering; kollapsade block är del av designen).
- Motorflikarna visar #DIV/0! i HELT tom mall tills första objektraden fylls i
  (normalt mitt-i-ifyllnad-tillstånd; presentationsytorna är rena).
- Grafer har ett synligt tomrum mellan diagram 1 och 2 på skärm (chart-ankare
  är print-kalibrerade; flytt riskerar 1-sidslayouten — inte värt det).
- FS "Summa investering 15 000 kvm" räknar in mark-arean (LM 371-arv, D-14).

**Redo att pusha** (Joakim granskar först; inget pushat till origin).
