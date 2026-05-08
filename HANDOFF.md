# Handoff: Investeringskalkyl Lejonfastigheter AB

**Syfte med dokumentet:** Ge en ny Claude-session (Claude Code eller annan miljö) all kontext, alla beslut och alla lärdomar som behövs för att fortsätta arbetet utan att tappa något av värde. Skrivet 2026-05-06.

---

## 0. STATUS — uppdaterat 2026-05-08

**Iter 8-filen finns i repot:** `Investeringskalkyl_iter8.xlsx` (~83 KB, 8 flikar, ~2 380 formler).

**Regressionsbaslinje verifierad direkt mot filens cached values:**
- `Resultat!D14` = **11 631 221,72 kr/år** ≈ 11 631 222 ✓
- `Lönsamhetskontroll!C45` = **0,07645** ≈ 7,65 % ✓
- `Översikt!C6` / `Översikt!C20` speglar samma värden ✓

**Build-scripten (Python/openpyxl) finns INTE.** Joakim har bekräftat att de inte är åtkomliga. Pågående arbete (fas 2): rekonstruera `build_iter8.py` från xlsx-filen så att pipelinen är reproducerbar igen.

**Det testfall som producerar baslinjen** är `Skola` (Nyb) 5 000 kvm × 40 000 kr/kvm = 200 Mkr, indexandel 70 %, avtal 2026-2045, belåning 37 %, IRR-krav 6,3 %.

**Doc-synk 2026-05-08:** Avvikelser mellan iter8-filen och DECISIONS/TECH_NOTES har dokumenterats i nya beslut D-16 → D-19. Filen är auktoritativ — dokumentation har synkats mot den.

---

## 1. Projekt i ett nötskal

**Beställare:** Lejonfastigheter AB — kommunalt fastighetsbolag i Linköping.
**Användare av verktyget:** fastighetsutvecklare, kund-/hyresgästansvariga, LCC-analytiker.
**Verktygets namn:** Investeringskalkyl (ersättare för det gamla **LM 371 Investeringskalkyl**).
**Syfte:** Värdera fastighetsinvesteringar (nybyggnation, ombyggnad, hyresgästanpassningar) genom att räkna fram **erforderlig hyresnivå** baserat på:
1. NPV (driftnetto-kassaflöde diskonterat)
2. Equity IRR (avkastning på eget kapital över 20 år)
3. Marknadsvärde vs. bokfört värde

**Kalkylhorisont:** 20 år.
**Teknik:** Pure Excel (inga makron) för bred tillgänglighet. Konstrueras programmatiskt med Python/openpyxl. Ingen VBA någonstans.

---

## 2. Projektets faser — vad som har hänt

### Fas 1 — Lathund-uppdatering
Initialt uppdrag: uppdatera dokumentationen ("lathunden") för det befintliga LM 371-verktyget.

### Fas 2 — Reverse engineering av LM 371
Eskalerade till djupgående analys av LM 371-arket:
- ~72 700 formler extraherade
- Buggar identifierade (se §6 nedan)
- **Formelkatalog** byggd som referensdokumentation av hur LM 371 faktiskt räknar

### Fas 3 — Bygg ersättning från grunden
Iter 1 → iter 8, varje iteration adderar struktur eller pedagogik.

**Iterationerna lever som programmatiska byggskript** (Python/openpyxl), inte som manuella Excel-redigeringar. Det innebär att hela arket kan återskapas deterministiskt och regressionstestas.

---

## 3. Aktuell arkitektur (iter 8 — verifierad mot fil)

> Strukturen nedan är verifierad direkt mot `Investeringskalkyl_iter8.xlsx` 2026-05-08.

| Flik | Roll | Innehåll |
|------|------|----------|
| **Översikt** | Dashboard | Innehållsförteckning, beskrivningar, hyperlänkar (interna) |
| **Indata** | Användarinput | Alla fält där användaren matar in projektdata, parametrar, antaganden |
| **Kassaflöde** | Driftnetto | Ren projektion: bruttohyra → vakans → nettohyra → drift/FS/TR/CA → DRIFTNETTO. **Stannar vid driftnetto** (matchar LM 371) |
| **Finansiering** | Block C | Dedikerad finansieringsflik. Lån, ränta, amortering, ränteupptrappning |
| **Resultat** | Output | Sammanställning av nyckeltal |
| **Lönsamhetskontroll** | Verifiering | NPV, equity-IRR, marknadsvärde vs. bokfört värde — godkänd/ej godkänd |
| **Beräkningslogik** | Motor | Delta-metod, Block A, B, D, E, F |
| **Dokumentation** | Förklaring | Hur kalkylen är uppbyggd, antaganden, metodik |

### Rollseparationen är viktig
LM 371 hade allt blandat. Iter 8 separerar ut:
- **Vad man matar in** (Indata)
- **Hur kassaflödet projiceras** (Kassaflöde, stannar vid driftnetto)
- **Hur finansieringen påverkar** (Finansiering, separat flik)
- **Hur det utvärderas** (Lönsamhetskontroll)
- **Vad användaren ser direkt** (Översikt + Resultat)
- **Den underliggande motorn** (Beräkningslogik)

---

## 4. LM 371 — referensens struktur (för jämförelse)

LM 371-filen i projektet (`LM_371_Investeringskalkyl_2.xlsx`, ~810 KB) har följande flikar:

| Flik | Storlek (rader×kol) |
|------|---------------------|
| 1. Framskrivningsunderlag | 132 × 25 |
| 2. Kalkyldata | 50 × 79 |
| 3. Finansiering | 61 × 61 |
| 4. NPV | 83 × 54 |
| 5. IRR | 50 × 53 |
| 6. Ack.kapitalutlägg | 14 × 53 |
| 7. Grafer | 57 × 16 384 |
| Kontoplan RR | 37 × 9 |

LM 371 är den **auktoritativa industrireferensen** för svensk fastighetsinvesteringskalkyl — alla strukturella och metodiska val i iter 8 ska kunna motiveras mot LM 371. Avvikelser kräver explicit motivering.

---

## 5. Nuvarande regressionsbaslinje (iter 8 mot iter 7)

Dessa siffror **MÅSTE** gå att reproducera efter varje vidare ändring. Om de avviker = regression.

| Mätvärde | Värde |
|----------|-------|
| Bindande erforderlig hyra | **11 631 222 kr/år** |
| Faktisk equity-IRR | **7,65 %** |
| Marginal mot IRR-krav | **+1,35 procentenheter** |

Använd LibreOffice headless för formelutvärdering (samma verktyg har använts hela vägen).

---

## 6. Bekräftade buggfixar i LM 371 (verifierade i iter 1)

Tre buggar i originalverktyget — alla **verifierade mot syntetiska testfall** med LibreOffice headless.

### Bug 1: NPV-intäktsindexering år 1
**Problem:** År 1 använde hårdkodad exponent **0** i stället för "år sedan kontraktsstart". Det innebar att indexuppräkningen skippades första året även när kontraktet startat tidigare.
**Fix:** Räkna ut `years-since-contract-start` dynamiskt och använd som exponent.

### Bug 2: Negativ exponent inflaterar färdiga investeringar
**Problem:** Sidoeffekt av samma indexeringsmodell. För investeringar som redan slutförts gav formeln **negativ exponent**, vilket inflaterade värdet i stället för att låta det vara.
**Fix:** Klampa exponenten vid noll (`MAX(0, exponent)`) för redan färdigställda investeringar.

### Bug 3: LOOKUP-vector storlek i tomtavgäld
**Problem:** Formeln för tomtavgäldsavgift använde `LOOKUP` med vektorer av olika längd → odefinierat beteende i kanten.
**Fix:** Matchande vektorlängder.

**Alla tre verifierade** mot syntetiska testfall, inkl. edge cases:
- Parallella kontrakt
- Investeringar i etapper

---

## 7. Tekniska principer och fallgropar

### Kassaflödets korrekthet
**Principen:** En kassaflödesflik ska **stanna vid driftnetto** enligt LM 371. Att inkludera avskrivningar och räntor i kassaflödesfliken är **konceptuellt felaktigt** för den fliktypen. Finansiering hör hemma på en **dedikerad Finansiering-flik**.

> Iter 8 följer denna princip: Kassaflöde-fliken stannar vid driftnetto, Finansiering-fliken hanterar lån/ränta/amortering/avskrivning separat (verifierat 2026-05-08).

### openpyxl XML-patchning
- `outlinePr summaryBelow=False` **persisterar inte** genom vanlig attributtilldelning i openpyxl. Krav: direkt XML-patch via Pythons `zipfile`-modul mot relevanta sheet-XML-filer.
- Excel row grouping (outline levels) implementeras med `outlineLevel=1` på rader + XML-patchen ovan. Används för kollapserbara sektioner (t.ex. inmatning av investeringar i etapper).

### Hyperlänkar i openpyxl
Måste använda `Hyperlink`-objektet med `location=`-parameter:

```python
from openpyxl.worksheet.hyperlink import Hyperlink
cell.hyperlink = Hyperlink(ref=cell.coordinate, location="'Indata'!A1", display="Gå till Indata")
```

**Funkar INTE:** sträng-tilldelning (`cell.hyperlink = "'Indata'!A1"`).

### Dokumentredigeringsprincip
**Modifiera alltid originalfiler kirurgiskt** — bygg INTE om från grunden om det går att undvika. (Undantag: hela iter 8 är ett undantag eftersom det är en avsiktlig replacement, men inom iter:n redigeras kirurgiskt mellan körningar.)

---

## 8. Arbetsstil & samarbetsmönster med Joakim

Detta är minst lika viktigt som det tekniska — Joakim har en specifik approach och avvikelser från den skapar friktion.

| Mönster | Innebörd |
|---------|----------|
| **Plan först, exekvering sen** | Föreslå explicit plan innan ändringar görs. Joakim godkänner i batch. |
| **Batch-godkännande** | Han debatterar inte enskilda val ett i taget — han approvar (eller invänder) i klump. |
| **Direkt pushback vid fel** | Vid faktafel: rätta omedelbart, inga utdragna ursäkter eller motiveringar. |
| **Tekniska sub-beslut är Claudes** | Småval inom en plan ägs av Claude. Flagga bara där Joakim faktiskt behöver välja. |
| **Motiveringar mot industri-referens** | Designval ska kunna försvaras mot LM 371 / branschpraxis. |
| **Svenska, koncist** | All kommunikation på svenska, hellre kort än långt. |
| **Regressionstest som validering** | Varje iteration verifieras mot kända baseline-värden — inte mot subjektiv bedömning. |

---

## 9. Verktyg & resurser

| Verktyg | Roll |
|---------|------|
| **Python + openpyxl** | Konstruktion av arbetsboken + XML-patchning |
| **LibreOffice headless** | Formelutvärdering, regressionstest |
| **LM 371** | Auktoritativ industrireferens — `LM_371_Investeringskalkyl_2.xlsx` ligger i projektet |
| **Formelkatalog** | Från reverse engineering — bör finnas lokalt hos Joakim |
| **Excel row grouping** | UX-mekanism för kollapserbara sektioner |

---

## 10. Att-göra-lista (loggad vid sessionsslut)

Punkter som var öppna när minnesbilden skrevs. **Ordningen är inte fixerad** — Joakim väljer nästa.

- [ ] **Pedagogisk omskrivning av Beräkningslogik-text** (round B)
- [ ] **Indata-fält renamning** — terminologi: `Avskrivningstakt`, `räntenivå` (matcha branschspråk)
- [ ] **Designcleanup av Lönsamhetskontroll** (round F)
- [ ] **Ersätt "År N" med "år 20"** (round E) — explicit horisont, lättare att läsa
- [ ] **Full design review** av hela dokumentet inkl. Översikt

Inga av dessa är blockerande för funktionalitet — de är kvalitets- och pedagogikförbättringar.

---

## 11. Setup-instruktion för ny Claude Code-session

Steg-för-steg för att komma igång på Joakims dator:

1. **Klona/öppna projektmappen** med Python-scripten (de borde ligga lokalt — minnet säger explicit "byggt i Python/openpyxl"). Mappen bör innehålla:
   - Build-script (något i stil med `build_iter8.py`)
   - Regressionstest-script
   - Senaste iter-xlsx
   - Eventuell Formelkatalog från reverse engineering-fasen

2. **Skapa virtualenv** och installera beroenden:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install openpyxl
   ```
   LibreOffice headless installeras separat (paketnamnet beror på OS).

3. **Kör regressionstestet** mot iter 8 — innan något ändras. Förväntad output:
   - Bindande hyra: **11 631 222 kr/år**
   - Equity-IRR: **7,65 %**
   - Marginal: **+1,35 pp**
   
   Om dessa inte stämmer: arbetsstationen är inte i synk med minnesbilden.

4. **Plocka första uppgift** från §10. Joakim väljer.

5. **Ny Claude-session börjar med:** "Läs `HANDOFF.md`, kör regressionstestet, rapportera resultat." Innan något annat.

---

## 12. Sanity-checks för ny session

Innan du tar tag i ny utveckling, verifiera:

- [ ] Iter 8-filen finns och har 8 flikar (Översikt, Indata, Kassaflöde, Finansiering, Resultat, Lönsamhetskontroll, Beräkningslogik, Dokumentation)
- [ ] Build-scripten kör utan fel
- [ ] Regressionstestet ger 11 631 222 / 7,65 % / +1,35 pp
- [ ] LM 371-referensen finns tillgänglig
- [ ] Du har läst denna handoff i sin helhet

Om någon av dessa fallerar: **stopp**. Reda ut innan vidare arbete.

---

## 13. Bilagor i detta paket

1. **HANDOFF.md** (det här dokumentet)
2. **DECISIONS.md** — designval och deras motiveringar, samlade på ett ställe
3. **TECH_NOTES.md** — tekniska detaljer (openpyxl-fallgropar, XML-patch, hyperlänkar) i kodbar form
4. **OPEN_QUESTIONS.md** — sådant som behöver klargöras med Joakim när ny session startar

---

*Slut på huvudhandoff. Se övriga filer för detaljer.*
