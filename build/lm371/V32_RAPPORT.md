# LM 371 v3.2 — användarvänlighet och design (morgonrapport)

Datum 2026-10-11/12. Produkt: `build/lm371/LM371_v3.xlsx` (nu v3.2). Kod: `step8_design.py`.
Utskrift (LibreOffice-rendering): 15 sidor (v3.1: 16); PDF:en ersatt av `print_preview_lo_v33.pdf` i v3.3.

## Kort

v3.2 är fortfarande LM 371. Flikar, rader, inmatningsceller, färghierarki och alla siffror är oförändrade.
Regressionsgaten är grön, och kravhyra 2 559 kr/kvm, IRR 8,08 % och NPV 10,3 Mkr är bit för bit som i v3.1.
Förbättringarna gäller tre områden:

- **Mallen förklarar sig själv.** Ungefär 70 inmatningsmeddelanden (gula rutor när man klickar i ett fält) finns
  på alla fält man fyller i. Flik 2 har fått en arbetsgång och underrubriker per sektion. Grå noter visar var
  siffror kommer ifrån och vart de går, till exempel "→ 5. IRR" på flik 3 och "Hämtas från 2. Kalkyldata" på flik 1.
- **Samma ord överallt.** Krav 1/2/3 namnges likadant på flik 1, 2, 4 och 5. Räntor, direktavkastning och IRR
  visas med två decimaler. Stavfel och engelska rester är rättade (Marknadsituation, Amotering, Driftentto,
  netnuvärde, "(+(", CASHFLOW, EBITA, "Indata sektion 5", "yield").
- **Utskriften blir begriplig.**
  - Flik 1 delar inte längre tabellen Kalkylantaganden mellan sidorna.
  - Flik 4 har årsrubrik på varje sida och sidordningen över-sedan-ned. Beslutssiffran, krav 1 på rad 82,
    är märkt.
  - 7.Grafer ryms på en sida och har rubrik och förklaring.
  - Kontoplan skrivs ut i läsbar skala.

## Så gick det till

1. **Granskning:** åtta agenter granskade var sin flik i detalj mot renderingar, strukturdump och originalmallen,
   och en granskade helheten ur nybörjar- och chefsperspektiv. Det gav cirka 150 fynd.
2. **Syntes:** en agent sammanställde 44 ändringar (V32-01 … V32-44, 20 med prioritet 1), 13 beslutspunkter och
   17 uppskjutna förslag. Varje fynd kontrollerades mot underlaget.
3. **Referenskontroll:** innan bygget kontrollerades att ingen formel, validering eller namngivet område läser
   någon av de 229 celler som ändras.
4. **Bygge och regression:** filen byggdes via XML (inga Excel-delar förlorade) och regressionen var grön.
5. **Verifiering:** sju agenter jämförde v3.1 och v3.2 sida för sida. De hittade cirka 30 nya småfel, som klipp,
   krockar och toner, och alla är åtgärdade. Därefter ombyggt, gaten grön igen, och jag har granskat alla sidor.

## Ändringar per flik (urval — fullständig lista i `step8_design.py`)

| Flik | Viktigast |
|---|---|
| 1. Framskrivningsunderlag | Sidbrytning före Kalkylantaganden. Typ av projekt-listan sitter på rätt cell (G11; ärvt fel). Inmatningsmeddelanden visas (de var avstängda på bladet). Krav 1–3 i resultatblocket och kolumnen "Utfall". IRR med två decimaler. Kravhyreraden i 11 pt fet. BILD/KARTA har blivit diskreta hjälptexter. Tomma objektrader visar inte 0. Bilagorna har flikarnas namn. |
| 2. Kalkyldata | Arbetsgång överst (rad 1–2, syns på skärm). Underrubriker i sektion 2 och 3. Krav 1–3 i nyckeltalsrutan; PV-raden i tkr. Projektinfo märkt "hämtas från flik 1" med grå speglar. Inmatningsmeddelanden i hyrestabell, drift, reinvestering, investeringsmatris och förutsättningar. Rätt toner i reinvesteringsblocket. Enheter i rubrikerna. Sektionsnumren med i utskriften. Antagandefältet är ett riktigt textfält. |
| 3. Finansiering | Ingress som säger var antagandena matas in. Inmatningsmeddelanden på C7–C10. Etiketter som säger vad raderna visar. "→ flik"-hänvisningar. Hjälprad (rad 12) och Residual (rad 36) dolda. Nollor visas som "–". |
| 4. NPV | Ingress som säger att beslutet läses på rad 82 (krav 1). Hela fastigheten = krav 3. Rad 59 heter "NPV hela fastigheten" och har en not om att den är informativ. Radnamn med enheter och rättade stavfel. Utskrift med årsrubrik, sidordning över-sedan-ned och utan den tomma kolumnen för år 22. |
| 5. IRR | Exit-blocket har rätt rubrik, vikterna har inmatningston och överstyrningen C11 förklaras ("ersätter helt"). Kassaflödesdelen är på svenska och har enheter. Känslighetstabellen har tydliga rubriker; hjälpraderna 58–61 är dolda. |
| 6. Ack.kapitalutlägg | Ingress och förklaring av tecknen. Raderna säger vad de innehåller. Den döda raden Räntebidrag är dold. Nollor visas som "–". |
| 7.Grafer | Rubrik och förklaring, enhetskolumn, exakta etiketter, källa per rad som inmatningsmeddelande, en utskriftssida och låst årsrubrik. |
| Kontoplan RR | Rubrik och syfte, kopplingen till rader på flik 2, rubrikerna på samma rad och läsbar utskriftsskala. |

## Togs bort efter verifieringen

- **Låsta etikettkolumner på flik 3–7** (rubrikkolumnen står kvar när man bläddrar åt höger). Det klipper
  sammanslagna ingresser och långa noter i kolumn B på skärmen. v3.1:s radlåsning behålls. Se beslut 6.
- **Att dölja år 22 och senare.** Kalkylperioden kan vara längre än 20 år (den blir avskrivningstiden när Q17 är
  tom), så kolumnerna visas. Nollor efter perioden visas tomma på flik 6.

## Beslut som ingår i v3.2 (presentationsbeslut, lätta att backa)

Syntesen lämnade 13 beslutspunkter. Följande gäller bara presentationen, och jag har genomfört dem enligt
rekommendationen:
- **Utskrift:** bara kolumn A:B upprepas på flik 4, så 10-årskolumnen står inte ensam på sida 2. 7.Grafer får
  plats på en sida.
- **Nollor:** dolda på flik 1 (objektrader) och visade som "–" på flik 3 och 6.
- **Legacy-rader dolda:** Räntebidrag (flik 6) och den ofullständiga Media/Övriga-uppdelningen (7.Grafer rad 53–56).
- **Toner:** vikterna på 5. IRR (A6:A8) har inmatningston.
- **Kravhyreraden på flik 1** är lyft till 11 pt fet.
- **NPV-flikens rad 59** heter "NPV hela fastigheten" och har noten "Informativ, ej krav".
- **5. IRR:** "EBITA" är rättat till "Resultat före avskrivningar (EBITDA)".
- **Källhänvisningar** i den tomma kolumn C på flik 3.
- **Flik 2:** rubriken i projektinforaden visar projektnamnet från flik 1 i stället för "0".
- **BILD/KARTA** på flik 1 har blivit diskreta grå hjälptexter.

## Beslut för dig

1. **Formelrunda.** Granskningen hittade fyra ärvda formelfel som designrundan inte fick röra:
   - Flik 1 G22 delar med tomma D22 och visar därför alltid 0 kr/kvm.
   - 2. Kalkyldata T34 räknar upp överstyrningen i U34 med inflation, men det gör inte T35:T39.
   - 7.Grafer C53 och C55 har omvänt tecken år 1.
   - Flik 6 rad 9 ingår i Ack. kapitalutlägg men inte i Kassaflöde.

   Rekommendation: en kort formelrunda med regression. Ingen av dem påverkar kraven.
2. **Platshållaren "Fyll i"** på flik 1 och dess speglar på flik 2 är kvar. Rekommendation: behåll den tills
   användarna har reagerat på v3.2.
3. **Terminologi som bör stämmas av mot er budgetmall:**
   - 2. Kalkyldata C60–C62 ("Avgifter & byggherrekostnader", "Finansiering", "Index").
   - Investeringsspecifikationens "dagens penningvärde", eftersom beloppet hanteras som nominellt vid
     produktionsavslut.
   - Listvalet "N/A" i Kontoplan.
4. **Flik 1 F24, "Därutöver planerat underhåll",** är en konstant 0 som går in i Total projektvolym. Ska den fyllas
   i (då ska den ha inmatningston) eller hämtas från flik 2 (då blir det en formelfråga)?
5. **Restvärdestabellen på flik 1:** ska de mörka kolumnerna (Restvärde och NPV) bli vanliga och raden "Bedömt"
   markeras i stället? Det är inte gjort, eftersom det är en smaksak.
6. **Låsta etikettkolumner (se ovan).** Det går att göra om de långa noterna i kolumn B flyttas eller bryts, men
   arbetet är större än ändringen är värd. Rekommendation: låt bli.
7. **Kontoplanens färger** (mörkblått med vit text, som i originalet) är behållna. Rekommendation: behåll.

## Verifiering och gränser

- **Regressionen** (`build_xml.py`, som kör `regression.py` via LibreOffice) är grön på alla 9 kontroller,
  inklusive källans baslinje. Den bygger inte om v3.1-gaten utan kör samma gate mot den nya filen.
- **Rendering:** alla 15 sidor är renderade och granskade, först av sju agenter och sedan av mig.
- **Kräver Excel för att bekräfta:**
  - Att filen öppnas utan reparationsdialog. v3.1 byggdes på samma sätt och öppnades rent.
  - Att inmatningsmeddelandena visas när man klickar.
  - Excels exakta sidbrytning.

  Excel räknar om hela boken när filen öppnas och frågar därför om du vill spara när du stänger. Det är väntat.
