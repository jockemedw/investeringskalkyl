# ONESHOT — Investeringskalkyl v2, autonom nattkörning

**Start:** 2026-06-13 (natt, UTC ~22:15). Mandat: bygg v2 från grunden, bättre än iter9, regression helig, fråga inget.
**Branch:** `oneshot-v2`. **Baseline vid start:** grön (11 631 221,72 / 7,6452 % / +1,35 pp, commit `91c5679`).

## Arkitekturbeslut (V2-01): spec-replay + v2-moduler

iter9.xlsx är inte reproducerbar från noll — `build/iter9.py` patchar binärbaselinen iter8.xlsx,
som i sin tur bär oåterskapbara lager (build-scripten för iter1-8 är förlorade, HANDOFF §0).

v2 bryter det beroendet: `extract_spec.py` extraherar HELA iter9 (20 137 celler, 2 469 formler,
stilar, merges, dimensioner, page setup, valideringar, kommentarer) till `spec_iter9.json`.
`build_v2.py` bygger arbetsboken från tom `Workbook()` genom spec-replay och applicerar sedan
v2-förbättringar som moduler. **v2 är därmed deterministiskt byggbar från enbart git-innehåll** —
ingen binär baseline krävs. Det är i sig en leverabel förbättring.

Motorflikarna replayas med **identiska celladresser** → `tests/regression.py` fungerar oförändrat
mot v2-filen (mappningen är identitet). `regression_v2.py` wrappar den + lägger till nollfelsscan
(inga #REF!/#DIV/0!/#VALUE!/#NAME?).

## Fynd under spec-arbetet

### F-1: TEXT()-formler är locale-trasiga i svensk Excel (objektiv defekt, ny)
7 formler använder `TEXT(x,"#,##0")`-mönster. Formatsträngen tolkas enligt Excels UI-locale —
i svensk Excel är `,` decimaltecken, så cached values visar skräp **i produktionsfilen idag**:
- `Översikt!B20`: "Investering 200,0 Mkr · … 5000,0 m²" (skulle vara "200 Mkr · 5 000 m²")
- `Översikt!J23`: "Marginal +0.00 pp · Krav 0.6%" (skulle vara "+1,35 pp · Krav 6,3 %")
- `Översikt!F42`: "11631221,719 kr/år"
- `Resultat!C19`: "hyresspann 10726244,344–12536199,095 kr/år"
- Även `Försättsblad!E63`, `Försättsblad!C73`, `Översikt!J29`.

**V2-regel:** aldrig TEXT() med talformat. Tal läggs i egna celler med `number_format`
(locale-oberoende i filformatet); etiketter separat.

### F-3: Två #NAME?-fel i produktions-iter9 (objektiv defekt, ny)
`Beräkningslogik!D18` (`=-npv_0 / b`) och `D30` (`=max av de tre`) — pedagogiska
anteckningar inskrivna som formler. Visas som `#NAME?` i filen idag. v2 lagrar dem som text.
(Upptäcktes av regression_v2:s nollfelsscan — iter9:s regressionstest scannar inte formelfel.)
Teknisk läxa: openpyxl skriver ALLA strängar som börjar med `=` som formler — även vid fix.

### F-4: Stale radreferenser i Beräkningslogiks pedagogik (objektiv defekt, ny)
Annoteringarna pekade på "Kassaflöde C70/C90/C91" och "Block A (rad 10–70)" — blocken ligger
på Beräkningslogik rad 59–204 sedan de flyttades. v2 korrigerar alla sju referenserna.
Dessutom: headerns merge B46:E46 högg texten vid kolumn E (breddad till K), B35 radhöjd
klippte wrappad text (43,5 → 60).

### F-5: INFASNING-tabellen på Kassaflöde är död i produktion (objektiv formelbugg, ny)
Rad 30–34 multiplicerar med `$E$7` = "Area Bef (aktiv) år 2" = 0 → hela tabellen visar 0
trots att Skola tillträder år 1. Felreferens (avsedd bas är ett kr-flöde — formelns
index-term bevisar det). v2: objektets andel × faktisk bruttohyra per år (rad 16) med
tillträdes-flagga; "Andel av fullt flöde" divideras per år. Inget refererar raderna →
regressionssäkert (verifierat: gate grön).

### F-6: Kassaflödes utskrift trasig i produktion (objektiv, ny — render_local dolde den)
####### i ALLA årskolumner utom 2026 (kolumn E..AB defaultbredd 8,43 < 8-siffriga belopp)
och print_area B1:AB156 täckte 120 tomma rader → 6 av 10 sidor blanka. v2: bredd 13,
print_area B1:W36 (år 21–25 syns på skärm, skrivs inte ut). 10 sidor → 2.

### F-2: Hero-bilden är rotorsaken till Översikts 18-sidiga utskrift (bekräftar NIGHTRUN)
Bilden (1786×765 px, OneCellAnchor B1) skalas inte av fit-to-page → sidblåsning.

## V2-beslut (subjektiva, tagna autonomt per mandat)

| # | Beslut | Motivering |
|---|--------|-----------|
| V2-01 | Spec-replay-arkitektur (ovan) | Reproducerbarhet från noll |
| V2-02 | Hero-bilden UTGÅR ur Översikt | Rotorsak till 18-sidersbuggen; tillför ingen beslutsinformation. Cellbaserad accentbanner ersätter. Bilden finns kvar i assets/hero.png om Joakim vill återinföra på Försättsblad. |
| V2-03 | Översikt byggs om: 1 sida liggande, print_area över hela layouten | Löser både sidblåsning och print_area-klippet (B1:F50 klippte högerspalten) |
| V2-04 | Alla 7 TEXT()-ställen ersätts locale-säkert | F-1 |
| V2-05 | Beräkningslogiks tekniska block (rad 48+) visas i **tkr** (`#,##0,`), märkt "Belopp i tkr" per blockrubrik | Löser ########; värden oförändrade (bara format). Pedagogiska delen (rad 1-47) behåller kr. |
| V2-06 | Ny flik **Grafer** (efter Resultat): driftnetto år 1-20, ack. kassaflöde, hyresspann | LM 371 har Grafer-flik; iter9 saknar visualisering. D-14-förankrat. |
| V2-07 | Sidnav regenereras med 10 poster (inkl Grafer) | tools/sidenav.py, NAV_ITEMS utökas |

## Sanningskriterium
Testfall Skola (Nyb) 5 000 kvm × 40 000 kr/kvm = 200 Mkr →
`Resultat!D14` = 11 631 221,72 · `Lönsamhetskontroll!C45` = 7,6452 % · marginal +1,35 pp.
Recalc: `tools/recalc.py` (Excel COM). Gate: `python build/oneshot/regression_v2.py`.

## Ledger

| Milstolpe | Status | Not |
|-----------|--------|-----|
| 1. Spec extraherad + designspec | ✅ | spec_iter9.json: 9 flikar, 20 137 celler, 2 469 formler. Ankare verifierade. |
| 2. Motor (replay) grön | ✅ | Exakt baslinje på första recalc. Nollfelsscan grön (iter9 har 2 #NAME? — fixade i v2, F-3). COM-krångel löst: strö-EXCEL + openpyxl "="-strängfällan. |
| 3. Design: Översikt om, tkr-fix, Grafer, locale-fix | ✅ | Översikt 18→1 sida (ny design, statusblock, ingen hero). Beräkningslogik tkr (1242 celler) + F-4-polish. Kassaflöde F-5/F-6 (10→2 sidor). Grafer-flik med 3 diagram. Sidnav 10 poster. |
| 4. Print: page setup + trogen PDF per flik | ⬜ | |
| 5. Slutrapport + jämförelse | ⬜ | |

## Kända brister / öppet
_(fylls på under natten)_

## Jämförelse v2 vs iter9
_(skrivs i milstolpe 5)_
