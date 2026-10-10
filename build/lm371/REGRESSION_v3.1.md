# LM 371 v3.1 — regressionsrapport, restvärdesmodell efter granskning

Datum 2026-10-09. Beslut A–E (Joakim), dokumenterade som ANALYS_OMTAG beslut 11–15 och D-27.
Kod: `step7_restvarde.py` (körs efter fas 2, före design), `regression.py` (nytt test `check_restvarde_v31`).

**Status:** `LM371_v3.xlsx` är byggd till v3.1 utan Excel (`build_xml.py`, se sist). Regressionsgaten är grön
mot den färdiga filen: alla 9 kontroller, inklusive källans baslinje.

## Antaganden — mallexemplet (oförändrade)

| Parameter | Cell | Värde |
|---|---|---|
| Objekt | Kalkyldata rad 23 | Skola, nybyggnad, 6 100 kvm BRA |
| Investering | sektion 4 | 240 000 000 kr |
| Hyra | I23 | 2 650 kr/kvm, index 70 % av inflation |
| Avtal | M23:N23 | 2026–2045 (= kalkylperioden) |
| Kalkylperiod | P17 | 20 år |
| Extern direktavkastning | M6 | 6,50 % |
| Kalkylränta driftnetto | M7 | 4,55 % |
| Inflation | M9 | 2,00 % |
| Kalkylränta restvärde | M8 | 8,50 % (v3: M11 + M9, v3.1: M6 + M9 — samma vid Neutralt) |
| IRR-krav | F14 | 6,34 % |
| Restvärdesbedömning | sektion 5 | Neutralt/Normal på alla fyra faktorer |

## Resultat — mallexemplet

| Mått | v3 | v3.1 | Ändring från |
|---|---|---|---|
| Kravhyra krav 1–2 | 2 559 kr/kvm | 2 559 kr/kvm | — |
| Bindande krav | NPV (krav 1) | NPV (krav 1) | — |
| NPV tillkommande, krav 1 | 10 300 tkr | 10 300 tkr | — |
| IRR före skatt, krav 2 | 8,032 % | **8,080 %** | C |
| PV mot bokfört värde, krav 3 | 10 300 tkr | 10 300 tkr | — |
| Restvärde år 20 (IRR-exit) | 268,50 Mkr | **271,94 Mkr** | C |
| Restvärde år 20 (krav 1 och 3) | 271,94 Mkr | 271,94 Mkr | — |
| PV restvärde | 53,20 Mkr | 53,20 Mkr | — |
| IRR vid kravhyran | — | 7,524 % | — |

NPV, PV-kravet och kravhyran är oförändrade eftersom NPV-spåret redan kapitaliserade år N+1. Exemplet
har varken vakans eller extraordinärt UH, och M8 är densamma vid Neutralt. IRR stiger 0,048 pp eftersom
exit nu är samma restvärde som i krav 1, alltså år 21:s driftnetto i stället för år 20:s.

## Känslighet ±1 pp på exit-yield (flik 5 rad 53–57, flik 1 rad 64–70)

| Scenario | Yield | Exit v3 | Exit v3.1 | IRR v3 | IRR v3.1 | NPV v3 | NPV v3.1 | Kravhyra v3 | Kravhyra v3.1 |
|---|---|---|---|---|---|---|---|---|---|
| Optimistisk | 5,5 % | 317,32 Mkr | 321,38 Mkr | 8,675 % | 8,725 % | 32,76 Mkr | **19,97 Mkr** | 2 384 | **2 480** |
| Bedömt | 6,5 % | 268,50 Mkr | 271,94 Mkr | 8,032 % | 8,080 % | 10,30 Mkr | 10,30 Mkr | 2 559 | 2 559 |
| Pessimistisk | 7,5 % | 232,70 Mkr | 235,68 Mkr | 7,505 % | 7,551 % | −4,53 Mkr | **3,20 Mkr** | 2 693 | **2 621** |

Exit och IRR flyttas av C (ca +0,05 pp i alla scenarier). NPV och kravhyra flyttas av A: scenariot ändrar
bara kapitaliseringen, inte diskonteringen. NPV-spannet krymper från 37,3 till 16,8 Mkr. Spannet var
tidigare dubbelräknat.

## Vakansfaktorns skala i exemplet (v3.1, B)

| Bedömning | Δ yield | Justerad yield | Motsv. vakans (kolumn I) | Exit år 20 | IRR | NPV | Kravhyra |
|---|---|---|---|---|---|---|---|
| Låg | −0,25 pp | 6,25 % | ≈ −3 % | 282,81 Mkr | 8,229 % | 12,42 Mkr | 2 541 |
| Normal | 0 | 6,50 % | ≈ 0 % | 271,94 Mkr | 8,080 % | 10,30 Mkr | 2 559 |
| Förhöjd | +0,50 pp | 7,00 % | ≈ 6 % | 252,51 Mkr | 7,803 % | 6,50 Mkr | 2 592 |
| Hög | +1,00 pp | 7,50 % | ≈ 11 % | 235,68 Mkr | 7,551 % | 3,20 Mkr | 2 621 |
| Mycket hög | +2,00 pp | 8,50 % | ≈ 19 % | 207,95 Mkr | 7,105 % | −2,22 Mkr | 2 671 |

Kontroll av kalibreringen: i v3 gav 10 % vakansgrad efter avtalet en NPV på 3,87 Mkr via NPV-spårets
restvärde. I v3.1 ger Hög (≈ 11 % vakans) 3,20 Mkr. Faktorn och motsvarande vakans i driftnettot ger nu
samma storleksordning i alla tre kraven, vilket inte gick med kopplad M8.

## Avvikelser mot baslinjen — en rad per avvikelse

| # | Kontroll / cell | v3-pinne | v3.1 | Beslut | Förklaring |
|---|---|---|---|---|---|
| 1 | IRR `D14` (`V3_IRR`) | 8,0324 % | 8,0804 % | C | IRR-exit = samma restvärde som krav 1 (driftnetto år N+1 / yield), +3,44 Mkr. Avsiktligt brott. |
| 2 | Exit `5. IRR!C10` | 268 500 000 | 271 940 000 | C | `C7 = '4. NPV'!E79`. Avrundning till 10 000 som i NPV (tidigare 100 000). |
| 3 | Legacy-vikter 0/0,4/0,6 (`LEGACY_IRR`) | 6,3062 % | 6,3333 % (+0,027 pp) | C | 0,4 × MV-delen följer exit N+1: 164,10 → 165,48 Mkr. Återger inte längre källmallens 6,2983 % + CA-fix exakt. Accepterat (beslut C). Kravhyra med legacy-vikter 2 656 → 2 652. |
| 4 | `M8` vid en faktor +0,50 pp (`check_adjustment`) | 9,00 % | 8,50 % | A | M8 = M6 + M9 påverkas inte av bedömningen. |
| 5 | Läge Mycket negativt: NPV / IRR / kravhyra | 2 160 tkr / 7,756 % / 2 631 | 6 500 tkr / 7,803 % / 2 592 | A (NPV), C (IRR) | Ingen dubbel diskontering. IRR +0,05 pp från exit N+1. |
| 6 | Känslighet NPV opt / pess | 32,76 / −4,53 Mkr | 19,97 / 3,20 Mkr | A | Se tabellen ovan. |
| 7 | Känslighet kravhyra opt / pess | 2 384 / 2 693 | 2 480 / 2 621 | A | NPV binder i alla scenarier. Mindre NPV-spann ger mindre kravhyrespann. |
| 8 | Känslighet IRR opt / bedömt / pess | 8,675 / 8,032 / 7,505 % | 8,725 / 8,080 / 7,551 % | C | Exit N+1. |
| 9 | Bef-testet, krav 3 `D15` | 7 610 tkr | 8 170 tkr | D | Bef-raden har 5 % vakans och avtal t.o.m. 2035. v3 drog vakansen ur restvärdet för hela fastigheten (320,49 Mkr), v3.1 räknar fullt uthyrt (323,34 Mkr). Krav 1/2 fortsatt oförändrade av Bef. |
| 10 | Formeltext `4. NPV!E50`, `5. IRR!C7` | `$M$11` / `ROUND((` | `$D$4:$BA$89` / `'4. NPV'!$E$79` | C | E50 läser definitionsrad 89, C7 läser E79. |
| 11 | `4. NPV` rad 34 (informativ) | 6,50 % | ca 6,42 % | C | Årets driftnetto / restvärdet vid årets slut (= nästa års driftnetto / yield). Binder inget krav. |
| 12 | `2. Kalkyldata!M8`-kontrollen | "M11 + inflation" | "M6 + inflation" | A | Samma värde (0,085) i exemplet, ny definition. |

Oförändrat och grönt: källans baslinje (10 300 / 6,2983 % / 10 300), NPV/PV-krav, tom-matris-vakten,
Bef → krav 1/2, känslighet bedömt == huvudresultat bit för bit, kravhyra exakt (insatt → NPV 20 tkr),
fas 2 F1–F6, 0 felvärden i synliga områden.

## Beteendeändringar utanför exemplets bas (scenariokörningar)

| Scenario | v3 | v3.1 | Beslut |
|---|---|---|---|
| Vakansgrad 10 %, avtal = kalkylperioden | NPV 3 870 tkr, IRR 8,032 % (bara NPV-spåret påverkat) | NPV 10 300 tkr, IRR 8,080 %, notis "Vakansgrad utan effekt", fältet spärrat | D (+C) |
| Vakansgrad 10 %, avtal t.o.m. 2035 | NPV −6 090 tkr, IRR 7,119 %, exit 236,1 Mkr | NPV 340 tkr, IRR 7,676 %, exit 271,94 Mkr (fullt uthyrt). Vakansen 2036–2045 ligger kvar i kassaflödet. | D |
| Extraordinärt UH 500 kr/kvm år 2046 (N+1) | NPV 8 520 tkr (UH år 21 räknades in i E49/E78) | NPV 10 300 tkr (utanför horisonten, restvärdet normaliserat) | E |
| Extraordinärt UH 500 kr/kvm år 2040 | 10-årsvarianten C78 inkluderade det | C49/C78 = 0, E78 < 0. Restvärde år 10 och 20 oförändrat. | E |

**Korrigering av min tidigare uppgift:** jag skrev att E49 summerade extraordinärt UH "alla år upp till
år 50". Det stämmer inte. Årsraden i flik 4 slutar vid år N+1 (kolumnerna därefter är tomma), så felet var
år N+1 i 20-årsvarianten (E49/E78) och år 11–21 i 10-årsvarianten (C49/C78). Fixen är densamma.

## Verifiering — metod och begränsningar

Steg 1 körde den oförändrade `regression.py` med `comrun` ersatt av en LibreOffice-motor (`lo_gate.py`).
Ops:arna appliceras med openpyxl på en kopia, och kopian räknas om i LibreOffice 24.2 via UNO. Två saker
kontrollerades först:

1. **LibreOffice räknar om v3 till Excels cachade värden**, inklusive kravhyran (D13/D14/D15/D17,
   AD15:AI18). Avvikelserna ligger på 1e-14.
2. **Datatabellen emuleras.** LibreOffice räknar inte om alla kolumner i Excels datatabell AC14:AI16 (AF–AI
   behöll cachade värden). Gaten skriver därför in rad 15/16 som rad 14 vid m = 0 respektive m = 1, vilket
   är Excels definition. Emuleringen reproducerar Excels v3-värden exakt.

Kan bara verifieras i Excel (steg 2):
- datavalideringen på P23:P27 (`probe_dv`, hoppas över i LibreOffice)
- dropdownen G76 och inmatningsmeddelanden
- avmergning/mergning I:U i sektion 5
- format och färger
- utskrift: NPV-fliken får rad 84–90, och sidantalet ska kontrolleras

## Hur filen byggdes — utan Excel

Det ordinarie bygget (`build.py`) öppnar mallen i Excel via COM och kan bara köras på Windows med Excel.
Molnmiljön har ingen Excel, och openpyxl förstör mallens x14-dropdowns. v3.1 byggs därför inkrementellt
direkt i xlsx-paketets XML:

```bash
python build/lm371/build_xml.py <v3-fil> <ut-fil>     # ops från step7_restvarde → XML, cachade värden, gate
```

1. `apply_xml.py` utför samma ops som `apply.ps1`: formler, format, merge, validering och utskriftsområde.
   Bara berörda delar skrivs om. Kommentarer, VML, x14-valideringar, customXml och skrivarinställningar
   kopieras byte för byte. Delade formelgrupper vars master skrivs över materialiseras först.
2. Filen räknas om i LibreOffice med svenskt språkläge och emulerad datatabell. Resultaten skrivs in som
   cachade värden i alla 57 327 formelceller, så filen visar rätt siffror även i läsare som inte räknar om.
3. `fullCalcOnLoad` är satt, så Excel räknar om hela boken när filen öppnas. calcChain är borttagen och
   byggs om av Excel.

Hur tillförlitligt är det? På v3-filen ger LibreOffice samma värde som Excel i 57 018 av 57 022
formelceller. De fyra avvikelserna var decimaltecken i texter, som är åtgärdat med svensk locale, och
dagens datum (NOW). Strukturkontrollen (cellordning, delade formler, merge-överlapp) och gaten körs i
`build_xml.py`.

Det här kan inte kontrolleras utan Excel:
- att Excel öppnar filen utan reparationsdialog
- att utskriften blir exakt Excels sidbrytning
- hur valideringsdialogerna beter sig

Skärmbilden är granskad i LibreOffice-rendering. När bygget nästa gång körs med Excel (`build.py` från
orörd mall) reproduceras samma innehåll av samma ops.
