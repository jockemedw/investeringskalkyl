# Analys — omtag: förbättringarna in i LM 371 i stället för ersättare

Datum: 2026-09-30. Underlag: skill `lf-investeringskalkyl` (lathund v2.0, vanliga fel, kända
brister), DECISIONS D-01–D-25, `restvardets_kanslighet_handover.md`, v2-koden, samt
direktinspektion + COM-omräkning av `LM 371 Investeringskalkyl.xlsx` (72 690 formler).

## Slutsats i tre rader

1. Av allt som byggts i iter 1–v2 är det **fyra** saker som flyttar beslutsunderlaget: restvärdesmodellen, kravhyra som output, känslighet på restvärdet och verifierade formelfixar. Resten (designsystem, sidenav, Grafer, print 18 sidor, öppningsvyer, bladskydd) är det hörn vi målat in oss i och lämnas.
2. Alla fyra går in i LM 371 utan att ändra flikstruktur, cellplacering för inmatning eller arbetsflödet i lathunden. Restvärdesmodellen är den billigaste: 52 formelceller pekas om plus en ny inmatningsruta på flik 2.
3. Byte till rent marknadsvärde är **inte** neutralt: på mallens eget exempel går IRR från 6,30 % till 8,03 % vid samma hyra. Det måste beslutas medvetet, inte smygas in.

## Nuläge

| | LM 371 (produktion) | v2 (repo) |
|---|---|---|
| Restvärde IRR | 0 × ack.kapital + 0,4 × MV + 0,6 × BV (vikter i `5. IRR!A6:A8`, manuell override `C11`) | Rent MV via Gordon, yield-kalibrerad (Indata sektion 9) |
| Restvärde NPV / PV-krav | Rent MV (`4. NPV!C50/E50` med `M6`) | Samma yield som IRR |
| Kravhyra | Goal Seek × 3, manuellt | Analytisk, MAX av tre krav |
| Känslighet restvärde | Saknas | Opt/bedömt/pess med kravhyra + IRR per scenario |
| Kända formelfel | År 1-exponent ^0, CA-indexhopp, 16 438 `#REF!` (IFERROR-inkapslade) | Fixade |
| Skydd, validering | Inget skydd, 1 datavalidering + x14-dropdowns | Skydd + 51 valideringar |

Inkonsistensen mellan NPV-spåret (rent MV) och IRR-spåret (viktat) i LM 371 är verifierad i
formlerna. Samma projekt värderas olika i krav 1/3 och krav 2 idag.

## Rankning — vad ger mest per insats i LM 371

| # | Förbättring | Värde | Insats i LM 371 | Hur | Beslut |
|---|---|---|---|---|---|
| 1 | **Restvärdesmodell** (rent MV, yield-kalibrering, samma yield i alla tre krav) | Högst — enda krav Joakim ställt | Liten | Ny ruta "Restvärdesbedömning" på flik 2 (4 dropdowns → justerad yield). Peka om `7.Grafer!C44:AZ44` (50 celler), `4. NPV!C50/E50`, visning `1. Framskr.!G68/K68`. Vikter `A6:A8` → 0/1/0 som default, kvar som inmatning. `C11`-override behålls. | Ja: vikter, stegstorlek, M8 |
| 2 | **Känslighet restvärde** (opt/bedömt/pess: exit-värde, IRR) | Högt — visar hur mycket beslutet hänger på yielden | Liten | Tre rader under `5. IRR` rad 37: rad 37 + Δexit i exit-årets kolumn (samma grepp som v2 D-20). IRR per rad. Tabell på flik 1 i befintligt block "Kontroll restvärde" (L108–L112). | Nej |
| 3 | **Verifierade formelfixar** | Medel — påverkar bara vissa fall | Liten per fix | Se lista nedan. Bara det som är verifierat i denna fil. Rad i flik 1:s uppdateringshistorik (kol R). | Nej |
| 4 | **Kravhyra som output** (analytisk, MAX av tre) | Högt — tar bort Goal Seek × 3 | Medel | Hyresmultiplikator m in i `L23:L27`; en envariabels datatabell (What-If, ren xlsx) ger kassaflöden vid m=0 och m=1 → alla tre krav linjära i m. Kräver ingrepp i kärnformler; MROUND på `4. NPV!C48/C50` och `5. IRR!C6:C10` bryter strikt linjäritet (räkna oavrundat, avrunda i visning). Fem objekt med olika typ kräver definition av vad som skalas (beslut 6). Datatabeller kräver "Automatiskt utom datatabeller" i en 72 700-formlers bok. | Ja: fas 3 eller ej |
| 5 | Kravhyra per scenario i tabellen ovan | Medel | Medel (kräver 4) | Följer gratis när 4 finns. | — |
| 6 | Bladskydd + valideringar | Låg | Medel, med kostnad | Skydd dödar outline-grupperna på flik 1 ("Tryck på +") och Goal Seek-kulturen. Rekommenderar **nej**. | Nej |
| 7 | Design, sidenav, Grafer-flik, print, öppningsvyer, tom mall-guards | Låg för beslutet | Hög | Lämnas. LM 371 har redan sitt ljusblå input-språk. | Nej |

## Restvärdesmodellen — konsekvens och koppling

**Kvantifierat på mallens exempel (Skola 6 100 kvm, 240 Mkr, yield 6,5 %):**

| | Exit-värde år 20 | IRR före skatt | NPV | PV vs BV |
|---|---|---|---|---|
| Idag (0/0,4/0,6) | 164,1 Mkr | 6,30 % | 10 300 tkr | 10 300 tkr |
| Rent MV (0/1/0) | 268,5 Mkr | **8,03 %** | oförändrat | oförändrat |

Mallens eget IRR-krav i exemplet (`2. Kalkyldata!F14`) är 6,34 %. Idag faller exemplet på kravet med en hårsmån; med rent MV passerar det med +1,7 pp marginal.

Krav 1 och 3 använder redan rent MV, därför oförändrade. Krav 2 lättar med ~1,7 pp. Det betyder
att projekt som idag faller på IRR kommer att passera. Hanteras genom att yield-kalibreringen
finns (kan höja yielden upp till +2,00 pp) och genom att avkastningskravet stäms av mot
intranätet, inte genom att vi låter viktningen ligga kvar.

**Koppling, konkret:**
- Ny cell `Justerad direktavkastning = M6 + Σ(effekt av 4 bedömningar)`. `M6` förblir extern yield (input).
- `7.Grafer!C44:AZ44` (marknadsvärde per år, matar `5. IRR!C7` och `4. NPV!D34`), `4. NPV!C50/E50`: byt `$M$6` → justerad cell.
- `M8` (kalkylränta restvärde = yield + inflation): **förslag att den följer justerad yield** — en fastighet med högre riskyield ska också diskonteras hårdare. v2 lät den ligga på extern yield; det var ett oreflekterat val.
- `2. Kalkyldata!G14` varnar redan "Eget restvärde ifyllt i 5. IRR" — utöka med "Yield justerad ±x pp".
- Motivering per bedömning i fritext bredvid dropdownen (auditerbarhet, spec-fråga 2).

**Regression:** pinna `D13 = 10 300`, `D14 = 6,298 %`, `D15 = 10 300` från dagens fil. Med
vikterna satta till 0/0,4/0,6 och alla bedömningar Neutralt ska nya versionen reproducera dessa
exakt. Det är både säkerhetsnät och "känner igen sig"-argumentet.

## Formelfel — vad som är verifierat i denna fil

| Fel | Status | Effekt |
|---|---|---|
| År 1-indexering `^(D$4-$D$4)` = ^0 (`4. NPV!D12:D16`) | Verifierat i formeltext | Fel bara när avtal startat före kalkylstart (Bef-objekt) |
| CA-index: `D26` utan index, `E26` `^2` (år 2 hoppar från ^0 till ^2) | Verifierat mot cachade värden (488 000 → 507 715 = ×1,02²) | CA överskattas ~2 % från år 2, konservativt (skill-referensen anger motsatt riktning — uppdateras) |
| 16 438 `#REF!`-formler (`5. IRR!D10:BA10`, `7.Grafer!54`, `2. Kalkyldata!L39:N39`, m.fl.) | Verifierat | IFERROR-inkapslade, visar 0/tomt. Städas men är inte ett arbetspaket |
| Negativ exponent vid färdig investering, LOOKUP-vektorer olika längd, MROUND på negativa tal | **Ej verifierat** i denna version | Kontrolleras innan fix påstås |

## Teknisk förutsättning som styr allt

openpyxl-rundtur av LM 371 är **destruktiv**: x14-datavalideringar (Typ-dropdowns, Kontoplan RR-
listor) tas bort vid spara och diagrammen på 7.Grafer försvinner. Lathund-sessionen lärde samma
sak ("fyllt via COM, ej openpyxl"). Patch-vägen är därför **Excel COM via PowerShell** (finns i
`tools/recalc.py`-mönstret) eller zip-nivå-XML, styrt från ett git-versionerat skript. Det är
huvudkostnaden för varje punkt ovan och skälet till att insatserna är "liten" snarare än "trivial".

## Föreslagen ordning

1. **Fas 1 — baslinje + restvärdesmodell.** Pinna D13/D14/D15. COM-patchskript. Ny ruta flik 2, ompekning 52 celler, vikter 0/1/0, känslighetsrader på flik 5, tabell i "Kontroll restvärde" på flik 1, historikrad. Regression: 0/0,4/0,6 + Neutralt ⇒ dagens siffror.
2. **Fas 2 — verifierade fixar.** År 1-exponent, CA-index, `#REF!`-städning. Verifiera de tre oklara innan de rörs.
3. **Fas 3 — kravhyra som output.** Multiplikator + datatabell, oavrundade mellanled. Bara om Joakim vill.

## Beslut tagna 2026-09-30 (Joakim, batch)

1. Vikter `5. IRR!A6:A8` default 0/1/0, kvar som synlig inmatning med not.
2. `M8` kalkylränta restvärde följer justerad yield.
3. Steg ±0,25/±0,50 pp symmetrisk.
4. Kravhyra som output: **ska med**.
5. Bladskydd/valideringar: nej.
6. Kravhyra löses för alla rader utom typ Bef. Bef-rader ligger fast.
7. Restvärdesbedömningen placeras på flik 2 (marknadsförutsättning). Flik 4 och 5 räknar, flik 1 redovisar. Ingen ny inmatning på flik 4/5 utöver dagens `C11`-override.
8. **Bef = nollalternativ.** Krav 1 och 2 räknas på tillkommande flöde (icke-Bef-rader, deras DoU och restvärde). Krav 3 räknas på hela fastigheten (alla rader, hela restvärdet) mot `P9` + investering − `P10`.
9. Krav 3 hela fastigheten **redovisas och flaggas** som nedskrivningsrisk men binder inte kravhyran (extern parameter: värdering mot bokföring). Kravhyran binder på krav 1, 2 och tillkommande del av krav 3 (utrangering `P10`).
10. Befintligt bokfört värde `P9` får egen återstående avskrivningstid (ny inmatning flik 2). Investeringen behåller sin.

## Fas 1 — plan

Pipeline: `build/lm371/patch.ps1` (Excel COM) styrd av `build/lm371/build.py`; källa = `LM 371 Investeringskalkyl.xlsx` (orörd, läggs in i repot), utdata `build/lm371/LM371_v3.xlsx`. Regression `build/lm371/regression.py`: (a) med vikter 0/0,4/0,6 + Neutralt + gamla flödesdefinitionen ⇒ `D13 = 10 300`, `D14 = 6,298 %`, `D15 = 10 300`; (b) nollfelsscan på synliga celler; (c) exempel utan Bef-rad ⇒ krav 1/2 oförändrade av Bef-omläggningen.

1. ✅ (9261d7d) Baslinje: kopiera mallen in i repot, pinna D13/D14/D15, COM-skelett + regression grön på oförändrad fil.
2. ✅ (9261d7d) Flik 2: rad under `M6` "Justering direktavkastning" + "Justerad direktavkastning"; `M8` → justerad. Ny sektion 4 "Restvärdesbedömning" rad 44+: 4 dropdowns (list-DV via COM), effekt, motivering, summa.
3. ✅ (9261d7d) Ompekning yield: `7.Grafer!C44:AZ44`, `4. NPV!C50/E50`, `1. Framskr.!G68/K68` (+ ny rad för justering). Vikter 0/1/0 med not. `G14`-varningen utökas med yield-justering.
4. Bef-omläggning flik 4: Bef-del av intäkter/DoU/restvärde särskiljs (SUMIFS på typ, `D$7`-schablonrader). Rad "Driftnetto tillkommande" och "Driftnetto hela fastigheten". Krav 1 (`E59`) och flik 5 rad 23 på tillkommande; krav 3 (`D15`) på hela. Ny inmatning "Återstående avskrivningstid befintligt" vid `P9`; flik 3 skriver av `P9` på egen tid.
5. Känslighet flik 5: tre rader = rad 37 + Δexit i exit-årskolumnen (yield −1/±0/+1 pp), IRR per rad. Flik 1 "Kontroll restvärde" (L108–L112) visar extern yield, justering, justerad yield, tre scenarier; krav 3-flagga.
6. Kravhyra: multiplikator m på icke-Bef-raders `L23:L27`; envariabels datatabell m=0/m=1 → kassaflöden; analytisk kravhyra per krav (1, 2, 3-tillkommande), MAX. Oavrundade mellanled (`C48/C50`, `5. IRR!C6:C10` avrundas i visning). Resultat på flik 2 nyckeltalsrutan + flik 1.
7. Uppdateringshistorik flik 1 kol R, lathund-notering, regression grön, COM-omräkning av exempel + skärmkoll av flik 1/2/5.

Steg 1–3 = restvärdesmodellen (Joakims krav). Steg 4 = Bef/bokfört värde. Steg 5–6 = känslighet + kravhyra. Varje steg committas separat med grön regression.
