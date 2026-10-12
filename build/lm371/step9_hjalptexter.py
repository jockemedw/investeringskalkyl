"""Steg 9 (v3.3): städning av hjälptexter 2026-10-12.

v3.2 la vägledning på tre sätt samtidigt: grå kursiva rader i bladet, ett sjuttiotal inmatningsmeddelanden
(gula rutor vid klick) och mallens egna anteckningar. Resultatet blev plottrigt. Regel i v3.3:

- **I bladet:** bara rubriker, enheter och korta enradiga underrubriker/källhänvisningar ("Hämtas från …",
  "← Krav 1: ska vara ≥ 0").
- **Anteckningar (Excel-noter, röd hörnmarkering, visas när musen vilar på cellen):** förklaringar. En not per
  begrepp, på rubriken eller radnamnet — inte på varje inmatningscell.
- **Inmatningsmeddelanden:** bara kvar där valet behöver förklaras i stunden: rullistorna Typ av projekt
  (flik 1 G11), Typ (flik 2 E23:E27), Långsiktig vakansrisk (G76) och Objekt för re-investering (Q34:Q39).
  Valideringar med felmeddelande (spärren P23:P27) behåller felmeddelandet.
- Mallens trådade kommentarer (flik 2) och anteckningar (flik 3 B7/B8) rörs inte. Mallens felplacerade not på
  flik 2 S22 (handlade om marknadens avkastningskrav, som har en egen kommentar på M6) får rätt innehåll.

Inga formler, värden eller valideringsregler ändras.
"""
from __future__ import annotations

from step8_design import ACK, ALL, FIN, FMT, FS, GRAF, IRR, KD, KONTO, NPV, S, op


def N(sheet, ref, text, title=None, w=300):
    d = {"op": "note", "sheet": sheet, "ref": ref, "text": text, "w": w}
    if title:
        d["title"] = title
    return d


def OFF(sheet, *refs):
    return [op("dv_prompt_off", sheet, r) for r in refs]


def CLR(sheet, *refs):
    return [op("clear_contents", sheet, r) for r in refs]


def ops_hjalptexter_v33() -> list[dict]:
    o: list[dict] = []
    for sh in ALL:
        o.append(op("footer_replace", sh, "A1", find="Investeringskalkyl LM 371 v3.2&",
                    repl="Investeringskalkyl LM 371 v3.3&"))

    # ── 1. Framskrivningsunderlag ──
    # Självklara fält (namn, utförd av, projektnummer, fritextfält) behöver ingen hjälp; resten blir noter
    # på radnamnen i kolumn B. Rullistan G11 (x14) behåller sitt meddelande.
    o += OFF(FS, "B3", "G4", "G5", "G6", "G7", "G8", "G9", "G10", "G12:G13", "B28", "B92", "B117", "G55", "G107")
    o += [N(FS, "B8", "Markareal i kvm. Används på 2. Kalkyldata för minsta markvärde."),
          N(FS, "B9", "Bruksarea i kvm. Ska stämma med summa BRA i hyrestabellen på 2. Kalkyldata."),
          N(FS, "B12", "Andel av ytan som hyrs ut momspliktigt, 0–100 %, före respektive efter åtgärden."),
          N(FS, "B55", "Enligt senaste värdering. Informativt – ingår inte i beräkningen."),
          N(FS, "B66", "Direktavkastningen vid exit ±1,00 pp. Raden Bedömt är kalkylens restvärde "
                       "(4. NPV rad 84–90). Underlag: 5. IRR rad 53–57.", title="Känslighet restvärde"),
          N(FS, "B107", "Lägsta hyra år 1 (viktad över tillkommande yta) som klarar både IRR-kravet (krav 2) och "
                        "NPV ≥ 0 (krav 1). Kolumn D visar vilket krav som binder. Källa: 2. Kalkyldata rad 17.",
            title="Kravhyra")]
    o += CLR(FS, "B71")

    # ── 2. Kalkyldata ──
    # Arbetsgången: en rad i bladet, detaljerna i en not på INMATNING.
    o += [S(KD, "C1", "Fyll i blå fält uppifrån och ned, sektion 1–5. Röd kolumnrubrik = obligatoriskt fält "
                      "som saknar värde."),
          N(KD, "B1", "1) Projektuppgifter på 1. Framskrivningsunderlag.\n2) Blå fält här uppifrån och ned: "
                      "kalkylstart och förutsättningar (1) → ytor, hyror och avtal (2) → drift (3) → investering "
                      "per objekt (4) → restvärdesbedömning (5).\n3) Läs resultatet i Nyckeltal kalkyl (rad 12–17) "
                      "och på flik 1.\nFlik 3–7 räknas automatiskt. 3. Finansiering ändras bara efter OK från "
                      "finansansvarig.", title="Arbetsgång", w=340)]
    o += CLR(KD, "C2") + [op("rowheight", KD, "2:2", h=6)]

    # Projektinfo (rad 6–11) speglar flik 1 — rubriken C4 säger det redan.
    o += OFF(KD, "E6", "G6", "H6", "I6", "J6", "E8", "F10", "F11")

    # Förutsättningar: noter på radnamnen i kolumn K/N i stället för meddelanden på värdecellerna.
    o += OFF(KD, "M10", "Q13", "Q15", "Q17", "D15")
    o += CLR(KD, "R13")
    o += [N(KD, "K10", "Summan av bedömningarna i sektion 5 Restvärdesbedömning (H79). Ändras där."),
          N(KD, "K11", "Extern direktavkastning (M6) + justering (M10). Kapitaliserar restvärdet i alla tre "
                       "kraven (4. NPV rad 88–89)."),
          N(KD, "N13", "Valfritt: återstående avskrivningstid i år för befintligt bokfört värde (P9). Anges i Q13. "
                       "Tomt = Avskrivningstid (P16).", title="Avskrivningstid befintligt"),
          N(KD, "N15", "Valfritt: årlig avskrivning för investeringen, anges i Q15. Tomt = viktat värde från "
                       "Avskrivning-tabellen (T10)."),
          N(KD, "N17", "Kalkylperiod i år, anges i Q17 (normalt 20). Tomt = avskrivningstiden (P16).")]

    # Nyckeltal: en not per krav.
    o += [N(KD, "C13", "NPV av tillkommande flöde (hela fastigheten minus Bef-objekt) inklusive restvärde, minus "
                       "investeringen, tkr. Ska vara > 0. Beräknas på 4. NPV rad 82.", title="Krav 1"),
          N(KD, "C14", "IRR på eget kapital före skatt för tillkommande flöde (5. IRR D40). Ska vara minst kravet "
                       "i F14.", title="Krav 2"),
          N(KD, "C15", "Nuvärde för hela fastigheten minus bokfört värde inkl. investering, tkr. Ska vara > 0. "
                       "Bef-objekt och bokfört värde (P9) ingår bara i detta krav.", title="Krav 3"),
          N(KD, "C17", "Lägsta hyra år 1 (viktad över tillkommande yta) som klarar både krav 1 och krav 2. "
                       "F17 visar vilket krav som binder.", title="Kravhyra")]

    # Sektion 2 Ytor: underrubrik kortad; noter på kolumnrubrikerna.
    o += [S(KD, "C20", "Ett objekt per rad (högst 5). Belopp i kalkylstartens penningvärde.")]
    o += OFF(KD, "F23:G27", "I23:I27", "J23:J27", "K23:K27", "M23:N27", "O23:O27", "Q23:R27", "S23:S27",
             "U23:U27", "P23", "P24", "P25", "P26", "P27")
    o += [S(KD, "P22", "efter avtal"),
          S(KD, "U22", "Justering,\nkr/kvm"), op("wrap", KD, "U22", on=True),
          N(KD, "F21", "Area BRA i kvm: Befintligt (F) respektive Nybyggnad/Tillbyggnad (G). Summa BRA ska stämma "
                       "med 1. Framskrivningsunderlag."),
          N(KD, "I22", "Avtalad årshyra i kr/kvm BRA, kalkylstartens penningvärde. Bef-rad: faktisk hyra i dag."),
          N(KD, "J22", "Andel av inflationen (M9) som hyran räknas upp med per år. 100 % = full indexering, "
                       "70 % = 70 % av inflationen."),
          N(KD, "K22", "Tilläggshyra i kr/kvm, t.ex. investeringstillägg."),
          N(KD, "M21", "Avtalets första och sista år, t.ex. 2026 och 2045. Därefter gäller marknadshyran (O) och "
                       "eventuell vakansgrad (P)."),
          N(KD, "O22", "Hyra efter avtalsslut, kr/kvm i kalkylstartens penningvärde (räknas upp automatiskt). "
                       "Förifylld med avtalshyran i kolumn I – skriv över vid annan bedömning."),
          N(KD, "P21", "Vakans för år efter avtalsslut inom kalkylperioden (påverkar kassaflödet). Kan bara anges "
                       "när avtalet slutar före kalkylperiodens sista år. Restvärdet räknas alltid fullt uthyrt – "
                       "restvärdets vakansrisk anges i sektion 5.", title="Vakansgrad efter avtal"),
          N(KD, "S22", "Beräknas: objektets summa i sektion 4 delat med area (H). Ändra i sektion 4 eller överstyr i "
                       "Justering (U).", title="Budget kr/kvm")]

    # Sektion 3 Kostnader.
    o += [S(KD, "C32", "Belopp i kalkylstartens penningvärde.")]
    o += OFF(KD, "G34:H38", "L33:M33", "L34:M38", "P34:P39", "R34:R39", "S34:S39", "U34:U39")
    o += [N(KD, "C33", "Kr/kvm och år i kalkylstartens penningvärde. Befintligt (G) gäller objekt av typ Bef, Omb, "
                       "Ins eller Mar; Nybyggnad/Tillbyggnad (H) gäller typ Nyb (kolumn E i Ytor)."),
          N(KD, "K33", "Stegvis höjning av posten i %, utöver inflation. L33/M33 = antal år efter kalkylstart då "
                       "höjningen börjar gälla (ändras normalt inte). 15-årshöjningen läggs ovanpå 10-årshöjningen."),
          N(KD, "P33", "En åtgärd per rad: beskrivning (P, påverkar inte beräkningen), objekt (Q), år (R) och "
                       "kostnad i kr/kvm i kalkylstartens penningvärde (S). Kostnaden räknas upp med inflation till "
                       "åtgärdsåret (T) och belastar kassaflödet det året (4. NPV rad 25). Justering (U) överstyr "
                       "kostnaden.", w=340)]

    # Sektion 4 Investeringsspecifikation.
    o += [S(KD, "C53", "Kr per objekt i dagens penningvärde.")]
    o += OFF(KD, "G55:K63", "G66:K68", "G64:K64", "G65:K65")
    o += [N(KD, "C54", "Belopp i kr per objekt (kolumnerna = objekten i Ytor, rad 23–27). Summeras i rad 69 och styr "
                       "Budget kr/kvm (S23:S27). Manuell överstyrning görs i Justering (U23:U27)."),
          N(KD, "C64", "Ange bidrag och intäkter som negativt belopp (t.ex. −2 000 000). Raden summeras med "
                       "kostnaderna i Summa investering."),
          N(KD, "C65", "Engångskostnad som ingår i investeringen. Löpande hyra för ersättningslokaler anges i stället "
                       "i P11 (kr/år). Ange inte samma kostnad på båda ställena.")]

    # Sektion 5 Restvärdesbedömning: rullistan G76 behåller meddelandet (skalan behövs när man väljer).
    o += [S(KD, "C73", "Fyra bedömningar justerar direktavkastningen för restvärdet (M11).")]
    o += OFF(KD, "G75", "G77", "G78")
    o += [N(KD, "C72", "Extern direktavkastning (M6) är utgångsläget. Bedömningarna justerar kapitaliseringen (M11) "
                       "i alla tre kraven, ±0,25/±0,50 pp per faktor (vakansrisk −0,25…+2,00 pp). Kalkylränta "
                       "restvärde = M6 + inflation.", title="Restvärdesbedömning", w=320),
          N(KD, "C75", "Kollektivtrafik, serviceunderlag, demografi, planerad infrastruktur.",
            title="Bedömningsgrund"),
          N(KD, "C77", "Planlösning, bjälklagshöjd, bärande stomme, installationskapacitet.", title="Bedömningsgrund"),
          N(KD, "C78", "Stomme, klimatskal, installationer, energiprestanda mot framtida krav.",
            title="Bedömningsgrund")]

    # ── 3. Finansiering ──
    o += [S(FIN, "B4", "Antagandena i C7:C10 ändras endast efter OK från ansvarig för bolagets finansiering."),
          op("rowheight", FIN, "4:4", h=20),
          N(FIN, "B3", "Fliken räknar fram lån, räntekostnad och amortering per år samt bokfört värde och "
                       "avskrivning. Eget kapital, skuld, ränta, amortering, avskrivning och bokfört värde används i "
                       "5. IRR; investeringen per år i 4. NPV.")]
    o += OFF(FIN, "C7", "C8", "C9", "C10")
    o += [N(FIN, "B9", "Rak amortering i % av upptaget lån per år. D9 visar motsvarande amorteringstid i år. "
                       "Påverkar rad 20–22 och finansieringskostnaden."),
          N(FIN, "B10", "Andel av investeringen som lånefinansieras vid start. Resten blir eget kapital (rad 16), "
                        "som används i 5. IRR."),
          N(FIN, "B23", "Räntan för byggnadskreditiv (C7) gäller kalkylår 1 och trappas linjärt upp till långsiktig "
                        "ränta (C8), som nås kalkylår 11 och gäller resten av kalkylperioden.")]
    # Pilarna "→ flik" i kolumn C blir en not på årsraden.
    o += CLR(FIN, "C15", "C16", "C20", "C22", "C24", "C27", "C47", "C57", "C39", "C49")
    o += [N(FIN, "B14", "Rader som används på andra flikar:\nrad 15 Investering → 4. NPV\nrad 16, 20, 22, 24, 47 "
                        "och 57 → 5. IRR\nrad 27 Belåningsgrad → 7.Grafer", title="Vart siffrorna går"),
          N(FIN, "B39", "Rak avskrivning; avskrivningstider enligt 2. Kalkyldata P13 (befintligt) och P16. Per objekt "
                        "i grupperade rader 40–46."),
          N(FIN, "B49", "Bokfört värde minus ackumulerad avskrivning. Används i 5. IRR (rad 8) och 7.Grafer (rad 48).")]

    # ── 4. NPV ──
    o += [S(NPV, "B2", "Beslutet läses på rad 82: NPV tillkommande (krav 1) ska vara ≥ 0."),
          op("rowheight", NPV, "2:2", h=20),
          N(NPV, "B1", "Nuvärdet av investeringens framtida driftnetton under kalkylperioden plus restvärdet vid "
                       "periodens slut, diskonterat med kalkylräntan, minus investeringen. Rad 47–62 visar hela "
                       "fastigheten (krav 3), rad 64–82 tillkommande flöde (krav 1) och rad 84–91 hur restvärdet "
                       "räknas.", title="Nettonuvärde (NPV)", w=320),
          N(NPV, "B5", "Året efter kalkylperioden (år N+1, vid 20 år kolumn X) visas bara som underlag för "
                       "restvärdet.")]
    o += CLR(NPV, "X10", "F47", "F59", "F62", "B75", "B91")
    o += OFF(NPV, "C47", "E47", "E52", "C55:E55", "C56:E56", "E82")
    o += [N(NPV, "B47", "Hela fastigheten inkl. Bef-objekt. PV TOTALT (rad 52) prövas mot bokfört värde i krav 3 "
                        "(2. Kalkyldata D15). C47 = kort horisont i år (standard 10, kan ändras); kolumn E = hela "
                        "kalkylperioden (2. Kalkyldata P17).", title="Krav 3", w=320),
          N(NPV, "B55", "Summa tillkommande investering från 2. Kalkyldata W28 (sektion 4). Rad 56 = investeringen "
                        "diskonterad till kalkylstart (X28)."),
          N(NPV, "B59", "Informativ, ej krav. Lika med NPV tillkommande (rad 82) när Bef-rader saknas."),
          N(NPV, "B62", "År tills ackumulerat driftnetto (ej diskonterat) täcker investeringen. Informativ, förs till "
                        "2. Kalkyldata D16."),
          N(NPV, "B64", "Bef-objekt (befintliga avtal som löper vidare) är nollalternativet och dras ifrån. Utan "
                        "Bef-rader är tillkommande flöde = hela fastigheten, och blocket visar samma siffror som "
                        "rad 48–59."),
          S(NPV, "F82", "← Krav 1: ska vara ≥ 0"),
          N(NPV, "B82", "Nuvärde av tillkommande driftnetto + restvärde (rad 88) minus PV investering över hela "
                        "kalkylperioden (kolumn E). Ska vara ≥ 0. Förs till 2. Kalkyldata D13; kravhyran (rad 17 där) "
                        "är den hyra som ger 0 här.", title="Krav 1", w=320),
          N(NPV, "B84", "Kapitaliseringsbart driftnetto = fullt uthyrt (vakans efter avtal återlagd), exkl. "
                        "extraordinärt UH. Restvärdet vid slutet av år t = driftnetto år t+1 / justerad "
                        "direktavkastning (2. Kalkyldata M11).", title="Restvärde", w=320)]

    # ── 5. IRR ──
    o += [N(IRR, "B1", "Beräknar den årliga avkastningen på eget kapital för det tillkommande kassaflödet. Jämförs "
                       "med IRR-kravet på 2. Kalkyldata rad 14. Restvärdet vid kalkylperiodens slut förs in som "
                       "försäljningsvärde: standard är rent marknadsvärde (driftnetto år N+1 / justerad "
                       "direktavkastning).", title="Internräntemetoden (IRR)", w=320)]
    o += CLR(IRR, "B2", "D10", "B13", "I55")
    o += OFF(IRR, "C11", "A6:A8")
    o += [S(IRR, "D11", "← Lämnas normalt tomt"),
          N(IRR, "A5", "Vikter i kolumn A, summa 100 %. Standard 0 % / 100 % / 0 % = rent marknadsvärde vid "
                       "justerad direktavkastning. 0 % / 40 % / 60 % = tidigare LM 371-viktning mot bokfört värde. "
                       "Ändra bara med motivering. Rad 9 (lån) viktas inte.", title="Vikter i exit-värdet"),
          N(IRR, "B10", "Vikt × värde, rad 6–8. Är C11 ifyllt gäller det beloppet."),
          N(IRR, "B11", "Lämna tomt = beräknat värde (rad 10). Ett ifyllt belopp i kr ERSÄTTER exit-värdet helt, det "
                        "är inget tillägg. Bara i undantagsfall; motivera i kommentaren. 2. Kalkyldata visar en notis "
                        "när fältet är ifyllt.", title="Eget exit-värde (undantag)"),
          S(IRR, "E40", "← Krav 2"),
          N(IRR, "B40", "IRR på eget kapital före skatt för tillkommande kassaflöde (rad 37). Jämförs med "
                        "IRR-kravet på 2. Kalkyldata F14.", title="Krav 2"),
          N(IRR, "B53", "Restvärdets direktavkastning ±1,00 pp kring bedömt läge (2. Kalkyldata M11), tillkommande "
                        "flöde. IRR per scenario räknas i dolda hjälprader 59–61.")]

    # ── 6. Ack.kapitalutlägg ──
    o += [S(ACK, "A2", "Kvar att tjäna in av investeringen inkl. räntekrav, år för år."),
          op("rowheight", ACK, "2:2", h=15),
          N(ACK, "A1", "Visar hur mycket av investeringen, inklusive räntekrav på bundet kapital, som återstår att "
                       "tjäna in år för år. När Ack. kapitalutlägg (rad 13) når 0 är kapitalet återvunnet. Används i "
                       "5. IRR rad 6; med standardvikterna (0 % / 100 % / 0 %) påverkar fliken inte exit-värdet.",
            w=320),
          N(ACK, "B3", "Tecknen (+)/(−) i radnamnen anger effekten på Ack. kapitalutlägg.")]
    o += CLR(ACK, "B15:B17") + [op("rowheight", ACK, "15:17", h=15), op("print_area", ACK, "A1", area="$A$1:$W$13")]
    o += OFF(ACK, "B8", "B9", "B11", "B13")
    o += [N(ACK, "B8", "Investering (4. NPV rad 9) minus projektbelastat bokfört värde (flik 2). Negativt belopp = "
                       "utbetalning, som ökar kapitalutlägget."),
          N(ACK, "B9", "Ersättningslokaler och tillfälliga intäkter (flik 2 P11–P12, via 5. IRR rad 24–25). Ingår i "
                       "Ack. kapitalutlägg men inte i Kassaflöde (rad 10)."),
          N(ACK, "B10", "Rad 6 + 7 + 8. Räntebidrag (rad 7, dold) ingår i Kassaflöde men inte i Ack. kapitalutlägg."),
          N(ACK, "B11", "År 1: investeringen × kalkylränta driftnetto (flik 2 M7). Därefter: föregående års Ack. "
                        "kapitalutlägg × samma ränta."),
          N(ACK, "B13", "Kvar att återvinna inklusive räntekrav. Det år raden blir 0 eller negativ är investeringen "
                        "återbetald.")]

    # ── 7.Grafer ──
    o += [S(GRAF, "A37", "Inget fylls i här – raderna hämtas från flik 3–5."),
          N(GRAF, "A36", "Raderna kan markeras för egna diagram (Infoga › Diagram). Rad 44 och 48 används av 4. NPV "
                         "rad 34–35 – ändra inte. Rad 40, 44 och 50 avser tillkommande flöde (Bef = nollalternativ), "
                         "övriga rader hela fastigheten. Utan Bef-objekt sammanfaller de.", w=320)]
    o += CLR(GRAF, "B49")
    kalla = {
        "B40": "4. NPV rad 71 − rad 72. Driftnetto för tillkommande flöde (Bef = nollalternativ), exkl. "
               "extraordinärt UH.",
        "B41": "4. NPV rad 19 Nettohyra (efter vakans), hela fastigheten.",
        "B42": "4. NPV rad 26 (drift och underhåll utan re-investering rad 25) + rad 27–28 (fastighetsskatt, "
               "tomträttsavgäld), hela fastigheten. Visas som positivt belopp.",
        "B43": "4. NPV rad 25 Re-investering/hyressatt underhåll, hela fastigheten. Visas som positivt belopp.",
        "B44": "4. NPV rad 88 = kapitaliseringsbart driftnetto år t+1 / justerad direktavkastning. Är driftnettot 0 "
               "visas bokfört värde (rad 48). Används av 4. NPV rad 34.",
        "B45": "5. IRR rad 16 (investering per år, positivt belopp).",
        "B46": "5. IRR rad 17.",
        "B47": "Löpande summa av 4. NPV rad 29 Driftnetto (hela fastigheten, inkl. extraordinärt UH).",
        "B48": "3. Finansiering rad 57. Används av 4. NPV rad 35 (direktavkastning bokfört värde).",
        "B50": "Rad 40, 41 och 42 delade med total area (4. NPV rad 8).",
        "B57": "3. Finansiering rad 27.",
    }
    o += OFF(GRAF, "B40", "B41", "B42", "B43", "B44", "B45", "B46", "B47", "B48", "B50:B52", "B57")
    o += [N(GRAF, ref, text, title="Källa") for ref, text in kalla.items()]

    # ── Kontoplan RR ──
    o += [S(KONTO, "E2", "Konton som ingår i driftkostnadsposterna på 2. Kalkyldata (rad 34–38)."),
          N(KONTO, "F5", "= rullistan Typ på 2. Kalkyldata (E23:E27). Ändra inte texterna – formler läser listorna."),
          N(KONTO, "H5", "= rullistan Typ av projekt på 1. Framskrivningsunderlag (G11). Ändra inte texterna – "
                         "formler läser listorna.")]
    o += CLR(KONTO, "E12:E14")
    return o


def ops_historik_v33() -> list[dict]:
    return [FMT(FS, "R26", "R27"), FMT(FS, "S26", "S27"),
            S(FS, "R27", "19. 2026-10-12 v3.3: Hjälptexter städade — förklaringar flyttade från grå textrader och "
                         "inmatningsmeddelanden till anteckningar (röd hörnmarkering) på rubriker och radnamn; "
                         "inmatningsmeddelanden kvar bara på rullistor; avklippta kolumnrubriker på flik 2 rättade; "
                         "sidfot v3.3. Inga beräkningar ändrade."),
            S(FS, "S27", "Claude/JW")]
