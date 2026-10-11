"""Steg 8 (v3.2): användarvänlighet och design efter agentgranskning 2026-10-11.

Underlag: åtta flikgranskare + helhetsgranskare + syntes (V32-01 … V32-44) och syntesens beslutspunkter med
presentationskaraktär (D1 utskrift, D2 nollor, D5 legacy-rader, D6 inmatningston, D7 kravhyrerad, D8 NPV rad 59,
D9a EBITDA, D10 källhänvisningar flik 3, D11 projektnamn flik 2, D12 BILD/KARTA). Inga beräkningsformler ändras;
bara etiketter, etikettformler (ren text), talformat, toner, vägledning, döljning och utskrift. Formelfel som
granskningen hittade ligger i en separat formelrunda (beslut för Joakim).
"""
from __future__ import annotations

KD, FS, FIN, NPV, IRR = "2. Kalkyldata", "1. Framskrivningsunderlag", "3. Finansiering", "4. NPV", "5. IRR"
ACK, GRAF, KONTO = "6. Ack.kapitalutlägg", "7.Grafer", "Kontoplan RR"
ALL = [FS, KD, FIN, NPV, IRR, ACK, GRAF, KONTO]
LEFT, CENTER, RIGHT = -4131, -4108, -4152
K = "'2. Kalkyldata'!"


def S(sheet, ref, text):
    return {"op": "set", "sheet": sheet, "ref": ref, "value": text}


def F(sheet, ref, formula):
    return {"op": "set", "sheet": sheet, "ref": ref, "formula": formula}


def P(sheet, ref, title, text):
    assert len(title) <= 32 and len(text) <= 255, (ref, len(title), len(text))
    return {"op": "dv_prompt", "sheet": sheet, "ref": ref, "prompt_title": title, "prompt": text}


def NF(sheet, ref, fmt):
    return {"op": "numfmt", "sheet": sheet, "ref": ref, "fmt": fmt}


def FMT(sheet, src, dst, src_sheet=None):
    d = {"op": "fmt_from", "sheet": sheet, "src": src, "dst": dst}
    if src_sheet:
        d["src_sheet"] = src_sheet
    return d


def FILL(sheet, ref, src, src_sheet=None):
    d = {"op": "fill_from", "sheet": sheet, "ref": ref, "src": src}
    if src_sheet:
        d["src_sheet"] = src_sheet
    return d


def note(sheet, ref, text):
    """Grå kursiv förklaring i en tom cell."""
    return [S(sheet, ref, text), {"op": "italic", "sheet": sheet, "ref": ref},
            {"op": "fontcolor", "sheet": sheet, "ref": ref, "color": "grey"}]


def op(name, sheet, ref, **kw):
    return {"op": name, "sheet": sheet, "ref": ref, **kw}


NOLL_STRECK = '#,##0;-#,##0;"–"'


def ops_design_v32() -> list[dict]:
    o: list[dict] = []

    # ── V32-01 sidfot v3.2 ──
    for sh in ALL:
        o.append(op("footer_replace", sh, "A1", find="Investeringskalkyl LM 371 v3&", repl="Investeringskalkyl LM 371 v3.2&"))

    # ── V32-02 krav 1/2/3 med samma namn överallt ──
    o += [op("colwidth", KD, "C:C", w=29),
          S(KD, "C13", "NPV tillkommande (krav 1)"), S(KD, "C14", "IRR före skatt (krav 2)"),
          S(KD, "C15", "PV mot bokfört värde (krav 3)"), S(KD, "C17", "Kravhyra krav 1–2, kr/kvm"),
          F(FS, "B105", '="IRR före skatt ("&ROUND(G102*100,0)&" % belåning, krav 2)"'),
          S(FS, "B108", "NPV tillkommande, över-/undervärde (krav 1)"),
          S(FS, "B111", "PV jmf bokfört värde, hela fastigheten (krav 3)"),
          S(IRR, "B40", "IRR eget kapital, före skatt"), S(IRR, "B41", "IRR eget kapital, efter skatt")]
    o += note(IRR, "E40", "← Krav 2: jämförs med IRR-kravet på 2. Kalkyldata (F14)")

    # ── V32-03 två decimaler på räntor, direktavkastning och IRR ──
    o += [NF(FS, r, "0.00%") for r in ("G105", "G75", "G77", "G78", "G88", "G101", "G102")]
    o += [NF(KD, "D14", "0.00%"), NF(IRR, "D40:D41", "0.00%"), NF(FIN, "E23:BA23", "0.00%")]

    # ── V32-04 stavfel och språk ──
    o += [S(FS, "B54", "Marknadssituation"), S(FS, "B23", "Därutöver re-investeringar/hyressatt underhåll"),
          S(FS, "B25", "Total projektvolym"), S(FS, "F16", "Investering"),
          S(FS, "B58", "Hur har restvärdet bedömts (direktavkastning, kalkylränta)? Har värderaren tagit hänsyn till "
                       "kommande investeringar för att upprätthålla hyresnivån? Vad händer med marknadsvärdet om vi inte "
                       "investerar? Kommentera sannolikt scenario, t.ex. sjukt hus, tekniskt utdömt eller vakanser. Är "
                       "byggrätterna optimerade? Förändras momsregistreringsgraden?"),
          S(KD, "C42", "Beskriv antaganden: motivera otydliga poster och särskilda antaganden."),
          S(NPV, "B25", "Re-investering/hyressatt underhåll"), S(NPV, "B37", "Driftnetto exkl extraordinärt UH"),
          S(NPV, "B40", "Beräkning nettonuvärde 10 år"), S(NPV, "B41", "Beräkning nettonuvärde 20 år"),
          S(FIN, "B30", "Befintligt bokfört värde"), S(FIN, "B40", "Befintligt bokfört värde"),
          S(FIN, "B50", "Befintligt bokfört värde"),
          S(IRR, "B35", "Amortering återstående skuld (vid exit)"),
          S(ACK, "B5", "(+) Ingående balans (bokfört värde)"),
          S(KONTO, "A26", " " * 6 + "4540 Planerat underhåll installationer")]

    # ── V32-05 sidbrytning flik 1 före Kalkylantaganden ──
    o.append(op("row_break", FS, "A72", after=71))

    # ── V32-06 projektinfo flik 1: rullista på rätt cell, vägledning ──
    o += [op("x14_dv", FS, "G11", enable_prompts=True, match="E14:F14 E11:F11", sqref="G11",
             prompt_title="Typ av projekt", prompt="Välj i listan. Visas även i 2. Kalkyldata (Typ av åtgärd).",
             error="Välj projekttyp i listan."),
          P(FS, "B3", "Projektnamn", "Skriv projektets namn (ersätt texten Projektnamn)."),
          P(FS, "G4", "Kalkyl utförd av", "Namn på den som gjort kalkylen."),
          P(FS, "G5", "Fastighet", "Fastighetsbeteckning, t.ex. Linköping Lejonet 1. Visas även i 2. Kalkyldata."),
          P(FS, "G6", "Objekt", "Objekt/byggnad som kalkylen avser. Visas även i 2. Kalkyldata."),
          P(FS, "G7", "Projektnummer", "Projektnummer i ekonomisystemet."),
          P(FS, "G8", "Markareal", "Fastighetens markareal i kvm. Används i 2. Kalkyldata för minsta markvärde."),
          P(FS, "G9", "BRA", "Bruksarea i kvm. Ska stämma med summa BRA i hyrestabellen på 2. Kalkyldata (visas i E22 nedan)."),
          P(FS, "G10", "BTA", "Bruttoarea i kvm enligt ritning. Visas i 2. Kalkyldata."),
          P(FS, "G12:G13", "Momsregistreringsgrad", "Andel av ytan som hyrs ut momspliktigt, 0–100 %, före respektive efter åtgärden."),
          P(FS, "B28", "Projektbeskrivning", "Beskriv projektet: omfattning, hyresgäst, tidplan."),
          P(FS, "B92", "Kommentar", "Kommentera antaganden och avvikelser."),
          P(FS, "B117", "Kommentar", "Kommentera resultatet och ge rekommendation."),
          NF(FS, "G8:G10", '#,##0" kvm"'),
          S(FS, "B8", "Fastighetsarea (markareal):"), S(FS, "B12", "Momsregistreringsgrad före:"),
          S(FS, "B13", "Momsregistreringsgrad efter:"),
          op("bold", FS, "B4"),
          FILL(FS, "G55", "G4"), op("bold", FS, "G55", on=False),
          P(FS, "G55", "Marknadsvärde", "Marknadsvärde i kr enligt senaste värdering. Informativt, ingår inte i beräkningen.")]

    # ── V32-07 restvärdestabell och resultatblock i klartext ──
    o += [F(FS, "B65", '="Restvärdesbedömning — direktavkastning vid exit "&ROUND(' + K + '$M$11*100,2)&" % (marknad "'
                       '&ROUND(' + K + '$M$6*100,2)&" % "&IF(' + K + '$M$10>=0,"+","")&ROUND(' + K + '$M$10*100,2)&" pp)"'),
          S(FS, "C66", "Direktavk."), S(FS, "E66", "Restvärde (exit)"), S(FS, "F66", "Kravhyra, kr/kvm"),
          op("wrap", FS, "B66:G66"), op("rowheight", FS, "66", h=30)]
    o += note(FS, "B71", "Känslighet: direktavkastningen vid exit ±1,00 pp. Raden Bedömt är kalkylens restvärde (4. NPV rad 84–90).")
    o += [op("bold", FS, "B71", on=False),
          S(FS, "G104", "Utfall"), op("bold", FS, "G104"), op("align", FS, "G104", h=RIGHT),
          S(FS, "D114", "–"),
          P(FS, "G107", "Kravhyra", "Lägsta hyra år 1 (viktad över tillkommande yta) som klarar både IRR-kravet och NPV ≥ 0. "
                                    "Kolumn D visar vilket krav som binder. Källa: 2. Kalkyldata rad 17.")]
    # D7: kravhyreraden lyfts från 9 pt kursiv till 11 pt, vit fyllning behålls
    o += [op("font_from", FS, "B107:G107", src="B88"), op("bold", FS, "B107"), op("bold", FS, "G107")]

    # ── V32-08 etiketter i Kalkylantaganden/Finansiering/kommentarer ──
    o += [S(FS, "B75", "Justerad direktavkastning (restvärde)"),
          S(FS, "B80", "Hyresvärde år 1, befintligt (inkl. ombyggnad)"),
          S(FS, "B81", "kr/kvm"), S(FS, "B83", "kr/kvm"), S(FS, "B85", "kr/kvm"), S(FS, "B87", "kr/kvm"),
          S(FS, "B84", "Summa hyresvärde, år 1"), S(FS, "B86", "Marknadshyra / långsiktigt hyresvärde"),
          S(FS, "E89", "Exkl. extraordinärt underhåll"),
          S(FS, "B99", "Ränta år 1 (byggnadskreditiv)"), S(FS, "B100", "Långsiktig ränta (från år 11)"),
          S(FS, "B57", "Kommentar marknadssituation och restvärde:"), S(FS, "B91", "Kommentar kalkylantaganden:"),
          S(FS, "B116", "Kommentar resultat och rekommendation:"),
          FMT(FS, "B117", "B92")]
    o += note(FS, "E72", "Hämtas från 2. Kalkyldata") + note(FS, "E98", "Hämtas från 3. Finansiering")
    o.append(op("align", FS, "E98", h=CENTER))

    # ── V32-09 bilagor ──
    o += [S(FS, "B135", "2. Kalkyldata"), S(FS, "B136", "3. Finansiering"),
          S(FS, "B137", "4. NPV (nettonuvärde)"), S(FS, "B138", "5. IRR (internränta)"),
          op("dv_list", FS, "E135:E138", source='"Ja,Nej"')]

    # ── D2 flik 1: tomma objektrader visar inget i stället för 0 ──
    o += [NF(FS, "E17:E21", '#,##0" kvm";-#,##0" kvm";'), NF(FS, "F17:F21", '#,##0" kr";-#,##0" kr";'),
          NF(FS, "G17:G22", '#,##0" kr/kvm";-#,##0" kr/kvm";')]

    # ── D12 BILD/KARTA blir diskreta hjälptexter ──
    for ref, txt in (("B35", "Bild – klistra in här"), ("E35", "Karta – klistra in här")):
        o += [S(FS, ref, txt), op("font_from", FS, ref, src="B88"), op("italic", FS, ref),
              op("fontcolor", FS, ref, color="grey")]

    # ── V32-10 nyckeltalsrutan ──
    o += [NF(KD, "D15", '#,##0" tkr"'),
          P(KD, "D15", "Krav 3", "Nuvärde för hela fastigheten minus bokfört värde inkl. investering, tkr. Ska vara > 0. "
                                 "Bef-objekt och bokfört värde (P9) ingår bara i detta krav."),
          FMT(KD, "I12", "G12"), S(KD, "G12", "Notiser"), op("clear_contents", KD, "I12")]

    # ── V32-11 projektinfo flik 2 hämtas från flik 1 (+ D11 projektnamn) ──
    o += [S(KD, "C4", "Projektinfo — hämtas från flik 1 (fyll i där)"),
          F(KD, "C5", "='1. Framskrivningsunderlag'!$B$3"), NF(KD, "C5", "0;-0;;@"),
          S(KD, "G5", "Objekt"), op("align", KD, "G5", h=LEFT), op("align", KD, "E8", h=LEFT)]
    for ref in ("E6", "G6", "H6", "I6", "J6", "E8", "F10", "F11", "S17"):
        o.append(op("fontcolor", KD, ref, color="grey"))
    for ref in ("E6", "G6", "H6", "I6", "J6", "E8", "F10", "F11"):
        o.append(P(KD, ref, "Från flik 1", "Hämtas från 1. Framskrivningsunderlag – fyll i där, skriv inte över här."))

    # ── V32-12 hyrestabellen: inmatningsmeddelanden ──
    o += [op("x14_dv", KD, "E23:E27", match="E23:E27", prompt_title="Typ",
             prompt="Bef = befintligt avtal som löper vidare (nollalternativ, ingår bara i krav 3). Omb = ombyggnad, "
                    "Nyb = nybyggnad/tillbyggnad, Ins = installationer, Mar = mark. Typen styr också driftkolumnen i sektion 3."),
          P(KD, "F23:G27", "Area", "Area BRA i kvm för befintlig del resp. ny-/tillbyggnad."),
          P(KD, "I23:I27", "Hyra", "Avtalad årshyra i kr/kvm BRA, kalkylstartens penningvärde. Bef-rad: faktisk hyra i dag."),
          P(KD, "J23:J27", "Index", "Andel av inflationen (M9) som hyran räknas upp med per år. 100 % = full indexering, "
                                    "70 % = 70 % av inflationen."),
          P(KD, "K23:K27", "Tillägg", "Tilläggshyra i kr/kvm, t.ex. investeringstillägg."),
          P(KD, "M23:N27", "Avtalsperiod", "Avtalets första och sista år, t.ex. 2026 och 2045. Därefter gäller "
                                            "marknadshyran (O) och eventuell vakansgrad (P)."),
          P(KD, "O23:O27", "Marknadshyra", "Hyra efter avtalsslut, kr/kvm i kalkylstartens penningvärde (räknas upp "
                                           "automatiskt). Förifylld med formel från kolumn I – skriv över vid annan bedömning."),
          P(KD, "Q23:R27", "Produktion", "År då produktionen startar och blir klar. Styr när investeringen belastar kalkylen."),
          P(KD, "S23:S27", "Budget kr/kvm", "Beräknas: objektets summa i sektion 4 delat med area (H). Ändra i sektion 4 "
                                            "eller överstyr i Justering (U)."),
          P(KD, "U23:U27", "Justering", "Lämna tomt. Ange kr/kvm bara för att överstyra budgeten från sektion 4.")]

    # ── V32-13 reinvesteringsblocket: rätt toner + vägledning ──
    o += [FILL(KD, "P34:S39", "I23"), FILL(KD, "U34:U39", "I23"), FILL(KD, "T34:T39", "H23"),
          P(KD, "P34:P39", "Åtgärd", "Kort beskrivning av åtgärden (fritext, påverkar inte beräkningen)."),
          P(KD, "Q34:Q39", "Objekt", "Välj objekt från Ytor (kolumn C). Kostnaden räknas på objektets area."),
          P(KD, "R34:R39", "År", "År då åtgärden genomförs, t.ex. 2041. Belastar kassaflödet det året (4. NPV rad 25)."),
          P(KD, "S34:S39", "Kostnad", "Kostnad i kr/kvm i kalkylstartens penningvärde. Räknas upp med inflation till "
                                      "åtgärdsåret (kolumn T)."),
          P(KD, "U34:U39", "Justering", "Valfritt: överstyr kostnaden i kr/kvm.")]

    # ── V32-14 investeringsspecifikation ──
    o += [S(KD, "C64", "Bidrag och intäkter (ange negativt)"),
          P(KD, "G64:K64", "Bidrag", "Ange bidrag och intäkter som negativt belopp (t.ex. −2 000 000). Raden summeras "
                                     "med kostnaderna i Summa investering."),
          P(KD, "G65:K65", "Ersättningslokaler", "Engångskostnad som ingår i investeringen. Löpande hyra för "
                                                 "ersättningslokaler anges i stället i P11 (kr/år). Ange inte samma kostnad på båda ställena."),
          P(KD, "G55:K63", "Projektbudget", "Belopp i kr per objekt (kolumnerna = objekten i Ytor, rad 23–27). Summeras i "
                                            "rad 69 och styr Budget kr/kvm (S23:S27)."),
          P(KD, "G66:K68", "Projektbudget", "Belopp i kr per objekt (kolumnerna = objekten i Ytor, rad 23–27). Summeras i "
                                            "rad 69 och styr Budget kr/kvm (S23:S27).")]

    # ── V32-15 arbetsgång och underrubriker ──
    o += [op("clear_contents", KD, "M2"), FMT(KD, "C53", "C1:C2"), FMT(KD, "C53", "C20"), FMT(KD, "C53", "C32"),
          S(KD, "C1", "Arbetsgång: 1) Projektuppgifter på 1. Framskrivningsunderlag. 2) Blå fält här uppifrån och ned: "
                      "kalkylstart och förutsättningar (1) → ytor, hyror och avtal (2) → drift (3) → investering per objekt (4) "
                      "→ restvärdesbedömning (5). 3) Läs resultatet i Nyckeltal kalkyl (rad 12–17) och på flik 1."),
          S(KD, "C2", "Flik 3–7 räknas automatiskt; 3. Finansiering ändras bara efter OK från finansansvarig. "
                      "Röd kolumnrubrik = obligatoriskt fält som saknar värde."),
          op("merge", KD, "C1:Q1"), op("merge", KD, "C2:Q2"), op("wrap", KD, "C1:C2"),
          op("align", KD, "C1:C2", h=LEFT, v="top"),
          S(KD, "C20", "Ett objekt per rad (högst 5). Hyror i kr/kvm och år, kalkylstartens penningvärde. "
                       "Budget kr/kvm räknas fram från sektion 4."),
          S(KD, "C32", "Kr/kvm och år i kalkylstartens penningvärde. Befintligt gäller objekt av typ Bef/Omb/Ins/Mar, "
                       "Nybyggnad/Tillbyggnad gäller typ Nyb. Kostnadshöjning = stegvis höjning utöver inflation.")]

    # ── V32-16 drift och kostnadshöjning ──
    o += [S(KD, "C33", "Drift och underhåll, kr/kvm och år"),
          P(KD, "G34:H38", "Drift och underhåll", "Kr/kvm och år i kalkylstartens penningvärde. Befintligt = objekt med typ "
                                                  "Bef, Omb, Ins eller Mar. Nybyggnad/Tillbyggnad = typ Nyb (kolumn E i Ytor)."),
          NF(KD, "G34:H39", "#,##0"),
          S(KD, "K33", "Kostnadshöjning utöver inflation"), op("wrap", KD, "K33"),
          NF(KD, "L33:M33", '"efter "0" år"'),
          P(KD, "L33:M33", "Höjningsår", "Antal år efter kalkylstart då höjningen börjar gälla. Ändras normalt inte."),
          P(KD, "L34:M38", "Kostnadshöjning", "Stegvis höjning av posten i %, utöver inflation, från året i rubriken. "
                                              "15-årshöjningen läggs ovanpå 10-årshöjningen."),
          S(KD, "K34", "Fastighetsskötsel"), S(KD, "K35", "Reparationer"), S(KD, "K36", "Planerat underhåll"),
          S(KD, "K37", "Media"), S(KD, "K38", "Övr. förvaltning")]

    # ── V32-17 rubriker i hyrestabellen ──
    o += [op("wrap", KD, "I22:L22"), op("wrap", KD, "U22"),
          S(KD, "C22", "Byggnad/lokal/objekt"), S(KD, "I22", "Hyra, kr/kvm"), S(KD, "J22", "Index, andel av inflation"),
          S(KD, "L22", "Totalhyra, kr/år"), S(KD, "M21", "Avtalsperiod"), S(KD, "M22", "Fr.o.m. år"),
          S(KD, "N22", "T.o.m. år"), S(KD, "U22", "Justering, kr/kvm"), S(KD, "Q10", "Fr.o.m. år"), S(KD, "R10", "T.o.m. år")]

    # ── V32-18 förutsättningar ──
    o += [S(KD, "N4", "Förutsättningar fastighet"), S(KD, "N6", "Tomträttsavgäld, kr/år"), S(KD, "N7", "Taxeringsvärde, kr"),
          S(KD, "N8", "Fastighetsskatt, %"), S(KD, "N9", "Bokfört värde, kr"), S(KD, "K10", "Justering (från sektion 5)"),
          P(KD, "M10", "Justering", "Summan av bedömningarna i sektion 5 Restvärdesbedömning (H79). Ändras där."),
          S(KD, "R13", "← egen tid, år (valfri)"), op("italic", KD, "R13"), op("align", KD, "R13", h=LEFT),
          NF(KD, "Q13", '0" år"'), NF(KD, "Q17", '0" år"'),
          P(KD, "Q13", "Avskrivningstid befintligt", "Valfritt: återstående avskrivningstid i år för befintligt bokfört "
                                                     "värde (P9). Tomt = Avskrivningstid (P16)."),
          P(KD, "Q15", "Årlig avskrivning", "Årlig avskrivning för investeringen. Tomt = viktat värde från "
                                            "Avskrivning-tabellen (T10)."),
          P(KD, "Q17", "Kalkylperiod", "Kalkylperiod i år (normalt 20). Tomt = avskrivningstiden (P16).")]

    # ── V32-19 antagandefältet som textfält ──
    o += [op("merge", KD, "C43:N50"), op("wrap", KD, "C43"), op("align", KD, "C43", h=LEFT, v="top")]

    # ── V32-20 totalrad och avskrivningsetiketter ──
    o += [NF(KD, "L28", '#,##0" kr"'), NF(KD, "O28", '#,##0" kr"'), NF(KD, "T28", '#,##0" kr"'), NF(KD, "F28", "#,##0"),
          FMT(KD, "D29", "C29"), S(KD, "C29", "Per kvm, viktat"), op("align", KD, "C29", h=LEFT),
          S(KD, "T5", "Avskrivning, %/år"), S(KD, "S8", "Övrigt 1"), S(KD, "S9", "Övrigt 2")]

    # ── V32-21 sektionsnumren med i utskriften ──
    o.append({"op": "print_area", "sheet": KD, "area": "$A$3:$U$81"})

    # ── V32-22 tipset på G76 ──
    o.append(P(KD, "G76", "Långsiktig vakansrisk",
               "Vakans efter avtalsslut utöver vad extern direktavkastning prisar. Låg −0,25 / Normal 0 / Förhöjd +0,50 / "
               "Hög +1,00 / Mycket hög +2,00 pp. Kolumn I visar motsvarande vakans. Extremfall: överstyr restvärdet i 5. IRR C11."))

    # ── V32-23/24/25 3. Finansiering ──
    o += [P(FIN, "C7", "Ränta byggnadskreditiv", "Aktuell ränta för byggnadskreditiv inkl. marginal. Gäller kalkylår 1 och "
                                                 "trappas sedan linjärt upp till långsiktig ränta (C8) på 10 år, se rad 23. "
                                                 "Ändras endast efter OK från finansansvarig."),
          P(FIN, "C8", "Långsiktig ränta", "Beslutad kalkylränta för långsiktig ränteprognos. Nås kalkylår 11 och gäller "
                                           "sedan resten av kalkylperioden (rad 23). Ändras endast efter OK från finansansvarig."),
          P(FIN, "C9", "Amortering per år", "Rak amortering i % av upptaget lån per år. D9 visar motsvarande amorteringstid "
                                            "i år. Påverkar rad 20–22 och finansieringskostnaden."),
          P(FIN, "C10", "Belåningsgrad", "Andel av investeringen som lånefinansieras vid start. Resten blir eget kapital "
                                         "(rad 16), som används i 5. IRR. Ändras endast efter OK från finansansvarig."),
          S(FIN, "B4", "Finansieringsantagandena (ränta, amortering, belåningsgrad) matas in i C7:C10 och ändras endast efter "
                       "OK från ansvarig för bolagets finansiering. Fliken räknar fram lån, räntekostnad och amortering per år "
                       "samt bokfört värde och avskrivning. Eget kapital, skuld, ränta, amortering, avskrivning och bokfört "
                       "värde används i 5. IRR; investeringen per år i 4. NPV."),
          op("rowheight", FIN, "4", h=68),
          S(FIN, "B7", "Ränta byggnadskreditiv (år 1)"), S(FIN, "B9", "Amortering per år"),
          S(FIN, "D6", "Innebär"), op("align", FIN, "D6", h=CENTER), S(FIN, "D8", "Nås år 11"),
          op("colwidth", FIN, "B:B", w=31),
          S(FIN, "B14", "Kalkylår · belopp i kr"), S(FIN, "B15", "Investering (från 2. Kalkyldata)"),
          S(FIN, "B22", "Återstående skuld vid årets slut"), S(FIN, "B23", "Räntesats (trappas upp år 1–11)"),
          S(FIN, "B26", "Finansieringskostnad (ränta + amortering)"), S(FIN, "B27", "Belåningsgrad vid årets slut"),
          NF(FIN, "D27:BA27", "0.0%"), S(FIN, "C29", "Högsta värde"),
          S(FIN, "B47", "Årets avskrivning"), S(FIN, "B49", "ÅTERSTÅENDE BOKFÖRT VÄRDE"),
          S(FIN, "B57", "Bokfört värde vid årets slut")]
    o += note(FIN, "C39", "Rak avskrivning; avskrivningstider enligt 2. Kalkyldata P13/P16. Per objekt i grupperade rader 40–46.")
    o += note(FIN, "C49", "Bokfört värde minus ackumulerad avskrivning (används i 5. IRR)")
    # D10 källhänvisningar i den tomma kolumn C
    for r, txt in ((15, "→ 4. NPV"), (16, "→ 5. IRR"), (20, "→ 5. IRR"), (22, "→ 5. IRR"), (24, "→ 5. IRR"),
                   (27, "→ 7.Grafer"), (47, "→ 5. IRR"), (57, "→ 5. IRR")):
        o += note(FIN, f"C{r}", txt)
    o += [op("fontcolor", FIN, "C22", color="D9D9D9"), op("fontcolor", FIN, "C27", color="D9D9D9"),   # på mörkblått band
          op("bold", FIN, "C57", on=False)]
    # ── V32-26 dölj hjälprad/legacy-rad, dölj "0" i objektnamn; D2 nollor som streck ──
    o += [op("hide_rows", FIN, "12"), op("hide_rows", FIN, "36"),
          NF(FIN, "B31:B35", "0;-0;;@"), NF(FIN, "B41:B45", "0;-0;;@"), NF(FIN, "B51:B55", "0;-0;;@"),
          NF(FIN, "D15:BA17", NOLL_STRECK), NF(FIN, "D30:BA35", NOLL_STRECK)]
    # V32-27 (lås etikettkolumner) utgår efter verifiering: sammanslagna ingresser och långa noter i kolumn B
    # klipps vid låsningslinjen på skärm. v3.1:s låsning behålls — tas upp som beslut.

    # ── V32-28 4. NPV: beslutssiffran synlig ──
    o += [S(NPV, "B47", "NPV HELA FASTIGHETEN, krav 3"), op("wrap", NPV, "B47")]
    o += note(NPV, "F47", "Hela fastigheten inkl. Bef-objekt. PV TOTALT (rad 52) prövas mot bokfört värde i krav 3 (2. Kalkyldata D15).")
    o += note(NPV, "F82", "← Krav 1: NPV vid kalkylperiodens slut (kolumn E) ska vara ≥ 0. Förs till 2. Kalkyldata D13.")
    o += [op("bold", NPV, "F82"),
          P(NPV, "E82", "Krav 1 – NPV tillkommande", "Nuvärde av tillkommande driftnetto + restvärde (rad 88) minus PV "
                                                     "investering över hela kalkylperioden. Ska vara ≥ 0. Kravhyran på 2. Kalkyldata rad 17 är den hyra som ger 0 här."),
          P(NPV, "E52", "PV hela fastigheten (krav 3)", "Jämförs mot bokfört värde i 2. Kalkyldata D15. Inkluderar Bef-objekt."),
          S(NPV, "B50", "Restvärde (rad 89)"), S(NPV, "B52", "PV TOTALT (krav 3)"), S(NPV, "B79", "Restvärde tillkommande (rad 88)"),
          P(NPV, "C55:E55", "Investering", "Summa tillkommande investering, 2. Kalkyldata W28 (sektion 4)."),
          P(NPV, "C56:E56", "PV investering", "Investeringen diskonterad till kalkylstart, 2. Kalkyldata X28."),
          P(NPV, "C47", "Kort NPV-horisont", "Antal år för den korta NPV-kolumnen (standard 10). Kan ändras. Den långa "
                                             "kolumnen följer kalkylperioden."),
          P(NPV, "E47", "Kalkylperiod", "Hämtas från 2. Kalkyldata P17 (ändras där).")]
    o += note(NPV, "F62", "År tills ackumulerat driftnetto (ej diskonterat) täcker investeringen. Informativ, förs till 2. Kalkyldata D16.")
    # D8: rad 59 är inte svaret
    o += [S(NPV, "B59", "NPV hela fastigheten")]
    o += note(NPV, "F59", "Informativ, ej krav. Lika med NPV tillkommande (rad 82) när Bef-rader saknas.")

    # ── V32-29 NPV-utskrift (+ D1: upprepa bara A:B) ──
    o += [{"op": "print_area", "sheet": NPV, "area": "$B$1:$X$91"},
          op("print_titles", NPV, "A1", cols="$A:$B", rows="$4:$4"),
          op("page_order", NPV, "A1", order="overThenDown"),
          op("hide_rows", NPV, "3")]
    o += note(NPV, "X10", "Sista kolumnen (år N+1): bara underlag för restvärdet") + [op("align", NPV, "X10", h=RIGHT)]
    o += [op("align", NPV, r, h=LEFT, indent=1) for r in ("F47", "F59", "F62", "F82")]
    o.append(op("colwidth", NPV, "B:B", w=40))          # etiketterna ryms när bara A:B upprepas i utskriften
    # V32-30 (lås etikettkolumn) utgår efter verifiering — B2:K2 är sammanslagen över låsningslinjen.

    # ── V32-31 NPV-rubriker och radnamn ──
    o += [S(NPV, "B1", "Nettonuvärde (NPV)"),
          S(NPV, "B2", "Nettonuvärdet (NPV) är nuvärdet av investeringens framtida driftnetton under kalkylperioden plus "
                       "restvärdet vid periodens slut, diskonterat med kalkylräntan, minus investeringen. Beslutet läses på rad "
                       "82: NPV tillkommande (krav 1, ska vara ≥ 0). Rad 47–62 visar hela fastigheten (krav 3) och rad 84–91 "
                       "hur restvärdet räknas."),
          op("rowheight", NPV, "2", h=50),
          S(NPV, "B5", "Kalkylår nr"), S(NPV, "B6", "Area nybyggnad, kvm"), S(NPV, "B7", "Area befintligt, kvm"),
          S(NPV, "B8", "Total area, kvm"), NF(NPV, "D6:BA7", "#,##0"),
          FMT(NPV, "C11", "B11"), S(NPV, "B11", "kr per år (löpande priser)"), op("align", NPV, "B11", h=LEFT),
          S(NPV, "B18", "Vakans (efter avtal)"), S(NPV, "B31", "Nettohyra, kr/kvm"),
          S(NPV, "B32", "Drift och underhåll exkl extraordinärt UH, kr/kvm"),
          S(NPV, "B53", "kr/kvm"), S(NPV, "B60", "kr/kvm"),
          S(NPV, "B64", "TILLKOMMANDE FLÖDE — hela fastigheten minus Bef-objekt (nollalternativet)"),
          FMT(NPV, "E64", "F64:X64"),
          S(NPV, "B73", "PV driftnetto tillkommande"), S(NPV, "B74", "PV extraordinärt UH tillkommande"),
          S(NPV, "B75", "Utan Bef-rader (befintliga avtal som löper vidare) är tillkommande flöde = hela fastigheten, och "
                        "blocket nedan visar samma siffror som rad 48–59."),
          op("fontcolor", NPV, "B75", color="grey"),
          S(NPV, "B84", "RESTVÄRDE — driftnetto år t+1 / justerad direktavkastning (2. Kalkyldata M11)"),
          FMT(NPV, "E84", "F84:X84"),
          S(NPV, "B90", "Bruttohyra tillk. (underlag Motsv. vakans)"),
          S(NPV, "B91", "Kapitaliseringsbart driftnetto = fullt uthyrt (vakans efter avtal återlagd), exkl. extraordinärt UH. "
                        "Restvärdet vid slutet av år t = driftnetto år t+1 / justerad direktavkastning."),
          op("fontcolor", NPV, "B91", color="grey")]

    # ── V32-32 5. IRR exit-blocket ──
    o += [S(IRR, "B5", "BERÄKNING AV EXIT-VÄRDE VID KALKYLPERIODENS SLUT"), S(IRR, "A5", "Vikt"),
          FMT(IRR, "B9", "B6:B7"), FMT(IRR, "C9", "C6:C7"), FMT(IRR, "D9", "D6:W7"),
          S(IRR, "B6", "Återstående ack. kapitalutlägg"),
          F(IRR, "B7", '="Marknadsvärde, direktavk. "&ROUND(' + K + '$M$11*100,2)&" %"'),
          S(IRR, "B9", "Återstående lån (löses rad 35)"),
          FMT(IRR, "C9", "C10"), op("bold", IRR, "C10"), op("bold", IRR, "C7"),
          S(IRR, "B11", "Justering av försäljningsvärde"),
          P(IRR, "C11", "Eget exit-värde (undantag)", "Lämna tomt = beräknat värde (rad 10). Ett ifyllt belopp i kr ERSÄTTER "
                                                     "exit-värdet helt, det är inget tillägg. Motivera i kommentaren. "
                                                     "2. Kalkyldata visar en notis när fältet är ifyllt."),
          op("hide_rows", IRR, "12"),
          S(IRR, "B13", "Vikter i kolumn A (summa 100 %): 0 % / 100 % / 0 % = rent marknadsvärde vid justerad "
                        "direktavkastning (standard). 0 % / 40 % / 60 % = tidigare LM 371-viktning mot bokfört värde."),
          P(IRR, "A6:A8", "Vikt i exit-värdet", "Standard 0 % / 100 % / 0 % (rent marknadsvärde). Summan av A6:A8 ska vara "
                                                 "100 %. Ändra bara med motivering. Rad 9 (lån) viktas inte."),
          FILL(IRR, "A6:A8", "I23", src_sheet=KD)]                       # D6: vikterna är inmatning
    o += note(IRR, "D10", "= vikt × värde, rad 6–8 (eller C11 om ifyllt). Förs in som Försäljningsvärde (rad 18) exit-året.")
    o += note(IRR, "D11", "← Lämnas normalt tomt. Ett belopp här ersätter exit-värdet i C10 helt – bara i undantagsfall och med motivering.")

    # ── V32-33 kassaflödesdelen på svenska ──
    o += [S(IRR, "B1", "Internräntemetoden (IRR)"), S(IRR, "B14", "RESULTATRÄKNING / KASSAFLÖDE"),
          S(IRR, "C15", "Belopp i kr"), op("align", IRR, "C15", h=RIGHT),
          S(IRR, "C16", "(ingår ej i IRR)"), S(IRR, "C17", "(ingår ej i IRR)"),
          op("italic", IRR, "C16:C17"), op("align", IRR, "C16:C17", h=RIGHT),
          S(IRR, "B20", "Investering, eget kapital"), S(IRR, "B21", "Bokfört värde – projektbelastning (avskrivning)"),
          S(IRR, "B23", "Driftnetto tillkommande (4. NPV)"), S(IRR, "B28", "Resultat före avskrivningar (EBITDA)"),
          S(IRR, "B33", "Resultat efter skatt"),
          S(IRR, "B37", "KASSAFLÖDE, före skatt"), S(IRR, "B38", "KASSAFLÖDE, efter skatt"),
          S(IRR, "B59", "Kassaflöde före skatt — optimistisk"), S(IRR, "B60", "Kassaflöde före skatt — bedömt (= rad 37)"),
          S(IRR, "B61", "Kassaflöde före skatt — pessimistisk")]

    # ── V32-34 känslighetstabellen ──
    o += [S(IRR, "B56", "Bedömt (2. Kalkyldata sektion 5)"), S(IRR, "C54", "Direktavk. vid exit"),
          S(IRR, "F54", "NPV tillk., kr"), S(IRR, "G54", "Mot IRR-krav"), S(IRR, "H54", "Kravhyra krav 1–2, kr/kvm"),
          op("hide_rows", IRR, "58:61")]
    o += note(IRR, "I55", "IRR per scenario räknas i dolda hjälprader 59–61.")

    # ── V32-35 dold introtext + låsta rutor (+ D1 upprepa A:B) ──
    o += [S(IRR, "B2", "Internräntemetoden (IRR) beräknar den årliga avkastningen på eget kapital för det tillkommande "
                       "kassaflödet och jämförs med IRR-kravet på 2. Kalkyldata rad 14.\n\nRestvärdet vid kalkylperiodens slut "
                       "förs in som försäljningsvärde. Standard är rent marknadsvärde (driftnetto år N+1 / justerad "
                       "direktavkastning). Vikterna i kolumn A kan väga in återstående kapitalutlägg eller bokfört värde."),
          op("align", IRR, "I55", h=LEFT, indent=1), op("align", IRR, "E40", h=LEFT, indent=1)]
    # D1 (upprepa bara A:B) gäller endast 4. NPV; flik 3 och 5 behåller $A:$C efter verifieringen

    # ── V32-36/37/38 6. Ack.kapitalutlägg ──
    o += [op("colwidth", ACK, "B:B", w=46),
          FMT(ACK, "A13", "B15:B17"), op("clear_contents", ACK, "A13"),
          FMT(ACK, "C3", "B3"), FMT(ACK, "C4", "B4"), op("align", ACK, "B3:B4", h=LEFT),
          S(ACK, "B3", "Belopp i kr"), S(ACK, "B4", "År i kalkylen"),
          S(ACK, "A2", "Visar hur mycket av investeringen, inklusive räntekrav på bundet kapital, som återstår att tjäna in "
                       "år för år. När Ack. kapitalutlägg (sista raden) når 0 är kapitalet återvunnet."),
          op("merge", ACK, "A2:B2"), op("wrap", ACK, "A2"), op("align", ACK, "A2", h=LEFT, v="top"),
          op("rowheight", ACK, "2", h=36),
          op("bold", ACK, "A2", on=False), op("italic", ACK, "A2"),
          S(ACK, "B6", "(-) Driftnetto (hela fastigheten, 4. NPV)"),
          S(ACK, "B9", "(-) Övriga intäkter/kostnader (5. IRR rad 24–25)"),
          S(ACK, "B11", "(+) Räntekrav (kalkylränta driftnetto, flik 2)"),
          S(ACK, "B12", "(=) Återvunnet kapital (kassaflöde − räntekrav)"),
          S(ACK, "B13", "(=) Ack. kapitalutlägg (kvar att återvinna)"),
          op("italic", ACK, "B15:B17"), op("fontcolor", ACK, "B15:B17", color="grey"),
          S(ACK, "B15", "Tecknen (+)/(−) anger effekten på Ack. kapitalutlägg. Kassaflöde = rad 6 + 7 + 8. Räntekrav = "
                        "föregående års Ack. kapitalutlägg × kalkylränta driftnetto (2. Kalkyldata M7); år 1 investeringen "
                        "× samma ränta."),
          op("wrap", ACK, "B15:B17"), op("align", ACK, "B15:B17", v="top"),
          op("rowheight", ACK, "15", h=48), op("rowheight", ACK, "16", h=36), op("rowheight", ACK, "17", h=36),
          S(ACK, "B16", "Räntebidrag (rad 7, dold) ingår i Kassaflöde men inte i Ack. kapitalutlägg. Övriga intäkter/"
                        "kostnader (rad 9) ingår i Ack. kapitalutlägg men inte i Kassaflöde."),
          S(ACK, "B17", "Ack. kapitalutlägg används i 5. IRR rad 6. Med standardvikterna (0 % / 100 % / 0 %) påverkar "
                        "fliken inte exit-värdet – den är information."),
          {"op": "print_area", "sheet": ACK, "area": "$A$1:$W$17"},
          P(ACK, "B8", "Investeringar", "Investering (4. NPV rad 9) minus projektbelastat bokfört värde (flik 2). "
                                        "Negativt belopp = utbetalning, som ökar kapitalutlägget."),
          P(ACK, "B9", "Övriga intäkter/kostnader", "Ersättningslokaler och tillfälliga intäkter (flik 2 P11–P12, via 5. IRR "
                                                    "rad 24–25). Ingår i Ack. kapitalutlägg men inte i Kassaflöde (rad 10)."),
          P(ACK, "B11", "Räntekrav", "År 1: investeringen × kalkylränta driftnetto (flik 2 M7). Därefter: föregående års "
                                     "Ack. kapitalutlägg × samma ränta."),
          P(ACK, "B13", "Ack. kapitalutlägg", "Kvar att återvinna inklusive räntekrav. Det år raden blir 0 eller negativ "
                                              "är investeringen återbetald."),
          FMT(ACK, "B6", "B9"), FMT(ACK, "C6", "C9:BA9"),
          # D2 nollor som streck, D5 legacy-rad och år 22+ dolda
          NF(ACK, "C5:W12", NOLL_STRECK), NF(ACK, "X5:BA13", "#,##0;-#,##0;"), op("hide_rows", ACK, "7")]
    # år 22+ döljs inte: kalkylperioden kan vara längre än 20 år (P17 = avskrivningstid när Q17 är tom)

    # ── V32-39/40/41/42 7.Grafer (+ D1 en sida, D5 legacy dolt) ──
    o += [op("hide_rows", GRAF, "1:35"),
          S(GRAF, "A36", "Diagramunderlag per kalkylår"), FMT(GRAF, "A1", "A36", src_sheet=ACK),
          op("rowheight", GRAF, "36", h=23.25),
          {"op": "print_area", "sheet": GRAF, "area": "$A$36:$V$57"},
          op("page_setup", GRAF, "A1", fit_wide=1, fit_tall=1)]
    o += note(GRAF, "A37", "Inget fylls i här. Raderna hämtas från flik 3–5 och kan markeras för egna diagram (Infoga › "
                           "Diagram). Rad 44 och 48 används av 4. NPV rad 34–35 – ändra inte.")
    o += [op("colwidth", GRAF, "B:B", w=42),
          S(GRAF, "B40", "Driftnetto tillkommande exkl. extraordinärt UH"), S(GRAF, "B42", "Kostnader exkl. extraordinärt UH"),
          S(GRAF, "B44", "Restvärde tillkommande, årets slut"), S(GRAF, "B47", "Ackumulerat driftnetto inkl. extraordinärt UH"),
          S(GRAF, "B50", "Driftnetto (= rad 40 per kvm)"), S(GRAF, "B51", "Intäkter (= rad 41 per kvm)"),
          S(GRAF, "B52", "Kostnader exkl. extraordinärt UH (= rad 42 per kvm)")]
    o += note(GRAF, "B49", "Rad 40, 44 och 50 avser tillkommande flöde (Bef = nollalternativ), övriga rader hela fastigheten. "
                           "Utan Bef-objekt sammanfaller de.")
    o += [P(GRAF, "B40", "Källa", "4. NPV rad 71 − rad 72. Driftnetto för tillkommande flöde (Bef = nollalternativ), exkl. extraordinärt UH."),
          P(GRAF, "B41", "Källa", "4. NPV rad 19 Nettohyra (efter vakans), hela fastigheten."),
          P(GRAF, "B42", "Källa", "4. NPV rad 26 (drift och underhåll utan re-investering rad 25) + rad 27–28 "
                                  "(fastighetsskatt, tomträttsavgäld), hela fastigheten. Visas som positivt belopp."),
          P(GRAF, "B43", "Källa", "4. NPV rad 25 Re-investering/hyressatt underhåll, hela fastigheten. Visas som positivt belopp."),
          P(GRAF, "B44", "Källa", "4. NPV rad 88 = kapitaliseringsbart driftnetto år t+1 / justerad direktavkastning. Är "
                                  "driftnettot 0 visas bokfört värde (rad 48). Används av 4. NPV rad 34."),
          P(GRAF, "B45", "Källa", "5. IRR rad 16 (investering per år, positivt belopp)."),
          P(GRAF, "B46", "Källa", "5. IRR rad 17."),
          P(GRAF, "B47", "Källa", "Löpande summa av 4. NPV rad 29 Driftnetto (hela fastigheten, inkl. extraordinärt UH)."),
          P(GRAF, "B48", "Källa", "3. Finansiering rad 57. Används av 4. NPV rad 35 (direktavkastning bokfört värde)."),
          P(GRAF, "B50:B52", "Källa", "Rad 40/41/42 delad med total area, 4. NPV rad 8."),
          P(GRAF, "B57", "Källa", "3. Finansiering rad 27.")]
    for r in range(40, 49):
        o.append(S(GRAF, f"A{r}", "kr"))
    o += [S(GRAF, "A57", "%"), op("colwidth", GRAF, "A:A", w=8),
          S(GRAF, "B38", "Kalenderår"), S(GRAF, "B39", "Kalkylår"),
          FMT(GRAF, "D13", "A38:W38", src_sheet=FIN), FMT(GRAF, "D14", "A39:W39", src_sheet=FIN),
          FMT(GRAF, "B51", "A51"), FMT(GRAF, "A53", "A52"), FMT(GRAF, "B44", "A44"),
          op("freeze", GRAF, "A40"),                      # bara årsrubrikerna låses (radlåsning som flik 4 och 6)
          op("hide_rows", GRAF, "53:56")]

    # ── V32-43/44 Kontoplan RR ──
    o += [op("rowheight", KONTO, "1", h=23.25), FMT(KONTO, "A1", "E1", src_sheet=ACK), S(KONTO, "E1", "KONTOPLAN RR"),
          op("merge", KONTO, "E2:I3"),
          S(KONTO, "E2", "Konton som ingår i driftkostnadsposterna på 2. Kalkyldata (rad 34–38). Tabellerna nedan är "
                         "källor till rullistor – ändra inte texterna."),
          op("wrap", KONTO, "E2"), op("italic", KONTO, "E2"),
          op("clear", KONTO, "E4:F4"), FMT(KONTO, "H5", "E5:F5"),
          S(KONTO, "E5", "Typ av åtgärd"), S(KONTO, "F5", "Förkortning")]
    o += note(KONTO, "E12", "Förkortning = rullistan Typ på 2. Kalkyldata (E23:E27).")
    o += note(KONTO, "E13", "Typ av projekt = rullistan på 1. Framskrivningsunderlag (G11).")
    o += note(KONTO, "E14", "Ändra inte texterna – formler på flik 1 läser listorna.")
    o += [S(KONTO, "A1", "Media → flik 2, rad 37"), S(KONTO, "A9", "Fastighetsskötsel → flik 2, rad 34"),
          S(KONTO, "A16", "Reparationer → flik 2, rad 35"), S(KONTO, "A24", "Planerat underhåll → flik 2, rad 36"),
          S(KONTO, "A33", "Övriga förvaltningskostnader → flik 2, rad 38"),
          op("colwidth", KONTO, "B:D", w=2), op("colwidth", KONTO, "G:G", w=2),
          op("page_setup", KONTO, "A1", fit_wide=1, fit_tall=1)]
    return o


def ops_historik_v32() -> list[dict]:
    return [FMT(FS, "R25", "R26"), FMT(FS, "S25", "S26"),
            S(FS, "R26", "18. 2026-10-11 v3.2: Användarvänlighet och design efter agentgranskning — inmatningsmeddelanden på "
                         "alla inmatningsfält, krav 1/2/3 namngivna överallt, två decimaler på räntor/IRR, stavfel och "
                         "engelska rester rättade, arbetsgång på flik 2, källhänvisningar mellan flikar, hjälprader dolda, "
                         "utskrift (sidbrytning flik 1, årsrubrik och sidordning flik 4), sidfot v3.2. Inga beräkningar ändrade."),
            S(FS, "S26", "Claude/JW")]
