# Designbeslut — Investeringskalkyl Lejonfastigheter

Beslut som har tagits genom projektet, med motivering. Varje beslut är låst om inte explicit ompröva (Joakim flaggar det).

---

## D-01: Ren Excel utan makron
**Beslut:** Verktyget byggs som ren `.xlsx` utan VBA/makron.
**Motivering:** Bred tillgänglighet bland användarroller (utvecklare, kund-/hyresgästansvariga, LCC-analytiker). Minskar säkerhetsfriktion (makro-varningar) och underlättar versionshantering.
**Konsekvens:** All logik måste uttryckas i Excel-formler. Programmatisk konstruktion sker via Python (openpyxl) i build-tid, inte i runtime.

---

## D-02: Programmatisk konstruktion via openpyxl
**Beslut:** Arket byggs deterministiskt från Python-script.
**Motivering:** Möjliggör versionskontroll i git, regressionstest, reproducerbara byggen och iterationer som kan jämföras programmatiskt.
**Konsekvens:** Alla ändringar görs i scripten, inte direkt i Excel. Filen är en artefakt, inte en källa.

---

## D-03: 20-års kalkylhorisont
**Beslut:** Kassaflöden projiceras 20 år framåt.
**Motivering:** Matchar LM 371 + svensk branschpraxis för fastighetsinvesteringskalkyl + typisk avtalslängd för längre kontrakt.
**Konsekvens:** Restvärdesberäkning år 20 är kritisk komponent.

---

## D-04: Rollseparerad arkitektur (8 flikar)
**Beslut:** Tydlig uppdelning: Översikt, Indata, Kassaflöde, Finansiering, Resultat, Lönsamhetskontroll, Beräkningslogik, Dokumentation.
**Motivering:** LM 371 blandar allt. Separation gör arket lättare att förstå för olika användarroller — en utvecklare behöver inte läsa Beräkningslogik för att fylla i Indata.
**Konsekvens:** Mer cross-flik-referenser. Hyperlänkar i Översikt blir viktiga.

---

## D-05: Kassaflöde stannar vid driftnetto
**Beslut:** Fliken `Kassaflöde` projicerar bruttohyra → vakans → nettohyra → drift/FS/TR/CA → DRIFTNETTO. Stoppar där.
**Motivering:** Matchar LM 371-strukturen exakt. Driftnetto är det rätta kassaflödesbegreppet för värderingen — räntor och avskrivningar är finansierings-/redovisningskonstrukt och hör hemma på sin egen flik.
**Konsekvens:** Finansieringsflik måste finnas separat (D-06). NPV och IRR konsumerar driftnetto-strömmen.

> Detta är en av projektets viktigaste principer. Den nuvarande filen i `/mnt/project/` följer INTE denna princip — flik 3 där blandar driftnetto med EBITA, skatt, finansiering och avskrivning.

---

## D-06: Dedikerad Finansiering-flik (Block C)
**Beslut:** Lån, ränta, ränteupptrappning, amortering ligger på egen flik.
**Motivering:** Spegling av LM 371. Möjliggör att variera finansieringsantaganden utan att röra kassaflödesmodellen. Renar Beräkningslogik.
**Konsekvens:** Equity-IRR blir möjlig att räkna ut utan att förorena driftnetto-strömmen.

---

## D-07: Tre lönsamhetskriterier
**Beslut:** Lönsamhet utvärderas mot tre kriterier:
1. **NPV** av driftnetto > 0 (med given kalkylränta)
2. **Equity-IRR** > avkastningskrav (typ 6,3 %)
3. **Marknadsvärde > bokfört värde** vid horisontens slut

**Motivering:** Standardansats i svensk fastighetsekonomi. Varje kriterium fångar olika dimension: NPV = värdeskapande, IRR = kapitalavkastning, MV/BV = nedskrivningsrisk.
**Konsekvens:** Lönsamhetskontroll-fliken visar alla tre + status (godkänd/ej godkänd) per rad.

---

## D-08: Delta-metoden i Beräkningslogik
**Beslut:** Ändringar mot nuläget räknas som deltan, inte som absoluta nya nivåer.
**Motivering:** Investeringskalkyl handlar om **inkrementell** lönsamhet — vad betalar sig av denna åtgärd ovanpå vad som redan finns. Delta-metoden gör det explicit.
**Konsekvens:** Beräkningslogik-fliken implementerar Block A, B, D, E, F som deltan.

---

## D-09: Restvärdesviktning 40/60 (MV/BV)
**Beslut:** Standardrestvärde år 20 beräknas som `0,4 × marknadsvärde + 0,6 × bokfört värde`.
**Motivering:** Konservativ ansats — viktar mer mot bokfört värde än mot framtida marknadsvärde. Speglar Lejonfastigheters interna policy.
**Konsekvens:** Restvärdet blir lägre än ren marknadsvärdering, vilket pressar ner IRR och kräver högre hyra för godkänd kalkyl.

---

## D-10: Belåningsgrad 37 % (default)
**Beslut:** Default belåningsgrad i Indata är 37 %.
**Motivering:** Portföljbaserad, ej objektspecifik — speglar Lejonfastigheters faktiska kapitalstruktur. Användare kan justera per kalkyl.
**Konsekvens:** Eget kapital-andel = 63 %. IRR-beräkning baseras på den.

---

## D-11: DoU-schablon 405 kr/kvm + 60 kr/kvm CA
**Beslut:** Default driftkostnader hämtas från Bilaga C: el/media 405 kr/kvm, central administration 60 kr/kvm.
**Motivering:** Standardiserad bilaga internt hos Lejonfastigheter.
**Konsekvens:** Avvikelser måste dokumenteras i Antaganden-fliken.

---

## D-12: Kalkylränta 4,55 % på driftnetto
**Beslut:** Kalkylräntan för NPV på driftnetto är 4,55 % (default).
**Motivering:** Reflekterar långsiktig WACC för Lejonfastigheter / kommunalt fastighetsbolag.
**Konsekvens:** Lägre ränta → högre NPV → lättare att passa kravet → lägre erforderlig hyra. Justeras per kalkyl vid behov.

---

## D-13: Reverse engineering före replacement
**Beslut:** LM 371 reverse-engineerades fullständigt (~72 700 formler, Formelkatalog) **innan** ersättaren byggdes.
**Motivering:** Förstå *varför* LM 371 räknar som den gör innan beslut tas att göra annorlunda. Bevarar institutionell kunskap.
**Konsekvens:** Iter 1 byggdes med dokumenterad förståelse + tre konfirmerade buggfixar (se HANDOFF §6).

---

## D-14: LM 371 är auktoritativ referens
**Beslut:** Vid varje strukturellt eller metodiskt val är LM 371 referensen. Avvikelser kräver explicit motivering i denna fil.
**Motivering:** Verktyget måste accepteras av branschvana användare. Drastiskt avvikande logik skulle skapa förtroendeproblem.
**Konsekvens:** "Vi gör som LM 371" är giltigt argument utan vidare diskussion. "Vi gör annorlunda än LM 371" kräver dokumenterat skäl.

---

## D-15: Regressionstest som accept-kriterium
**Beslut:** Varje iteration måste passera regressionstest mot iter 7-baseline (11 631 222 / 7,65 % / +1,35 pp) innan den anses klar.
**Motivering:** Skydd mot subtila buggintroduktioner. Pedagogisk omarbetning får inte ändra siffror.
**Konsekvens:** Vid avvikelse måste antingen ändringen rullas tillbaka eller baseline uppdateras med dokumenterad anledning.

---

## D-16: Hyresspann ±X % på investeringen
**Beslut:** Resultat-fliken redovisar tre kravhyror: vid Lägsta utfall (−X %), Mål-utfall, Högsta utfall (+X %). X anges i Indata sektion 8 (default 10 %).
**Motivering:** Investeringsbudgetar är osäkra på beslutspunkten. Att redovisa ett hyresspann snarare än en enskild punktsiffra ger förhandlaren rätt verktyg när hyresavtalet skrivs.
**Konsekvens:** Översikt visar både mål-utfall och spann. Lönsamhetskontroll utvärderar mål-utfallet. Hyresavtalets villkor bör formuleras med spannet i åtanke.
**Verifierat:** Resultat!C14:E14, Resultat!E16, Indata!C68 (filen 2026-05-08).

---

## D-17: Analytisk lösning av kravhyran
**Beslut:** Beräkningslogik löser kravhyran analytiskt via NPV(hyra) = NPV_0 + b × hyra → hyra = −NPV_0 / b. Två referenspunkter (hyra=0 och hyra=1 Mkr) ger lutningen b.
**Motivering:** Excel-goal-seek/iteration kan vara opålitligt och skapar fragilitet. Analytisk lösning är deterministisk, snabb och läsbar (steg 1-3 syns direkt i Beräkningslogik).
**Konsekvens:** Kassaflöde måste vara linjärt i kravhyran (verifierat: bruttohyra, vakans, nettohyra skalas linjärt). Tre separata analytiska lösningar — en per kriterium (NPV, IRR, MV/BV) — i Beräkningslogik C119, C177, C204.
**Verifierat:** Beräkningslogik!C7-C18, Resultat!D10:D12.

---

## D-18: OMPRÖVAR D-11 — Drift och underhåll som matris bef/nyb
**Beslut:** Drift och underhåll matas in som matris i Indata sektion 4: 5 kostnadsposter (Fastighetsskötsel, Reparationer, Planerat underhåll, Media, Övriga förvaltningskostnader) × 2 kategorier (Befintligt, Nybyggnad), enhet kr/kvm/år. CA (central administration) ligger separat i sektion 6.
**Motivering:** Den gamla schablonen "405 kr/kvm el/media + 60 kr/kvm CA" var för grov för differentierad kalkyl mellan befintligt bestånd och nybyggnation. LM 371 skiljer normalt på dessa.
**Konsekvens:** D-11:s schablon är ersatt. Nya defaults i iter 8 (för Nyb): FS 50, Rep 30, PU 40, Media 150, Övr 130 kr/kvm = totalt 400 kr/kvm/år. CA = 80 kr/kvm. Befintligt-kolumnen är 0 i baslinjefallet (rent nybyggnationsprojekt).
**Verifierat:** Indata!B36-D40, Indata!C53.

---

## D-19: OMPRÖVAR D-12 — Kalkylränta driftnetto 4,0 %
**Beslut:** Kalkylränta driftnetto = 4,0 % i iter 8 (var 4,55 % i tidigare dokumentation).
**Motivering:** Reflekterar uppdaterad WACC-bedömning för Lejonfastigheter på rådande räntemarknad. Kan justeras per kalkyl.
**Konsekvens:** Lägre kalkylränta → högre NPV → lägre kravhyra. Baslinjen 11 631 222 kr/år är räknad mot 4,0 %.
**Verifierat:** Indata!C7 (filen 2026-05-08).

---

*Lägg till nya beslut längst ner med löpande nummer. Ändra ALDRIG befintliga — markera som "OMPRÖVAD i D-XX" om de ersätts.*
