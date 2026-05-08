# Öppna frågor inför fortsättning

Dessa frågor måste redas ut **innan** ny utveckling påbörjas i Claude Code. Sortera in svaren här när Joakim svarat och uppdatera HANDOFF.md i konsekvens.

---

## Q-01 LÖST 2026-05-08

Iter 8-filen `Investeringskalkyl_iter8.xlsx` finns i repot. 8 flikar verifierade. Cached values matchar baslinjen (11 631 222 / 7,65 %). Build-scripten är inte åtkomliga — fas 2 rekonstruerar dem från xlsx-filen.

## Q-01b (NY): Defaults vs konkret testfall

Värdena i iter8-filen (skatt 20,6 %, kalkylränta 4 %, CA 80 kr/kvm, direktavkastning 5 %) — är dessa de **defaults Lejonfastigheter ska se i en tom mall**, eller är de **specifika för Skola-testfallet**? Påverkar hur build-scriptet ska initiera filen.

---

## Q-02: Vilken är prio-ordningen i att-göra-listan?

Fem öppna punkter (HANDOFF §10):
1. Pedagogisk omskrivning Beräkningslogik (round B)
2. Indata-fält renamning (Avskrivningstakt, räntenivå)
3. Designcleanup Lönsamhetskontroll (round F)
4. "År N" → "år 20" (round E)
5. Full design review inkl. Översikt

Ingen är blockerande — Joakim väljer.

---

## Q-03: Finns Formelkatalogen från reverse engineering?

Den byggdes i fas 2 från LM 371 (~72 700 formler). Bör finnas lokalt. Användbar som referens vid framtida frågor om "varför räknar LM 371 så här". Om den finns: ladda upp till projektet.

---

## Q-04: Är Lejonfastigheters Bilaga C uppdaterad?

DoU-schablon 405 kr/kvm + 60 kr/kvm CA är defaults i kalkylen. Stämmer dessa fortfarande? Om Bilaga C reviderats måste defaultvärdena uppdateras.

---

## Q-05: Är IRR-kravet 6,3 % stillastående?

Default i Indata. Om Lejonfastigheter höjt/sänkt kravet (t.ex. som följd av räntehöjningar) måste defaulten justeras.

---

## Q-06: Build-miljö — vilken Python-version?

Trivialt men praktiskt: Joakims setup (Python 3.x, vilken openpyxl-version, LibreOffice-version på OS-nivå). Viktigt för reproducerbarhet om Claude Code ska köra scripten lokalt.

---

## Q-07: Lagring av iterationerna

Ligger iter 1-7 sparade någonstans (för historisk regression)? Eller skrivs varje iter över? Påverkar hur regression mot **äldre** baseline ska hanteras (om man någon gång vill verifiera att en bug-fix inte regression-trasat något långt bak).

---

## Q-08: Användartest

Har iter 8 visats för faktiska användare (utvecklare, kundansvariga, LCC) ännu? Om ja: feedback samlad någonstans? Om nej: är användartest planerat innan launch?

---

## Q-09: Hur "frys" ser ut

När anses iter 8 (eller en framtida iter) "klar"? Acceptkriterier utöver regressionsbaseline? Vem signerar?

---

*Lägg till nya frågor löpande. När en fråga är besvarad: stryk den och uppdatera HANDOFF.md / DECISIONS.md i konsekvens.*
