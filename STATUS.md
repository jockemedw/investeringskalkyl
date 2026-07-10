---
goal: "Ersättare för LM 371 Investeringskalkyl — 20-årig kalkyl i xlsx för Lejonfastigheter"
tags: [python, openpyxl, excel, fastigheter]
---

# Investeringskalkyl

## Läge (2026-07-10)

**Produkten är klar för granskning:** `build/oneshot/Investeringskalkyl_v2.xlsx` — 10 flikar, regression grön (11 631 221,72 kr/år / 7,6452 % / +1,35 pp), 18 print-sidor (trogen-PDF-verifierad). ONESHOT-POLISH klar: ett blått input-språk med 51 svenska valideringar, bladskydd med Tab-vandring, öppningsvyer, ifyllnadsguide + status, ren tom mall. Slutrapporter i [FINAL.md](FINAL.md) och [POLISH.md](POLISH.md). Main är redo att pushas (Joakim pushar efter granskning).

## Uppgifter
- [x] Iter 8 baseline — kravhyra 11,6 Mkr, IRR EK 7,65 %, marginal +1,35 pp
- [x] Iter 9 designrundor B–AE (pedagogik, branschterminologi, designsystem, känslighetstabeller)
- [x] ONESHOT v2 — ombyggd motor från spec-replay, exakt baslinje, Grafer-flik, Översikt-redesign
- [x] FINAL m1 — utskriftsformat v2: 31 → 18 sidor, all data kvar (D-22)
- [x] FINAL m2 — Faktisk IRR EK per yield-scenario + timing-fix MV-exit (D-20, D-21)
- [x] FINAL m3 — design-excellens-pass, 2 hela PDF-granskningsvarv
- [x] POLISH m1–m5 — ifyllnadsupplevelsen: input-språk, validering, bladskydd, öppningsvyer, tom mall, 2 skärm-granskningsvarv (D-23–D-25)
- [ ] Joakim: granska + pusha main till origin

## Anteckningar
Bygg: `python build/oneshot/build_v2.py` (replay + rundor + recalc + regressionsgate). Print/design verifieras ENDAST med trogen export (`build/oneshot/export_pdf.py`) — render_local//xlsx-review och PageSetup.Pages.Count är opålitliga. Testfall: Skola (Nyb) 5 000 kvm × 40 000 kr/kvm = 200 Mkr.
