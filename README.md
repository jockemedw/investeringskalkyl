# Handoff-paket: Investeringskalkyl Lejonfastigheter

**För Claude Code (eller annan ny session) — börja här.**

## Läsordning

1. **HANDOFF.md** ← BÖRJA HÄR. Övergripande projektkontext, arkitektur, kritisk diskrepans mellan minne och projektfil.
2. **OPEN_QUESTIONS.md** ← Q-01 är BLOCKER. Reda ut den med Joakim innan något annat.
3. **DECISIONS.md** — varför arket är konstruerat som det är.
4. **TECH_NOTES.md** — kodbara mönster (openpyxl-fallgropar, XML-patch, regressionstest).

## Pipelinestruktur (etablerad 2026-05-08)

```
investeringskalkyl/
├── Investeringskalkyl_iter8.xlsx   # iter 8 baseline (auktoritativ)
├── tools/
│   ├── patches.py    # load/save/diff/rename/replace utilities (TECH "kirurgiskt")
│   └── recalc.py     # tvingar formelutvärdering via Excel COM eller LibreOffice
├── tests/
│   └── regression.py # verifierar 11 631 222 / 7,65 % / +1,35 pp
└── build/
    └── demo_roundtrip.py   # round-trip-test som verifierar att pipeline fungerar
```

## Snabbstart

```bash
pip install openpyxl

# Verifiera baseline direkt (cached values från iter8):
python tests/regression.py

# Round-trip-test (load → save → recalc → regression):
python build/demo_roundtrip.py
```

## Arbetsflöde för iter 8 → iter 9

```python
from tools.patches import load_iter, save_iter, ITER8
from tools.recalc import recalc

wb = load_iter(ITER8)
ws = wb["Indata"]
ws["B18"] = "Avskrivningstakt"     # round B / round E patches här
save_iter(wb, "Investeringskalkyl_iter9.xlsx")
recalc("Investeringskalkyl_iter9.xlsx")  # uppdaterar cached values via Excel COM

# Verifiera att inga siffror flyttat sig:
python tests/regression.py Investeringskalkyl_iter9.xlsx
```

`recalc` använder LibreOffice headless om installerat, annars Excel COM via PowerShell (kräver Excel installerat — fungerar på Joakims maskin).

## Joakims arbetsstil — sammanfattat

- Plan först, exekvering sen. Batch-godkännande.
- Direkta korrigeringar vid fel — inga utdragna ursäkter.
- Svenska, koncist.
- Beslut motiveras mot LM 371 / branschpraxis.
- Regressionstest = sanningens domstol.

## Kritiska principer

- **Kassaflöde stannar vid driftnetto** (D-05). Räntor och avskrivningar tillhör Finansiering-fliken.
- **LM 371 är auktoritativ referens** (D-14). Avvik bara med dokumenterat skäl.
- **Modifiera kirurgiskt, inte återskapa från grunden** — undantag: själva iter-replacement.

---

*Lycka till — Joakim är en bra samarbetspartner, var direkt och resultatorienterad.*
