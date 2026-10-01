"""Design- och städrunda (körs sist, före gotos). Behåller LM 371:s formspråk men gör det konsekvent.

Färger: mallens tretons-hierarki behålls (mörkblå+vit = nyckelantagande, ljusblå = tabellinmatning,
mellanblå = beräknat i tabell). Nya celler ärver format från motsvarande mallceller (fmt_from).
Struktur:
  hjälpområden döljs (flik 2 AC:CB), fel-/diagnostikceller rensas, årskolumner breddas så ##### försvinner,
  fönsterlås på etikettkolumner, 100 % zoom, rutnät av.
Utskrift:
  flik 1 + Kontoplan A4 stående, övriga A3 liggande; alltid 1 sida bred; enhetliga marginaler; sidfot
  "Lejonfastigheter · LM 371 v3 · <flik> · sida x av y · datum". Grafer får print_area (var 3 282 sidor).
"""
from __future__ import annotations
from step4_bef import _set, _fmt, KD, NPV, IRR, GRAF, FS, FIN

ACK, KONTO = "6. Ack.kapitalutlägg", "Kontoplan RR"
INPUT, CALC = 0.4, 0.8
WHITE, BLACK = 1, 2                     # empiriskt i denna bok: Font.ThemeColor 1 → vit, 2 → svart
PORTRAIT, LANDSCAPE, A3, A4 = 1, 2, 8, 9
MARGINS = [1.0, 1.0, 1.5, 1.5, 0.6, 0.6]   # cm: L R T B header footer
# Svenska huvud/fot-koder via COM: &S = sidnummer, &A = antal sidor, &D = datum (&P/&N betyder annat)
FOOT_L, FOOT_R = "Lejonfastigheter · Investeringskalkyl LM 371 v3", "Sida &S av &A · &D"
YEAR_W = 13.0


def _fill(sheet, ref, tint):
    return {"op": "fill", "sheet": sheet, "ref": ref, "tint": tint}


def _font(sheet, ref, theme):
    return {"op": "fontcolor", "sheet": sheet, "ref": ref, "theme": theme}


def _page(sheet, orient, paper, area=None, tall=0, titles_rows=None, titles_cols=None, wide=1):
    d = {"op": "page_setup", "sheet": sheet, "orient": orient, "paper": paper, "fit_wide": wide, "fit_tall": tall,
         "margins_cm": MARGINS, "foot_l": FOOT_L, "foot_c": sheet, "foot_r": FOOT_R}
    if area:
        d["area"] = area
    if titles_rows:
        d["titles_rows"] = titles_rows
    if titles_cols:
        d["titles_cols"] = titles_cols
    return d


def ops_design() -> list[dict]:
    o: list[dict] = []
    # Färger: mallens egen tretons-hierarki behålls orörd (mörk+vit = nyckelantagande,
    # ljus = tabellinmatning, mellan = beräknat). Nya celler ärver via fmt_from i stegfilerna.
    # Diagnostikceller i rad 2 (LEN/ISBLANK) — ingen refererar dem
    o += [{"op": "clear", "sheet": KD, "ref": "F2"}, {"op": "clear", "sheet": KD, "ref": "V2"}]
    # Hjälpområden: dropdown-skala, datatabell, årsvisa hjälptabeller → dolda kolumner
    o += [{"op": "clear_outline", "sheet": KD, "ref": "AC:AP"},   # mallens +-knapp skulle annars visa hjälpområdet
          {"op": "hide_cols", "sheet": KD, "ref": "AC:CB"}]
    # ── Årskolumner: ##### bort ───────────────────────────────────────────────
    o += [{"op": "colwidth", "sheet": KD, "ref": "D:D", "w": 11.0},          # "10 300 tkr" i D13:E13 vid utskrift
          {"op": "colwidth", "sheet": KD, "ref": "P:P", "w": 14.0},          # "80 kr/kvm" klipptes i utskrift
          {"op": "colwidth", "sheet": NPV, "ref": "D:BA", "w": YEAR_W},
          {"op": "colwidth", "sheet": IRR, "ref": "D:BA", "w": YEAR_W},
          {"op": "colwidth", "sheet": FIN, "ref": "D:BA", "w": 12.0},
          {"op": "colwidth", "sheet": ACK, "ref": "C:BA", "w": 12.0}]
    # ── Fönsterlås, zoom, rutnät ──────────────────────────────────────────────
    # Radlås på årsrubriken (kolumnlås ger dubbelritad rubriktext över delningen)
    for sh, ref in ((NPV, "A5"), (ACK, "A5")):
        o.append({"op": "freeze", "sheet": sh, "ref": ref})
    for sh, zoom in ((FS, 80), (KD, 100), (FIN, 100), (NPV, 100), (IRR, 100), (ACK, 100), (GRAF, 100), (KONTO, 100)):
        o.append({"op": "view", "sheet": sh, "zoom": zoom, "gridlines": False})
    # ── Utskrift ──────────────────────────────────────────────────────────────
    o += [_page(FS, PORTRAIT, A4, "$A$1:$H$138"),
          _page(KD, LANDSCAPE, A3, "$C$3:$U$81", tall=1),
          _page(FIN, LANDSCAPE, A3, "$A$3:$W$57", titles_cols="$A:$C", wide=2),
          _page(NPV, LANDSCAPE, A3, "$B$1:$Y$82", titles_cols="$A:$C", wide=2),
          _page(IRR, LANDSCAPE, A3, "$A$1:$W$61", titles_cols="$A:$C", wide=2),
          _page(ACK, LANDSCAPE, A3, "$A$1:$W$13", tall=1, titles_cols="$A:$B", wide=2),
          _page(GRAF, LANDSCAPE, A3, "$A$38:$W$57", tall=1, titles_cols="$A:$B", wide=2),
          _page(KONTO, PORTRAIT, A4, "$A$1:$I$37", tall=1)]
    # Flik 5: två etiketter som klipps av grannkolumnen (mallen)
    o += [_set(IRR, "B10", v="Bedömt försäljningsvärde, exit"), _set(IRR, "C37", v="(avskr. återförda)")]
    # NPV: kolumn C i tillkommande-blocket var ofylld (vit remsa)
    o += [_fmt(NPV, "C20", f"C65:C74")]
    # NPV: sidbrytning vid blockgräns (Tillkommande flöde) i stället för mitt i blocket
    o += [{"op": "hpagebreak", "sheet": NPV, "ref": "A64"}]
    # Grafer: diagramunderlaget står i ljusgrå text — läsbart i utskrift
    o += [_font(GRAF, "A38:W57", BLACK)]
    return o
