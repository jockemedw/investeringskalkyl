"""v2-förbättringsmoduler ovanpå spec-replayen. En funktion per V2-beslut (ONESHOT.md).

Milstolpe 2: tom pipeline (ren replay valideras först).
Milstolpe 3 fyller på: Översikt-redesign, tkr-fix, Grafer, locale-fix, sidnav.
"""
from __future__ import annotations
from pathlib import Path


def fix_iter9_name_errors(wb) -> None:
    """F-3: två pedagogiska anteckningar i iter9 är inskrivna som formler
    ('=-npv_0 / b', '=max av de tre') → #NAME? i produktionsfilen.
    v2 lagrar dem som text — samma innehåll, inget formelfel."""
    ws = wb["Beräkningslogik"]
    # OBS: får inte börja med "=" — openpyxl skriver det som formel igen.
    ws["D18"].value = "dvs −NPV₀ / b"
    ws["D30"].value = "dvs max av de tre"


def apply_all(wb) -> None:
    """Körs på workbook-objektet innan save."""
    fix_iter9_name_errors(wb)


def post_save(out_path: Path) -> None:
    """Körs på den sparade filen (XML-nivå) efter save + summaryBelow-patch."""
    pass
