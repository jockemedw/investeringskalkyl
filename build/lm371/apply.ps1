# Generisk COM-applier: öppnar Source, kör ops (JSON), räknar om, sparar Out (om angiven),
# skriver probes/scan-resultat till Probes (JSON). Formler i ops är en-US-syntax (COM .Formula).
param(
    [Parameter(Mandatory)][string]$Source,
    [string]$Out = "",
    [Parameter(Mandatory)][string]$Ops,
    [Parameter(Mandatory)][string]$Probes
)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$opsList = Get-Content -Raw -Encoding UTF8 $Ops | ConvertFrom-Json

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($Source)
try {
    foreach ($o in $opsList) {
        $sheet = [string]$o.sheet; $ref = [string]$o.ref
        $ws = $wb.Worksheets.Item($sheet)
        switch ([string]$o.op) {
            "set" {
                $r = $ws.Range($ref)
                if ($null -ne $o.formula) { $r.Formula = [string]$o.formula }
                elseif ($o.value -is [string]) { $r.Value2 = [string]$o.value }
                else { $r.Value2 = [double]$o.value }
            }
            "numfmt"     { $ws.Range($ref).NumberFormat = [string]$o.fmt }
            "fmt_from"   { $ws.Range([string]$o.src).Copy() | Out-Null; $ws.Range([string]$o.dst).PasteSpecial(-4122) | Out-Null }
            "merge"      { $ws.Range($ref).Merge() | Out-Null }
            "clear"      { $ws.Range($ref).Clear() | Out-Null }
            "clear_contents" { $ws.Range($ref).ClearContents() | Out-Null }
            "clear_fmt"  { $ws.Range($ref).ClearFormats() | Out-Null }
            "replace"    { $ws.Range($ref).Replace([string]$o.find, [string]$o.repl, 2) | Out-Null }
            "dv_list"    { $r = $ws.Range($ref); $r.Validation.Delete(); $r.Validation.Add(3, 1, 1, [string]$o.source) | Out-Null; $r.Validation.InCellDropdown = $true
                           if ($null -ne $o.prompt) { $r.Validation.InputTitle = [string]$o.prompt_title; $r.Validation.InputMessage = [string]$o.prompt } }
            "print_area" { $ws.PageSetup.PrintArea = [string]$o.area }
            "align"      { $ws.Range($ref).HorizontalAlignment = [int]$o.h }
            "wrap"       { $ws.Range($ref).WrapText = $true }
            "bold"       { $ws.Range($ref).Font.Bold = $true }
            "italic"     { $ws.Range($ref).Font.Italic = $true }
            "colwidth_from" { $ws.Columns([string]$o.dst).ColumnWidth = $ws.Columns([string]$o.src).ColumnWidth }
            "numfmt_local" { $ws.Range($ref).NumberFormatLocal = [string]$o.fmt }
            "insert_rows" { $ws.Rows($ref).Insert() | Out-Null }
            "datatable"  { $ws.Range($ref).Table([System.Reflection.Missing]::Value, $ws.Range([string]$o.col_input)) | Out-Null }
            "fill"       { $r = $ws.Range($ref); $r.Interior.Pattern = 1; $r.Interior.ThemeColor = 5; $r.Interior.TintAndShade = [double]$o.tint }
            "hpagebreak" { $ws.HPageBreaks.Add($ws.Range($ref)) | Out-Null }
            "inside_h_none" { $ws.Range($ref).Borders(12).LineStyle = -4142 }
            "nofill"     { $ws.Range($ref).Interior.Pattern = -4142 }
            "fontcolor"  { $ws.Range($ref).Font.ThemeColor = [int]$o.theme }
            "clear_outline" { $ws.Range($ref).EntireColumn.ClearOutline() | Out-Null }
            "hide_cols"  { $ws.Range($ref).EntireColumn.Hidden = $true }
            "colwidth"   { $ws.Range($ref).EntireColumn.ColumnWidth = [double]$o.w }
            "freeze"     { $ws.Activate() | Out-Null; $xl.ActiveWindow.FreezePanes = $false; $ws.Range($ref).Select() | Out-Null; $xl.ActiveWindow.FreezePanes = $true }
            "view"       { $ws.Activate() | Out-Null; $xl.ActiveWindow.Zoom = [int]$o.zoom; $xl.ActiveWindow.DisplayGridlines = [bool]$o.gridlines }
            "page_setup" {
                $ps = $ws.PageSetup
                if ($null -ne $o.area) { $ps.PrintArea = [string]$o.area }
                $ps.Orientation = [int]$o.orient; $ps.PaperSize = [int]$o.paper
                $ps.Zoom = $false; $ps.FitToPagesWide = [int]$o.fit_wide
                if ([int]$o.fit_tall -gt 0) { $ps.FitToPagesTall = [int]$o.fit_tall } else { $ps.FitToPagesTall = $false }
                $m = $o.margins_cm
                $ps.LeftMargin = $xl.CentimetersToPoints([double]$m[0]); $ps.RightMargin = $xl.CentimetersToPoints([double]$m[1])
                $ps.TopMargin = $xl.CentimetersToPoints([double]$m[2]); $ps.BottomMargin = $xl.CentimetersToPoints([double]$m[3])
                $ps.HeaderMargin = $xl.CentimetersToPoints([double]$m[4]); $ps.FooterMargin = $xl.CentimetersToPoints([double]$m[5])
                $ws.ResetAllPageBreaks()
                $ps.OddAndEvenPagesHeaderFooter = $false
                $ps.LeftFooter = [string]$o.foot_l; $ps.CenterFooter = [string]$o.foot_c; $ps.RightFooter = [string]$o.foot_r
                $ps.LeftHeader = ""; $ps.CenterHeader = ""; $ps.RightHeader = ""
                # Mallens förstasideshuvud: behåll logotypen (&G), släng platshållartexterna, ge samma sidfot
                if ($ps.DifferentFirstPageHeaderFooter) {
                    $fp = $ps.FirstPage
                    if (([string]$fp.LeftHeader.Text) -match "&G") { $fp.LeftHeader.Text = '&"Arial,Regular"&8&G' } else { $fp.LeftHeader.Text = "" }
                    $fp.CenterHeader.Text = ""; $fp.RightHeader.Text = ""
                    $fp.LeftFooter.Text = [string]$o.foot_l; $fp.CenterFooter.Text = [string]$o.foot_c; $fp.RightFooter.Text = [string]$o.foot_r
                }
                if ($null -ne $o.titles_rows) { $ps.PrintTitleRows = [string]$o.titles_rows }
                if ($null -ne $o.titles_cols) { $ps.PrintTitleColumns = [string]$o.titles_cols }
            }
            "rowheight"  { $ws.Rows($o.row).RowHeight = [double]$o.h }
            "goto"       { $ws.Activate() | Out-Null; $xl.Goto($ws.Range($ref), $true) | Out-Null }
            "probe"      { }
            "scan"       { }
            default      { throw "Okänd op: $($o.op)" }
        }
    }
    $xl.CalculateFullRebuild()
    $probeOut = @{}; $scanOut = @{}
    foreach ($o in $opsList) {
        $sheet = [string]$o.sheet; $ref = [string]$o.ref
        $ws = $wb.Worksheets.Item($sheet)
        if ($o.op -eq "probe") { $probeOut[$sheet + "!" + $ref] = $ws.Range($ref).Value2 }
        if ($o.op -eq "scan") {
            $n = 0
            foreach ($c in $ws.Range($ref).Cells) { $v = $c.Value2; if ($v -is [int] -and $v -lt -2146826000) { $n++ } }
            $scanOut[$sheet + "!" + $ref] = $n
        }
    }
    if ($Out) { $wb.SaveAs($Out, 51) }
    @{ probes = $probeOut; errors = $scanOut } | ConvertTo-Json -Depth 3 | Out-File -Encoding UTF8 $Probes
}
finally {
    $wb.Close($false); $xl.Quit()
}
