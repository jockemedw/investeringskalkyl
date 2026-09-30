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
            "clear_fmt"  { $ws.Range($ref).ClearFormats() | Out-Null }
            "replace"    { $ws.Range($ref).Replace([string]$o.find, [string]$o.repl, 2) | Out-Null }
            "dv_list"    { $r = $ws.Range($ref); $r.Validation.Delete(); $r.Validation.Add(3, 1, 1, [string]$o.source) | Out-Null; $r.Validation.InCellDropdown = $true }
            "print_area" { $ws.PageSetup.PrintArea = [string]$o.area }
            "align"      { $ws.Range($ref).HorizontalAlignment = [int]$o.h }
            "wrap"       { $ws.Range($ref).WrapText = $true }
            "bold"       { $ws.Range($ref).Font.Bold = $true }
            "italic"     { $ws.Range($ref).Font.Italic = $true }
            "colwidth_from" { $ws.Columns([string]$o.dst).ColumnWidth = $ws.Columns([string]$o.src).ColumnWidth }
            "numfmt_local" { $ws.Range($ref).NumberFormatLocal = [string]$o.fmt }
            "insert_rows" { $ws.Rows($ref).Insert() | Out-Null }
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
