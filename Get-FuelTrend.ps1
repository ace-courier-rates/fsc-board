#Requires -Version 5.1
<#
.SYNOPSIS
    Builds the 12-month trend of BC diesel prices against ACE Courier's fuel surcharge.

.DESCRIPTION
    Inputs:
      Natural Resources Canada weekly average retail diesel prices (cents per litre,
        taxes included) for every BC city NRCan surveys: Abbotsford, Fort St. John,
        Kamloops, Kelowna, Prince George, Vancouver and Victoria. The BC figure is the
        average of the cities reporting that week.
      data/ace-fsc-history.json - ACE's published surcharge changes, each confirmed from
        its effective date through confirmed_through.
      data/latest.json - today's scrape; a new ACE rate is added to the history, and the
        current rate's confirmed_through is extended to today.

    Output: site/trend.json (public) and site/trend.js (the same data for the dashboard
    opened off disk).

    Both series are weekly, which matches how often ACE changes its surcharge. The
    correlation compares ACE's day-weighted average BC surcharge for each week with that
    week's BC diesel price, over weeks where ACE's rate is on record for at least half the
    days. Weeks where it isn't are left empty rather than filled in.

.PARAMETER Months
    Months of history to show. Default 12.

.PARAMETER LocalOnly
    Write only site/trend.js. Leaves the committed history and trend.json untouched, so
    a local run can't conflict with the cloud run that owns them.
#>
[CmdletBinding()]
param(
    [int]    $Months = 12,
    [switch] $LocalOnly
)

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

$dataDir    = Join-Path $PSScriptRoot 'data'
$siteDir    = Join-Path $PSScriptRoot 'site'
$histPath   = Join-Path $dataDir 'ace-fsc-history.json'
$latestPath = Join-Path $dataDir 'latest.json'
$trendPath  = Join-Path $siteDir 'trend.json'
$trendJs    = Join-Path $siteDir 'trend.js'
$utf8NoBom  = New-Object System.Text.UTF8Encoding($false)
$ci         = [Globalization.CultureInfo]::InvariantCulture

$now   = (Get-Date).ToUniversalTime()
$today = $now.AddHours(-8).ToString('yyyy-MM-dd')   # Pacific date

function ConvertTo-Day { param([string] $Iso) [datetime]::ParseExact($Iso, 'yyyy-MM-dd', $ci) }

#------------------------------------------------------------------------------
# ACE surcharge history
#------------------------------------------------------------------------------

$hist    = Get-Content $histPath -Raw | ConvertFrom-Json
$changes = New-Object System.Collections.Generic.List[object]
foreach ($c in $hist.changes) { $changes.Add($c) }

$historyChanged = $false
if (Test-Path $latestPath) {
    $latest = Get-Content $latestPath -Raw | ConvertFrom-Json
    $bc = $latest.rates | Where-Object { $_.carrier -eq 'ACE Courier' -and $_.service -eq 'British Columbia' -and $_.status -eq 'ok' } | Select-Object -First 1

    if ($bc -and $bc.effective_from) {
        $current = $changes | Where-Object { $_.effective -eq $bc.effective_from } | Select-Object -First 1
        if ($current) {
            if ([string]::CompareOrdinal([string]$current.confirmed_through, $today) -lt 0) {
                $current.confirmed_through = $today
                $historyChanged = $true
            }
        }
        else {
            $changes.Add([pscustomobject]@{
                effective         = $bc.effective_from
                bc                = $bc.percent
                ab                = $null
                confirmed_through = $today
                source            = 'https://www.acecourier.ca/faq/'
            })
            $historyChanged = $true
            Write-Host "ACE surcharge change recorded: BC $($bc.percent)% effective $($bc.effective_from)"
        }
    }
}

[object[]] $sorted = @($changes | Sort-Object { $_.effective })

#------------------------------------------------------------------------------
# Diesel prices - NRCan weekly averages for BC cities
#------------------------------------------------------------------------------

# NRCan location IDs for every BC city in the survey.
$script:BcCities = '2,3,4,5,6,70,90'   # Vancouver, Victoria, Prince George, Kamloops, Kelowna, Fort St. John, Abbotsford
$script:BcCityNames = 'Abbotsford, Fort St. John, Kamloops, Kelowna, Prince George, Vancouver and Victoria'
$script:UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'

# One year of weekly rows: week-ending date -> average price across the reporting cities.
function Get-NrcanWeek {
    param([int] $Year)
    $url = 'https://www2.nrcan.gc.ca/eneene/sources/pripri/prices_bycity_e.cfm' +
           "?ProductID=5&locationID=$($script:BcCities)&frequency=W&priceYear=$Year"
    $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -UserAgent $script:UserAgent -TimeoutSec 60
    $html = [string]$resp.Content

    $out = @{}
    foreach ($row in [regex]::Matches($html, '(?is)<tr[^>]*>(.*?)</tr>')) {
        $cells = $row.Groups[1].Value
        $d = [regex]::Match($cells, '(?<d>\d{4}-\d{2}-\d{2})')
        if (-not $d.Success) { continue }

        # Each city contributes four columns; the price column is headers="header4_<n>_1".
        $prices = foreach ($c in [regex]::Matches($cells, '(?is)<td[^>]*headers="header4_\d+_1[^"]*"[^>]*>(.*?)</td>')) {
            $v = ([regex]::Replace($c.Groups[1].Value, '<[^>]+>', '')).Trim()
            if ($v -match '^\d+(\.\d+)?$') { [double]$v }
        }
        [object[]] $vals = @($prices)
        if ($vals.Count -eq 0) { continue }   # future or blank week
        $out[$d.Groups['d'].Value] = [pscustomobject]@{
            price  = [math]::Round(($vals | Measure-Object -Average).Average, 1)
            cities = $vals.Count
        }
    }
    if ($out.Count -eq 0) { throw "NRCan returned no weekly rows for $Year" }
    return $out
}

$diesel = @{}
try {
    $years = @($now.Year, $now.AddMonths(-$Months).Year) | Sort-Object -Unique
    foreach ($y in $years) {
        foreach ($kv in (Get-NrcanWeek -Year $y).GetEnumerator()) { $diesel[$kv.Key] = $kv.Value }
    }
}
catch {
    # Keep the last published prices rather than blanking the chart.
    Write-Warning "NRCan prices unavailable ($($_.Exception.Message)); reusing previous trend data."
    if (-not (Test-Path $trendPath)) { throw }
    foreach ($p in (Get-Content $trendPath -Raw | ConvertFrom-Json).diesel_all) {
        $diesel[$p.date] = [pscustomobject]@{ price = [double]$p.diesel; cities = $p.cities }
    }
}

$cutoff   = $now.AddMonths(-$Months).ToString('yyyy-MM-dd')
$allWeeks = @($diesel.Keys | Sort-Object)
$shown    = @($allWeeks | Where-Object { [string]::CompareOrdinal($_, $cutoff) -ge 0 -and [string]::CompareOrdinal($_, $today) -le 0 })

#------------------------------------------------------------------------------
# Combine
#------------------------------------------------------------------------------

# ACE's rate on a given day, only if that day falls inside a confirmed span.
function Get-AceOn {
    param([datetime] $Day)
    foreach ($c in $sorted) {
        if ($Day -ge (ConvertTo-Day $c.effective) -and $Day -le (ConvertTo-Day $c.confirmed_through)) { return [double]$c.bc }
    }
    return $null
}

# ACE's day-weighted average for the week ending on $WeekEnding (7 days), when its rate is
# on record for at least half of them.
function Get-AceWeekAverage {
    param([string] $WeekEnding)
    $end   = ConvertTo-Day $WeekEnding
    $start = $end.AddDays(-6)
    $sum = 0.0; $covered = 0
    for ($d = $start; $d -le $end; $d = $d.AddDays(1)) {
        $v = Get-AceOn $d
        if ($null -ne $v) { $sum += $v; $covered++ }
    }
    if ($covered -ge 4) { return [pscustomobject]@{ average = [math]::Round($sum / $covered, 1); covered = $covered } }
    return $null
}

$aceByWeek = @{}
foreach ($w in $allWeeks) {
    $a = Get-AceWeekAverage $w
    if ($a) { $aceByWeek[$w] = $a }
}

$points = foreach ($w in $shown) {
    $avg = $null; $cov = $null
    if ($aceByWeek.ContainsKey($w)) { $avg = $aceByWeek[$w].average; $cov = "$($aceByWeek[$w].covered)/7" }
    [pscustomobject]@{
        date         = $w
        diesel       = $diesel[$w].price
        cities       = $diesel[$w].cities
        ace_bc       = $avg
        ace_coverage = $cov
    }
}

$allRows = foreach ($w in $allWeeks) {
    [pscustomobject]@{ date = $w; diesel = $diesel[$w].price; cities = $diesel[$w].cities }
}

# Correlate over every week with both a diesel price and ACE's rate on record.
$pairs = foreach ($w in $shown) {
    if ($aceByWeek.ContainsKey($w)) {
        [pscustomobject]@{ date = $w; bc = $aceByWeek[$w].average; diesel = $diesel[$w].price }
    }
}
[object[]] $pairArray = @($pairs)

$r = $null
if ($pairArray.Count -ge 3) {
    $xs = $pairArray | ForEach-Object { $_.diesel }
    $ys = $pairArray | ForEach-Object { $_.bc }
    $mx = ($xs | Measure-Object -Average).Average
    $my = ($ys | Measure-Object -Average).Average
    $sxy = 0.0; $sxx = 0.0; $syy = 0.0
    for ($k = 0; $k -lt $pairArray.Count; $k++) {
        $dx = $xs[$k] - $mx; $dy = $ys[$k] - $my
        $sxy += $dx * $dy; $sxx += $dx * $dx; $syy += $dy * $dy
    }
    if ($sxx -gt 0 -and $syy -gt 0) { $r = [math]::Round($sxy / [math]::Sqrt($sxx * $syy), 2) }
}

[object[]] $pointArray   = @($points)
[object[]] $allRowArray  = @($allRows)
[object[]] $segmentArray = @($sorted | ForEach-Object {
    [pscustomobject]@{ from = $_.effective; to = $_.confirmed_through; bc = [double]$_.bc }
})

$trend = [pscustomobject]@{
    generated_at = $now.ToString('o')
    interval     = 'weekly'
    window       = [pscustomobject]@{ from = $pointArray[0].date; to = $today }
    points       = $pointArray
    ace_segments = $segmentArray
    correlation  = [pscustomobject]@{
        r     = $r
        n     = $pairArray.Count
        from  = $(if ($pairArray.Count) { $pairArray[0].date } else { $null })
        basis = "ACE's average BC surcharge vs the same week's BC diesel price, over weeks with ACE's rate on record for at least four days"
    }
    diesel_all   = $allRowArray
    sources      = [pscustomobject]@{
        diesel = "Natural Resources Canada, weekly average retail diesel prices for $($script:BcCityNames)"
        ace    = 'ACE Courier fuel surcharge schedule and FAQ page'
    }
}

$json = $trend | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText($trendJs, "window.FSC_TREND = $json;`r`n", $utf8NoBom)

if (-not $LocalOnly) {
    [IO.File]::WriteAllText($trendPath, $json, $utf8NoBom)
    if ($historyChanged) {
        $hist.changes = $sorted
        [IO.File]::WriteAllText($histPath, ($hist | ConvertTo-Json -Depth 4), $utf8NoBom)
    }
}

Write-Host ("Trend: {0} weeks ({1} to {2}), {3} ACE changes, r={4} over {5} weeks" -f `
            $pointArray.Count, $pointArray[0].date, $pointArray[-1].date, $sorted.Count, $r, $pairArray.Count)
