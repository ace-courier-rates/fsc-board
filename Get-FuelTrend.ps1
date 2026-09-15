#Requires -Version 5.1
<#
.SYNOPSIS
    Builds the 12-month trend of BC diesel prices against ACE Courier's fuel surcharge.

.DESCRIPTION
    Inputs:
      Statistics Canada table 18-10-0001-01 - monthly average retail price of diesel at
        self-service stations, Vancouver and Victoria (cents per litre).
      data/ace-fsc-history.json - ACE's published surcharge changes, each confirmed from
        its effective date through confirmed_through.
      data/latest.json - today's scrape; a new ACE rate is added to the history, and the
        current rate's confirmed_through is extended to today.

    Output: site/trend.json (public) and site/trend.js (the same data for the dashboard
    opened off disk).

    The correlation compares ACE's day-weighted average BC surcharge for each month with
    that month's Vancouver diesel price, over months where ACE's rate is on record for at
    least half the days. Months where it isn't are left empty rather than filled in.

.PARAMETER Months
    Diesel months to show. Default 12.

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
    $ab = $latest.rates | Where-Object { $_.carrier -eq 'ACE Courier' -and $_.service -eq 'Alberta' -and $_.status -eq 'ok' } | Select-Object -First 1

    if ($bc -and $bc.effective_from) {
        $current = $changes | Where-Object { $_.effective -eq $bc.effective_from } | Select-Object -First 1
        if ($current) {
            if ([string]::CompareOrdinal([string]$current.confirmed_through, $today) -lt 0) {
                $current.confirmed_through = $today
                $historyChanged = $true
            }
        }
        else {
            $abPct = $null
            if ($ab) { $abPct = $ab.percent }
            $changes.Add([pscustomobject]@{
                effective         = $bc.effective_from
                bc                = $bc.percent
                ab                = $abPct
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
# Diesel prices (Statistics Canada)
#------------------------------------------------------------------------------

$cities = [ordered]@{
    vancouver = '16.6.0.0.0.0.0.0.0.0'   # Vancouver, diesel at self-service stations
    victoria  = '17.6.0.0.0.0.0.0.0.0'   # Victoria,  diesel at self-service stations
}

$diesel = @{}   # city -> @{ 'yyyy-MM' = cents }
try {
    $requests = @(foreach ($k in $cities.Keys) { @{ productId = 18100001; coordinate = $cities[$k]; latestN = $Months + 12 } })
    $body     = ConvertTo-Json -InputObject $requests -Compress
    $resp     = Invoke-RestMethod -Method Post -Uri 'https://www150.statcan.gc.ca/t1/wds/rest/getDataFromCubePidCoordAndLatestNPeriods' `
                                  -ContentType 'application/json' -Body $body -TimeoutSec 60
    $i = 0
    foreach ($k in $cities.Keys) {
        $series = @{}
        foreach ($p in $resp[$i].object.vectorDataPoint) {
            if ($null -ne $p.value) { $series[$p.refPer.Substring(0, 7)] = [double]$p.value }
        }
        if ($series.Count -eq 0) { throw "Statistics Canada returned no data for $k" }
        $diesel[$k] = $series
        $i++
    }
}
catch {
    # Keep the last published prices rather than blanking the chart.
    Write-Warning "Diesel prices unavailable ($($_.Exception.Message)); reusing previous trend data."
    if (-not (Test-Path $trendPath)) { throw }
    $prev = Get-Content $trendPath -Raw | ConvertFrom-Json
    foreach ($k in $cities.Keys) { $diesel[$k] = @{} }
    foreach ($m in $prev.diesel_all) {
        foreach ($k in $cities.Keys) { if ($null -ne $m.$k) { $diesel[$k][$m.month] = [double]$m.$k } }
    }
}

$allMonths = @($diesel['vancouver'].Keys | Sort-Object)
$shown     = @($allMonths | Select-Object -Last $Months)

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

# ACE's day-weighted average BC surcharge for a month - only when its rate is on record
# for at least half the month's days. Diesel is monthly and ACE changes its rate almost
# weekly, so a monthly average is the like-for-like comparison.
function Get-AceMonthAverage {
    param([string] $Month)
    $start = [datetime]::ParseExact("$Month-01", 'yyyy-MM-dd', $ci)
    $end   = $start.AddMonths(1).AddDays(-1)
    $days  = ($end - $start).Days + 1
    $sum = 0.0; $covered = 0
    for ($d = $start; $d -le $end; $d = $d.AddDays(1)) {
        $v = Get-AceOn $d
        if ($null -ne $v) { $sum += $v; $covered++ }
    }
    if ($covered -ge $days / 2) {
        return [pscustomobject]@{ average = [math]::Round($sum / $covered, 1); covered = $covered; days = $days }
    }
    return $null
}

$aceByMonth = @{}
foreach ($m in $allMonths) {
    $a = Get-AceMonthAverage $m
    if ($a) { $aceByMonth[$m] = $a }
}

$monthRows = foreach ($m in $shown) {
    $vic = $null
    if ($diesel['victoria'].ContainsKey($m)) { $vic = $diesel['victoria'][$m] }
    $avg = $null; $cov = $null
    if ($aceByMonth.ContainsKey($m)) { $avg = $aceByMonth[$m].average; $cov = "$($aceByMonth[$m].covered)/$($aceByMonth[$m].days)" }
    [pscustomobject]@{
        month        = $m
        vancouver    = $diesel['vancouver'][$m]
        victoria     = $vic
        ace_bc       = $avg
        ace_coverage = $cov
    }
}

$allRows = foreach ($m in $allMonths) {
    $vic = $null
    if ($diesel['victoria'].ContainsKey($m)) { $vic = $diesel['victoria'][$m] }
    [pscustomobject]@{ month = $m; vancouver = $diesel['vancouver'][$m]; victoria = $vic }
}

# Correlate over every month with both a diesel price and ACE's rate on record.
$pairs = foreach ($m in $allMonths) {
    if ($aceByMonth.ContainsKey($m)) {
        [pscustomobject]@{ month = $m; bc = $aceByMonth[$m].average; diesel = $diesel['vancouver'][$m] }
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

$windowFrom = "$($shown[0])-01"
$lastMonthEnd = ([datetime]::ParseExact("$($shown[-1])-01", 'yyyy-MM-dd', $ci)).AddMonths(1).AddDays(-1).ToString('yyyy-MM-dd')
$windowTo = $today
if ([string]::CompareOrdinal($lastMonthEnd, $today) -gt 0) { $windowTo = $lastMonthEnd }

[object[]] $monthArray   = @($monthRows)
[object[]] $allRowArray  = @($allRows)
[object[]] $segmentArray = @($sorted | ForEach-Object {
    [pscustomobject]@{ from = $_.effective; to = $_.confirmed_through; bc = [double]$_.bc }
})

$trend = [pscustomobject]@{
    generated_at = $now.ToString('o')
    window       = [pscustomobject]@{ from = $windowFrom; to = $windowTo }
    months       = $monthArray
    ace_segments = $segmentArray
    correlation  = [pscustomobject]@{
        r     = $r
        n     = $pairArray.Count
        from  = $(if ($pairArray.Count) { $pairArray[0].month } else { $null })
        basis = "ACE's average BC surcharge vs the same month's Vancouver diesel price, over months with ACE's rate on record for at least half the month"
    }
    diesel_all   = $allRowArray
    sources      = [pscustomobject]@{
        diesel = 'Statistics Canada, table 18-10-0001-01'
        ace    = 'ACE Courier fuel surcharge notices and FAQ page, with earlier rates from Internet Archive copies'
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

Write-Host ("Trend: {0} months ({1} to {2}), {3} ACE changes, r={4} over {5} months" -f $monthArray.Count, $shown[0], $shown[-1], $sorted.Count, $r, $pairArray.Count)
