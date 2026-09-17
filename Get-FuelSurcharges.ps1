#Requires -Version 5.1
<#
.SYNOPSIS
    Collects published fuel surcharge (FSC) rates for ACE Courier and its BC competitors.

.DESCRIPTION
    Scrapes each carrier's public fuel surcharge page, normalises the result, and writes:
      data/latest.json   - full snapshot incl. notes and errors (local only, never committed)
      site/data.js       - the same full snapshot for the local dashboard (local only)
      site/latest.json   - public snapshot: live published rates only, no notes or errors
      data/history.jsonl - append-only log of published rates, one row per carrier/service/effective-date

    Carriers that do not publish a machine-readable rate are read from data/manual.json
    and flagged stale once they pass the age threshold, so a missing number is always
    visible rather than silently absent.

.PARAMETER DataDir
    Output directory. Defaults to .\data next to this script.

.PARAMETER StaleAfterDays
    Manual entries older than this many days are marked "stale". Default 14.

.PARAMETER NoHistory
    Skip appending to history.jsonl (useful when testing).

.PARAMETER PublicSnapshotPath
    Where to write the public snapshot, relative to this script. Default site/latest.json.
    The local publisher writes data/local-public.json instead, so it never touches the
    files the cloud run owns.

.PARAMETER FallbackSnapshot
    A public snapshot (relative to this script) to borrow rows from when a carrier fails
    to scrape here - used in the cloud, where some carrier sites block requests, with the snapshot
    this PC publishes. Borrowed rows are used only while the rate is still in effect.

.EXAMPLE
    .\Get-FuelSurcharges.ps1
    .\Get-FuelSurcharges.ps1 -Verbose
    .\Get-FuelSurcharges.ps1 -FallbackSnapshot data/local-public.json
#>
[CmdletBinding()]
param(
    [string] $DataDir,
    [int]    $StaleAfterDays = 14,
    [switch] $NoHistory,
    [string] $PublicSnapshotPath,
    [string] $FallbackSnapshot
)

$ErrorActionPreference = 'Stop'
if (-not $DataDir) { $DataDir = Join-Path $PSScriptRoot 'data' }
if (-not (Test-Path $DataDir)) { New-Item -ItemType Directory -Path $DataDir | Out-Null }

try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

$script:UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
$script:RxOpts    = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline
$script:Manual    = $null

#------------------------------------------------------------------------------
# Helpers
#------------------------------------------------------------------------------

function Get-PageHtml {
    param([Parameter(Mandatory = $true)][string] $Url)
    # Several carrier sites (Purolator, UPS) reject requests that carry only a
    # User-Agent, so send the rest of a normal browser's request headers too.
    $headers = @{
        'Accept'          = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'
        'Accept-Language' = 'en-CA,en;q=0.9'
        'Cache-Control'   = 'no-cache'
        'Pragma'          = 'no-cache'
        'Sec-Fetch-Dest'  = 'document'
        'Sec-Fetch-Mode'  = 'navigate'
        'Sec-Fetch-Site'  = 'none'
        'Sec-Fetch-User'  = '?1'
        'Upgrade-Insecure-Requests' = '1'
    }
    try {
        $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -UserAgent $script:UserAgent `
                                  -Headers $headers -TimeoutSec 45 -MaximumRedirection 5
        return [string]$resp.Content
    }
    catch {
        # Purolator (and others behind the same WAF) reject .NET's TLS/header
        # fingerprint but accept curl, which ships with Windows 10+ and every
        # GitHub Actions runner. Fall back rather than losing the carrier.
        Write-Verbose "  Invoke-WebRequest failed ($($_.Exception.Message)); retrying with curl.exe"
        # curl.exe on Windows, curl on the Linux CI runner.
        $curl = Get-Command curl.exe, curl -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $curl) { throw }

        $tmp = [IO.Path]::GetTempFileName()
        try {
            # Do not name this $args - that is an automatic variable and splatting it
            # would pass the function's own arguments instead.
            $curlArgs = @(
                '-sSL', '--compressed', '--max-time', '45',
                '-A', $script:UserAgent,
                '-H', 'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                '-H', 'Accept-Language: en-CA,en;q=0.9',
                '-o', $tmp,
                '-w', '%{http_code}',
                $Url
            )
            $code = & $curl.Source @curlArgs
            if ($LASTEXITCODE -ne 0) { throw "curl.exe exited $LASTEXITCODE for $Url" }
            if ("$code" -notmatch '^2\d\d$') { throw "curl.exe got HTTP $code for $Url" }
            return [IO.File]::ReadAllText($tmp)
        }
        finally {
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}

# Strip a page down to readable text: drop script/style/comments, drop tags,
# decode entities, collapse all whitespace (including &nbsp;) to single spaces.
function Get-PageText {
    param([Parameter(Mandatory = $true)][string] $Url)
    $html = Get-PageHtml -Url $Url
    $html = [regex]::Replace($html, '(?is)<(script|style|noscript)\b[^>]*>.*?</\1\s*>', ' ')
    $html = [regex]::Replace($html, '(?s)<!--.*?-->', ' ')
    $html = [regex]::Replace($html, '(?s)<[^>]+>', ' ')
    $html = [System.Net.WebUtility]::HtmlDecode($html)
    $html = [regex]::Replace($html, '\s+', ' ')
    return $html.Trim()
}

function ConvertTo-IsoDate {
    param([string] $Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t  = ([regex]::Replace($Text, '\s+', ' ')).Trim()
    $ci = [Globalization.CultureInfo]::InvariantCulture
    $d  = [datetime]::MinValue
    $formats = @('yyyy-MM-dd', 'MMMM d, yyyy', 'MMM d, yyyy', 'MMMM d yyyy', 'MMM d yyyy')
    foreach ($f in $formats) {
        if ([datetime]::TryParseExact($t, $f, $ci, [Globalization.DateTimeStyles]::None, [ref]$d)) {
            return $d.ToString('yyyy-MM-dd')
        }
    }
    if ([datetime]::TryParse($t, $ci, [Globalization.DateTimeStyles]::None, [ref]$d)) {
        return $d.ToString('yyyy-MM-dd')
    }
    return $null
}

function New-FscRecord {
    param(
        [Parameter(Mandatory = $true)][string] $Carrier,
        [Parameter(Mandatory = $true)][string] $Service,
        [double] $Percent,
        [string] $From,
        [string] $To,
        [string] $Status = 'ok',
        [string] $Source,
        [string] $Note,
        [string] $Segment = 'parcel'
    )
    [pscustomobject]@{
        carrier        = $Carrier
        service        = $Service
        segment        = $Segment
        percent        = [math]::Round($Percent, 2)
        effective_from = $From
        effective_to   = $To
        status         = $Status
        source         = $Source
        note           = $Note
    }
}

function Assert-Match {
    param($Match, [string] $Carrier, [string] $What)
    if (-not $Match.Success) { throw "$Carrier - could not locate $What (page layout may have changed)" }
}

#------------------------------------------------------------------------------
# Carrier adapters
# Each returns one or more FSC records, or throws. The runner catches failures and
# records them, so a broken scraper shows on the dashboard instead of vanishing.
#------------------------------------------------------------------------------

function Get-FscAce {
    $url = 'https://www.acecourier.ca/faq/'
    $t   = Get-PageText $url

    $pattern = 'Effective\s+(?<d>[A-Za-z]+\s+\d{1,2},?\s+\d{4}).{0,120}?Fuel Surcharge rates are\s*' +
               '(?<bc>\d+(?:\.\d+)?)\s*%?\s*for British Columbia\s*and\s*(?<ab>\d+(?:\.\d+)?)\s*%?\s*for Alberta'
    $m = [regex]::Match($t, $pattern, $script:RxOpts)
    Assert-Match $m 'ACE Courier' 'the BC/Alberta fuel surcharge sentence'

    $from = ConvertTo-IsoDate $m.Groups['d'].Value

    New-FscRecord -Carrier 'ACE Courier' -Service 'British Columbia' -Segment 'ace' -Percent ([double]$m.Groups['bc'].Value) -From $from -Source $url
    New-FscRecord -Carrier 'ACE Courier' -Service 'Alberta'          -Segment 'ace' -Percent ([double]$m.Groups['ab'].Value) -From $from -Source $url

    $f = [regex]::Match($t, 'FTL\s*/\s*Direct Drive FSC is\s*(?<v>\d+(?:\.\d+)?)\s*%', $script:RxOpts)
    if ($f.Success) {
        New-FscRecord -Carrier 'ACE Courier' -Service 'FTL / Direct Drive' -Segment 'ace' -Percent ([double]$f.Groups['v'].Value) -From $from -Source $url
    }
}

function Get-FscComoxPacific {
    $url = 'https://www.comoxpacific.com/'
    $t   = Get-PageText $url

    # "FUEL SURCHARGE (2026-09-03) Under 10,000lbs: 64.4% Over 10,000lbs: 74.4%"
    $pattern = 'FUEL SURCHARGE\s*\((?<d>\d{4}-\d{2}-\d{2})\)\s*Under\s*10,?000\s*lbs:\s*(?<u>\d+(?:\.\d+)?)\s*%' +
               '\s*Over\s*10,?000\s*lbs:\s*(?<o>\d+(?:\.\d+)?)\s*%'
    $m = [regex]::Match($t, $pattern, $script:RxOpts)
    Assert-Match $m 'Comox Pacific' 'the homepage fuel surcharge block'

    $from = $m.Groups['d'].Value
    New-FscRecord -Carrier 'Comox Pacific Express' -Service 'LTL under 10,000 lb' -Segment 'ltl' -Percent ([double]$m.Groups['u'].Value) -From $from -Source $url
    New-FscRecord -Carrier 'Comox Pacific Express' -Service 'LTL over 10,000 lb'  -Segment 'ltl' -Percent ([double]$m.Groups['o'].Value) -From $from -Source $url

    # Comox posts next week's rate alongside the current one.
    $nextPattern = 'As of\s*(?<d>\d{4}-\d{2}-\d{2})\s*Under\s*10,?000\s*lbs:\s*(?<u>\d+(?:\.\d+)?)\s*%' +
                   '\s*Over\s*10,?000\s*lbs:\s*(?<o>\d+(?:\.\d+)?)\s*%'
    $n = [regex]::Match($t, $nextPattern, $script:RxOpts)
    if ($n.Success) {
        $nf   = $n.Groups['d'].Value
        $note = "Announced in advance; takes effect $nf"
        New-FscRecord -Carrier 'Comox Pacific Express' -Service 'LTL under 10,000 lb' -Segment 'ltl' -Percent ([double]$n.Groups['u'].Value) -From $nf -Status 'upcoming' -Source $url -Note $note
        New-FscRecord -Carrier 'Comox Pacific Express' -Service 'LTL over 10,000 lb'  -Segment 'ltl' -Percent ([double]$n.Groups['o'].Value) -From $nf -Status 'upcoming' -Source $url -Note $note
    }
}

#------------------------------------------------------------------------------
# Manual entries - carriers with no public, machine-readable rate
#------------------------------------------------------------------------------

# Latest rate for a carrier from ACE's internal competitor comparisons
# (data/competitor-reports.json, local only - never committed or published).
function Get-ReportRate {
    param([string] $Carrier)
    $path = Join-Path $DataDir 'competitor-reports.json'
    if (-not (Test-Path $path)) { return $null }
    $doc = Get-Content $path -Raw | ConvertFrom-Json
    $best = $null
    foreach ($rep in @($doc.reports)) {
        foreach ($row in @($rep.rates)) {
            if ($row.carrier -ne $Carrier -or $null -eq $row.percent) { continue }
            if (-not $best -or [string]::CompareOrdinal([string]$rep.date, [string]$best.date) -gt 0) {
                $best = [pscustomobject]@{ date = [string]$rep.date; percent = [double]$row.percent }
            }
        }
    }
    return $best
}

function Get-FscManual {
    $path = Join-Path $DataDir 'manual.json'
    if (-not (Test-Path $path)) { return }

    $script:Manual = Get-Content $path -Raw | ConvertFrom-Json
    if (-not ($script:Manual.PSObject.Properties.Name -contains 'carriers')) { return }

    $today = Get-Date
    foreach ($e in $script:Manual.carriers) {
        $status = 'manual'
        $note   = $e.note
        $asOf   = ConvertTo-IsoDate $e.as_of
        $pct    = $null
        if ($null -ne $e.percent -and -not [string]::IsNullOrWhiteSpace([string]$e.percent)) { $pct = [double]$e.percent }

        # A comparison report newer than the manual entry wins.
        $fromReport = $false
        $report = Get-ReportRate ([string]$e.carrier)
        if ($report -and ($null -eq $pct -or -not $asOf -or [string]::CompareOrdinal($report.date, $asOf) -gt 0)) {
            $pct  = $report.percent
            $asOf = $report.date
            $note = "From ACE's competitor comparison of $(Get-Date ([datetime]::ParseExact($asOf, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)) -Format 'MMM d, yyyy'); compared with ACE's rate on that date."
            $fromReport = $true
        }

        if ($null -eq $pct) {
            $status = 'unavailable'
            if (-not $note) { $note = 'No rate on file. Add one in data/manual.json when you learn it.' }
        }
        elseif ($asOf) {
            $age = ($today - [datetime]::ParseExact($asOf, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)).Days
            if ($age -gt $StaleAfterDays) {
                $status = 'stale'
                $note   = "$age days old. $note".Trim()
            }
        }

        $seg = 'ltl'
        if ($e.PSObject.Properties.Name -contains 'segment' -and $e.segment) { $seg = [string]$e.segment }

        $pctValue = 0.0
        if ($null -ne $pct) { $pctValue = $pct }

        $rec = New-FscRecord -Carrier ([string]$e.carrier) -Service ([string]$e.service) -Segment $seg `
                             -Percent $pctValue -From $asOf -Status $status -Source ([string]$e.source) -Note $note
        # Dated rates are compared with ACE's rate on the same date, not today's.
        if ($fromReport) { $rec | Add-Member -NotePropertyName 'benchmark_date' -NotePropertyValue $asOf }
        $rec
    }
}

# ACE's BC surcharge on a given date, from the committed schedule.
function Get-AceBcOn {
    param([string] $Date)
    $path = Join-Path $DataDir 'ace-fsc-history.json'
    if (-not (Test-Path $path)) { return $null }
    foreach ($c in @((Get-Content $path -Raw | ConvertFrom-Json).changes)) {
        if ([string]::CompareOrdinal($Date, [string]$c.effective) -ge 0 -and [string]::CompareOrdinal($Date, [string]$c.confirmed_through) -le 0) {
            return [double]$c.bc
        }
    }
    return $null
}

#------------------------------------------------------------------------------
# Runner
#------------------------------------------------------------------------------

$adapters = [ordered]@{
    # Direct competitors only: carriers moving heavy LTL freight in BC/AB. Parcel and
    # courier networks (FedEx, Purolator, Canada Post, Canpar, UPS) were dropped - they
    # don't compete for ACE's freight.
    'ACE Courier'           = 'Get-FscAce'
    'Comox Pacific Express' = 'Get-FscComoxPacific'
}

# Manual entries first; scraped adapters follow.
$records = New-Object System.Collections.Generic.List[object]
$errors  = New-Object System.Collections.Generic.List[object]

try {
    foreach ($r in @(Get-FscManual)) { $records.Add($r) }
}
catch {
    $errors.Add([pscustomobject]@{ carrier = 'manual.json'; error = $_.Exception.Message })
    Write-Warning "manual.json - $($_.Exception.Message)"
}

foreach ($name in $adapters.Keys) {
    $fn = $adapters[$name]
    Write-Verbose "Fetching $name ..."
    try {
        foreach ($r in @(& $fn)) { $records.Add($r) }
        Write-Verbose "  ok"
    }
    catch {
        $msg = $_.Exception.Message
        $errors.Add([pscustomobject]@{ carrier = $name; error = $msg })
        $records.Add((New-FscRecord -Carrier $name -Service 'unknown' -Percent 0 -Status 'error' -Note $msg))
        Write-Warning "$name - $msg"
    }
}

$now = (Get-Date).ToUniversalTime()

# Borrow rows for carriers that failed here from another machine's public snapshot,
# but only while the borrowed rate is still in effect. A row with an end date is valid
# through that date; a row without one is valid if the snapshot is under 3 days old.
if ($FallbackSnapshot) {
    $fbPath = Join-Path $PSScriptRoot $FallbackSnapshot
    if (Test-Path $fbPath) {
        $fb = Get-Content $fbPath -Raw | ConvertFrom-Json
        $fbAge = $null
        try { $fbAge = ($now - ([datetime]$fb.generated_at).ToUniversalTime()).TotalDays } catch { }
        $todayPacific = $now.AddHours(-8).Date
        $failed = @($errors | ForEach-Object { $_.carrier })

        foreach ($carrier in $failed) {
            $borrowed = @()
            foreach ($row in @($fb.rates | Where-Object { $_.carrier -eq $carrier })) {
                $valid = $false
                $end = ConvertTo-IsoDate ([string]$row.effective_to)
                if ($end) {
                    $valid = [datetime]::ParseExact($end, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) -ge $todayPacific
                }
                elseif ($null -ne $fbAge) {
                    $valid = $fbAge -le 3
                }
                if ($valid) {
                    $borrowed += New-FscRecord -Carrier $row.carrier -Service $row.service -Segment $row.segment `
                                     -Percent ([double]$row.percent) -From $row.effective_from -To $row.effective_to `
                                     -Status $row.status -Source $row.source `
                                     -Note ("Supplied by the local scrape of " + $fb.generated_date + "; unreachable from here.")
                }
            }

            if ($borrowed.Count) {
                # Replace this carrier's error placeholder with the borrowed rows.
                $placeholder = @($records | Where-Object { $_.carrier -eq $carrier -and $_.status -eq 'error' })
                foreach ($p in $placeholder) { [void]$records.Remove($p) }
                foreach ($b in $borrowed) { $records.Add($b) }
                Write-Host "$carrier - using $($borrowed.Count) row(s) from $FallbackSnapshot ($($fb.generated_date))"
            }
            else {
                Write-Warning "$carrier - no still-valid rows in $FallbackSnapshot"
            }
        }
    }
    else {
        Write-Verbose "Fallback snapshot $fbPath not found"
    }
}

# Compare like with like. A competitor's truckload surcharge belongs against ACE's
# FTL/Direct Drive rate, not against its BC LTL/parcel rate - benchmarking a 99.3%
# TL number against 50.5% would invent an advantage that isn't real.
function Get-AcePercent {
    param([string] $Service)
    $row = $records | Where-Object { $_.carrier -eq 'ACE Courier' -and $_.service -eq $Service } | Select-Object -First 1
    if ($row) { return [double]$row.percent }
    return $null
}

$aceBcPct  = Get-AcePercent 'British Columbia'
$aceFtlPct = Get-AcePercent 'FTL / Direct Drive'

foreach ($r in $records) {
    $delta = $null
    $rel   = $null
    $benchName = $null
    $benchPct  = $null

    $comparable = ($r.carrier -ne 'ACE Courier') -and ($r.percent -gt 0) -and
                  ($r.status -ne 'error') -and ($r.status -ne 'unavailable')

    $benchDate = $null
    if ($r.PSObject.Properties.Name -contains 'benchmark_date') { $benchDate = $r.benchmark_date }

    if ($comparable -and $benchDate) {
        $benchPct  = Get-AceBcOn $benchDate
        $benchName = "ACE British Columbia on $benchDate"
    }
    elseif ($comparable) {
        # Truckload / full-load services measure against ACE's FTL rate.
        if ($r.service -match '(^|\s)TL$' -or $r.service -match 'truckload|full[- ]load|FTL') {
            $benchPct  = $aceFtlPct
            $benchName = 'ACE FTL / Direct Drive'
        }
        else {
            $benchPct  = $aceBcPct
            $benchName = 'ACE British Columbia'
        }
    }

    if ($null -ne $benchPct -and $benchPct -gt 0) {
        $delta = [math]::Round($r.percent - $benchPct, 2)
        $rel   = [math]::Round((($r.percent - $benchPct) / $benchPct) * 100, 1)
    }
    else {
        $benchName = $null
    }

    $r | Add-Member -NotePropertyName 'delta_points'  -NotePropertyValue $delta     -Force
    $r | Add-Member -NotePropertyName 'delta_percent' -NotePropertyValue $rel       -Force
    $r | Add-Member -NotePropertyName 'benchmark'     -NotePropertyValue $benchName -Force
}

# NOTE: PowerShell 5.1 throws "Argument types do not match" when a generic List is
# wrapped in @() inside a [pscustomobject] cast. Materialise real arrays first.
[object[]] $rateArray  = $records.ToArray()
[object[]] $errorArray = $errors.ToArray()
$okCount = @($rateArray | Where-Object { $_.status -eq 'ok' }).Count

$snapshot = [pscustomobject]@{
    generated_at   = $now.ToString('o')
    generated_date = $now.ToString('yyyy-MM-dd')
    ace_benchmark  = $aceBcPct
    ok_count       = $okCount
    error_count    = $errorArray.Count
    errors         = $errorArray
    rates          = $rateArray
}

$json = $snapshot | ConvertTo-Json -Depth 6

# Public snapshot for the hosted site. Allowlist, not blocklist: only live rates the
# carriers publish themselves, only the fields the page renders, and no maintenance
# notes, error messages or manually entered (possibly rep-sourced) numbers.
$publicStatuses = @('ok', 'upcoming')
$publicRates = foreach ($r in $rateArray) {
    if ($publicStatuses -notcontains $r.status) { continue }
    $publicNote = $null
    if ($r.status -eq 'upcoming') { $publicNote = "Announced in advance; takes effect $($r.effective_from)" }
    [pscustomobject]@{
        carrier        = $r.carrier
        service        = $r.service
        segment        = $r.segment
        percent        = $r.percent
        effective_from = $r.effective_from
        effective_to   = $r.effective_to
        status         = $r.status
        source         = $r.source
        note           = $publicNote
        delta_points   = $r.delta_points
        delta_percent  = $r.delta_percent
        benchmark      = $r.benchmark
    }
}
[object[]] $publicRateArray = @($publicRates)
$publicSnapshot = [pscustomobject]@{
    generated_at   = $now.ToString('o')
    generated_date = $now.ToString('yyyy-MM-dd')
    ace_benchmark  = $aceBcPct
    ok_count       = $okCount
    rates          = $publicRateArray
}
$publicJson = $publicSnapshot | ConvertTo-Json -Depth 6

# Write without a BOM - PowerShell 5.1's -Encoding UTF8 adds one, which breaks JSON.parse.
$utf8NoBom  = New-Object System.Text.UTF8Encoding($false)
$latestPath = Join-Path $DataDir 'latest.json'
[IO.File]::WriteAllText($latestPath, $json, $utf8NoBom)
Write-Verbose "Wrote $latestPath"

# Build paths segment-by-segment: a literal 'site\data.js' would become a filename
# containing a backslash when this runs on Linux CI.
$siteDir = Join-Path (Split-Path $DataDir -Parent) 'site'
if (Test-Path $siteDir) {
    # Full data for the local dashboard. Loaded via <script src> so the page works when
    # opened straight off disk (file:// blocks fetch). Gitignored - never published.
    $dataJsPath = Join-Path $siteDir 'data.js'
    [IO.File]::WriteAllText($dataJsPath, "window.FSC_DATA = $json;`r`n", $utf8NoBom)
    Write-Verbose "Wrote $dataJsPath"
}

if ($PublicSnapshotPath) {
    $publicPath = Join-Path $PSScriptRoot $PublicSnapshotPath
}
else {
    $publicPath = Join-Path $siteDir 'latest.json'
}
$publicDir = Split-Path $publicPath -Parent
if (Test-Path $publicDir) {
    [IO.File]::WriteAllText($publicPath, $publicJson, $utf8NoBom)
    Write-Verbose "Wrote $publicPath"
}

if (-not $NoHistory) {
    $histPath = Join-Path $DataDir 'history.jsonl'
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    if (Test-Path $histPath) {
        foreach ($line in (Get-Content $histPath)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try {
                $o = $line | ConvertFrom-Json
                [void]$seen.Add(('{0}|{1}|{2}|{3}' -f $o.carrier, $o.service, $o.effective_from, $o.percent))
            }
            catch { }
        }
    }

    $new = 0
    foreach ($r in $records) {
        # History is committed, so it only records rates carriers publish themselves -
        # never manual, stale or derived numbers.
        if ($publicStatuses -notcontains $r.status) { continue }
        $key = '{0}|{1}|{2}|{3}' -f $r.carrier, $r.service, $r.effective_from, $r.percent
        if ($seen.Add($key)) {
            $row = [pscustomobject]@{
                carrier        = $r.carrier
                service        = $r.service
                segment        = $r.segment
                percent        = $r.percent
                effective_from = $r.effective_from
                effective_to   = $r.effective_to
                status         = $r.status
                recorded_at    = $now.ToString('o')
            }
            Add-Content -Path $histPath -Value ($row | ConvertTo-Json -Depth 4 -Compress) -Encoding UTF8
            $new++
        }
    }
    Write-Verbose "history.jsonl: $new new row(s)"
}

# Console summary
$records |
    Sort-Object @{ Expression = { if ($_.carrier -eq 'ACE Courier') { 0 } else { 1 } } }, carrier, service |
    Format-Table @{ L = 'Carrier'; E = { $_.carrier }; W = 24 },
                 @{ L = 'Service'; E = { $_.service }; W = 24 },
                 @{ L = 'FSC %';   E = { if ($_.percent -gt 0) { '{0:N2}' -f $_.percent } else { '-' } }; A = 'Right'; W = 8 },
                 @{ L = 'vs ACE';  E = { if ($null -ne $_.delta_points) { '{0:+0.0;-0.0;0.0}' -f $_.delta_points } else { '' } }; A = 'Right'; W = 8 },
                 @{ L = 'From';    E = { $_.effective_from }; W = 11 },
                 @{ L = 'Status';  E = { $_.status }; W = 12 } -AutoSize

if ($errors.Count -gt 0) {
    Write-Host ''
    Write-Warning "$($errors.Count) carrier(s) failed to scrape:"
    foreach ($e in $errors) { Write-Host ("  - {0}: {1}" -f $e.carrier, $e.error) -ForegroundColor Yellow }
}

Write-Host ''
Write-Host ("Snapshot written to {0} ({1} rows, ACE BC benchmark {2}%)" -f $latestPath, $records.Count, $aceBcPct) -ForegroundColor Green

# One line per run, so an unattended scheduled run leaves evidence it happened.
$logLine = '{0}  rows={1}  live={2}  errors={3}  aceBC={4}{5}' -f `
           $now.ToString('yyyy-MM-dd HH:mm:ss'), $records.Count, $okCount, $errorArray.Count, $aceBcPct,
           $(if ($errorArray.Count) { '  [' + (($errorArray | ForEach-Object { $_.carrier }) -join ', ') + ']' } else { '' })
Add-Content -Path (Join-Path $DataDir 'run.log') -Value $logLine -Encoding UTF8
