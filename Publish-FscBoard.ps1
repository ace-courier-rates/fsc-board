#Requires -Version 5.1
<#
.SYNOPSIS
    Daily local run: refresh the board on this PC and publish this PC's public snapshot
    so the hosted site can use carriers the cloud can't reach.

.DESCRIPTION
    1. Pulls the latest repo so code changes and cloud data flow down.
    2. Runs Get-FuelSurcharges.ps1, writing the full local data as usual, but the public
       snapshot to data/local-public.json instead of site/latest.json. The cloud run owns
       site/latest.json and data/history.jsonl; this script never modifies them, so the
       two can't conflict.
    3. Commits and pushes data/local-public.json only - an allowlisted file holding live
       published rates, with no notes, errors or manual entries. The push triggers the
       cloud workflow, which republishes the site.

    Any git or network failure is logged to data/run.log; the local scrape still runs.

.EXAMPLE
    .\Publish-FscBoard.ps1
    .\Publish-FscBoard.ps1 -NoPush     # refresh locally without publishing
#>
[CmdletBinding()]
param(
    [switch] $NoPush,
    [string] $MsgFolder = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'FSC Inbox')
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$log = Join-Path $PSScriptRoot 'data\run.log'
function Write-RunLog {
    param([string] $Message)
    $line = '{0}  publish: {1}' -f (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -Path $log -Value $line -Encoding UTF8
    Write-Host $line
}

# git writes progress to stderr; in Windows PowerShell 5.1 that surfaces as errors, so
# run it through cmd and judge success by exit code alone.
function Invoke-Git {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]] $GitArgs)
    $quoted = ($GitArgs | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } }) -join ' '
    $out = cmd /c "git $quoted 2>&1"
    [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = ($out -join "`n") }
}

$snapshot = 'data/local-public.json'

# 1. Pull
$pull = Invoke-Git pull --ff-only --quiet origin main
if ($pull.ExitCode -ne 0) { Write-RunLog "pull failed, continuing with local copy - $($pull.Output)" }

# 2. Emailed rates, ahead of the scrape: a new comparison feeds the carriers that have no
# public page, and a new ACE change has to be on record before the trend is rebuilt.
$history = 'data/ace-fsc-history.json'
if (Test-Path -LiteralPath $MsgFolder) {
    $msgs = @(Get-ChildItem -LiteralPath $MsgFolder -Filter *.msg -File)
    if ($msgs.Count) {
        try {
            $out = & (Join-Path $PSScriptRoot 'Get-FscEmail.ps1') -MsgFolder $MsgFolder
            foreach ($line in @($out | Where-Object { "$_" -match '->' })) { Write-RunLog ("email: " + "$line".Trim()) }
            Write-RunLog "read $($msgs.Count) saved message(s) from $MsgFolder"
        }
        catch { Write-RunLog "email step failed - $($_.Exception.Message)" }
    }
}

# A new rate from email changes a tracked file. Left uncommitted it would block the next
# --ff-only pull, and the cloud run needs it anyway to redraw the chart.
$historyDirty = (Invoke-Git status --porcelain -- $history).Output.Trim().Length -gt 0

# 3. Scrape
& (Join-Path $PSScriptRoot 'Get-FuelSurcharges.ps1') -NoHistory -PublicSnapshotPath $snapshot | Out-Null

# Diesel trend for the local dashboard only; the cloud run owns the committed copy.
try { & (Join-Path $PSScriptRoot 'Get-FuelTrend.ps1') -LocalOnly | Out-Null }
catch { Write-RunLog "trend failed - $($_.Exception.Message)" }

if ($NoPush) { Write-RunLog 'scraped; push skipped (-NoPush)'; return }

# 3. Publish the snapshot only
$snapFull = Join-Path $PSScriptRoot $snapshot
if (-not (Test-Path $snapFull)) { Write-RunLog "no $snapshot produced; nothing to publish"; return }

$rows = @((Get-Content $snapFull -Raw | ConvertFrom-Json).rates).Count
if ($rows -eq 0) { Write-RunLog "$snapshot has no rows; not publishing"; return }

# The snapshot is regenerated every run with a fresh timestamp. Only publish when the
# rates themselves changed, or the last published copy is a day old - otherwise every
# run would push a commit that differs by timestamp alone.
$committed = Invoke-Git show "HEAD:$snapshot"
$needsPush = $true
if ($committed.ExitCode -eq 0) {
    try {
        $old = $committed.Output | ConvertFrom-Json
        $new = Get-Content $snapFull -Raw | ConvertFrom-Json
        $sig = { param($s) (@($s.rates) | Sort-Object carrier, service | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.carrier, $_.service, $_.percent, $_.effective_from, $_.effective_to }) -join ';' }
        $sameRates = (& $sig $old) -eq (& $sig $new)
        $ageHours  = ((Get-Date).ToUniversalTime() - ([datetime]$old.generated_at).ToUniversalTime()).TotalHours
        if ($sameRates -and $ageHours -lt 20) { $needsPush = $false }
    }
    catch { }
}
if (-not $needsPush) {
    # Drop the regenerated snapshot, which differs by timestamp alone.
    Invoke-Git checkout -- $snapshot | Out-Null
    if (-not $historyDirty) {
        Write-RunLog 'rates unchanged since last publish; nothing pushed'
        return
    }
}

$files = @()
if ($needsPush)     { $files += $snapshot }
if ($historyDirty)  { $files += $history }

$message = if ($historyDirty -and $needsPush) { 'Local snapshot and emailed ACE rate ' }
           elseif ($historyDirty)             { 'Emailed ACE rate ' }
           else                               { 'Local snapshot ' }

$add    = Invoke-Git add -- @files
$commit = Invoke-Git commit --quiet -m ($message + (Get-Date -Format 'yyyy-MM-dd')) -- @files
if ($commit.ExitCode -ne 0) { Write-RunLog "commit failed - $($commit.Output)"; return }

$push = Invoke-Git push --quiet origin main
if ($push.ExitCode -ne 0) {
    # The cloud run may have pushed in the meantime. Rebase onto it and try once more.
    $rebase = Invoke-Git pull --rebase --quiet origin main
    if ($rebase.ExitCode -eq 0) { $push = Invoke-Git push --quiet origin main }
}

if ($push.ExitCode -eq 0) { Write-RunLog ("pushed " + ($files -join ', ') + " ($rows rows)") }
else { Write-RunLog "push failed - $($push.Output)" }
