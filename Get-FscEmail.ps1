<#
.SYNOPSIS
    Reads ACE's internal fuel surcharge emails from Exchange and feeds them into the board.

.DESCRIPTION
    ACE's own FSC changes reach staff by email days before the public FAQ page is updated,
    and the competitor comparisons are never published anywhere else. This pulls both out
    of the mailbox over EWS and writes them into the same files the scraper uses:

      data/ace-fsc-history.json    ACE's BC / Alberta / FTL rates, one entry per change
      data/competitor-reports.json one entry per competitor comparison, dated by the email

    Two message shapes are understood, both sent as plain prose:

      "BC's FSC is 48.8 % and Alberta's FSC is 46.4 %. Direct Drive/FTL is 63.8 %"
      "Manitoulin - 65.8%  Vankam/Mustang - 69.75%  Overland - 69.9%  ..."

    Attachments are never parsed. Spreadsheets, PDFs and photos are saved to disk and
    listed in the output so a person can look at them - a scanned rate sheet is not
    something to OCR and then quote to a customer.

.PARAMETER SetupCredential
    Prompt for the mailbox credential and save it, DPAPI-encrypted, to CredentialPath.
    Run this once. The file can only be read back by the same Windows account on this
    machine; it is not portable and is not a plaintext password store.

.PARAMETER MsgFolder
    Read saved .msg files from this folder instead of the mailbox, for when EWS is closed
    to the account. Drag the email out of Outlook or save it from OWA into the folder;
    everything downstream is identical. Files that were read are moved to processed\ so
    the same rates can't be booked twice. No credential is needed in this mode.

.PARAMETER DryRun
    Parse and report what would change, writing nothing. Run this first.

.PARAMETER Days
    How far back to look. Defaults to 30.

.EXAMPLE
    .\Get-FscEmail.ps1 -SetupCredential
    .\Get-FscEmail.ps1 -DryRun
    .\Get-FscEmail.ps1

.EXAMPLE
    .\Get-FscEmail.ps1 -MsgFolder 'C:\Users\ace\Desktop\FSC Inbox' -DryRun
    .\Get-FscEmail.ps1 -MsgFolder 'C:\Users\ace\Desktop\FSC Inbox'
#>
[CmdletBinding()]
param(
    [switch] $SetupCredential,
    [switch] $DryRun,
    [int]    $Days = 30,
    [string] $MsgFolder,
    [string] $Folder = 'FSC',
    [string] $Endpoint       = 'https://mail.acecourier.ca/EWS/Exchange.asmx',
    [string] $CredentialPath,
    [string[]] $Senders      = @('dverbin@acecourier.ca', 'jcromwell@acecourier.ca', 'jmcmullin@acecourier.ca')
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# $PSScriptRoot is empty while param defaults are bound under "powershell -File", so the
# script's own folder is resolved here and every path hangs off this rather than the
# caller's working directory.
$ScriptDir      = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $CredentialPath) { $CredentialPath = Join-Path $ScriptDir 'data\ews-cred.xml' }

$DataDir        = Join-Path $ScriptDir 'data'
$AttachmentDir  = Join-Path $DataDir 'email-attachments'
$HistoryPath    = Join-Path $DataDir 'ace-fsc-history.json'
$ReportsPath    = Join-Path $DataDir 'competitor-reports.json'
$utf8NoBom      = New-Object System.Text.UTF8Encoding($false)
$ci             = [Globalization.CultureInfo]::InvariantCulture
$RxOpts         = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Singleline

# Outlook autocorrects the separator to an en dash, phones send a plain hyphen, and the
# occasional em dash turns up. Built from char codes because Windows PowerShell 5.1 has
# no `u{....} escape and this file is read as UTF-8 without a BOM.
$Dash           = '[' + [char]0x2013 + [char]0x2014 + '\-]'

# The comparison emails use ACE's shorthand for each carrier. Comox and Overland are
# scraped from their own sites, so Comox maps to nothing on purpose - the scraped rate
# wins and the emailed one is kept only as a label for the record.
$CarrierMap = @{
    'manitoulin'       = 'Manitoulin Transport'
    'vankam/mustang'   = 'Van-Kam Freightways'
    'vankam'           = 'Van-Kam Freightways'
    'comox/freightways'= $null
    'comox'            = $null
    'overland'         = 'Overland West Freight Lines'
    'hi-way 9'         = 'Hi-Way 9 Express'
    'hiway 9'          = 'Hi-Way 9 Express'
    'bandstra'         = 'Bandstra Transportation'
    "steele's"         = "Steele's Transfer"
    'steeles'          = "Steele's Transfer"
    'roseneau'         = 'Rosenau Transport'
    'rosenau'          = 'Rosenau Transport'
    'grimshaw'         = 'Grimshaw Trucking'
    'clark'            = 'Clark Freightways'
}

#------------------------------------------------------------------------------
# Credential
#------------------------------------------------------------------------------

if ($SetupCredential) {
    $dir = Split-Path $CredentialPath -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $c = Get-Credential -Message "Mailbox credential for $Endpoint (DOMAIN\user or user@acecourier.ca)"
    $c | Export-Clixml -Path $CredentialPath
    Write-Host "Saved to $CredentialPath (readable only by $env:USERNAME on $env:COMPUTERNAME)."
    return
}

if (-not $MsgFolder) {
    if (-not (Test-Path $CredentialPath)) {
        throw "No saved credential. Run: .\Get-FscEmail.ps1 -SetupCredential"
    }
    $Credential = Import-Clixml -Path $CredentialPath
}

#------------------------------------------------------------------------------
# EWS
#------------------------------------------------------------------------------

function Invoke-Ews {
    param([string] $Body)

    $envelope = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header><t:RequestServerVersion Version="Exchange2013_SP1" /></soap:Header>
  <soap:Body>$Body</soap:Body>
</soap:Envelope>
"@

    try {
        $r = Invoke-WebRequest -Uri $Endpoint -Method Post -Credential $Credential `
                               -ContentType 'text/xml; charset=utf-8' -Body $envelope `
                               -UseBasicParsing -TimeoutSec 60
    }
    catch {
        $resp = $_.Exception.Response
        if ($resp -and [int]$resp.StatusCode -eq 401) {
            throw "EWS rejected the credential (401). Re-run with -SetupCredential, or confirm the account is allowed EWS access."
        }
        throw "EWS request failed: $($_.Exception.Message)"
    }

    [xml] $xml = $r.Content
    $ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('s', 'http://schemas.xmlsoap.org/soap/envelope/')
    $ns.AddNamespace('m', 'http://schemas.microsoft.com/exchange/services/2006/messages')
    $ns.AddNamespace('t', 'http://schemas.microsoft.com/exchange/services/2006/types')

    $fault = $xml.SelectSingleNode('//s:Fault/faultstring', $ns)
    if ($fault) { throw "EWS fault: $($fault.InnerText)" }

    $bad = $xml.SelectSingleNode("//m:ResponseMessages/*[@ResponseClass='Error']/m:MessageText", $ns)
    if ($bad) { throw "EWS error: $($bad.InnerText)" }

    [pscustomobject]@{ Xml = $xml; Ns = $ns }
}

# Resolves a mail folder by the name it shows under in Outlook, searching the whole tree
# so a folder nested under another is still found. Returns the XML that names it in a
# request; falls back to the Inbox when no name is given or nothing matches.
function Get-FolderRef {
    param([string] $Name)

    if (-not $Name) { return '<t:DistinguishedFolderId Id="inbox" />' }

    $body = @"
    <m:FindFolder Traversal="Deep">
      <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties><t:FieldURI FieldURI="folder:DisplayName" /></t:AdditionalProperties>
      </m:FolderShape>
      <m:Restriction>
        <t:IsEqualTo>
          <t:FieldURI FieldURI="folder:DisplayName" />
          <t:FieldURIOrConstant><t:Constant Value="$([Security.SecurityElement]::Escape($Name))" /></t:FieldURIOrConstant>
        </t:IsEqualTo>
      </m:Restriction>
      <m:ParentFolderIds><t:DistinguishedFolderId Id="msgfolderroot" /></m:ParentFolderIds>
    </m:FindFolder>
"@
    $res = Invoke-Ews $body
    $id  = $res.Xml.SelectSingleNode('//t:Folders/*/t:FolderId', $res.Ns)
    if (-not $id) {
        Write-Warning "No mail folder named '$Name' - searching the Inbox instead."
        return '<t:DistinguishedFolderId Id="inbox" />'
    }
    Write-Host "Scanning the '$Name' folder."
    '<t:FolderId Id="{0}" ChangeKey="{1}" />' -f $id.Id, $id.ChangeKey
}

function Find-FscMessages {
    param([datetime] $Since, [string] $FolderRef)

    $cutoff = $Since.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $found  = New-Object System.Collections.Generic.List[object]
    $offset = 0

    do {
        $body = @"
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="item:DateTimeReceived" />
          <t:FieldURI FieldURI="message:From" />
          <t:FieldURI FieldURI="message:InternetMessageId" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="50" Offset="$offset" BasePoint="Beginning" />
      <m:Restriction>
        <t:And>
          <t:IsGreaterThanOrEqualTo>
            <t:FieldURI FieldURI="item:DateTimeReceived" />
            <t:FieldURIOrConstant><t:Constant Value="$cutoff" /></t:FieldURIOrConstant>
          </t:IsGreaterThanOrEqualTo>
          <t:Contains ContainmentMode="Substring" ContainmentComparison="IgnoreCase">
            <t:FieldURI FieldURI="item:Subject" />
            <t:Constant Value="FSC" />
          </t:Contains>
        </t:And>
      </m:Restriction>
      <m:ParentFolderIds>$FolderRef</m:ParentFolderIds>
    </m:FindItem>
"@
        $res   = Invoke-Ews $body
        $items = $res.Xml.SelectNodes('//t:Items/t:Message', $res.Ns)

        foreach ($i in $items) {
            $id     = $i.SelectSingleNode('t:ItemId', $res.Ns)
            $sender = $i.SelectSingleNode('t:From/t:Mailbox/t:EmailAddress', $res.Ns)
            $found.Add([pscustomobject]@{
                Id        = $id.Id
                ChangeKey = $id.ChangeKey
                Subject   = $i.SelectSingleNode('t:Subject', $res.Ns).InnerText
                Received  = [datetime]$i.SelectSingleNode('t:DateTimeReceived', $res.Ns).InnerText
                From      = if ($sender) { $sender.InnerText } else { '' }
                MessageId = $i.SelectSingleNode('t:InternetMessageId', $res.Ns).InnerText
            })
        }

        $root     = $res.Xml.SelectSingleNode('//m:RootFolder', $res.Ns)
        $lastOne  = $root.IncludesLastItemInRange -eq 'true'
        $offset  += 50
    } while (-not $lastOne -and $items.Count -gt 0)

    # Keep only the people who actually publish these, so a forwarded thread from
    # elsewhere with "FSC" in the subject doesn't get parsed as an announcement.
    $found | Where-Object { $Senders -contains $_.From.ToLower() }
}

function Get-MessageBody {
    param([string] $Id, [string] $ChangeKey)

    $body = @"
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:BodyType>Text</t:BodyType>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Body" />
          <t:FieldURI FieldURI="item:Attachments" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemIds><t:ItemId Id="$Id" ChangeKey="$ChangeKey" /></m:ItemIds>
    </m:GetItem>
"@
    $res  = Invoke-Ews $body
    $text = $res.Xml.SelectSingleNode('//t:Body', $res.Ns)

    $attachments = foreach ($a in $res.Xml.SelectNodes('//t:Attachments/t:FileAttachment', $res.Ns)) {
        [pscustomobject]@{
            Id   = $a.SelectSingleNode('t:AttachmentId', $res.Ns).Id
            Name = $a.SelectSingleNode('t:Name', $res.Ns).InnerText
        }
    }

    [pscustomobject]@{
        Text        = if ($text) { $text.InnerText } else { '' }
        Attachments = @($attachments)
    }
}

function Save-Attachment {
    param([string] $Id, [string] $Name, [string] $Stamp)

    $res     = Invoke-Ews "<m:GetAttachment><m:AttachmentShape /><m:AttachmentIds><t:AttachmentId Id=`"$Id`" /></m:AttachmentIds></m:GetAttachment>"
    $content = $res.Xml.SelectSingleNode('//t:FileAttachment/t:Content', $res.Ns)
    if (-not $content) { return $null }

    if (-not (Test-Path $AttachmentDir)) { New-Item -ItemType Directory -Path $AttachmentDir -Force | Out-Null }
    $safe = ($Name -replace '[\\/:*?"<>|]', '_')
    $out  = Join-Path $AttachmentDir "$Stamp-$safe"
    [IO.File]::WriteAllBytes($out, [Convert]::FromBase64String($content.InnerText))
    $out
}

#------------------------------------------------------------------------------
# Saved messages (.msg), for when EWS is closed
#------------------------------------------------------------------------------

# A .msg is an OLE compound file. Rather than implement that format, this takes the two
# text encodings Outlook writes its streams in and reads both: UTF-16 for the body and
# single-byte for the transport headers and HTML. Each is parsed separately and in its
# own reading order, because a forwarded chain carries older rates further down and the
# parsers deliberately take the first match - concatenating the two would scramble that.
function Read-MsgFile {
    param([string] $Path)

    $bytes = [IO.File]::ReadAllBytes($Path)
    $texts = @(
        [Text.Encoding]::Unicode.GetString($bytes)
        [Text.Encoding]::GetEncoding(28591).GetString($bytes)
    ) | ForEach-Object {
        $t = $_ -replace '<[^>]{1,400}>', ' '
        [Net.WebUtility]::HtmlDecode($t) -replace '[^\S\r\n]+', ' '
    }

    # Outlook stores the transport headers as a Unicode stream, so they surface in the
    # UTF-16 decoding, not the single-byte one. Take whichever actually carries them.
    $headers = $texts | Where-Object { $_ -match 'Date:\s*(?:\w{3},\s*)?\d{1,2}\s+\w{3}\s+\d{4}' } | Select-Object -First 1
    if (-not $headers) { $headers = $texts[1] }

    # The RFC822 Date: header, which only the real message has - quoted replies use
    # Outlook's "Sent: August 6, 2026 7:50 AM" instead.
    $received = $null
    $dm = [regex]::Match($headers, 'Date:\s*(?:\w{3},\s*)?(?<d>\d{1,2}\s+\w{3}\s+\d{4}\s+\d{2}:\d{2}:\d{2})\s*(?<z>[+-]\d{4})')
    if ($dm.Success) {
        try { $received = [datetime]::ParseExact("$($dm.Groups['d'].Value) $($dm.Groups['z'].Value)", 'd MMM yyyy HH:mm:ss zzz', $ci) } catch { }
    }
    if (-not $received) { $received = (Get-Item $Path).LastWriteTime }

    $fm = [regex]::Match($headers, 'From:\s*[^<\r\n]{0,80}<(?<a>[^>@\s]+@[^>\s]+)>(?!\s*<mailto)')
    $sm = [regex]::Match($headers, 'Subject:\s*(?<s>[^\r\n]{1,150})')

    [pscustomobject]@{
        Id        = $Path
        Subject   = if ($sm.Success) { $sm.Groups['s'].Value.Trim() } else { [IO.Path]::GetFileNameWithoutExtension($Path) }
        Received  = $received
        From      = if ($fm.Success) { $fm.Groups['a'].Value.ToLower() } else { '' }
        Texts     = $texts
        Files     = @([regex]::Matches($headers, '[A-Za-z0-9_\-]{1,40}\.(?:xlsx|xls|pdf|jpg|jpeg|png|docx)') |
                      ForEach-Object { $_.Value } | Sort-Object -Unique)
    }
}

#------------------------------------------------------------------------------
# Parsing
#------------------------------------------------------------------------------

# "BC's FSC is 48.8 % and Alberta's FSC is 46.4 %. Direct Drive/FTL is 63.8 %"
# The apostrophe arrives as either ' or a curly U+2019 depending on who typed it.
function Read-AceRates {
    param([string] $Text)

    $bc = [regex]::Match($Text, "BC.{0,3}s\s+FSC\s+is\s*(?<v>\d+(?:\.\d+)?)\s*%", $RxOpts)
    if (-not $bc.Success) { return $null }

    $ab  = [regex]::Match($Text, "Alberta.{0,3}s\s+FSC\s+is\s*(?<v>\d+(?:\.\d+)?)\s*%", $RxOpts)
    $ftl = [regex]::Match($Text, "Direct\s*Drive\s*/\s*FTL\s+is\s*(?<v>\d+(?:\.\d+)?)\s*%", $RxOpts)

    [pscustomobject]@{
        bc  = [double]$bc.Groups['v'].Value
        ab  = if ($ab.Success)  { [double]$ab.Groups['v'].Value }  else { $null }
        ftl = if ($ftl.Success) { [double]$ftl.Groups['v'].Value } else { $null }
    }
}

# "Manitoulin - 65.8%", "Vankam/Mustang - 69.75 %", "Hi-Way 9 (Alberta) - 54.25%"
# The dash is an en dash in Outlook's autocorrect and a hyphen when typed on a phone.
# Driven off $CarrierMap rather than capturing whatever sits left of the dash: the same
# list carries ACE's own rate as "ours is currently BC - 50.5% and Alberta - 40.1%", and
# a free-form label capture reads that as a carrier called "ours is currently BC".
# Longest key first, so "Vankam/Mustang" is tried before "Vankam".
function Read-CompetitorRates {
    param([string] $Text)

    $rows = New-Object System.Collections.Generic.List[object]
    $seen = @{}

    # "Steele's" reaches us as &#8217;, which decodes to a curly apostrophe and then
    # matches nothing in $CarrierMap. Fold both quote styles down to ASCII first.
    $Text = $Text -replace [char]0x2019, "'" -replace [char]0x2018, "'"

    foreach ($key in ($CarrierMap.Keys | Sort-Object { $_.Length } -Descending)) {
        $pattern = "(?<label>" + [regex]::Escape($key) + "(?:\s*\([^)]{1,20}\))?)\s*" + $Dash + "\s*(?<pct>\d+(?:\.\d+)?)\s*%"
        $m = [regex]::Match($Text, $pattern, $RxOpts)
        if (-not $m.Success) { continue }

        $carrier = $CarrierMap[$key]
        $id = if ($carrier) { $carrier } else { "label:$key" }
        if ($seen.ContainsKey($id)) { continue }
        $seen[$id] = $true

        $rows.Add([pscustomobject]@{
            label   = ($m.Groups['label'].Value -replace '\s+', ' ').Trim()
            carrier = $carrier
            percent = [double]$m.Groups['pct'].Value
        })
    }

    # A forwarded chain quotes every earlier comparison below the current one, so counting
    # "<name> - NN%" patterns and comparing with what matched cries wolf on every forward.
    # Only the useless case is worth a warning; a carrier newly added to the comparison
    # needs a $CarrierMap entry and will otherwise just be absent from the report.
    $all = ([regex]::Matches($Text, "[A-Za-z0-9)]\s*" + $Dash + "\s*\d+(?:\.\d+)?\s*%", $RxOpts)).Count
    if ($rows.Count -eq 0 -and $all -ge 3) {
        Write-Warning "Found $all rate(s) but matched no known carrier - check `$CarrierMap against this message."
    }

    $rows
}

#------------------------------------------------------------------------------
# Run
#------------------------------------------------------------------------------

if ($MsgFolder) {
    if (-not (Test-Path $MsgFolder)) { throw "Folder not found: $MsgFolder" }
    $messages = @(Get-ChildItem -Path $MsgFolder -Filter *.msg -File | ForEach-Object { Read-MsgFile $_.FullName } | Sort-Object Received)
    if (-not $messages) {
        Write-Host "No .msg files in $MsgFolder. Save the email there from Outlook or OWA, then run this again."
        return
    }
}
else {
    $messages = @(Find-FscMessages -Since (Get-Date).AddDays(-$Days) -FolderRef (Get-FolderRef $Folder) | Sort-Object Received)
    if (-not $messages) {
        Write-Host "No FSC messages from $($Senders -join ', ') in the last $Days days."
        return
    }
}

$history       = Get-Content $HistoryPath -Raw | ConvertFrom-Json
$changes       = New-Object System.Collections.Generic.List[object]
foreach ($c in $history.changes) { $changes.Add($c) }

$reports       = Get-Content $ReportsPath -Raw | ConvertFrom-Json
$reportList    = New-Object System.Collections.Generic.List[object]
foreach ($r in $reports.reports) { $reportList.Add($r) }

$historyDirty  = $false
$reportsDirty  = $false
$attachments   = New-Object System.Collections.Generic.List[string]

foreach ($msg in $messages) {
    $date = $msg.Received.ToString('yyyy-MM-dd')
    if ($MsgFolder) {
        $texts = $msg.Texts
        if ($msg.From -and $Senders -notcontains $msg.From) {
            Write-Warning "$(Split-Path $msg.Id -Leaf) is from $($msg.From), not one of the usual senders - reading it anyway."
        }
    }
    else {
        $item  = Get-MessageBody -Id $msg.Id -ChangeKey $msg.ChangeKey
        $texts = @($item.Text)
    }
    Write-Host "`n$date  $($msg.From)  `"$($msg.Subject)`""

    # A rate change is announced under the bare subject "FSC". Every other message -
    # a competitor comparison, a forwarded chain - quotes older announcements below the
    # current text, and reading those files last month's rate under this month's date.
    # Missing a change is recoverable; back-dating a wrong one quietly is not.
    $isAnnouncement = $msg.Subject -match '^\s*(?:RE:|FW:|FWD:)?\s*FSC\s*$'

    $ace = $null
    foreach ($t in $texts) { $ace = Read-AceRates $t; if ($ace) { break } }
    if ($ace -and -not $isAnnouncement) {
        Write-Host "  ACE: BC $($ace.bc)% appears here but is quoted from an older message - ignored"
        $ace = $null
    }
    if ($ace) {
        $existing = $changes | Where-Object { $_.effective -eq $date } | Select-Object -First 1
        if ($existing) {
            Write-Host "  ACE: BC $($ace.bc)% - already on record for $date"
        }
        else {
            Write-Host "  ACE: BC $($ace.bc)%  AB $($ace.ab)%  FTL $($ace.ftl)%  -> new change effective $date"
            $changes.Add([pscustomobject]@{
                effective         = $date
                bc                = $ace.bc
                ab                = $ace.ab
                ftl               = $ace.ftl
                confirmed_through = $date
                source            = "Email: $($msg.From), `"$($msg.Subject)`", $date"
            })
            $historyDirty = $true
        }
    }

    $comp = @()
    foreach ($t in $texts) { $comp = @(Read-CompetitorRates $t); if ($comp.Count -ge 3) { break } }
    if ($comp.Count -ge 3) {
        if ($reportList | Where-Object { $_.date -eq $date }) {
            Write-Host "  Competitors: $($comp.Count) rates - a report for $date is already on file"
        }
        else {
            Write-Host "  Competitors: $($comp.Count) rates -> new report dated $date"
            foreach ($r in $comp) { Write-Host ("    {0,-22} {1}%" -f $r.label, $r.percent) }
            $reportList.Add([pscustomobject]@{ date = $date; rates = $comp })
            $reportsDirty = $true
        }
    }

    if ($MsgFolder) {
        # Pulling a file out of a .msg means implementing the compound-file format, and a
        # half-parsed spreadsheet is worse than none. The name is reported so you know
        # there is something to go and save yourself.
        foreach ($n in $msg.Files) { Write-Host "  Attached: $n (not extracted - save it from Outlook if you need it)" }
    }
    else {
        foreach ($a in $item.Attachments) {
            if ($DryRun) { Write-Host "  Attachment (not saved in -DryRun): $($a.Name)"; continue }
            $saved = Save-Attachment -Id $a.Id -Name $a.Name -Stamp $date
            if ($saved) { $attachments.Add($saved); Write-Host "  Attachment saved: $saved" }
        }
    }
}

# Close off the previous rate the day before a newer one starts, so the two spans
# never overlap and the chart doesn't have to guess which was in effect.
if ($historyDirty) {
    $sorted = @($changes | Sort-Object { $_.effective })
    for ($i = 0; $i -lt $sorted.Count - 1; $i++) {
        $limit = ([datetime]::ParseExact($sorted[$i + 1].effective, 'yyyy-MM-dd', $ci)).AddDays(-1).ToString('yyyy-MM-dd')
        if ([string]::CompareOrdinal([string]$sorted[$i].confirmed_through, $limit) -gt 0) {
            $sorted[$i].confirmed_through = $limit
        }
    }
    $history.changes = $sorted
}

if ($DryRun) {
    Write-Host "`n-DryRun: nothing written."
    return
}

if ($historyDirty) {
    [IO.File]::WriteAllText($HistoryPath, ($history | ConvertTo-Json -Depth 5), $utf8NoBom)
    Write-Host "`nUpdated $HistoryPath"
}
if ($reportsDirty) {
    $reports.reports = @($reportList | Sort-Object { $_.date })
    [IO.File]::WriteAllText($ReportsPath, ($reports | ConvertTo-Json -Depth 6), $utf8NoBom)
    Write-Host "Updated $ReportsPath"
}
if (-not $historyDirty -and -not $reportsDirty) {
    Write-Host "`nNothing new."
}
if ($attachments.Count) {
    Write-Host "`n$($attachments.Count) attachment(s) saved to $AttachmentDir - these are not parsed, open them yourself:"
    foreach ($a in $attachments) { Write-Host "  $a" }
}

# Move what was read into processed\ so the folder shows only what still needs doing,
# and a second run can't book the same rates twice under a different date.
if ($MsgFolder) {
    $done = Join-Path $MsgFolder 'processed'
    if (-not (Test-Path $done)) { New-Item -ItemType Directory -Path $done -Force | Out-Null }
    foreach ($m in $messages) {
        $leaf = Split-Path $m.Id -Leaf
        $dest = Join-Path $done $leaf
        if (Test-Path $dest) { $dest = Join-Path $done ("{0}-{1}" -f $m.Received.ToString('yyyyMMdd'), $leaf) }
        Move-Item -LiteralPath $m.Id -Destination $dest -Force
    }
    Write-Host "`nMoved $($messages.Count) file(s) to $done"
}
