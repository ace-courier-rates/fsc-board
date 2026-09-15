#Requires -Version 5.1
<#
.SYNOPSIS
    Serves the dashboard in .\site over http://localhost so you can view it locally
    exactly as it will appear on GitHub Pages.

.EXAMPLE
    .\Start-Preview.ps1
    .\Start-Preview.ps1 -Port 8090 -NoBrowser
#>
[CmdletBinding()]
param(
    [int]    $Port = 8080,
    [string] $Root,
    [switch] $NoBrowser
)

$ErrorActionPreference = 'Stop'
if (-not $Root) { $Root = Join-Path $PSScriptRoot 'site' }
$Root = (Resolve-Path $Root).Path

# Over http the page reads site/latest.json - the public snapshot the scraper writes.
# This preview therefore shows exactly what customers see. For the full internal view,
# open site/index.html straight off disk instead.

$prefix   = "http://localhost:$Port/"
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)

try { $listener.Start() }
catch { throw "Could not bind $prefix - is another process using port $Port? ($($_.Exception.Message))" }

Write-Host "Serving $Root at $prefix" -ForegroundColor Green
Write-Host 'Press Ctrl+C to stop.' -ForegroundColor DarkGray
if (-not $NoBrowser) { Start-Process $prefix }

$types = @{
    '.html' = 'text/html; charset=utf-8'
    '.js'   = 'application/javascript; charset=utf-8'
    '.json' = 'application/json; charset=utf-8'
    '.css'  = 'text/css; charset=utf-8'
    '.svg'  = 'image/svg+xml'
    '.ico'  = 'image/x-icon'
}

try {
    while ($listener.IsListening) {
        $ctx  = $listener.GetContext()
        $path = [Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath).TrimStart('/')
        if ([string]::IsNullOrWhiteSpace($path)) { $path = 'index.html' }

        $full = Join-Path $Root $path
        # Keep requests inside the served directory.
        $ok = $false
        try { $ok = ((Resolve-Path $full -ErrorAction Stop).Path).StartsWith($Root, [StringComparison]::OrdinalIgnoreCase) } catch { }

        if ($ok -and (Test-Path $full -PathType Leaf)) {
            $bytes = [IO.File]::ReadAllBytes($full)
            $ext   = [IO.Path]::GetExtension($full).ToLowerInvariant()
            $ctx.Response.ContentType = $(if ($types.ContainsKey($ext)) { $types[$ext] } else { 'application/octet-stream' })
            $ctx.Response.Headers.Add('Cache-Control', 'no-store')
            $ctx.Response.ContentLength64 = $bytes.Length
            $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
            Write-Host ("  200  /{0}" -f $path) -ForegroundColor DarkGray
        }
        else {
            $ctx.Response.StatusCode = 404
            $msg = [Text.Encoding]::UTF8.GetBytes("404 - /$path")
            $ctx.Response.OutputStream.Write($msg, 0, $msg.Length)
            Write-Host ("  404  /{0}" -f $path) -ForegroundColor Yellow
        }
        $ctx.Response.OutputStream.Close()
    }
}
finally {
    $listener.Stop()
    $listener.Close()
    Write-Host 'Stopped.' -ForegroundColor DarkGray
}
