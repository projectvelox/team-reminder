<#
  build.ps1 — package the Teams app for sideloading / org upload.

  Usage:
    .\build.ps1                                          # uses contentUrl as-is in manifest.json
    .\build.ps1 -BaseUrl 'https://myname.github.io/team-reminder'
                                                          # rewrites contentUrl + validDomains before zipping

  Output: dist\team-reminder.zip
#>
[CmdletBinding()]
param(
  [string]$BaseUrl
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$dist = Join-Path $root 'dist'
$zip  = Join-Path $dist 'team-reminder.zip'

if (-not (Test-Path $dist)) { New-Item -ItemType Directory -Path $dist | Out-Null }
if (Test-Path $zip) { Remove-Item $zip -Force }

# Stage files in a temp dir so a rewritten manifest doesn't clobber the source
$stage = Join-Path $dist 'pkg'
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Path $stage | Out-Null

Copy-Item (Join-Path $root 'manifest.json') $stage
Copy-Item (Join-Path $root 'color.png')     $stage
Copy-Item (Join-Path $root 'outline.png')   $stage

if ($BaseUrl) {
  $base = $BaseUrl.TrimEnd('/')
  $uri    = [System.Uri]$base
  $domain = $uri.Host
  $manifestPath = Join-Path $stage 'manifest.json'
  $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
  $manifest.staticTabs[0].contentUrl = "$base/index.html"
  $manifest.staticTabs[0].websiteUrl = "$base/index.html"
  if ($manifest.validDomains -notcontains $domain) {
    $manifest.validDomains = @($manifest.validDomains + $domain)
  }
  $manifest | ConvertTo-Json -Depth 10 | Set-Content -Path $manifestPath -Encoding UTF8
  Write-Host "Rewrote manifest: contentUrl=$base/index.html, ensured $domain in validDomains"
}

# Sanity-check the staged manifest
$m = Get-Content (Join-Path $stage 'manifest.json') -Raw | ConvertFrom-Json
if ($m.staticTabs[0].contentUrl -match 'REPLACE-ME') {
  Write-Warning "manifest.json still contains REPLACE-ME placeholders. Either edit manifest.json or re-run with -BaseUrl 'https://your-host/path'."
}

# Zip — Teams requires manifest.json + icons at the package root (no parent folder)
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip -Force
Remove-Item $stage -Recurse -Force

Write-Host "Built $zip"
Get-Item $zip | Select-Object FullName,Length
