# פרסום: clasp push + עדכון ה-deployment הקיים (כתובת /exec לא משתנה) + git push לדף
param([switch]$ServerOnly, [switch]$WebOnly)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

# כלים שהותקנו לאחרונה לא תמיד נמצאים ב-PATH של חלון פתוח
$env:Path = 'C:\Program Files\nodejs;' + $env:APPDATA + '\npm;' +
  [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
  [Environment]::GetEnvironmentVariable('Path', 'User')

function Run($exe, [string[]]$argv) {
  & $exe @argv
  if ($LASTEXITCODE -ne 0) { throw "$exe $($argv -join ' ') נכשל (קוד $LASTEXITCODE)" }
}

# מפרסמים רק קוד שנשמר ב-git, כדי שמה שרץ יהיה תמיד זהה למאגר
if (git status --porcelain) {
  throw 'יש שינויים שלא נשמרו ב-git. צריך לבצע commit לפני הפרסום.'
}
$rev = (git rev-parse --short HEAD).Trim()

if (-not $WebOnly) {
  $cfgPath = Join-Path $root 'deploy.config.json'
  if (-not (Test-Path $cfgPath)) {
    throw 'חסר deploy.config.json. מעתיקים את deploy.config.example.json ומדביקים את מזהה ה-deployment.'
  }
  $cfg = Get-Content $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
  if (-not $cfg.deploymentId -or $cfg.deploymentId -notmatch '^AKfy') {
    throw 'deploymentId ב-deploy.config.json לא תקין.'
  }

  Write-Host '== Apps Script: push ==' -ForegroundColor Cyan
  Push-Location (Join-Path $root 'server')
  try {
    Run 'clasp' @('push', '-f')
    Write-Host '== Apps Script: עדכון ה-deployment הקיים ==' -ForegroundColor Cyan
    Run 'clasp' @('update-deployment', $cfg.deploymentId, '-d', "git $rev")
  } finally { Pop-Location }
}

if (-not $ServerOnly) {
  Write-Host '== GitHub Pages: git push ==' -ForegroundColor Cyan
  Run 'git' @('push', 'origin', 'main')
}

Write-Host "הפרסום הסתיים (git $rev)." -ForegroundColor Green
