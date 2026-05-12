$ErrorActionPreference = 'Stop'

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

if (-not (Test-Path (Join-Path $ProjectRoot 'dist\index.html'))) {
  npm run build | Out-Null
}

$Node = (Get-Command node -ErrorAction Stop).Source

Start-Process `
  -FilePath $Node `
  -ArgumentList @('scripts\launch-electron.cjs', '.') `
  -WorkingDirectory $ProjectRoot `
  -WindowStyle Hidden
