param([string]$Environment = '.venv', [string]$OutputDir = 'release/FINDEX-v3.0.1')
$ErrorActionPreference = 'Stop'
$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
Set-Location -LiteralPath $root
$python = Join-Path (Join-Path $root $Environment) 'Scripts/python.exe'
& $python scripts/build_release.py --environment $Environment --output-dir $OutputDir
if ($LASTEXITCODE -ne 0) { throw 'FINDEX release build failed' }
