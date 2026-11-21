param([string]$Manifest = (Join-Path $PSScriptRoot '..\RELEASE_MANIFEST.json'))
$ErrorActionPreference = 'Stop'
$manifestFile = (Resolve-Path -LiteralPath $Manifest).Path
$root = [System.IO.Path]::GetDirectoryName($manifestFile)
$release = Get-Content -LiteralPath $manifestFile -Raw | ConvertFrom-Json
foreach ($entry in $release.files) {
    $path = [System.IO.Path]::GetFullPath((Join-Path $root $entry.path))
    $prefix = $root.TrimEnd('\') + '\'
    if (-not $path.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe manifest path' }
    $actual = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actual -ne $entry.sha256) { throw "SHA-256 mismatch: $($entry.path)" }
    Write-Output "Verified $($entry.path): $actual"
}
Write-Output 'All listed SHA-256 checksums match. This does not authenticate the publisher; obtain the manifest from a trusted source.'
