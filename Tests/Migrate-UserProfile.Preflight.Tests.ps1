$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path $PSScriptRoot '..\Migrate-UserProfile.ps1'
$source = Get-Content -Raw -Path $scriptPath

$matches = [regex]::Matches(
    $source,
    '(?s)\$openBrowsers\s*=\s*\$browserDefs\s*\|\s*Where-Object\s*\{(?<filter>.*?)\}'
)

if ($matches.Count -ne 2) {
    throw "Expected 2 browser pre-flight filters, found $($matches.Count)."
}

foreach ($match in $matches) {
    $filter = $match.Groups['filter'].Value
    if ($filter -notmatch '\$_\.Selected\s+-and\s+\(Get-Process') {
        throw 'Browser pre-flight must only check selected browser checkboxes.'
    }
}

Write-Host 'Migrate-UserProfile pre-flight tests passed.'
