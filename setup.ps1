param(
    [Parameter(Mandatory, Position = 0)]
    [string]$GitHubUsername
)

$file = Join-Path $PSScriptRoot 'manifest.xml'
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

if ($content -notmatch 'YOUR_GITHUB_USERNAME') {
    Write-Warning "manifest.xml has already been configured. Edit it manually to change the username."
    exit 0
}

$updated = $content.Replace('YOUR_GITHUB_USERNAME', $GitHubUsername)
[System.IO.File]::WriteAllText($file, $updated, [System.Text.Encoding]::UTF8)

Write-Host "manifest.xml configured for '$GitHubUsername'."
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. git push"
Write-Host "  2. GitHub: Settings -> Pages -> Source: main branch, /docs folder"
Write-Host "  3. Word: Insert -> Get Add-ins -> Upload My Add-in -> manifest.xml"
