$source = "$PSScriptRoot"
$destRoot = Join-Path $env:APPDATA "pyRevit\Extensions"
$dest = Join-Path $destRoot "SHN_Tools.extension"

Write-Host "Source: $source"
Write-Host "Destination: $dest"

if (-not (Test-Path $destRoot)) {
    New-Item -ItemType Directory -Path $destRoot -Force | Out-Null
}

if (Test-Path $dest) {
    Write-Host "Removing existing extension at $dest..."
    Remove-Item -Recurse -Force $dest
}

Write-Host "Copying extension..."
Copy-Item -Recurse -Force $source $dest

Write-Host "Done. Restart Revit / pyRevit."
