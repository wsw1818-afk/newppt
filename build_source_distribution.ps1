$ErrorActionPreference = "Stop"

$repoDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputDir = [Text.Encoding]::UTF8.GetString(
    [Convert]::FromBase64String("RDpcT25lRHJpdmVc7L2U65Oc7J6R7JeFXOqysOqzvOusvFxuZXdwcHQ=")
)
$targetDir = Join-Path $outputDir "DocumentExtractor_v3_source"

New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
if (Test-Path -LiteralPath $targetDir) {
    Remove-Item -LiteralPath $targetDir -Recurse -Force
}
New-Item -ItemType Directory -Path $targetDir -Force | Out-Null

$files = @(
    "ppt_extractor_v3.py",
    "requirements_runtime.txt",
    "run_from_source.bat",
    "install_requirements.bat",
    "SECURITY_PC_README.txt"
)

foreach ($file in $files) {
    Copy-Item -LiteralPath (Join-Path $repoDir $file) -Destination $targetDir -Force
}

Write-Host "Source distribution copied:"
Write-Host $targetDir
