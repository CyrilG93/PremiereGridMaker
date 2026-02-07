param(
    [ValidateSet("User", "System")]
    [string]$Scope = "User",
    [switch]$SkipDebugMode
)

$ErrorActionPreference = "Stop"

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Copy-ExtensionFiles {
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string]$DestinationRoot
    )

    $excludeNames = @(".git", "node_modules")

    if (Test-Path -LiteralPath $DestinationRoot) {
        Remove-Item -LiteralPath $DestinationRoot -Recurse -Force
    }
    New-Item -ItemType Directory -Path $DestinationRoot -Force | Out-Null

    Get-ChildItem -LiteralPath $SourceRoot -Force | ForEach-Object {
        if ($excludeNames -contains $_.Name) {
            return
        }
        Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $DestinationRoot $_.Name) -Recurse -Force
    }
}

function Enable-CepDebugMode {
    $versions = 8..11
    foreach ($version in $versions) {
        $keyPath = "HKCU:\Software\Adobe\CSXS.$version"
        New-Item -Path $keyPath -Force | Out-Null
        New-ItemProperty -Path $keyPath -Name "PlayerDebugMode" -Value "1" -PropertyType String -Force | Out-Null
    }
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path (Join-Path $scriptDir "..")).Path
$extensionName = "PremiereGridMaker"

if ($Scope -eq "System" -and -not (Test-IsAdministrator)) {
    throw "System scope requires an elevated PowerShell session (Run as Administrator)."
}

$basePath = if ($Scope -eq "System") {
    "${env:ProgramFiles(x86)}\Common Files\Adobe\CEP\extensions"
} else {
    Join-Path $env:APPDATA "Adobe\CEP\extensions"
}

New-Item -ItemType Directory -Path $basePath -Force | Out-Null
$installPath = Join-Path $basePath $extensionName

Copy-ExtensionFiles -SourceRoot $repoRoot -DestinationRoot $installPath

if (-not $SkipDebugMode) {
    Enable-CepDebugMode
}

Write-Host "Installed '$extensionName' to: $installPath"
if ($SkipDebugMode) {
    Write-Host "Skipped CEP debug mode changes."
} else {
    Write-Host "CEP debug mode enabled for CSXS.8 to CSXS.11 (HKCU)."
}
Write-Host "Open Premiere Pro: Window > Extensions (Legacy) > Premiere Grid Maker"
