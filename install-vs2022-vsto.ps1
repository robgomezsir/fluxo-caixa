# Instala / repara Visual Studio 2022 Community para o projeto VSTO Fluxo de Caixa
# IMPORTANTE: execute este script como Administrador (botao direito > Executar como administrador)

$ErrorActionPreference = "Stop"
$vsconfig = Join-Path $PSScriptRoot "fluxo_caixa_crud\.vsconfig"
$setup = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\setup.exe"
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

Write-Host "=== Verificacao Visual Studio 2022 ===" -ForegroundColor Cyan

if (-not (Test-Admin)) {
    Write-Warning "Execute como Administrador para instalar/modificar workloads VSTO."
}

if (Test-Path $vswhere) {
    $path = & $vswhere -latest -property installationPath
    Write-Host "Instalado: $(& $vswhere -latest -property displayName)" -ForegroundColor Green
    Write-Host "Caminho: $path"
} else {
    Write-Host "Visual Studio 2022 nao encontrado. Instalando..." -ForegroundColor Yellow
    if (-not (Test-Admin)) { throw "Instalacao inicial requer Administrador." }
    winget install --id Microsoft.VisualStudio.2022.Community `
        --accept-package-agreements --accept-source-agreements `
        --override "--passive --norestart --config `"$vsconfig`""
}

if (-not (Test-Path $setup)) { throw "setup.exe do Visual Studio Installer nao encontrado." }

$installPath = & $vswhere -latest -property installationPath
$officeTargets = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Microsoft\VisualStudio\v17.0\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets"

if (-not (Test-Path $officeTargets)) {
    Write-Host "Adicionando workload Office/VSTO (requer Administrador)..." -ForegroundColor Cyan
    if (-not (Test-Admin)) {
        Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Wait
        exit $LASTEXITCODE
    }
    & $setup modify --installPath $installPath --passive --norestart `
        --add Microsoft.VisualStudio.Workload.Office `
        --add Microsoft.VisualStudio.Workload.ManagedDesktop `
        --add Microsoft.Net.Component.4.8.TargetingPack `
        --includeRecommended
}

Write-Host "`n=== Pos-instalacao ===" -ForegroundColor Cyan
& $vswhere -latest -property displayName,installationPath
$msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
Write-Host "OfficeTools: $(Test-Path $officeTargets)"
Write-Host "MSBuild: $msbuild"

if ($msbuild) {
    & $msbuild "$PSScriptRoot\fluxo_caixa_crud.sln" /p:Configuration=Release /p:Platform="Any CPU" /v:minimal
    if ($LASTEXITCODE -eq 0) {
        Write-Host "`nBuild OK: fluxo_caixa_crud\bin\Release\Fluxo_Caixa.xlsx" -ForegroundColor Green
    }
}

Write-Host "`nSe o instalador pediu reinicio, reinicie o PC antes de abrir o Excel." -ForegroundColor Yellow
