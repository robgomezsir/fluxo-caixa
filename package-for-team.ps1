# Gera pacote pronto para distribuir aos colegas (sem Visual Studio)
param(
    [switch]$SkipBuild,
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$releaseDir = Join-Path $root "fluxo_caixa_crud\bin\$Configuration"
$assetsDir = Join-Path $root "dist-assets"
$outDir = Join-Path $root "dist\FluxoCaixa"
$zipPath = Join-Path $root "dist\FluxoCaixa.zip"

if (-not $SkipBuild) {
    Write-Host "Compilando Release..." -ForegroundColor Cyan
    & (Join-Path $root "build.ps1") -Configuration $Configuration
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if (-not (Test-Path (Join-Path $releaseDir "Fluxo_Caixa.xlsx"))) {
    Write-Error "Build nao encontrado em $releaseDir. Feche o Excel e rode .\build.ps1"
}

Write-Host "Montando pacote em $outDir ..." -ForegroundColor Cyan
if (Test-Path $outDir) {
    try {
        Remove-Item $outDir -Recurse -Force -ErrorAction Stop
    } catch {
        # Excel pode estar usando arquivos da dist; remove o que der e mantem o resto.
        Get-ChildItem $outDir -Recurse -Force | Sort-Object { $_.PSIsContainer } -Descending | ForEach-Object {
            try { Remove-Item $_.FullName -Recurse -Force -ErrorAction Stop } catch { }
        }
    }
}
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

function Copy-ItemSafe {
    param([string]$Src, [string]$Dest)
    try {
        Copy-Item -LiteralPath $Src -Destination $Dest -Force -ErrorAction Stop
        return $true
    } catch {
        # Arquivo em uso: se o destino ja tiver o mesmo conteudo, nada a fazer.
        if (Test-Path -LiteralPath $Dest) {
            try {
                $srcHash = (Get-FileHash -LiteralPath $Src -Algorithm SHA256).Hash
                $dstHash = (Get-FileHash -LiteralPath $Dest -Algorithm SHA256).Hash
                if ($srcHash -eq $dstHash) { return $true }
            } catch { }
        }
        Write-Warning "Nao foi possivel atualizar '$Dest' (em uso pelo Excel e com conteudo diferente)."
        return $false
    }
}

$excludeExt = @('.pdb')
Get-ChildItem $releaseDir -Recurse -File | Where-Object {
    $excludeExt -notcontains $_.Extension
} | ForEach-Object {
    $rel = $_.FullName.Substring($releaseDir.Length).TrimStart('\')
    $dest = Join-Path $outDir $rel
    $destParent = Split-Path $dest -Parent
    if (-not (Test-Path $destParent)) { New-Item -ItemType Directory -Path $destParent -Force | Out-Null }
    Copy-ItemSafe -Src $_.FullName -Dest $dest | Out-Null
}

foreach ($name in @("configurar.ps1", "abrir-fluxo.ps1", "Fluxo de Caixa.vbs", "Atalho - Fluxo de Caixa.vbs", "Abrir Fluxo de Caixa.bat", "Configurar (como administrador).bat", "LEIA-ME.txt")) {
    Copy-ItemSafe -Src (Join-Path $assetsDir $name) -Dest (Join-Path $outDir $name) | Out-Null
}

$pfx = Join-Path $root "fluxo_caixa_crud\fluxo_caixa_crud_TemporaryKey.pfx"
if (Test-Path $pfx) {
    Copy-ItemSafe -Src $pfx -Dest (Join-Path $outDir "fluxo_caixa_crud_TemporaryKey.pfx") | Out-Null
}

if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path $outDir -DestinationPath $zipPath -Force

Write-Host ""
Write-Host "Pacote pronto!" -ForegroundColor Green
Write-Host "  Pasta: $outDir" -ForegroundColor White
Write-Host "  ZIP:   $zipPath" -ForegroundColor White
Write-Host ""
Write-Host "Envie o ZIP aos colegas. Eles devem:" -ForegroundColor Yellow
Write-Host "  1. Extrair o ZIP" -ForegroundColor Yellow
Write-Host "  2. Clicar em 'Fluxo de Caixa.vbs'" -ForegroundColor Yellow
Write-Host ""
