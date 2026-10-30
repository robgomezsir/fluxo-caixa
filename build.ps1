# Build script — requer Visual Studio 2022 com carga Office/SharePoint
param(
    [switch]$ForcePartial,
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release'
)

function Test-FileLocked {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $stream.Close()
        $stream.Dispose()
        return $false
    } catch {
        return $true
    }
}

$paths = @(
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
)
$msbuild = $paths | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $msbuild) {
    Write-Error "MSBuild nao encontrado. Execute install-vs2022-vsto.ps1 como Administrador."
    exit 1
}

$outDir = Join-Path $PSScriptRoot "fluxo_caixa_crud\bin\$Configuration"
$hostXlsx = Join-Path $outDir 'Fluxo_Caixa.xlsx'
$excelOpen = Test-FileLocked -Path $hostXlsx
$skipHost = $ForcePartial -or $excelOpen

$msbuildArgs = @(
    "$PSScriptRoot\fluxo_caixa_crud.sln",
    "/p:Configuration=$Configuration",
    '/p:Platform=Any CPU',
    '/v:minimal'
)
if ($skipHost) {
    $msbuildArgs += '/p:SkipVstoHostCustomization=true'
}

if ($excelOpen) {
    Write-Host "Excel esta usando $hostXlsx" -ForegroundColor Yellow
    Write-Host "Build parcial: DLL, manifesto e frontend serao atualizados; o .xlsx nao sera recustomizado." -ForegroundColor Yellow
    Write-Host "Feche o Excel e rode .\build.ps1 novamente para build completo." -ForegroundColor Yellow
} elseif ($ForcePartial) {
    Write-Host "Build parcial forcado (-ForcePartial)." -ForegroundColor Yellow
}

& $msbuild @msbuildArgs
if ($LASTEXITCODE -eq 0) {
    $assetsDir = Join-Path $PSScriptRoot 'dist-assets'
    foreach ($name in @('abrir-fluxo.ps1', 'configurar.ps1', 'Fluxo de Caixa.vbs', 'Atalho - Fluxo de Caixa.vbs', 'Abrir Fluxo de Caixa.bat')) {
        $src = Join-Path $assetsDir $name
        if (Test-Path $src) {
            Copy-Item $src (Join-Path $outDir $name) -Force
        }
    }

    if ($skipHost) {
        Write-Host "Build parcial OK. Feche o Excel, rode .\build.ps1 e reabra Fluxo_Caixa.xlsx para carregar a nova DLL." -ForegroundColor Green
    } else {
        Write-Host "Build OK. Abra com: fluxo_caixa_crud\bin\$Configuration\Abrir Fluxo de Caixa.bat" -ForegroundColor Green
    }
} else {
    if (-not $skipHost -and (Test-FileLocked -Path $hostXlsx)) {
        Write-Host ""
        Write-Host "Falha provavelmente porque o Excel esta com a planilha aberta." -ForegroundColor Yellow
        Write-Host "Feche o Excel ou use: .\build.ps1 -ForcePartial" -ForegroundColor Yellow
    }
    exit $LASTEXITCODE
}
