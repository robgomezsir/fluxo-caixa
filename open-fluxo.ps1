# Desenvolvimento — abre bin\Release normalmente (sem fechar outras planilhas)
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$launcher = Join-Path $PSScriptRoot 'dist-assets\abrir-fluxo.ps1'
$xlsx = Join-Path $PSScriptRoot "fluxo_caixa_crud\bin\$Configuration\Fluxo_Caixa.xlsx"

if (-not (Test-Path $launcher)) {
    Write-Error "Launcher nao encontrado: $launcher"
}
if (-not (Test-Path $xlsx)) {
    Write-Error "Planilha nao encontrada: $xlsx`nRode .\build.ps1 primeiro."
}

& $launcher -XlsxPath $xlsx
exit $LASTEXITCODE
