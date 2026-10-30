# Configuracao do pacote Fluxo de Caixa (automatica, idempotente)
param(
    [switch]$AllUsers,
    [switch]$Quiet,
    [string]$RootPath = ''
)
$ErrorActionPreference = "Stop"
if ($RootPath) {
    $root = ([System.IO.Path]::GetFullPath($RootPath)).TrimEnd('\') + '\'
} else {
    $root = (Resolve-Path $PSScriptRoot).Path.TrimEnd('\') + '\'
}

function Test-VstoRuntime {
    $keys = @(
        "HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R"
    )
    foreach ($k in $keys) {
        if (Test-Path $k) { return $true }
    }
    return $false
}

function Test-WebView2Runtime {
    $keys = @(
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
        "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
    )
    foreach ($k in $keys) {
        if (Test-Path $k) { return $true }
    }
    return $false
}

function Add-ExcelTrustedLocation {
    param([string]$FolderPath)
    $folderPath = $FolderPath.TrimEnd('\') + '\'
    $officeVersions = Get-ChildItem "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' } | Select-Object -ExpandProperty PSChildName

    foreach ($ver in $officeVersions) {
        $base = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security\Trusted Locations"
        if (-not (Test-Path $base)) { New-Item -Path $base -Force | Out-Null }

        $existing = Get-ChildItem $base -ErrorAction SilentlyContinue | ForEach-Object {
            (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).Path
        }
        if ($existing -contains $folderPath) { continue }

        $idx = 0
        while (Test-Path "$base\Location$idx") { $idx++ }
        $key = "$base\Location$idx"
        New-Item -Path $key -Force | Out-Null
        Set-ItemProperty -Path $key -Name Path -Value $folderPath
        Set-ItemProperty -Path $key -Name AllowSubfolders -Value 1 -Type DWord
        Set-ItemProperty -Path $key -Name Description -Value "Fluxo de Caixa CRUD"
    }
}

function Install-SigningCertificate {
    param([string]$PfxPath)
    if (-not (Test-Path $PfxPath)) { return }
    $pwd = ConvertTo-SecureString "fluxo_caixa_dev" -Force -AsPlainText
    $stores = @("Cert:\CurrentUser\TrustedPublisher")
    if ($AllUsers) { $stores += "Cert:\LocalMachine\TrustedPublisher" }
    foreach ($store in $stores) {
        try {
            Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation $store -Password $pwd -Exportable | Out-Null
        } catch { }
    }
}

if (-not (Test-Path (Join-Path $root "Fluxo_Caixa.xlsx"))) {
    Write-Error "Arquivo Fluxo_Caixa.xlsx nao encontrado em: $root"
}

if (-not (Test-VstoRuntime)) {
    if (-not $Quiet) {
        Write-Host ""
        Write-Host "ATENCAO: VSTO Runtime nao detectado." -ForegroundColor Yellow
        Write-Host "Instale (gratis, uma vez): https://aka.ms/vstoruntimedownload" -ForegroundColor Yellow
        Write-Host "Normalmente ja vem com o Microsoft Office." -ForegroundColor Yellow
        Write-Host ""
    }
}

if (-not (Test-WebView2Runtime)) {
    if (-not $Quiet) {
        Write-Host "ATENCAO: WebView2 Runtime nao detectado." -ForegroundColor Yellow
        Write-Host "Instale (gratis, uma vez): https://go.microsoft.com/fwlink/p/?LinkId=2124703" -ForegroundColor Yellow
        Write-Host ""
    }
}
Add-ExcelTrustedLocation -FolderPath $root
$pfx = Join-Path $root "fluxo_caixa_crud_TemporaryKey.pfx"
Install-SigningCertificate -PfxPath $pfx

$inclusion = "HKCU:\Software\Microsoft\VSTO\Security\Inclusion"
if (-not (Test-Path $inclusion)) { New-Item -Path $inclusion -Force | Out-Null }
Set-ItemProperty -Path $inclusion -Name $root -Value "Fluxo_Caixa" -Type String -Force

$fluxoKey = 'HKCU:\Software\FluxoCaixa'
if (-not (Test-Path $fluxoKey)) { New-Item -Path $fluxoKey -Force | Out-Null }
Set-ItemProperty -Path $fluxoKey -Name InstallPath -Value $root -Type String -Force

if (-not $Quiet) {
    Write-Host "Configuracao concluida." -ForegroundColor Green
}