# Configura confianca VSTO para Fluxo de Caixa (locais confiaveis + certificado)
# Execute como Administrador para instalar certificado em LocalMachine\TrustedPublisher

param(
    [switch]$AllUsers
)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
if (Test-Path (Join-Path $PSScriptRoot "fluxo_caixa_crud")) { $root = $PSScriptRoot }

$paths = @(
    (Join-Path $root "fluxo_caixa_crud\bin\Release"),
    (Join-Path $root "fluxo_caixa_crud\bin\Debug"),
    (Join-Path $root "fluxo_caixa_crud"),
    $root
) | ForEach-Object { try { [IO.Path]::GetFullPath($_) } catch { $_ } } | Select-Object -Unique

function Add-ExcelTrustedLocation {
    param([string]$FolderPath, [string]$Description)
    if (-not (Test-Path $FolderPath)) { return }
    $folderPath = $FolderPath.TrimEnd('\') + '\'
    $officeVersions = Get-ChildItem "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' } | Select-Object -ExpandProperty PSChildName

    foreach ($ver in $officeVersions) {
        $base = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security\Trusted Locations"
        if (-not (Test-Path $base)) { New-Item -Path $base -Force | Out-Null }

        $existing = Get-ChildItem $base -ErrorAction SilentlyContinue | ForEach-Object {
            (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).Path
        }
        if ($existing -contains $folderPath) {
            Write-Host "[OK] Ja confiavel (Office $ver): $folderPath" -ForegroundColor DarkGray
            continue
        }

        $idx = 0
        while (Test-Path "$base\Location$idx") { $idx++ }
        $key = "$base\Location$idx"
        New-Item -Path $key -Force | Out-Null
        Set-ItemProperty -Path $key -Name Path -Value $folderPath
        Set-ItemProperty -Path $key -Name AllowSubfolders -Value 1 -Type DWord
        Set-ItemProperty -Path $key -Name Description -Value $Description
        Write-Host "[+] Local confiavel (Office $ver): $folderPath" -ForegroundColor Green
    }
}

function Install-SigningCertificate {
    param([string]$PfxPath, [string]$StoreRoot)
    if (-not (Test-Path $PfxPath)) {
        Write-Warning "PFX nao encontrado: $PfxPath"
        return
    }
    $pwd = ConvertTo-SecureString "fluxo_caixa_dev" -Force -AsPlainText
    $stores = @("Cert:\CurrentUser\TrustedPublisher")
    if ($AllUsers) { $stores += "Cert:\LocalMachine\TrustedPublisher" }

    foreach ($store in $stores) {
        try {
            Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation $store -Password $pwd -Exportable | Out-Null
            Write-Host "[+] Certificado em $store" -ForegroundColor Green
        } catch {
            Write-Warning "Falha em ${store}: $($_.Exception.Message)"
        }
    }
}

Write-Host "=== Confianca VSTO - Fluxo de Caixa ===" -ForegroundColor Cyan

foreach ($p in $paths) {
    Add-ExcelTrustedLocation -FolderPath $p -Description "Fluxo de Caixa CRUD VSTO"
}

$pfx = Join-Path $root "fluxo_caixa_crud\fluxo_caixa_crud_TemporaryKey.pfx"
Install-SigningCertificate -PfxPath $pfx

# Inclusion list VSTO (desenvolvimento local)
$inclusion = "HKCU:\Software\Microsoft\VSTO\Security\Inclusion"
if (-not (Test-Path $inclusion)) { New-Item -Path $inclusion -Force | Out-Null }
$release = (Join-Path $root "fluxo_caixa_crud\bin\Release").TrimEnd('\') + '\'
if (Test-Path $release) {
    Set-ItemProperty -Path $inclusion -Name $release -Value "Fluxo_Caixa" -Type String -Force
    Write-Host "[+] VSTO Inclusion: $release" -ForegroundColor Green
}

# Limpa restos de instalacoes antigas: entradas de inclusao que apontam para pastas que
# ja nao existem mais. Entradas para pastas existentes (outras copias) sao preservadas.
$inclusionKey = Get-Item $inclusion
foreach ($entry in $inclusionKey.GetValueNames()) {
    if ($entry -eq 'PSPath' -or $entry -like 'PS*') { continue }
    if (Test-Path -LiteralPath $entry) { continue }

    try {
        Remove-ItemProperty -Path $inclusion -Name $entry -Force
        Write-Host "[ - ] Inclusao obsoleta removida: $entry" -ForegroundColor DarkYellow
    } catch {
        Write-Warning "Nao foi possivel remover a inclusao obsoleta ${entry}: $($_.Exception.Message)"
    }
}

Write-Host "`nConcluido. Feche o Excel se estiver aberto e abra:" -ForegroundColor Yellow
Write-Host "  $release\Fluxo_Caixa.xlsx" -ForegroundColor White
Write-Host ""
Write-Host "Se ainda falhar, execute como Administrador com -AllUsers" -ForegroundColor Yellow
