# Launcher unico do Fluxo de Caixa — configura, libera lock e abre normalmente (sem fechar outras planilhas)
param(
    [switch]$Silent,
    [string]$XlsxPath = '',
    [string]$InstallPath = ''
)

$ErrorActionPreference = 'Stop'
$script:RegKey = 'HKCU:\Software\FluxoCaixa'

function Show-FluxoError {
    param([string]$Message)
    if ($Silent) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            $Message, 'Fluxo de Caixa',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } else {
        Write-Host ""
        Write-Host "[Erro] $Message" -ForegroundColor Red
        Write-Host ""
    }
}

function Register-FluxoInstallPath {
    param([string]$Path)
    $normalized = ([System.IO.Path]::GetFullPath($Path)).TrimEnd('\') + '\'
    if (-not (Test-Path -LiteralPath (Join-Path $normalized 'Fluxo_Caixa.xlsx'))) { return }

    if (-not (Test-Path $script:RegKey)) {
        New-Item -Path $script:RegKey -Force | Out-Null
    }
    Set-ItemProperty -Path $script:RegKey -Name InstallPath -Value $normalized -Type String -Force
}

function Get-RegisteredInstallPath {
    try {
        $path = (Get-ItemProperty -Path $script:RegKey -Name InstallPath -ErrorAction Stop).InstallPath
        if ([string]::IsNullOrWhiteSpace($path)) { return $null }
        $path = $path.TrimEnd('\')
        if ((Test-Path -LiteralPath (Join-Path $path 'Fluxo_Caixa.xlsx')) -and
            (Test-Path -LiteralPath (Join-Path $path 'abrir-fluxo.ps1'))) {
            return $path
        }
    } catch { }
    return $null
}

function Get-InstallPathFromSidecar {
    param([string]$SidecarFolder)
    $ini = Join-Path $SidecarFolder 'fluxo-caixa.path'
    if (-not (Test-Path -LiteralPath $ini)) { return $null }

    $line = (Get-Content -LiteralPath $ini -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ([string]::IsNullOrWhiteSpace($line)) { return $null }

    $path = $line.Trim().Trim('"')
    if ((Test-Path -LiteralPath (Join-Path $path 'Fluxo_Caixa.xlsx'))) {
        return ([System.IO.Path]::GetFullPath($path)).TrimEnd('\')
    }
    return $null
}

function Resolve-FluxoInstallRoot {
    if (-not [string]::IsNullOrWhiteSpace($InstallPath)) {
        $candidate = ([System.IO.Path]::GetFullPath($InstallPath)).TrimEnd('\')
        if (Test-Path -LiteralPath (Join-Path $candidate 'Fluxo_Caixa.xlsx')) {
            Register-FluxoInstallPath -Path $candidate
            return $candidate
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($XlsxPath) -and (Test-Path -LiteralPath $XlsxPath)) {
        $candidate = ([System.IO.Path]::GetFullPath((Split-Path $XlsxPath -Parent))).TrimEnd('\')
        Register-FluxoInstallPath -Path $candidate
        return $candidate
    }

    $local = $PSScriptRoot
    if ((Test-Path -LiteralPath (Join-Path $local 'Fluxo_Caixa.xlsx')) -and
        (Test-Path -LiteralPath (Join-Path $local 'abrir-fluxo.ps1'))) {
        Register-FluxoInstallPath -Path $local
        return ([System.IO.Path]::GetFullPath($local)).TrimEnd('\')
    }

    $registered = Get-RegisteredInstallPath
    if ($registered) { return $registered }

    return $null
}

function Get-ExcelPath {
    try {
        return (Get-Command excel.exe -ErrorAction Stop).Source
    } catch {
        $candidates = @(
            "${env:ProgramFiles}\Microsoft Office\root\Office16\EXCEL.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE"
        )
        return ($candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1)
    }
}

function Close-StaleFluxoExcel {
    param([string]$XlsxFullPath)

    $fullPath = [System.IO.Path]::GetFullPath($XlsxFullPath)
    $fileName = [System.IO.Path]::GetFileName($fullPath)
    $folder = [System.IO.Path]::GetDirectoryName($fullPath)

    # Fecha APENAS a planilha do Fluxo, se ela estiver aberta, e preserva todos
    # os outros documentos Excel abertos e a propria janela do Excel.
    try {
        $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        foreach ($wb in @($xl.Workbooks)) {
            try {
                # Identidade por caminho completo: uma copia com o mesmo nome em outra
                # pasta NAO e fechada por engano.
                $wbPath = $null
                try { if ($wb.FullName) { $wbPath = [System.IO.Path]::GetFullPath($wb.FullName) } } catch { }
                if ($wbPath -and ($wbPath -ieq $fullPath)) {
                    try { $wb.Save() } catch { }
                    $wb.Close($false)
                }
            } catch { }
        }
    } catch { }

    # Remove apenas o lock orfao (~$Arquivo) deixado por um Excel que travou.
    $excelRunning = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0
    $lockFile = Join-Path $folder "~`$$fileName"
    if (-not $excelRunning -and (Test-Path -LiteralPath $lockFile)) {
        try { Remove-Item -LiteralPath $lockFile -Force } catch { }
    }
}

$root = Resolve-FluxoInstallRoot
if (-not $root) {
    Show-FluxoError @"
Nao foi possivel localizar a instalacao do Fluxo de Caixa.

Abra uma vez a partir da pasta onde estao os arquivos do programa
ou crie um arquivo fluxo-caixa.path ao lado deste atalho com o caminho da pasta.
"@
    exit 1
}

$xlsx = Join-Path $root 'Fluxo_Caixa.xlsx'
$xlsx = (Resolve-Path -LiteralPath $xlsx).Path

$configScript = Join-Path $root 'configurar.ps1'
if (Test-Path -LiteralPath $configScript) {
    try {
        & $configScript -Quiet -RootPath $root
    } catch {
        Show-FluxoError "Falha na configuracao automatica.`n`n$($_.Exception.Message)"
        exit 1
    }
}

$excel = Get-ExcelPath
if (-not $excel) {
    Show-FluxoError "Microsoft Excel nao encontrado neste computador."
    exit 1
}

Close-StaleFluxoExcel -XlsxFullPath $xlsx

# Instancia dedicada (/x): garante que a customizacao VSTO seja inicializada.
# Reaproveitar um Excel ja aberto faz o VSTO NAO carregar a customizacao nessa
# sessao, e ai a aba "Fluxo de Caixa" nao aparece de forma alguma.
# As demais planilhas do usuario ficam na instancia antiga, intactas.
$proc = Start-Process -FilePath $excel -ArgumentList @('/x', $xlsx) -PassThru
if (-not $proc) {
    Show-FluxoError "Nao foi possivel iniciar o Excel."
    exit 1
}

exit 0
