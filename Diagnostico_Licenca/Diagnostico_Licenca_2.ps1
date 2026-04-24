<#
.SYNOPSIS
    Diagnostico completo de licenciamento, versao e atualizacoes do Windows.
    Detecta divergencia entre a chave OEM (BIOS) e a chave instalada,
    identifica chaves genericas (GVLK) e aplica correcao automatica (opcional).
.DESCRIPTION
    Este script gera um relatorio detalhado que inclui:
    - Chave de produto OEM (BIOS) e chave instalada.
    - Canal de licenciamento (OEM, Retail, Volume).
    - Status de ativacao do Windows.
    - Deteccao de chave generica (GVLK) por sufixo e canal.
    - Edicao, versao e compilacao do Windows.
    - Data da ultima atualizacao instalada (dd/mm/yyyy).
    - Comparacao OEM vs. chave instalada com correcao automatica (se permitido).
.PARAMETER DiagnosticOnly
    Se especificado, realiza apenas o diagnostico, sem aplicar correcoes de chave.
.NOTES
    Requer: Execucao como Administrador.
    Compatibilidade: Windows 10/11
#>

param(
    [switch]$DiagnosticOnly
)

# ─────────────────────────────────────────────
#  Configuracao de log
# ─────────────────────────────────────────────
$ScriptName = "Diagnostico_Licenca"
if ($PSScriptRoot) {
    $logDir = Join-Path $PSScriptRoot "logs"
}
else {
    $logDir = Join-Path $env:TEMP "Diagnostico_Logs"
}
$logFile = Join-Path $logDir "$ScriptName.log"

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force -ErrorAction SilentlyContinue | Out-Null
}

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "$timestamp [$ScriptName] $Message" -Encoding UTF8 -ErrorAction SilentlyContinue
}

# ─────────────────────────────────────────────
#  Helpers de exibicao
# ─────────────────────────────────────────────
function Write-DiagnosticResult {
    param(
        [string]$Category,
        [string]$Property,
        $Value
    )
    Write-Host "[$Category] " -NoNewline -ForegroundColor Cyan
    Write-Host "$Property" -NoNewline
    Write-Host ": " -NoNewline -ForegroundColor Gray
    Write-Host $Value -ForegroundColor White
    Write-Log "[$Category] $Property`: $Value"
}

function Write-Section {
    param([string]$Title)
    Write-Host "`n-- $Title " -ForegroundColor DarkCyan -NoNewline
    Write-Host ("-" * (50 - $Title.Length)) -ForegroundColor DarkGray
    Write-Log "-- $Title"
}

function Invoke-Slmgr {
    param([string[]]$Arguments)
    $output = & cscript //B //Nologo "$env:SystemRoot\System32\slmgr.vbs" @Arguments 2>&1
    return ($output -join "`n").Trim()
}

# ─────────────────────────────────────────────
#  Verificacao de privilegios
# ─────────────────────────────────────────────
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERRO: Execute este script como Administrador." -ForegroundColor Red
    exit 1
}

Clear-Host
Write-Host "===== DIAGNOSTICO DE LICENCIAMENTO DO WINDOWS =====" -ForegroundColor Yellow
if ($DiagnosticOnly) {
    Write-Host "(Modo: Somente Diagnostico)" -ForegroundColor DarkYellow
}
Write-Host "Iniciando analise detalhada do sistema..." -ForegroundColor Gray
Write-Host ""
Write-Log "===== Inicio do diagnostico ====="

# ─────────────────────────────────────────────
#  1. CHAVE OEM NA BIOS/UEFI
# ─────────────────────────────────────────────
Write-Section "BIOS / UEFI"

$oemKey = $null
try {
    $oemKey = (Get-CimInstance -Query 'SELECT * FROM SoftwareLicensingService').OA3xOriginalProductKey

    if ($oemKey) {
        Write-DiagnosticResult "BIOS" "Chave OEM (BIOS)" $oemKey
    }
    else {
        Write-DiagnosticResult "BIOS" "Chave OEM (BIOS)" "Nenhuma chave OEM encontrada na BIOS/UEFI"
    }
}
catch {
    Write-DiagnosticResult "BIOS" "Erro" "Nao foi possivel acessar o firmware. Execute como Admin."
    Write-Log "ERRO ao consultar BIOS: $_"
}

# ─────────────────────────────────────────────
#  2. CANAL E STATUS DE LICENCIAMENTO
#     (Isola o produto Windows pelo ApplicationID)
# ─────────────────────────────────────────────
Write-Section "LICENCIAMENTO"

$windowsAppID = '55c92734-d682-4d71-983e-d6ec3f16059f'  # Windows OS
$licensingData = Get-CimInstance -ClassName SoftwareLicensingProduct |
Where-Object { $_.PartialProductKey -and $_.ApplicationID -eq $windowsAppID }

if (-not $licensingData) {
    $licensingData = Get-CimInstance -ClassName SoftwareLicensingProduct |
    Where-Object { $_.PartialProductKey -and $_.Name -like 'Windows*' }
}

$installedPartial = $null

if ($licensingData) {
    $licenseStatus = switch ($licensingData.LicenseStatus) {
        0 { "Nao Licenciado" }
        1 { "Licenciado (Ativado)" }
        2 { "Periodo de avaliacao" }
        3 { "Periodo de avaliacao expirado" }
        4 { "Notificacao de avaliacao" }
        default { "Desconhecido ($($licensingData.LicenseStatus))" }
    }

    $licenseType = switch -Wildcard ($licensingData.ProductKeyChannel) {
        "*OEM*" { "OEM - Vinculado ao hardware" }
        "*Retail*" { "Retail - Chave adquirida separadamente" }
        "*Volume*" { "Volume (GVLK/MAK) - Licenciamento corporativo" }
        default { "Outro (possivelmente Licenca Digital)" }
    }

    $installedPartial = $licensingData.PartialProductKey

    Write-DiagnosticResult "SISTEMA" "Status da Licenca"       $licenseStatus
    Write-DiagnosticResult "SISTEMA" "Canal de Licenca"        $licensingData.ProductKeyChannel
    Write-DiagnosticResult "SISTEMA" "Tipo de Licenca"         $licenseType
    Write-DiagnosticResult "SISTEMA" "Chave Parcial Instalada" $installedPartial

    # Leitura da chave de backup no registro
    try {
        $regKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
        $backupKey = Get-ItemPropertyValue -Path $regKeyPath -Name "BackupProductKeyDefault" -ErrorAction Stop
        if ($backupKey) {
            Write-DiagnosticResult "SISTEMA" "Chave Instalada (Backup Registro)" $backupKey
        }
    }
    catch {
        Write-DiagnosticResult "SISTEMA" "Chave Instalada (Backup Registro)" "Nao disponivel"
    }
}
else {
    Write-DiagnosticResult "SISTEMA" "ERRO" "Nenhum produto de licenciamento do Windows encontrado."
}

# ─────────────────────────────────────────────
#  3. DETECCAO DE CHAVE GENERICA (GVLK)
#     Lista de objetos para evitar duplicacao
# ─────────────────────────────────────────────
Write-Section "DETECCAO DE CHAVE GENERICA"

# Lista robusta de chaves genericas da Microsoft (Windows 10/11/Server)
$genericKeyList = @(
    # Windows 11
    [PSCustomObject]@{ Suffix = "W269N"; Edition = "Windows 11 Pro" }
    [PSCustomObject]@{ Suffix = "MH37W"; Edition = "Windows 11 Pro N" }
    [PSCustomObject]@{ Suffix = "NW6C2"; Edition = "Windows 11 Pro Education" }
    [PSCustomObject]@{ Suffix = "2WH4N"; Edition = "Windows 11 Pro Education N" }
    [PSCustomObject]@{ Suffix = "YNMGQ"; Edition = "Windows 11 Education" }
    [PSCustomObject]@{ Suffix = "84NGF"; Edition = "Windows 11 Education N" }
    [PSCustomObject]@{ Suffix = "DPH2V"; Edition = "Windows 11 Enterprise" }
    [PSCustomObject]@{ Suffix = "YYVX9"; Edition = "Windows 11 Enterprise N" }
    [PSCustomObject]@{ Suffix = "44RPN"; Edition = "Windows 11 Enterprise G" }
    [PSCustomObject]@{ Suffix = "FW7NV"; Edition = "Windows 11 Enterprise G N" }
    [PSCustomObject]@{ Suffix = "YTMG3"; Edition = "Windows 11 Home" }
    [PSCustomObject]@{ Suffix = "4NX46"; Edition = "Windows 11 Home N" }
    [PSCustomObject]@{ Suffix = "BT79Q"; Edition = "Windows 11 Home Single Language" }
    [PSCustomObject]@{ Suffix = "YBCNQ"; Edition = "Windows 11 Home Country Specific" }
    [PSCustomObject]@{ Suffix = "GVGXT"; Edition = "Windows 11 Pro for Workstations" }
    [PSCustomObject]@{ Suffix = "C64HJ"; Edition = "Windows 11 Pro N for Workstations" }
    # Windows 10
    [PSCustomObject]@{ Suffix = "TX9XD"; Edition = "Windows 10 Home" }
    [PSCustomObject]@{ Suffix = "3KHY7"; Edition = "Windows 10 Home N" }
    [PSCustomObject]@{ Suffix = "7HNRX"; Edition = "Windows 10 Home Single Language" }
    [PSCustomObject]@{ Suffix = "PVMJN"; Edition = "Windows 10 Home Country Specific" }
    [PSCustomObject]@{ Suffix = "VK7JG"; Edition = "Windows 10 Pro" }
    [PSCustomObject]@{ Suffix = "HMNMQ"; Edition = "Windows 10 Pro N" }
    [PSCustomObject]@{ Suffix = "W269N"; Edition = "Windows 10 Pro" }
    [PSCustomObject]@{ Suffix = "XKCNC"; Edition = "Windows 10 Pro for Workstations" }
    [PSCustomObject]@{ Suffix = "V3W6J"; Edition = "Windows 10 Pro N for Workstations" }
    [PSCustomObject]@{ Suffix = "NW6C2"; Edition = "Windows 10 Pro Education" }
    [PSCustomObject]@{ Suffix = "2WH4N"; Edition = "Windows 10 Pro Education N" }
    [PSCustomObject]@{ Suffix = "YNMGQ"; Edition = "Windows 10 Education" }
    [PSCustomObject]@{ Suffix = "84NGF"; Edition = "Windows 10 Education N" }
    [PSCustomObject]@{ Suffix = "NPPR9"; Edition = "Windows 10 Enterprise" }
    [PSCustomObject]@{ Suffix = "DPH2V"; Edition = "Windows 10 Enterprise" }
    [PSCustomObject]@{ Suffix = "YYVX9"; Edition = "Windows 10 Enterprise N" }
    [PSCustomObject]@{ Suffix = "44RPN"; Edition = "Windows 10 Enterprise G" }
    [PSCustomObject]@{ Suffix = "FW7NV"; Edition = "Windows 10 Enterprise G N" }
    # Server
    [PSCustomObject]@{ Suffix = "WX4NM"; Edition = "Windows Server 2022 Standard" }
    [PSCustomObject]@{ Suffix = "VDYBN"; Edition = "Windows Server 2022 Datacenter" }
    [PSCustomObject]@{ Suffix = "N69G4"; Edition = "Windows Server 2019 Standard" }
    [PSCustomObject]@{ Suffix = "WMDGN"; Edition = "Windows Server 2019 Datacenter" }
    [PSCustomObject]@{ Suffix = "DPCNP"; Edition = "Windows Server 2016 Standard" }
    [PSCustomObject]@{ Suffix = "CB7KF"; Edition = "Windows Server 2016 Datacenter" }
)

$isGeneric = $false
$genericReason = @()
$genericEdition = $null

if ($installedPartial) {
    $partialUpper = $installedPartial.ToUpper().Trim()

    # Verificacao 1: sufixo na lista de chaves genericas conhecidas
    $match = $genericKeyList | Where-Object { $_.Suffix -eq $partialUpper }
    if ($match) {
        $isGeneric = $true
        $genericEdition = $match.Edition
        $genericReason += "Sufixo '$partialUpper' eh GVLK conhecida para: $genericEdition"
    }

    # Verificacao 2: canal Volume:GVLK (chaves de ativacao KMS)
    if ($licensingData -and $licensingData.ProductKeyChannel -like "*Volume:GVLK*") {
        $isGeneric = $true
        $genericReason += "Canal de licenca eh 'Volume:GVLK' (ativacao por servidor KMS)"
    }

    # Verificacao 3: Windows nao ativado com chave parcial presente
    if ($licensingData -and $licensingData.LicenseStatus -ne 1) {
        $genericReason += "LicenseStatus=$($licensingData.LicenseStatus) - Windows nao esta ativado"
    }

    if ($isGeneric) {
        Write-Host "`n[ALERTA] CHAVE GENERICA DETECTADA" -ForegroundColor Red
        foreach ($reason in $genericReason) {
            Write-DiagnosticResult "GENERICA" "Motivo" $reason
        }
        if ($genericEdition) {
            Write-DiagnosticResult "GENERICA" "Edicao mapeada" $genericEdition
        }
        Write-DiagnosticResult "GENERICA" "Conclusao" "Esta chave NAO ativa o Windows. Eh apenas uma chave de instalacao padrao."
        Write-Log "CHAVE GENERICA detectada. Sufixo=$partialUpper | Motivos: $($genericReason -join ' | ')"
    }
    else {
        Write-Host "`n[OK] Chave instalada nao eh generica." -ForegroundColor Green
        Write-DiagnosticResult "GENERICA" "Sufixo verificado" $partialUpper
        Write-DiagnosticResult "GENERICA" "Conclusao" "Chave proprietaria - pode ser OEM, Retail ou Licenca Digital."
        Write-Log "Chave NAO generica. Sufixo=$partialUpper"
    }
}
else {
    Write-DiagnosticResult "GENERICA" "Verificacao" "Nenhuma chave parcial disponivel para analise."
    Write-Log "Verificacao de chave generica ignorada: sem chave parcial instalada."
}

# ─────────────────────────────────────────────
#  4. EDICAO, VERSAO E BUILD
# ─────────────────────────────────────────────
Write-Section "VERSAO DO SISTEMA"

$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
Write-DiagnosticResult "VERSAO" "Edicao do Windows" $osInfo.Caption
Write-DiagnosticResult "VERSAO" "Numero da Versao"  $osInfo.Version
Write-DiagnosticResult "VERSAO" "Build do Sistema"  $osInfo.BuildNumber

# ─────────────────────────────────────────────
#  5. DATA DA ULTIMA ATUALIZACAO
# ─────────────────────────────────────────────
Write-Section "ATUALIZACOES"

try {
    $latestHotfix = Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1
    if ($latestHotfix.InstalledOn) {
        $formattedDate = $latestHotfix.InstalledOn.ToString("dd/MM/yyyy")
        Write-DiagnosticResult "ATUALIZACAO" "Ultima (HotFix)" "$formattedDate - $($latestHotfix.HotFixID)"
    }
    else {
        Write-DiagnosticResult "ATUALIZACAO" "Ultima (HotFix)" "Nenhuma atualizacao encontrada"
    }
}
catch {
    Write-DiagnosticResult "ATUALIZACAO" "ERRO HotFix" "Falha ao consultar historico."
    Write-Log "ERRO HotFix: $_"
}

try {
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $historyCount = $updateSearcher.GetTotalHistoryCount()

    if ($historyCount -gt 0) {
        $lastComUpdate = $updateSearcher.QueryHistory(0, $historyCount) |
        Where-Object { $_.ResultCode -eq 2 } |
        Sort-Object Date -Descending |
        Select-Object -First 1

        if ($lastComUpdate) {
            $comDate = $lastComUpdate.Date.ToString("dd/MM/yyyy")
            Write-DiagnosticResult "ATUALIZACAO" "Ultima (COM)" "$comDate - $($lastComUpdate.Title)"
        }
    }
}
catch {
    Write-DiagnosticResult "ATUALIZACAO" "Ultima (COM)" "Nao disponivel para esta versao do Windows"
}

# ─────────────────────────────────────────────
#  6. COMPARACAO OEM vs. CHAVE INSTALADA
#     e correcao automatica (se permitido)
# ─────────────────────────────────────────────
Write-Section "COMPARACAO E CORRECAO DE CHAVE"

if ($oemKey -and $installedPartial) {

    $oemPartial = $oemKey.Split("-")[-1].ToUpper().Trim()
    $sysPartial = $installedPartial.ToUpper().Trim()

    Write-DiagnosticResult "COMPARACAO" "Sufixo OEM (BIOS)"    $oemPartial
    Write-DiagnosticResult "COMPARACAO" "Sufixo Instalado"     $sysPartial

    if ($oemPartial -eq $sysPartial) {
        Write-Host "`n[OK] A chave instalada corresponde a chave OEM da BIOS. Nenhuma acao necessaria." -ForegroundColor Green
        Write-Log "Chaves OEM e instalada sao compativeis. Nenhuma acao realizada."

    }
    else {
        Write-Host "`n[ALERTA] Divergencia detectada entre a chave OEM (BIOS) e a chave instalada." -ForegroundColor Yellow
        Write-Log "DIVERGENCIA: OEM=$oemPartial | Instalada=$sysPartial"

        if ($DiagnosticOnly) {
            Write-Host "Modo somente diagnostico ativo. A correcao NAO sera aplicada." -ForegroundColor DarkYellow
            Write-Log "DiagnosticOnly = true. Correcao ignorada."
        }
        else {
            # Solicita confirmacao do usuario
            Write-Host "Deseja substituir a chave atual pela OEM da BIOS? (S/N): " -ForegroundColor Yellow -NoNewline
            $resp = Read-Host
            if ($resp -ne 'S' -and $resp -ne 's') {
                Write-Host "Correcao cancelada pelo usuario." -ForegroundColor Yellow
                Write-Log "Correcao cancelada pelo usuario."
            }
            else {
                # 6a. Copiar chave OEM para o clipboard
                try {
                    Set-Clipboard -Value $oemKey
                    Write-Host "[OK] Chave OEM copiada para o clipboard." -ForegroundColor Green
                    Write-Log "Chave OEM copiada para o clipboard."
                }
                catch {
                    Write-Host "[ALERTA] Nao foi possivel copiar para o clipboard: $_" -ForegroundColor Yellow
                    Write-Log "AVISO clipboard: $_"
                }

                # 6b. Remover a chave atual (slmgr /upk)
                Write-Host "  Removendo a chave instalada atual..." -ForegroundColor Yellow
                try {
                    $r = Invoke-Slmgr "/upk"
                    Write-DiagnosticResult "CORRECAO" "Remocao da chave atual" $r
                }
                catch {
                    Write-Host "ERRO ao remover a chave atual: $_" -ForegroundColor Red
                    Write-Log "ERRO /upk: $_"
                    exit 1
                }

                # 6c. Limpar a chave do registro (slmgr /cpky)
                Write-Host "  Limpando chave do registro..." -ForegroundColor Yellow
                try {
                    $r = Invoke-Slmgr "/cpky"
                    Write-DiagnosticResult "CORRECAO" "Limpeza do registro" $r
                }
                catch {
                    Write-Log "AVISO /cpky: $_"
                }

                # 6d. Instalar a chave OEM da BIOS (via CIM/WMI)
                Write-Host "  Instalando chave OEM da BIOS..." -ForegroundColor Yellow
                try {
                    $svc = Get-CimInstance -ClassName SoftwareLicensingService
                    $svc.InstallProductKey($oemKey) | Out-Null
                    $svc.RefreshLicenseStatus() | Out-Null
                    Write-DiagnosticResult "CORRECAO" "Instalacao da chave OEM" "Sucesso (via CIM)"
                    Write-Log "Chave OEM instalada com sucesso via CIM."
                }
                catch {
                    Write-Host "ERRO ao instalar chave via CIM. Tentando slmgr..." -ForegroundColor Yellow
                    try {
                        $r = Invoke-Slmgr "/ipk", $oemKey
                        Write-DiagnosticResult "CORRECAO" "Instalacao da chave OEM (slmgr)" $r
                    }
                    catch {
                        Write-Host "ERRO ao instalar a chave OEM: $_" -ForegroundColor Red
                        Write-Log "ERRO /ipk: $_"
                        exit 1
                    }
                }

                # 6e. Ativar o Windows (slmgr /ato)
                Write-Host "  Ativando o Windows..." -ForegroundColor Yellow
                try {
                    $r = Invoke-Slmgr "/ato"
                    Write-DiagnosticResult "CORRECAO" "Ativacao" $r
                    Write-Log "Ativacao executada: $r"
                }
                catch {
                    Write-Host "ERRO durante a ativacao: $_" -ForegroundColor Red
                    Write-Log "ERRO /ato: $_"
                }

                # 6f. Aguardar e verificar novo status
                Write-Host "  Aguardando atualizacao do status de ativacao..." -ForegroundColor Yellow
                Start-Sleep -Seconds 3
                Write-Host "`n  Verificando novo status de ativacao..." -ForegroundColor Yellow
                try {
                    $dli = Invoke-Slmgr "/dli"
                    Write-Host "`n$dli" -ForegroundColor Gray
                    Write-Log "Status pos-correcao: $dli"
                }
                catch {
                    Write-Log "ERRO /dli: $_"
                }
            }
        }
    }
}
elseif ($oemKey -and -not $installedPartial) {
    Write-Host "`n[ALERTA] Chave OEM encontrada na BIOS, mas nenhuma chave esta instalada no sistema." -ForegroundColor Yellow
    Write-Log "OEM presente, nenhuma chave instalada. Intervencao manual recomendada."
}
elseif (-not $oemKey -and $installedPartial) {
    Write-Host "`n[INFO] Nenhuma chave OEM na BIOS. Chave instalada mantida sem alteracoes." -ForegroundColor Cyan
    Write-Log "Sem chave OEM. Chave instalada preservada."
}
else {
    Write-Host "`n[ALERTA] Nenhuma chave encontrada (BIOS ou sistema). Ativacao manual necessaria." -ForegroundColor Red
    Write-Log "Nenhuma chave detectada em nenhuma fonte."
}

# ─────────────────────────────────────────────
#  7. RESUMO FINAL
# ─────────────────────────────────────────────
Write-Host "`n===== RESUMO DO DIAGNOSTICO =====" -ForegroundColor Yellow
Write-Log "===== Resumo final ====="

$finalLic = Get-CimInstance -ClassName SoftwareLicensingProduct |
Where-Object { $_.PartialProductKey -and $_.ApplicationID -eq $windowsAppID }
if (-not $finalLic) {
    $finalLic = Get-CimInstance -ClassName SoftwareLicensingProduct |
    Where-Object { $_.PartialProductKey -and $_.Name -like 'Windows*' }
}

if ($finalLic.LicenseStatus -eq 1) {
    Write-Host "[OK] O Windows esta ATIVADO." -ForegroundColor Green
    Write-Log "Status final: ATIVADO"
}
else {
    Write-Host "[ALERTA] O Windows NAO esta ativado. Execute 'slmgr /dlv' para detalhes." -ForegroundColor Red
    Write-Log "Status final: NAO ATIVADO"
}

if ($oemKey) {
    Write-Host "[OK] Chave OEM da BIOS: presente e processada." -ForegroundColor Green
    if (-not $DiagnosticOnly) {
        Write-Host "   A chave esta disponivel no clipboard para uso manual, se necessario." -ForegroundColor Gray
    }
}
else {
    Write-Host "[INFO] Nenhuma chave OEM da BIOS detectada (comum em PCs montados)." -ForegroundColor Yellow
    Write-Host "   A ativacao pode depender de licenca digital vinculada a conta Microsoft." -ForegroundColor Gray
}

Write-Host "`nLog completo salvo em: $logFile" -ForegroundColor DarkCyan
Write-Host "Diagnostico concluido.`n" -ForegroundColor Cyan
Write-Log "===== Fim do diagnostico ====="