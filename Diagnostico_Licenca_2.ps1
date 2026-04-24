<#
.SYNOPSIS
    Diagnóstico completo de licenciamento, versão e atualizações do Windows.
    Detecta divergência entre a chave OEM (BIOS) e a chave instalada,
    identifica chaves genéricas (GVLK) e aplica correção automática (opcional).
.DESCRIPTION
    Este script gera um relatório detalhado que inclui:
    - Chave de produto OEM (BIOS) e chave instalada.
    - Canal de licenciamento (OEM, Retail, Volume).
    - Status de ativação do Windows.
    - Detecção de chave genérica (GVLK) por sufixo e canal.
    - Edição, versão e compilação do Windows.
    - Data da última atualização instalada (dd/mm/yyyy).
    - Comparação OEM vs. chave instalada com correção automática (se permitido).
.PARAMETER DiagnosticOnly
    Se especificado, realiza apenas o diagnóstico, sem aplicar correções de chave.
.NOTES
    Requer: Execução como Administrador.
    Compatibilidade: Windows 10/11
#>

param(
    [switch]$DiagnosticOnly
)

# ─────────────────────────────────────────────
#  Configuração de log
# ─────────────────────────────────────────────
$ScriptName = "Diagnostico_Licenca"
if ($PSScriptRoot) {
    $logDir = Join-Path $PSScriptRoot "logs"
} else {
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
#  Helpers de exibição
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
    Write-Host "`n── $Title " -ForegroundColor DarkCyan -NoNewline
    Write-Host ("─" * (50 - $Title.Length)) -ForegroundColor DarkGray
    Write-Log "── $Title"
}

function Invoke-Slmgr {
    param([string[]]$Arguments)
    $output = & cscript //B //Nologo "$env:SystemRoot\System32\slmgr.vbs" @Arguments 2>&1
    return ($output -join "`n").Trim()
}

# ─────────────────────────────────────────────
#  Verificação de privilégios
# ─────────────────────────────────────────────
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERRO: Execute este script como Administrador." -ForegroundColor Red
    exit 1
}

Clear-Host
Write-Host "===== DIAGNÓSTICO DE LICENCIAMENTO DO WINDOWS =====" -ForegroundColor Yellow
if ($DiagnosticOnly) {
    Write-Host "(Modo: Somente Diagnóstico)" -ForegroundColor DarkYellow
}
Write-Host "Iniciando análise detalhada do sistema...`n" -ForegroundColor Gray
Write-Log "===== Início do diagnóstico ====="

# ─────────────────────────────────────────────
#  1. CHAVE OEM NA BIOS/UEFI
# ─────────────────────────────────────────────
Write-Section "BIOS / UEFI"

$oemKey = $null
try {
    $oemKey = (Get-CimInstance -Query 'SELECT * FROM SoftwareLicensingService').OA3xOriginalProductKey

    if ($oemKey) {
        Write-DiagnosticResult "BIOS" "Chave OEM (BIOS)" $oemKey
    } else {
        Write-DiagnosticResult "BIOS" "Chave OEM (BIOS)" "Nenhuma chave OEM encontrada na BIOS/UEFI"
    }
} catch {
    Write-DiagnosticResult "BIOS" "Erro" "Não foi possível acessar o firmware. Execute como Admin."
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
        0 { "Não Licenciado" }
        1 { "Licenciado (Ativado)" }
        2 { "Período de avaliação" }
        3 { "Período de avaliação expirado" }
        4 { "Notificação de avaliação" }
        default { "Desconhecido ($($licensingData.LicenseStatus))" }
    }

    $licenseType = switch -Wildcard ($licensingData.ProductKeyChannel) {
        "*OEM*"    { "OEM — Vinculado ao hardware" }
        "*Retail*" { "Retail — Chave adquirida separadamente" }
        "*Volume*" { "Volume (GVLK/MAK) — Licenciamento corporativo" }
        default    { "Outro (possivelmente Licença Digital)" }
    }

    $installedPartial = $licensingData.PartialProductKey

    Write-DiagnosticResult "SISTEMA" "Status da Licença"       $licenseStatus
    Write-DiagnosticResult "SISTEMA" "Canal de Licença"        $licensingData.ProductKeyChannel
    Write-DiagnosticResult "SISTEMA" "Tipo de Licença"         $licenseType
    Write-DiagnosticResult "SISTEMA" "Chave Parcial Instalada" $installedPartial

    # Leitura da chave de backup no registro
    try {
        $regKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
        $backupKey  = Get-ItemPropertyValue -Path $regKeyPath -Name "BackupProductKeyDefault" -ErrorAction Stop
        if ($backupKey) {
            Write-DiagnosticResult "SISTEMA" "Chave Instalada (Backup Registro)" $backupKey
        }
    } catch {
        Write-DiagnosticResult "SISTEMA" "Chave Instalada (Backup Registro)" "Não disponível"
    }
} else {
    Write-DiagnosticResult "SISTEMA" "ERRO" "Nenhum produto de licenciamento do Windows encontrado."
}

# ─────────────────────────────────────────────
#  3. DETECÇÃO DE CHAVE GENÉRICA (GVLK)
#     Lista de objetos para evitar duplicação
# ─────────────────────────────────────────────
Write-Section "DETECÇÃO DE CHAVE GENÉRICA"

# Lista robusta de chaves genéricas da Microsoft (Windows 10/11/Server)
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

$isGeneric      = $false
$genericReason  = @()
$genericEdition = $null

if ($installedPartial) {
    $partialUpper = $installedPartial.ToUpper().Trim()

    # Verificação 1: sufixo na lista de chaves genéricas conhecidas
    $match = $genericKeyList | Where-Object { $_.Suffix -eq $partialUpper }
    if ($match) {
        $isGeneric      = $true
        $genericEdition = $match.Edition
        $genericReason += "Sufixo '$partialUpper' é GVLK conhecida para: $genericEdition"
    }

    # Verificação 2: canal Volume:GVLK (chaves de ativação KMS)
    if ($licensingData -and $licensingData.ProductKeyChannel -like "*Volume:GVLK*") {
        $isGeneric      = $true
        $genericReason += "Canal de licença é 'Volume:GVLK' (ativação por servidor KMS)"
    }

    # Verificação 3: Windows não ativado com chave parcial presente
    if ($licensingData -and $licensingData.LicenseStatus -ne 1) {
        $genericReason += "LicenseStatus=$($licensingData.LicenseStatus) — Windows não está ativado"
    }

    if ($isGeneric) {
        Write-Host "`n⚠ CHAVE GENÉRICA DETECTADA" -ForegroundColor Red
        foreach ($reason in $genericReason) {
            Write-DiagnosticResult "GENÉRICA" "Motivo" $reason
        }
        if ($genericEdition) {
            Write-DiagnosticResult "GENÉRICA" "Edição mapeada" $genericEdition
        }
        Write-DiagnosticResult "GENÉRICA" "Conclusão" "Esta chave NÃO ativa o Windows. É apenas uma chave de instalação padrão."
        Write-Log "CHAVE GENÉRICA detectada. Sufixo=$partialUpper | Motivos: $($genericReason -join ' | ')"
    } else {
        Write-Host "`n✔ Chave instalada não é genérica." -ForegroundColor Green
        Write-DiagnosticResult "GENÉRICA" "Sufixo verificado" $partialUpper
        Write-DiagnosticResult "GENÉRICA" "Conclusão" "Chave proprietária — pode ser OEM, Retail ou Licença Digital."
        Write-Log "Chave NÃO genérica. Sufixo=$partialUpper"
    }
} else {
    Write-DiagnosticResult "GENÉRICA" "Verificação" "Nenhuma chave parcial disponível para análise."
    Write-Log "Verificação de chave genérica ignorada: sem chave parcial instalada."
}

# ─────────────────────────────────────────────
#  4. EDIÇÃO, VERSÃO E BUILD
# ─────────────────────────────────────────────
Write-Section "VERSÃO DO SISTEMA"

$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
Write-DiagnosticResult "VERSÃO" "Edição do Windows" $osInfo.Caption
Write-DiagnosticResult "VERSÃO" "Número da Versão"  $osInfo.Version
Write-DiagnosticResult "VERSÃO" "Build do Sistema"  $osInfo.BuildNumber

# ─────────────────────────────────────────────
#  5. DATA DA ÚLTIMA ATUALIZAÇÃO
# ─────────────────────────────────────────────
Write-Section "ATUALIZAÇÕES"

try {
    $latestHotfix = Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1
    if ($latestHotfix.InstalledOn) {
        $formattedDate = $latestHotfix.InstalledOn.ToString("dd/MM/yyyy")
        Write-DiagnosticResult "ATUALIZAÇÃO" "Última (HotFix)" "$formattedDate — $($latestHotfix.HotFixID)"
    } else {
        Write-DiagnosticResult "ATUALIZAÇÃO" "Última (HotFix)" "Nenhuma atualização encontrada"
    }
} catch {
    Write-DiagnosticResult "ATUALIZAÇÃO" "ERRO HotFix" "Falha ao consultar histórico."
    Write-Log "ERRO HotFix: $_"
}

try {
    $updateSession  = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $historyCount   = $updateSearcher.GetTotalHistoryCount()

    if ($historyCount -gt 0) {
        $lastComUpdate = $updateSearcher.QueryHistory(0, $historyCount) |
            Where-Object { $_.ResultCode -eq 2 } |
            Sort-Object Date -Descending |
            Select-Object -First 1

        if ($lastComUpdate) {
            $comDate = $lastComUpdate.Date.ToString("dd/MM/yyyy")
            Write-DiagnosticResult "ATUALIZAÇÃO" "Última (COM)" "$comDate — $($lastComUpdate.Title)"
        }
    }
} catch {
    Write-DiagnosticResult "ATUALIZAÇÃO" "Última (COM)" "Não disponível para esta versão do Windows"
}

# ─────────────────────────────────────────────
#  6. COMPARAÇÃO OEM vs. CHAVE INSTALADA
#     e correção automática (se permitido)
# ─────────────────────────────────────────────
Write-Section "COMPARAÇÃO E CORREÇÃO DE CHAVE"

if ($oemKey -and $installedPartial) {

    $oemPartial = $oemKey.Split("-")[-1].ToUpper().Trim()
    $sysPartial = $installedPartial.ToUpper().Trim()

    Write-DiagnosticResult "COMPARAÇÃO" "Sufixo OEM (BIOS)"    $oemPartial
    Write-DiagnosticResult "COMPARAÇÃO" "Sufixo Instalado"     $sysPartial

    if ($oemPartial -eq $sysPartial) {
        Write-Host "`n✔ A chave instalada corresponde à chave OEM da BIOS. Nenhuma ação necessária." -ForegroundColor Green
        Write-Log "Chaves OEM e instalada são compatíveis. Nenhuma ação realizada."

    } else {
        Write-Host "`n⚠ Divergência detectada entre a chave OEM (BIOS) e a chave instalada." -ForegroundColor Yellow
        Write-Log "DIVERGÊNCIA: OEM=$oemPartial | Instalada=$sysPartial"

        if ($DiagnosticOnly) {
            Write-Host "Modo somente diagnóstico ativo. A correção NÃO será aplicada." -ForegroundColor DarkYellow
            Write-Log "DiagnosticOnly = true. Correção ignorada."
        } else {
            # Solicita confirmação do usuário
            Write-Host "Deseja substituir a chave atual pela OEM da BIOS? (S/N): " -ForegroundColor Yellow -NoNewline
            $resp = Read-Host
            if ($resp -ne 'S' -and $resp -ne 's') {
                Write-Host "Correção cancelada pelo usuário." -ForegroundColor Yellow
                Write-Log "Correção cancelada pelo usuário."
            } else {
                # 6a. Copiar chave OEM para o clipboard
                try {
                    Set-Clipboard -Value $oemKey
                    Write-Host "✔ Chave OEM copiada para o clipboard." -ForegroundColor Green
                    Write-Log "Chave OEM copiada para o clipboard."
                } catch {
                    Write-Host "⚠ Não foi possível copiar para o clipboard: $_" -ForegroundColor Yellow
                    Write-Log "AVISO clipboard: $_"
                }

                # 6b. Remover a chave atual (slmgr /upk)
                Write-Host "  Removendo a chave instalada atual..." -ForegroundColor Yellow
                try {
                    $r = Invoke-Slmgr "/upk"
                    Write-DiagnosticResult "CORREÇÃO" "Remoção da chave atual" $r
                } catch {
                    Write-Host "ERRO ao remover a chave atual: $_" -ForegroundColor Red
                    Write-Log "ERRO /upk: $_"
                    exit 1
                }

                # 6c. Limpar a chave do registro (slmgr /cpky)
                Write-Host "  Limpando chave do registro..." -ForegroundColor Yellow
                try {
                    $r = Invoke-Slmgr "/cpky"
                    Write-DiagnosticResult "CORREÇÃO" "Limpeza do registro" $r
                } catch {
                    Write-Log "AVISO /cpky: $_"
                }

                # 6d. Instalar a chave OEM da BIOS (via CIM/WMI)
                Write-Host "  Instalando chave OEM da BIOS..." -ForegroundColor Yellow
                try {
                    $svc = Get-CimInstance -ClassName SoftwareLicensingService
                    $svc.InstallProductKey($oemKey) | Out-Null
                    $svc.RefreshLicenseStatus() | Out-Null
                    Write-DiagnosticResult "CORREÇÃO" "Instalação da chave OEM" "Sucesso (via CIM)"
                    Write-Log "Chave OEM instalada com sucesso via CIM."
                } catch {
                    Write-Host "ERRO ao instalar chave via CIM. Tentando slmgr..." -ForegroundColor Yellow
                    try {
                        $r = Invoke-Slmgr "/ipk", $oemKey
                        Write-DiagnosticResult "CORREÇÃO" "Instalação da chave OEM (slmgr)" $r
                    } catch {
                        Write-Host "ERRO ao instalar a chave OEM: $_" -ForegroundColor Red
                        Write-Log "ERRO /ipk: $_"
                        exit 1
                    }
                }

                # 6e. Ativar o Windows (slmgr /ato)
                Write-Host "  Ativando o Windows..." -ForegroundColor Yellow
                try {
                    $r = Invoke-Slmgr "/ato"
                    Write-DiagnosticResult "CORREÇÃO" "Ativação" $r
                    Write-Log "Ativação executada: $r"
                } catch {
                    Write-Host "ERRO durante a ativação: $_" -ForegroundColor Red
                    Write-Log "ERRO /ato: $_"
                }

                # 6f. Aguardar e verificar novo status
                Write-Host "  Aguardando atualização do status de ativação..." -ForegroundColor Yellow
                Start-Sleep -Seconds 3
                Write-Host "`n  Verificando novo status de ativação..." -ForegroundColor Yellow
                try {
                    $dli = Invoke-Slmgr "/dli"
                    Write-Host "`n$dli" -ForegroundColor Gray
                    Write-Log "Status pós-correção: $dli"
                } catch {
                    Write-Log "ERRO /dli: $_"
                }
            }
        }
    }
} elseif ($oemKey -and -not $installedPartial) {
    Write-Host "`n⚠ Chave OEM encontrada na BIOS, mas nenhuma chave está instalada no sistema." -ForegroundColor Yellow
    Write-Log "OEM presente, nenhuma chave instalada. Intervenção manual recomendada."
} elseif (-not $oemKey -and $installedPartial) {
    Write-Host "`n⚙ Nenhuma chave OEM na BIOS. Chave instalada mantida sem alterações." -ForegroundColor Cyan
    Write-Log "Sem chave OEM. Chave instalada preservada."
} else {
    Write-Host "`n⚠ Nenhuma chave encontrada (BIOS ou sistema). Ativação manual necessária." -ForegroundColor Red
    Write-Log "Nenhuma chave detectada em nenhuma fonte."
}

# ─────────────────────────────────────────────
#  7. RESUMO FINAL
# ─────────────────────────────────────────────
Write-Host "`n===== RESUMO DO DIAGNÓSTICO =====" -ForegroundColor Yellow
Write-Log "===== Resumo final ====="

$finalLic = Get-CimInstance -ClassName SoftwareLicensingProduct |
    Where-Object { $_.PartialProductKey -and $_.ApplicationID -eq $windowsAppID }
if (-not $finalLic) {
    $finalLic = Get-CimInstance -ClassName SoftwareLicensingProduct |
        Where-Object { $_.PartialProductKey -and $_.Name -like 'Windows*' }
}

if ($finalLic.LicenseStatus -eq 1) {
    Write-Host "✔ O Windows está ATIVADO." -ForegroundColor Green
    Write-Log "Status final: ATIVADO"
} else {
    Write-Host "⚠ O Windows NÃO está ativado. Execute 'slmgr /dlv' para detalhes." -ForegroundColor Red
    Write-Log "Status final: NÃO ATIVADO"
}

if ($oemKey) {
    Write-Host "✔ Chave OEM da BIOS: presente e processada." -ForegroundColor Green
    if (-not $DiagnosticOnly) {
        Write-Host "   A chave está disponível no clipboard para uso manual, se necessário." -ForegroundColor Gray
    }
} else {
    Write-Host "⚙ Nenhuma chave OEM da BIOS detectada (comum em PCs montados)." -ForegroundColor Yellow
    Write-Host "   A ativação pode depender de licença digital vinculada à conta Microsoft." -ForegroundColor Gray
}

Write-Host "`nLog completo salvo em: $logFile" -ForegroundColor DarkCyan
Write-Host "Diagnóstico concluído.`n" -ForegroundColor Cyan
Write-Log "===== Fim do diagnóstico ====="
