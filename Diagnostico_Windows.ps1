<#
.SYNOPSIS
    Diagnóstico completo de licenciamento, versão e atualizações do Windows.
.DESCRIPTION
    Este script gera um relatório detalhado que inclui:
    - Chave de produto OEM (BIOS) e chave instalada.
    - Canal de licenciamento (OEM, Retail, Volume).
    - Status de ativação do Windows.
    - Edição, versão e compilação do Windows.
    - Data da última atualização instalada (dd/mm/yyyy).
.NOTES
    Autor: DeepSeek AI
    Requer: Execução como Administrador.
    Compatibilidade: Windows 10/11
#>

# Função para formatar a saída no console
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
}

Clear-Host
Write-Host "===== DIAGNÓSTICO DE LICENCIAMENTO DO WINDOWS =====" -ForegroundColor Yellow
Write-Host "Iniciando análise detalhada do sistema...`n" -ForegroundColor Gray

#------------------------------------------------------------
# 1. INFORMAÇÕES DA BIOS/FIRMWARE
#------------------------------------------------------------
Write-DiagnosticResult "BIOS" "Consultando chave OEM..." ""

try {
    $oemKey = (Get-CimInstance -Query 'SELECT * FROM SoftwareLicensingService').OA3xOriginalProductKey

    if ($oemKey) {
        Write-DiagnosticResult "BIOS" "Chave OEM (BIOS)" $oemKey
    } else {
        Write-DiagnosticResult "BIOS" "Chave OEM (BIOS)" "Nenhuma chave OEM encontrada na BIOS/UEFI"
    }
} catch {
    Write-DiagnosticResult "BIOS" "Erro ao consultar" "Não foi possível acessar o firmware. Execute como Admin."
}

#------------------------------------------------------------
# 2. CANAL DE LICENCIAMENTO E EDIÇÃO DO WINDOWS
#------------------------------------------------------------
$licensingData = Get-CimInstance -ClassName SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.Name -like 'Windows*' }

if ($licensingData) {
    $licenseStatus = switch ($licensingData.LicenseStatus) {
        0 { "Não Licenciado" }
        1 { "Licenciado (Ativado)" }
        2 { "Período de avaliação" }
        3 { "Período de avaliação expirado" }
        4 { "Notificação de avaliação" }
        default { "Desconhecido ($_)" }
    }
    Write-DiagnosticResult "SISTEMA" "Status da Licença" $licenseStatus
    Write-DiagnosticResult "SISTEMA" "Canal de Licença" $licensingData.ProductKeyChannel
    
    # Identificação do tipo de licença
    $licenseType = switch -wildcard ($licensingData.ProductKeyChannel) {
        "*OEM*"   { "OEM (Original Equipment Manufacturer) - Vinculado ao hardware" }
        "*Retail*" { "Retail (Varejo) - Chave adquirida separadamente" }
        "*Volume*" { "Volume (GVLK/MAK) - Licenciamento corporativo" }
        default    { "Outro (Provavelmente Digital License)" }
    }
    Write-DiagnosticResult "SISTEMA" "Tipo de Licença" $licenseType
    Write-DiagnosticResult "SISTEMA" "Chave Parcial Instalada" $licensingData.PartialProductKey

    # Chave de produto completa do registro (apenas para licenças instaladas)
    try {
        $regKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
        $backupKey = Get-ItemPropertyValue -Path $regKeyPath -Name "BackupProductKeyDefault" -ErrorAction Stop
        if ($backupKey) {
            Write-DiagnosticResult "SISTEMA" "Chave Instalada (Backup)" $backupKey
        }
    } catch {
        Write-DiagnosticResult "SISTEMA" "Chave Instalada (Backup)" "Não foi possível ler do registro"
    }
} else {
    Write-DiagnosticResult "SISTEMA" "ERRO" "Nenhum produto de licenciamento do Windows encontrado."
}

#------------------------------------------------------------
# 3. EDIÇÃO, VERSÃO E BUILD DO WINDOWS
#------------------------------------------------------------
$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
$edition = $osInfo.Caption
$version = $osInfo.Version
$build = $osInfo.BuildNumber

Write-DiagnosticResult "VERSÃO" "Edição do Windows" $edition
Write-DiagnosticResult "VERSÃO" "Número da Versão" $version
Write-DiagnosticResult "VERSÃO" "Build do Sistema" $build

#------------------------------------------------------------
# 4. DATA DA ÚLTIMA ATUALIZAÇÃO (DD/MM/YYYY)
#------------------------------------------------------------
try {
    # Método 1: Get-HotFix (para atualizações tradicionais)
    $latestHotfix = Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1
    $lastUpdate = $latestHotfix.InstalledOn
    
    if ($lastUpdate) {
        $formattedDate = $lastUpdate.ToString("dd/MM/yyyy")
        Write-DiagnosticResult "ATUALIZAÇÃO" "Última Instalada" "$formattedDate ($($latestHotfix.HotFixID))"
    } else {
        Write-DiagnosticResult "ATUALIZAÇÃO" "Última Instalada" "Nenhuma atualização encontrada via Get-HotFix"
    }

    # Método 2: COM Object (mais preciso em versões recentes)
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
                Write-DiagnosticResult "ATUALIZAÇÃO" "Última (COM)" "$comDate ($($lastComUpdate.Title))"
            }
        }
    } catch {
        Write-DiagnosticResult "ATUALIZAÇÃO" "Detalhes COM" "Não disponível para esta versão do Windows"
    }
} catch {
    Write-DiagnosticResult "ATUALIZAÇÃO" "ERRO" "Falha ao consultar histórico de atualizações."
}

#------------------------------------------------------------
# 5. RESUMO FINAL
#------------------------------------------------------------
Write-Host "`n===== RESUMO DO DIAGNÓSTICO =====" -ForegroundColor Yellow
if ($licensingData.LicenseStatus -eq 1) {
    Write-Host "✔ O Windows está ATIVADO." -ForegroundColor Green
} else {
    Write-Host "⚠ O Windows NÃO está ativado." -ForegroundColor Red
}

if ($oemKey) {
    Write-Host "✔ Chave OEM da BIOS: Presente e recuperada com sucesso." -ForegroundColor Green
    Write-Host "   Essa chave é a identidade original do seu dispositivo." -ForegroundColor Gray
} else {
    Write-Host "⚙ Nenhuma chave OEM da BIOS detectada (comum em PCs montados)." -ForegroundColor Yellow
    Write-Host "   A ativação pode depender de uma licença digital vinculada à sua conta." -ForegroundColor Gray
}

Write-Host "`nDiagnóstico concluído. Para solucionar problemas de ativação, execute 'slmgr /dlv'." -ForegroundColor Cyan
