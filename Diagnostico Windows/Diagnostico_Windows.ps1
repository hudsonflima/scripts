<#
.SYNOPSIS
    Diagnóstico completo de licenciamento, versão e atualizações do Windows.
.DESCRIPTION
    Este script gera um relatório detalhado que inclui:
    - Chave de produto OEM (BIOS) e chave instalada.
    - Canal de licenciamento (OEM, Retail, Volume).
    - Status de ativação do Windows.
    - Detecção de chave genérica (GVLK).
    - Edição, versão e compilação do Windows.
    - Data da última atualização instalada (dd/mm/yyyy).
.NOTES
    Autor: Hudson Lima
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
Write-Host "===== DIAGNOSTICO DE LICENCIAMENTO DO WINDOWS =====" -ForegroundColor Yellow
Write-Host "Iniciando analise detalhada do sistema..." -ForegroundColor Gray
Write-Host ""

#------------------------------------------------------------
# 1. INFORMACOES DA BIOS/FIRMWARE
#------------------------------------------------------------
Write-DiagnosticResult "BIOS" "Consultando chave OEM..." ""

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
    Write-DiagnosticResult "BIOS" "Erro ao consultar" "Nao foi possivel acessar o firmware. Execute como Admin."
}

#------------------------------------------------------------
# 2. CANAL DE LICENCIAMENTO E DETECCAO DE CHAVE GENERICA
#------------------------------------------------------------
$licensingData = Get-CimInstance -ClassName SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.Name -like 'Windows*' }

if ($licensingData) {
    $licenseStatus = switch ($licensingData.LicenseStatus) {
        0 { "Nao Licenciado" }
        1 { "Licenciado (Ativado)" }
        2 { "Periodo de avaliacao" }
        3 { "Periodo de avaliacao expirado" }
        4 { "Notificacao de avaliacao" }
        default { "Desconhecido ($($licensingData.LicenseStatus))" }
    }
    Write-DiagnosticResult "SISTEMA" "Status da Licenca" $licenseStatus
    Write-DiagnosticResult "SISTEMA" "Canal de Licenca" $licensingData.ProductKeyChannel
    
    # Identificacao do tipo de licenca
    $licenseType = switch -wildcard ($licensingData.ProductKeyChannel) {
        "*OEM*" { "OEM (Original Equipment Manufacturer) - Vinculado ao hardware" }
        "*Retail*" { "Retail (Varejo) - Chave adquirida separadamente" }
        "*Volume*" { "Volume (GVLK/MAK) - Licenciamento corporativo" }
        default { "Outro (Provavelmente Digital License)" }
    }
    Write-DiagnosticResult "SISTEMA" "Tipo de Licenca" $licenseType
    Write-DiagnosticResult "SISTEMA" "Chave Parcial Instalada" $licensingData.PartialProductKey

    # Deteccao de chave generica (GVLK)
    $genericSuffixes = @(
        "8HVX7",  # Home
        "6F4BT",  # Home Single Language
        "3V66T",  # Pro
        "7CFBY",  # Education
        "GVGXT",  # Pro Education
        "DPH2V",  # Enterprise
        "W269N",  # Pro (alternativa)
        "YTMG3",  # Home (outra)
        "BT79Q",  # Home Single Language (outra)
        "VK7JG"   # Pro (outra)
    )

    $partial = $licensingData.PartialProductKey
    $canal = $licensingData.ProductKeyChannel

    $isGeneric = ($genericSuffixes -contains $partial) -or ($canal -like "*Volume:GVLK*")
    
    if ($isGeneric) {
        Write-DiagnosticResult "SISTEMA" "Chave Generica" "SIM - Esta e uma chave de instalacao padrao (GVLK). Nao ativa o Windows."
    }
    else {
        Write-DiagnosticResult "SISTEMA" "Chave Generica" "NAO - Chave aparenta ser original/ativacao digital."
    }

    # Chave de produto completa do registro (apenas para licencas instaladas)
    try {
        $regKeyPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
        $backupKey = Get-ItemPropertyValue -Path $regKeyPath -Name "BackupProductKeyDefault" -ErrorAction Stop
        if ($backupKey) {
            Write-DiagnosticResult "SISTEMA" "Chave Instalada (Backup)" $backupKey
        }
    }
    catch {
        Write-DiagnosticResult "SISTEMA" "Chave Instalada (Backup)" "Nao foi possivel ler do registro"
    }
}
else {
    Write-DiagnosticResult "SISTEMA" "ERRO" "Nenhum produto de licenciamento do Windows encontrado."
}

#------------------------------------------------------------
# 3. EDICAO, VERSAO E BUILD DO WINDOWS
#------------------------------------------------------------
$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
$edition = $osInfo.Caption
$version = $osInfo.Version
$build = $osInfo.BuildNumber

Write-DiagnosticResult "VERSAO" "Edicao do Windows" $edition
Write-DiagnosticResult "VERSAO" "Numero da Versao" $version
Write-DiagnosticResult "VERSAO" "Build do Sistema" $build

#------------------------------------------------------------
# 4. DATA DA ULTIMA ATUALIZACAO (DD/MM/YYYY)
#------------------------------------------------------------
try {
    # Metodo 1: Get-HotFix (para atualizacoes tradicionais)
    $latestHotfix = Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1
    $lastUpdate = $latestHotfix.InstalledOn
    
    if ($lastUpdate) {
        $formattedDate = $lastUpdate.ToString("dd/MM/yyyy")
        Write-DiagnosticResult "ATUALIZACAO" "Ultima Instalada" "$formattedDate ($($latestHotfix.HotFixID))"
    }
    else {
        Write-DiagnosticResult "ATUALIZACAO" "Ultima Instalada" "Nenhuma atualizacao encontrada via Get-HotFix"
    }

    # Metodo 2: COM Object (mais preciso em versoes recentes)
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
                Write-DiagnosticResult "ATUALIZACAO" "Ultima (COM)" "$comDate ($($lastComUpdate.Title))"
            }
        }
    }
    catch {
        Write-DiagnosticResult "ATUALIZACAO" "Detalhes COM" "Nao disponivel para esta versao do Windows"
    }
}
catch {
    Write-DiagnosticResult "ATUALIZACAO" "ERRO" "Falha ao consultar historico de atualizacoes."
}

#------------------------------------------------------------
# 5. RESUMO FINAL
#------------------------------------------------------------
Write-Host ""
Write-Host "===== RESUMO DO DIAGNOSTICO =====" -ForegroundColor Yellow

if ($licensingData.LicenseStatus -eq 1) {
    Write-Host "[OK] O Windows esta ATIVADO." -ForegroundColor Green
}
else {
    Write-Host "[ALERTA] O Windows NAO esta ativado." -ForegroundColor Red
}

if ($oemKey) {
    Write-Host "[OK] Chave OEM da BIOS: Presente e recuperada com sucesso." -ForegroundColor Green
    Write-Host "   Essa chave e a identidade original do seu dispositivo." -ForegroundColor Gray
}
else {
    Write-Host "[INFO] Nenhuma chave OEM da BIOS detectada (comum em PCs montados)." -ForegroundColor Yellow
    Write-Host "   A ativacao pode depender de uma licenca digital vinculada a sua conta." -ForegroundColor Gray
}

Write-Host ""
Write-Host "Diagnostico concluido. Para solucionar problemas de ativacao, execute 'slmgr /dlv'." -ForegroundColor Cyan