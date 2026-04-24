<#
.SYNOPSIS
    Remove a chave Pro generica, detecta o tipo de licenca (OEM ou RTM)
    e realiza a ativacao do Windows 11 Home Single Language.
.DESCRIPTION
    O script executa as seguintes etapas:
    1. Verifica privilegios de Administrador.
    2. Remove a chave de produto Pro generica ativa.
    3. Limpa a chave do registro do Windows.
    4. Verifica se existe chave OEM gravada na BIOS/UEFI.
       - OEM: instala e ativa automaticamente.
       - RTM/Retail ou ausente: solicita a chave ao usuario.
    5. Realiza a ativacao e exibe o status final.
.NOTES
    Execute o PowerShell como Administrador.
    Pode ser necessaria reinicializacao para aplicar a troca de edicao.
#>

# -----------------------------------------------------------
#  Configuracao de log
# -----------------------------------------------------------
$ScriptName = "RemoveGenericKey"
$logDir = Join-Path $PSScriptRoot "logs"
$logFile = Join-Path $logDir "$ScriptName.log"

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$timestamp [$ScriptName] $Message"
    Add-Content -Path $logFile -Value $line -Encoding UTF8
    Write-Host $line -ForegroundColor $Color
}

function Invoke-Slmgr {
    param([string[]]$Arguments)
    $output = & cscript //B //Nologo "$env:SystemRoot\System32\slmgr.vbs" @Arguments 2>&1
    return ($output -join "`n").Trim()
}

# -----------------------------------------------------------
#  1. Verificacao de privilegios
# -----------------------------------------------------------
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "ERRO: Este script deve ser executado como Administrador." "Red"
    exit 1
}

Write-Log "===== Iniciando processo de downgrade para Windows 11 Home =====" "Cyan"

# -----------------------------------------------------------
#  2. Remover a chave Pro generica atual
# -----------------------------------------------------------
Write-Log "Removendo a chave de produto Pro atual..." "Yellow"
try {
    $result = Invoke-Slmgr "/upk"
    Write-Log "Chave removida: $result" "Green"
}
catch {
    Write-Log "ERRO ao remover chave: $_" "Red"
    exit 1
}

# -----------------------------------------------------------
#  3. Limpar a chave armazenada no registro
# -----------------------------------------------------------
Write-Log "Limpando chave do registro..." "Yellow"
try {
    $result = Invoke-Slmgr "/cpky"
    Write-Log "Registro limpo: $result" "Green"
}
catch {
    Write-Log "Aviso ao limpar registro: $_" "Yellow"
}

# -----------------------------------------------------------
#  4. Deteccao do tipo de licenca (OEM ou RTM)
# -----------------------------------------------------------
Write-Log "Verificando tipo de licenca e chave na BIOS/UEFI..." "Yellow"

$prodKey = $null
$tipoLic = $null

try {
    # Obtem a chave OEM embarcada na BIOS/UEFI (OA3)
    $sls = Get-WmiObject -Query "SELECT * FROM SoftwareLicensingService"
    $oemKey = $sls.OA3xOriginalProductKey

    # Obtem o canal da licenca ativa (ex: "OEM:NONSLP", "OEM:DM", "Retail", "Volume:GVLK")
    $slp = Get-WmiObject -Query @"
        SELECT ProductKeyChannel FROM SoftwareLicensingProduct
        WHERE ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f'
          AND LicenseStatus=1
"@
    $canal = if ($slp) { $slp.ProductKeyChannel } else { $null }

    Write-Log "Canal de licenca detectado: $(if ($canal) { $canal } else { 'N/A' })" "Gray"

    # -- Logica de decisao --
    # Eh OEM se: a chave existe na BIOS E o canal comeca com "OEM"
    if ($oemKey -and $canal -like "OEM*") {
        $tipoLic = "OEM"
        $prodKey = $oemKey
        Write-Log "Licenca OEM confirmada. Chave BIOS sera utilizada." "Green"

    }
    elseif ($oemKey -and -not $canal) {
        # Chave na BIOS mas canal nao identificado -- assume OEM
        $tipoLic = "OEM"
        $prodKey = $oemKey
        Write-Log "Chave OEM encontrada na BIOS (canal nao identificado). Sera utilizada." "Yellow"

    }
    else {
        # Canal Retail, Volume ou nenhuma chave na BIOS -> RTM
        $tipoLic = "RTM"
        Write-Log "Licenca RTM/Retail detectada ou chave OEM ausente na BIOS." "Yellow"
    }

}
catch {
    Write-Log "ERRO ao consultar WMI: $_" "Red"
    $tipoLic = "RTM"   # Falha segura: solicitar chave ao usuario
}

# -----------------------------------------------------------
#  5. Obter chave: automatico (OEM) ou manual (RTM)
# -----------------------------------------------------------
function Request-ProductKey {
    Write-Host "`nDigite a chave do produto Home Single Language" -ForegroundColor White
    Write-Host "Formato: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX" -ForegroundColor Gray
    $pattern = '^[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}$'
    do {
        $key = (Read-Host "Chave").ToUpper().Trim()
        if ($key -notmatch $pattern) {
            Write-Host "Formato invalido. Tente novamente." -ForegroundColor Red
        }
    } while ($key -notmatch $pattern)
    return $key
}

if ($tipoLic -eq "OEM") {
    Write-Log "Utilizando chave OEM da BIOS/UEFI automaticamente." "Green"
    # $prodKey ja definido acima

}
else {
    Write-Log "Tipo RTM: solicitando chave ao usuario." "Yellow"
    Write-Host "`nNenhuma chave OEM valida foi encontrada." -ForegroundColor Yellow
    Write-Host "Insira a chave do produto Windows 11 Home Single Language:" -ForegroundColor White
    $prodKey = Request-ProductKey
    Write-Log "Chave inserida manualmente pelo usuario." "Gray"
}

# -----------------------------------------------------------
#  6. Instalar a chave de produto
# -----------------------------------------------------------
Write-Log "Instalando a chave de produto..." "Yellow"
try {
    $result = Invoke-Slmgr "/ipk", $prodKey
    Write-Log "Chave instalada: $result" "Green"
}
catch {
    Write-Log "ERRO ao instalar chave: $_" "Red"
    exit 1
}

# -----------------------------------------------------------
#  7. Ativar o Windows
# -----------------------------------------------------------
Write-Log "Ativando o Windows junto aos servidores Microsoft..." "Yellow"
try {
    $result = Invoke-Slmgr "/ato"
    Write-Log "Ativacao concluida: $result" "Green"
}
catch {
    Write-Log "ERRO durante a ativacao: $_" "Red"
}

# -----------------------------------------------------------
#  8. Exibir status final de ativacao
# -----------------------------------------------------------
Write-Log "Verificando status final de ativacao..." "Yellow"
try {
    $dli = Invoke-Slmgr "/dli"
    Write-Log "--- Status de Ativacao ---" "Cyan"
    $dli -split "`n" | ForEach-Object {
        $linha = $_.Trim()
        if ($linha) { Write-Log $linha "Gray" }
    }
    Write-Log "--------------------------" "Cyan"
}
catch {
    Write-Log "ERRO ao verificar status: $_" "Red"
}

Write-Log "===== Processo concluido =====" "Cyan"
Write-Log "Se a edicao foi alterada, reinicie o computador para aplicar todas as mudancas." "Yellow"