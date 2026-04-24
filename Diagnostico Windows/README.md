# 📋 Diagnóstico de Licenciamento do Windows

**Script:** `Diagnostico_Windows.ps1`  
**Plataforma:** Windows 10/11  
**Requer:** Execução como Administrador  
**Versão:** 1.0

---

## 🔍 Objetivo

Fornecer um relatório completo sobre o estado de licenciamento, versão do Windows e histórico de atualizações, permitindo identificar rapidamente:

- A chave OEM gravada na BIOS/firmware.
- O canal e tipo de licença instalada (OEM, Retail, Volume, Digital).
- A edição exata do Windows, build e versão.
- A data da última atualização aplicada ao sistema.

## 📦 Requisitos

- Windows 10 ou 11.
- PowerShell 5.1 ou superior.
- Privilégios de **Administrador** (necessário para leitura da BIOS e chaves de registro protegidas).

## 🚀 Como utilizar

1. Salve o script em uma pasta local, por exemplo `C:\Scripts\Diagnostico_Windows.ps1`.
2. Abra o **PowerShell como Administrador** (clique com botão direito no Menu Iniciar → Terminal Admin).
3. Execute os comandos:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
C:\Scripts\Diagnostico_Windows.ps1
