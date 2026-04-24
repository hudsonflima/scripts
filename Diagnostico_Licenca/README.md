# 📋 Diagnóstico de Licenciamento do Windows

**Script:** `Diagnostico_Licenca_2.ps1`  
**Plataforma:** Windows 10 / 11  
**Requere:** Execução como Administrador  
**Versão:** 2.0  
**Autor:** Hudson Lima

---

## 🔍 Visão Geral

O `Diagnostico_Licenca_2.ps1` é um script PowerShell que realiza um diagnóstico completo do estado de licenciamento do Windows. Ele identifica chaves OEM gravadas na BIOS, verifica se a chave instalada é genérica (GVLK), compara a chave OEM com a instalada e pode corrigir automaticamente divergências.  

O relatório abrange:

- Chave de produto OEM (BIOS) e chave instalada.
- Canal de licenciamento (OEM, Retail, Volume) e status de ativação.
- Detecção de chave genérica (GVLK) por sufixo e canal.
- Edição, versão e build do Windows.
- Data da última atualização instalada (formato dd/mm/yyyy).
- Comparação entre a chave OEM e a chave instalada, com correção opcional.

---

## 📦 Requisitos

- Windows 10 ou 11.
- PowerShell 5.1 ou superior.
- Privilégios de **Administrador** (necessário para leitura da BIOS e manipulação de chaves).
- Conexão com a Internet (opcional; apenas se a ativação online for necessária).

---

## 🚀 Como Usar

1. **Salve o script** em uma pasta local, por exemplo, `C:\Scripts\Diagnostico_Licenca_2.ps1`.
2. **Abra o PowerShell como Administrador**:
   - Clique com o botão direito no Menu Iniciar e selecione **Terminal (Admin)** ou **Windows PowerShell (Admin)**.
3. **Execute o script**:
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
   C:\Scripts\Diagnostico_Licenca_2.ps1
