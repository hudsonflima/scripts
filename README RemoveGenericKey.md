# RemoveGenericKey.ps1 — Documentação

Script PowerShell para remover a chave de ativação Pro genérica, detectar o tipo de licença (OEM ou RTM) e realizar a ativação do Windows 11 Home Single Language.

---

## Pré-requisitos

- Windows 11 com PowerShell 5.1 ou superior
- Execução **obrigatória como Administrador**
- Conexão com a internet para ativação online
- Chave de produto Home Single Language válida (somente se RTM)

---

## Como executar

1. Clique com o botão direito sobre o arquivo `RemoveGenericKey.ps1`
2. Selecione **"Executar com o PowerShell como Administrador"**

Ou via terminal:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\RemoveGenericKey.ps1
```

---

## Fluxo de execução

```
Início
  │
  ├─ [1] Verifica privilégios de Administrador
  │       └─ Não é Admin → encerra com erro (exit 1)
  │
  ├─ [2] Remove a chave Pro genérica (slmgr /upk)
  │       └─ Falha → encerra com erro (exit 1)
  │
  ├─ [3] Limpa a chave do registro (slmgr /cpky)
  │
  ├─ [4] Consulta WMI: chave OEM na BIOS + canal da licença
  │       ├─ Chave na BIOS + canal "OEM*" → tipo = OEM ✅
  │       ├─ Chave na BIOS + canal indefinido → assume OEM ⚠️
  │       └─ Sem chave / canal Retail ou Volume → tipo = RTM 🔑
  │
  ├─ [5] Obtém a chave de produto
  │       ├─ OEM → usa chave da BIOS automaticamente
  │       └─ RTM → solicita entrada manual do usuário
  │
  ├─ [6] Instala a chave (slmgr /ipk)
  │       └─ Falha → encerra com erro (exit 1)
  │
  ├─ [7] Ativa o Windows (slmgr /ato)
  │
  └─ [8] Exibe e registra o status de ativação (slmgr /dli)
```

---

## Detecção OEM vs RTM

A distinção entre os tipos de licença é feita consultando dois valores via WMI:

| Propriedade WMI | Classe | Descrição |
|---|---|---|
| `OA3xOriginalProductKey` | `SoftwareLicensingService` | Chave gravada pelo fabricante na BIOS/UEFI |
| `ProductKeyChannel` | `SoftwareLicensingProduct` | Canal da licença ativa no sistema |

### Possíveis valores de `ProductKeyChannel`

| Valor | Tipo |
|---|---|
| `OEM:NONSLP` | OEM padrão |
| `OEM:DM` | OEM com ativação digital |
| `Retail` | RTM (varejo) |
| `Volume:GVLK` | Volume (KMS) |
| `Volume:MAK` | Volume (MAK) |

### Tabela de decisão

| Chave na BIOS | Canal detectado | Resultado |
|---|---|---|
| ✅ Presente | `OEM*` | Usa chave da BIOS automaticamente |
| ✅ Presente | Indefinido / `null` | Assume OEM, usa chave da BIOS |
| ❌ Ausente | Qualquer | Classifica como RTM, solicita chave ao usuário |
| ✅ Presente | `Retail` ou `Volume*` | Classifica como RTM, solicita chave ao usuário |
| Erro WMI | — | Falha segura: classifica como RTM |

---

## Funções internas

### `Write-Log`

Registra mensagens no console e no arquivo de log com timestamp.

```powershell
Write-Log -Message "Texto da mensagem" -Color "Green"
```

| Parâmetro | Tipo | Padrão | Descrição |
|---|---|---|---|
| `Message` | `string` | — | Texto a ser registrado |
| `Color` | `string` | `"White"` | Cor de exibição no console |

---

### `Invoke-Slmgr`

Centraliza as chamadas ao `slmgr.vbs`, retornando a saída como string.

```powershell
$resultado = Invoke-Slmgr "/upk"
$resultado = Invoke-Slmgr "/ipk", "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"
```

| Parâmetro | Tipo | Descrição |
|---|---|---|
| `Args` | `string[]` | Argumentos passados ao `slmgr.vbs` |

---

### `Request-ProductKey`

Solicita a chave de produto ao usuário via terminal, validando o formato antes de aceitar.

- Converte automaticamente a entrada para maiúsculas
- Rejeita qualquer entrada fora do padrão `XXXXX-XXXXX-XXXXX-XXXXX-XXXXX`
- Repete o prompt até receber uma entrada válida

---

## Comandos `slmgr.vbs` utilizados

| Comando | Ação |
|---|---|
| `slmgr /upk` | Remove a chave de produto instalada |
| `slmgr /cpky` | Apaga a chave armazenada no registro |
| `slmgr /ipk <chave>` | Instala uma nova chave de produto |
| `slmgr /ato` | Realiza a ativação online junto à Microsoft |
| `slmgr /dli` | Exibe informações resumidas da licença ativa |

---

## Registro de log

Todas as operações são registradas em:

```
<diretório do script>\logs\RemoveGenericKey.log
```

O arquivo é criado automaticamente caso não exista. Cada entrada segue o formato:

```
yyyy-MM-dd HH:mm:ss [RemoveGenericKey] Mensagem
```

O log é salvo em **UTF-8** para preservar acentuação e caracteres especiais.

---

## Tratamento de erros

| Etapa | Comportamento em caso de falha |
|---|---|
| Verificação de Administrador | Exibe erro e encerra com `exit 1` |
| Remoção da chave Pro | Exibe erro e encerra com `exit 1` |
| Limpeza do registro | Registra aviso e continua |
| Consulta WMI | Registra erro e assume RTM (falha segura) |
| Instalação da chave | Exibe erro e encerra com `exit 1` |
| Ativação online | Registra erro e continua |
| Verificação de status | Registra erro e continua |

---

## Observações importantes

> ⚠️ **Este script não reinstala o sistema operacional.** A troca de edição é feita exclusivamente via licenciamento de software. Em alguns casos, uma reinstalação completa pode ser necessária para migrar de Pro para Home de fato.

> 🔄 **Reinicialização recomendada** após a execução para que todas as mudanças de edição sejam aplicadas corretamente.

> 🔑 **Guarde a chave RTM** em local seguro antes de executar o script. Após a remoção da chave Pro, não será possível recuperá-la automaticamente.
