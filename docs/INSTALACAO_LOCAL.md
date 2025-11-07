# InstalaÃ§Ã£o Local - Chainsaw

## [INFO] MudanÃ§a Importante

O sistema Chainsaw agora funciona a partir da **pasta Documentos do usuÃ¡rio**, eliminando a necessidade de acesso Ã  rede corporativa durante a instalaÃ§Ã£o.

## [*] BenefÃ­cios

### Antes (Rede)
- [X] Dependia de acesso Ã  rede corporativa
- [X] Problemas com VPN e credenciais
- [X] LentidÃ£o na cÃ³pia de arquivos
- [X] Falhas por desconexÃ£o de rede

### Agora (Local)
- [OK] Funciona offline
- [OK] InstalaÃ§Ã£o mais rÃ¡pida
- [OK] Mais confiÃ¡vel
- [OK] Sem dependÃªncia de rede

## [PKG] Estrutura NecessÃ¡ria

A pasta `chainsaw` deve ser copiada para a pasta Documentos do usuÃ¡rio:

```
%USERPROFILE%\chainsaw\
â”œâ”€â”€ assets\
â”‚   â””â”€â”€ stamp.png
â”œâ”€â”€ configs\
â”‚   â””â”€â”€ Templates\
â”‚       â””â”€â”€ [todos os templates]
â”œâ”€â”€ install.ps1
â”œâ”€â”€ install.cmd
â””â”€â”€ [outros arquivos]
```

## [>>] InstalaÃ§Ã£o

### 1. Copiar Arquivos

Primeiro, copie a pasta completa `chainsaw` para:
- **Windows**: `C:\Users\[seu_usuario]\chainsaw`

### 2. Executar InstalaÃ§Ã£o

```cmd
cd "%USERPROFILE%\chainsaw"
install.cmd
```

Ou usando PowerShell:

```powershell
cd "$env:USERPROFILE\chainsaw"
.\install.ps1
```

## [CFG] Como Funciona

### DetecÃ§Ã£o AutomÃ¡tica de Origem

O script agora detecta automaticamente de onde estÃ¡ sendo executado:

```powershell
# O caminho de origem Ã© automaticamente definido como a pasta do script
$SourcePath = $PSScriptRoot
```

### VerificaÃ§Ã£o de Auto-CÃ³pia

Para evitar erros quando executado diretamente da pasta de destino, o script:

1. **Verifica se origem = destino** para `stamp.png`
   - Se sim, pula a cÃ³pia (jÃ¡ estÃ¡ instalado)
   - Se nÃ£o, copia normalmente

2. **Verifica se origem = destino** para `Templates`
   - Se sim, pula a cÃ³pia (jÃ¡ estÃ¡ instalado)
   - Se nÃ£o, copia normalmente

## [CHART] Exemplo de ExecuÃ§Ã£o

```
[SEC] Verificando polÃ­tica de execuÃ§Ã£o...
[OK] PolÃ­tica de execuÃ§Ã£o adequada

â•”â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•—
â•‘          CHAINSAW - InstalaÃ§Ã£o de ConfiguraÃ§Ãµes do Word       â•‘
â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

[i] Verificando acesso ao caminho: C:\Users\csantos\chainsaw
[OK] Arquivos de origem encontrados [OK]

[i] Origem: C:\Users\csantos\chainsaw\assets\stamp.png
[i] Destino: C:\Users\csantos\chainsaw\assets\stamp.png
[OK] Arquivo stamp.png copiado com sucesso [OK]

[i] Origem: C:\Users\csantos\chainsaw\configs\Templates
[i] Destino: C:\Users\csantos\AppData\Roaming\Microsoft\Templates
[OK] Pasta Templates copiada com sucesso (37 arquivos) [OK]

â•”â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•—
â•‘              INSTALAÃ‡ÃƒO CONCLUÃDA COM SUCESSO!                 â•‘
â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
```

## [SYNC] DistribuiÃ§Ã£o

### Para Distribuir para Outros UsuÃ¡rios

1. **Comprimir** a pasta `chainsaw` completa
2. **Enviar** por email, rede, ou USB
3. **Instruir** o usuÃ¡rio a:
   - Extrair para `Documentos\chainsaw`
   - Executar `install.cmd`

### Script de DistribuiÃ§Ã£o (Opcional)

VocÃª pode criar um script batch para automatizar a cÃ³pia:

```batch
@echo off
echo Copiando Chainsaw para Documentos...
xcopy /E /I /Y "\\servidor\compartilhado\chainsaw" "%USERPROFILE%\chainsaw\"
echo.
echo Instalando...
cd "%USERPROFILE%\chainsaw"
install.cmd
```

## ðŸ†š ComparaÃ§Ã£o

| Aspecto | Rede (Antes) | Local (Agora) |
|---------|-------------|---------------|
| **Velocidade** | Lenta (rede) | RÃ¡pida (disco local) |
| **Confiabilidade** | Depende da rede | 100% confiÃ¡vel |
| **Requisitos** | VPN/Rede corporativa | Nenhum |
| **Offline** | [X] NÃ£o funciona | [OK] Funciona |
| **DistribuiÃ§Ã£o** | Centralizada | Descentralizada |

## [SEC] SeguranÃ§a

### Mantida
- [OK] Bypass automÃ¡tico seguro
- [OK] Sem privilÃ©gios de administrador
- [OK] Backup automÃ¡tico
- [OK] Log completo
- [OK] Rollback em caso de erro

### Melhorada
- [OK] NÃ£o requer acesso Ã  rede corporativa
- [OK] Reduz superfÃ­cie de ataque (menos dependÃªncias externas)
- [OK] Verifica se origem = destino para evitar sobrescrever

## [LOG] Notas TÃ©cnicas

### ParÃ¢metro SourcePath

O parÃ¢metro `-SourcePath` ainda existe para casos especiais:

```powershell
# Se os arquivos estÃ£o em outro local
.\install.ps1 -SourcePath "C:\outro\local\chainsaw"

# Ou atÃ© mesmo de uma rede (se necessÃ¡rio)
.\install.ps1 -SourcePath "\\servidor\compartilhado\chainsaw"
```

### PSScriptRoot

O script usa `$PSScriptRoot` para detectar automaticamente sua localizaÃ§Ã£o:
- [OK] Funciona em PowerShell 3.0+
- [OK] Sempre aponta para o diretÃ³rio do script
- [OK] Funciona com caminhos UNC

## ðŸ› SoluÃ§Ã£o de Problemas

### Erro: "Arquivos de origem nÃ£o encontrados"

**Causa**: Pasta `chainsaw` nÃ£o estÃ¡ em Documentos ou estrutura incompleta.

**SoluÃ§Ã£o**:
1. Verifique se a pasta estÃ¡ em: `%USERPROFILE%\chainsaw`
2. Verifique se existe: `assets\stamp.png` e `configs\Templates\`

### Erro: "NÃ£o pode substituir o item por ele mesmo"

**Causa**: VersÃ£o antiga do script (jÃ¡ corrigido).

**SoluÃ§Ã£o**: Atualize para a versÃ£o mais recente do script.

## [OK] Checklist de InstalaÃ§Ã£o

Para usuÃ¡rios finais:

- [ ] Copiar pasta `chainsaw` para `Documentos`
- [ ] Fechar o Microsoft Word
- [ ] Abrir PowerShell ou Prompt de Comando
- [ ] Navegar para: `cd "%USERPROFILE%\chainsaw"`
- [ ] Executar: `install.cmd`
- [ ] Aguardar conclusÃ£o
- [ ] Verificar mensagem de sucesso

## ðŸ“ž Suporte

Se encontrar problemas:

1. Verifique o log: `%USERPROFILE%\chainsaw\logs\install_*.log`
2. Consulte `INSTALL.md` para documentaÃ§Ã£o completa
3. Entre em contato: chrmsantos@protonmail.com

---

**VersÃ£o:** 1.1.0 (InstalaÃ§Ã£o Local)  
**Data:** 05/11/2025  
**Autor:** Christian Martin dos Santos

