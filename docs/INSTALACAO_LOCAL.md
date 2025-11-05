# Instalação Local - Chainsaw

## 📋 Mudança Importante

O sistema Chainsaw agora funciona a partir da **pasta Documentos do usuário**, eliminando a necessidade de acesso à rede corporativa durante a instalação.

## 🎯 Benefícios

### Antes (Rede)
- ❌ Dependia de acesso à rede corporativa
- ❌ Problemas com VPN e credenciais
- ❌ Lentidão na cópia de arquivos
- ❌ Falhas por desconexão de rede

### Agora (Local)
- ✅ Funciona offline
- ✅ Instalação mais rápida
- ✅ Mais confiável
- ✅ Sem dependência de rede

## 📦 Estrutura Necessária

A pasta `chainsaw` deve ser copiada para a pasta Documentos do usuário:

```
%USERPROFILE%\Documents\chainsaw\
├── assets\
│   └── stamp.png
├── configs\
│   └── Templates\
│       └── [todos os templates]
├── install.ps1
├── install.cmd
└── [outros arquivos]
```

## 🚀 Instalação

### 1. Copiar Arquivos

Primeiro, copie a pasta completa `chainsaw` para:
- **Windows**: `C:\Users\[seu_usuario]\Documents\chainsaw`

### 2. Executar Instalação

```cmd
cd "%USERPROFILE%\Documents\chainsaw"
install.cmd
```

Ou usando PowerShell:

```powershell
cd "$env:USERPROFILE\Documents\chainsaw"
.\install.ps1
```

## 🔧 Como Funciona

### Detecção Automática de Origem

O script agora detecta automaticamente de onde está sendo executado:

```powershell
# O caminho de origem é automaticamente definido como a pasta do script
$SourcePath = $PSScriptRoot
```

### Verificação de Auto-Cópia

Para evitar erros quando executado diretamente da pasta de destino, o script:

1. **Verifica se origem = destino** para `stamp.png`
   - Se sim, pula a cópia (já está instalado)
   - Se não, copia normalmente

2. **Verifica se origem = destino** para `Templates`
   - Se sim, pula a cópia (já está instalado)
   - Se não, copia normalmente

## 📊 Exemplo de Execução

```
🔒 Verificando política de execução...
✓ Política de execução adequada

╔════════════════════════════════════════════════════════════════╗
║          CHAINSAW - Instalação de Configurações do Word       ║
╚════════════════════════════════════════════════════════════════╝

ℹ Verificando acesso ao caminho: C:\Users\csantos\Documents\chainsaw
✓ Arquivos de origem encontrados ✓

ℹ Origem: C:\Users\csantos\Documents\chainsaw\assets\stamp.png
ℹ Destino: C:\Users\csantos\chainsaw\assets\stamp.png
✓ Arquivo stamp.png copiado com sucesso ✓

ℹ Origem: C:\Users\csantos\Documents\chainsaw\configs\Templates
ℹ Destino: C:\Users\csantos\AppData\Roaming\Microsoft\Templates
✓ Pasta Templates copiada com sucesso (37 arquivos) ✓

╔════════════════════════════════════════════════════════════════╗
║              INSTALAÇÃO CONCLUÍDA COM SUCESSO!                 ║
╚════════════════════════════════════════════════════════════════╝
```

## 🔄 Distribuição

### Para Distribuir para Outros Usuários

1. **Comprimir** a pasta `chainsaw` completa
2. **Enviar** por email, rede, ou USB
3. **Instruir** o usuário a:
   - Extrair para `Documentos\chainsaw`
   - Executar `install.cmd`

### Script de Distribuição (Opcional)

Você pode criar um script batch para automatizar a cópia:

```batch
@echo off
echo Copiando Chainsaw para Documentos...
xcopy /E /I /Y "\\servidor\compartilhado\chainsaw" "%USERPROFILE%\Documents\chainsaw\"
echo.
echo Instalando...
cd "%USERPROFILE%\Documents\chainsaw"
install.cmd
```

## 🆚 Comparação

| Aspecto | Rede (Antes) | Local (Agora) |
|---------|-------------|---------------|
| **Velocidade** | Lenta (rede) | Rápida (disco local) |
| **Confiabilidade** | Depende da rede | 100% confiável |
| **Requisitos** | VPN/Rede corporativa | Nenhum |
| **Offline** | ❌ Não funciona | ✅ Funciona |
| **Distribuição** | Centralizada | Descentralizada |

## 🔐 Segurança

### Mantida
- ✅ Bypass automático seguro
- ✅ Sem privilégios de administrador
- ✅ Backup automático
- ✅ Log completo
- ✅ Rollback em caso de erro

### Melhorada
- ✅ Não requer acesso à rede corporativa
- ✅ Reduz superfície de ataque (menos dependências externas)
- ✅ Verifica se origem = destino para evitar sobrescrever

## 📝 Notas Técnicas

### Parâmetro SourcePath

O parâmetro `-SourcePath` ainda existe para casos especiais:

```powershell
# Se os arquivos estão em outro local
.\install.ps1 -SourcePath "C:\outro\local\chainsaw"

# Ou até mesmo de uma rede (se necessário)
.\install.ps1 -SourcePath "\\servidor\compartilhado\chainsaw"
```

### PSScriptRoot

O script usa `$PSScriptRoot` para detectar automaticamente sua localização:
- ✅ Funciona em PowerShell 3.0+
- ✅ Sempre aponta para o diretório do script
- ✅ Funciona com caminhos UNC

## 🐛 Solução de Problemas

### Erro: "Arquivos de origem não encontrados"

**Causa**: Pasta `chainsaw` não está em Documentos ou estrutura incompleta.

**Solução**:
1. Verifique se a pasta está em: `%USERPROFILE%\Documents\chainsaw`
2. Verifique se existe: `assets\stamp.png` e `configs\Templates\`

### Erro: "Não pode substituir o item por ele mesmo"

**Causa**: Versão antiga do script (já corrigido).

**Solução**: Atualize para a versão mais recente do script.

## ✅ Checklist de Instalação

Para usuários finais:

- [ ] Copiar pasta `chainsaw` para `Documentos`
- [ ] Fechar o Microsoft Word
- [ ] Abrir PowerShell ou Prompt de Comando
- [ ] Navegar para: `cd "%USERPROFILE%\Documents\chainsaw"`
- [ ] Executar: `install.cmd`
- [ ] Aguardar conclusão
- [ ] Verificar mensagem de sucesso

## 📞 Suporte

Se encontrar problemas:

1. Verifique o log: `%USERPROFILE%\chainsaw\logs\install_*.log`
2. Consulte `INSTALL.md` para documentação completa
3. Entre em contato: chrmsantos@protonmail.com

---

**Versão:** 1.1.0 (Instalação Local)  
**Data:** 05/11/2025  
**Autor:** Christian Martin dos Santos
