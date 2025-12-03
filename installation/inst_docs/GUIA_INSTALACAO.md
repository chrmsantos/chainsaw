# Guia de Instalação - CHAINSAW

## Visão Geral

Sistema automatizado para padronização de documentos legislativos no Microsoft Word.

## Requisitos

- Windows 10 ou superior
- PowerShell 5.1 ou superior
- Microsoft Word 2010 ou superior
- Conexão com a internet (para download inicial)

---

## Instalação Automática (Recomendado)

### ✨ Método Mais Simples - Um Único Arquivo

**Você precisa de apenas 1 arquivo:**

1. **Baixe**: `chainsaw_installer.cmd` da raiz do repositório
2. **Execute**: Dê duplo-clique no arquivo
3. **Pronto!** O instalador fará tudo sozinho

#### O que o instalador automático faz

```text
✓ Baixa o código-fonte completo do GitHub
✓ Valida a integridade dos arquivos baixados
✓ Cria backup da instalação existente (se houver)
✓ Extrai para: C:\Users\[seu_usuario]\chainsaw
✓ Executa a instalação automaticamente
✓ Valida a instalação final
✓ Cria logs detalhados
```

**Vantagens**:

- ✅ Sempre baixa a versão mais recente
- ✅ Não precisa clonar repositório completo
- ✅ Backup automático e validado
- ✅ Rollback em caso de falha
- ✅ Logs completos para diagnóstico


---

## Instalação Manual (Se já tem o repositório)

Se você já clonou/baixou o repositório completo:

### Passo 1: Posicionar na raiz do projeto

Navegue até a pasta onde está o `chainsaw_installer.cmd`

### Passo 2: Executar o Instalador

Dê duplo-clique em: `chainsaw_installer.cmd`

**OU** execute via linha de comando:

```batch
chainsaw_installer.cmd
```

### Passo 3: Aguardar Conclusão

O instalador executará automaticamente:

```text
[OK] ETAPA 1: Verificação de Pré-requisitos
[OK] ETAPA 2: Validação de Arquivos
[OK] ETAPA 3: Backup Automático
[OK] ETAPA 4: Instalação de Templates
[OK] ETAPA 5: Atualização do Módulo VBA
[OK] ETAPA 6: Importação de Personalizações (se disponível)
```

## Atualização do Módulo VBA

Para atualizar apenas o módulo VBA (sem reinstalar tudo):

```powershell
cd "$env:USERPROFILE\chainsaw\installation\inst_scripts"
.\update-vba-module.ps1
```

Ou dê duplo-clique em: `update-vba-module.cmd`

## Exportar Personalizações

Para fazer backup de suas personalizações do Word:

```cmd
cd %USERPROFILE%\chainsaw\installation\inst_scripts
exportar_configs.cmd
```

Isso criará uma pasta `exported-config` com:

- Faixa de Opções personalizada (Ribbon)
- Partes Rápidas (Quick Parts)
- Blocos de Construção (Building Blocks)
- Template Normal.dotm

## Importar Personalizações

Se você possui uma pasta `exported-config`:

1. Copie-a para: `chainsaw\installation\`
2. Execute `chainsaw_installer.cmd` normalmente
3. O instalador detectará e oferecerá importar automaticamente

## Opções Avançadas

### Instalação Silenciosa (sem confirmação)

```powershell
.\install.ps1 -Force
```

### Sem Backup Automático

```powershell
.\install.ps1 -NoBackup
```

### Sem Importar Personalizações

```powershell
.\install.ps1 -SkipCustomizations
```

## Logs

Todos os logs ficam em:

```text
chainsaw\installation\inst_docs\inst_logs\install_YYYYMMDD_HHMMSS.log
```

## Resolução de Problemas

### Word está aberto

**Problema:** Erro ao fazer backup ou copiar arquivos  
**Solução:** Feche o Word completamente antes de executar a instalação

### Política de Execução do PowerShell

**Problema:** Script não executa  
**Solução:** Use `chainsaw_installer.cmd` que possui bypass automático seguro

### Erro de Permissões

**Problema:** Acesso negado  
**Solução:** NÃO execute como administrador - use seu usuário normal

### Verificar Instalação

Para verificar se a instalação foi bem-sucedida:

1. Abra o Word
2. Pressione `Alt + F11` para abrir o VBA
3. Verifique se o módulo `monolithicMod` está presente
4. Verifique se a Faixa de Opções personalizada aparece

## Segurança

- [OK] Não requer privilégios de administrador
- [OK] Não modifica arquivos do sistema
- [OK] Backup automático antes de qualquer alteração
- [OK] Rollback em caso de erro
- [OK] Bypass temporário seguro (não altera configurações permanentes)
- [OK] Logs completos de todas as operações

## Localização dos Arquivos

| Item | Localização |
|------|-------------|
| **Scripts de instalação** | `chainsaw\installation\inst_scripts\` |
| **Templates** | `chainsaw\installation\inst_configs\Templates\` |
| **Módulo VBA** | `chainsaw\source\main\monolithicMod.bas` |
| **Logs** | `chainsaw\installation\inst_docs\inst_logs\` |
| **Configurações exportadas** | `chainsaw\installation\exported-config\` |
| **Normal.dotm instalado** | `%APPDATA%\Microsoft\Templates\Normal.dotm` |

## Documentação Adicional

- [README.md](../README.md) - Visão geral do projeto
- [CHANGELOG.md](../CHANGELOG.md) - Histórico de versões
- [IDENTIFICACAO_ELEMENTOS.md](IDENTIFICACAO_ELEMENTOS.md) - Sistema de identificação de elementos
- [NOVIDADES_v1.1.md](NOVIDADES_v1.1.md) - Novidades da versão 1.1

## Dicas

1. **Primeira instalação**: Execute sem opções adicionais
2. **Reinstalação**: Use `-Force` para instalação rápida
3. **Distribuição**: Compartilhe a pasta `chainsaw` completa
4. **Backup**: Exporte suas personalizações periodicamente
5. **Atualização**: Use `update-vba-module.ps1` para atualizar apenas o código

---

**Versão:** 2.0.1  
**Última atualização:** 8 de novembro de 2024
