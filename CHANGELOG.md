# Histórico de Mudanças - CHAINSAW

## [2.0.0] - 2024-01-15

### [NEW] Instalação Unificada

**Mudança Principal:** Unificação do processo de instalação em um único script que automaticamente detecta e importa personalizações do Word.

### [NEW] Adicionado

- **Detecção Automática de Personalizações**: `install.ps1` agora detecta automaticamente a pasta `exported-config` e oferece importar personalizações
- **Importação Integrada**: Todas as funções de importação foram integradas ao script principal
- **Novo Parâmetro**: `-SkipCustomizations` para instalar apenas templates sem importar personalizações
- **Backup de Personalizações**: Backup automático antes de importar personalizações
- **Documentação Nova**: `GUIA_INSTALACAO_UNIFICADA.md` com instruções completas do novo processo

### [SYNC] Modificado

- **install.ps1**: Agora versão 2.0.0 com importação integrada
  - Detecta pasta `exported-config` automaticamente
  - Importa personalizações se confirmado pelo usuário
  - Mantém todas as funcionalidades anteriores
  - Adiciona nova etapa (ETAPA 6: Importação de Personalizações)

### [DEL] Removido

Scripts legados consolidados ou obsoletos:

- `import-config.ps1` - Funcionalidade integrada ao `install.ps1`
- `import-config.cmd` - Não é mais necessário
- `start-install.ps1` - Substituído por `install.cmd` desde v1.x
- `test-simple.ps1` - Script de teste legado
- `test-permissions.ps1` - Script de teste legado
- `test-install.ps1` - Script de teste legado

### [PKG] Mantido

Scripts essenciais que permanecem:

- `install.ps1` - **Instalador unificado** (agora v2.0.0)
- `install.cmd` - Launcher seguro com bypass automático
- `export-config.ps1` - Exportação de personalizações (mantido separado por design)
- `export-config.cmd` - Launcher seguro para exportação

### [INFO] Documentação Atualizada

- `README.md` - Atualizado com informações sobre importação automática
- `GUIA_INSTALACAO_UNIFICADA.md` - Novo guia completo do processo unificado
- `CHANGELOG.md` - Este arquivo

### [CFG] Melhorias

- **Experiência do Usuário**: Processo mais simples e intuitivo
- **Menos Arquivos**: Redução de scripts de 9 para 4 (instalação + exportação)
- **Manutenção**: Código consolidado facilita manutenção futura
- **Flexibilidade**: Mantém opções para usuários avançados (`-SkipCustomizations`, `-Force`, `-NoBackup`)

### [TOOL] Notas Técnicas

#### Fluxo de Instalação Unificado

```
install.ps1 v2.0.0
├── ETAPA 1: Verificação de Pré-requisitos
├── ETAPA 2: Validação de Arquivos
├── ETAPA 3: Cópia do Arquivo de Imagem
├── ETAPA 4: Backup da Pasta Templates
├── ETAPA 5: Cópia da Pasta Templates
└── ETAPA 6: Importação de Personalizações ← NOVO
    ├── Detecta exported-config/
    ├── Solicita confirmação (modo interativo)
    ├── Cria backup de personalizações
    └── Importa:
        ├── Normal.dotm
        ├── Building Blocks
        ├── Document Themes
        ├── Ribbon Customization
        └── Office Custom UI
```

#### Compatibilidade

- [OK] Scripts v1.x continuam funcionando
- [OK] Pasta `exported-config` é opcional (não causa erro se ausente)
- [OK] Todos os parâmetros anteriores mantidos
- [OK] Formato de log inalterado

### [SEC] Segurança

- Nenhuma alteração nas políticas de segurança
- Bypass temporário continua limitado ao script
- Backups automáticos protegem contra perda de dados
- Sem necessidade de privilégios de administrador

### [*] Uso Recomendado

**Instalação Nova:**
```cmd
install.cmd
```

**Com Personalizações (workflow completo):**
```cmd
# Máquina de origem
export-config.cmd

# [Copiar pasta CHAINSAW completa]

# Máquina de destino
install.cmd
```

**Apenas Templates:**
```cmd
install.cmd -SkipCustomizations
```

**Modo Automático (Deploy):**
```cmd
install.cmd -Force
```

### 📊 Estatísticas da Mudança

- **Linhas de Código Adicionadas**: ~450 (funções de importação integradas)
- **Scripts Removidos**: 6
- **Scripts Mantidos**: 4
- **Funcionalidades Adicionadas**: 8
- **Breaking Changes**: 0 (totalmente retrocompatível)

---

## [1.0.0] - 2024-01-10

### [NEW] Adicionado

- Script de instalação automatizado (`install.ps1`)
- Launcher seguro com bypass (`install.cmd`)
- Exportação de personalizações (`export-config.ps1` e `.cmd`)
- Importação de personalizações (`import-config.ps1` e `.cmd`)
- Sistema de logs detalhado
- Backup automático de templates
- Documentação completa

### [CFG] Funcionalidades

- Instalação de templates do Word
- Cópia de arquivo de imagem (stamp.png)
- Backup automático com rotação (mantém 5 mais recentes)
- Verificação de pré-requisitos
- Detecção de política de execução
- Auto-relançamento com bypass seguro

### [INFO] Documentação

- `README.md` - Visão geral do sistema
- `INSTALL.md` - Guia de instalação detalhado
- `GUIA_RAPIDO_EXPORT_IMPORT.md` - Guia de exportação/importação
- `docs/INSTALACAO_LOCAL.md` - Instalação a partir da pasta Documentos
- `docs/EXPORTACAO_IMPORTACAO.md` - Documentação técnica completa
- `docs/BYPASS_SEGURO.md` - Explicação do mecanismo de bypass

### [*] Objetivos Alcançados

- [OK] Instalação sem necessidade de privilégios de administrador
- [OK] Funcionamento com qualquer política de execução
- [OK] Backup automático de configurações
- [OK] Log detalhado de todas as operações
- [OK] Exportação e importação de personalizações do Word
- [OK] Instalação a partir da pasta Documentos do usuário

---

## Formato do Changelog

Este changelog segue o formato [Keep a Changelog](https://keepachangelog.com/pt-BR/1.0.0/),
e este projeto adere ao [Versionamento Semântico](https://semver.org/lang/pt-BR/).

### Categorias

- **Adicionado** - Para novas funcionalidades
- **Modificado** - Para mudanças em funcionalidades existentes
- **Obsoleto** - Para funcionalidades que serão removidas em breve
- **Removido** - Para funcionalidades removidas
- **Corrigido** - Para correções de bugs
- **Segurança** - Para questões de segurança
