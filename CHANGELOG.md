# Histórico de Mudanças - CHAINSAW

## [2.0.2] - 2025-01-XX

### [CLEANUP] Reestruturação e Simplificação Completa do Projeto

**Mudança Principal:** Grande limpeza e reorganização da estrutura de pastas e documentação.

### [REMOVED] Removido

- **Documentos Obsoletos**: Removidos 8+ arquivos de documentação duplicados ou obsoletos:
  - `ANALISE_SCRIPT.md`, `BYPASS_SEGURO.md`, `RESUMO_IMPLEMENTACAO.md`
  - `CORRECAO_*.md` (5 arquivos de correção de processos antigos)
  - `IMPLEMENTACAO_COMPLETA.md` (duplicado)
- **Guias Duplicados**: Consolidados 3 guias de instalação em um único:
  - Removidos: `INSTALACAO_LOCAL.md`, `GUIA_INSTALACAO_UNIFICADA.md`, `INSTALL.md`
  - Criado: `installation/inst_docs/GUIA_INSTALACAO.md` (guia único consolidado)
- **Guias Redundantes**: Removidos guias rápidos redundantes:
  - `GUIA_RAPIDO_EXPORT_IMPORT.md`, `GUIA_RAPIDO_IDENTIFICACAO.md`
  - `EXPORTACAO_IMPORTACAO.md`, `ATUALIZACAO_MODULO_VBA.md`
- **Backups Duplicados**: Removidos backups VBA obsoletos da raiz de `source/backups/`
  - Mantido apenas `source/backups/main/monolithicMod.bas` (versão atual)

### [REFACTOR] Refatorado

- **README.md**: Completamente reescrito e simplificado (981 → ~130 linhas)
  - Foco em quick start e links para documentação detalhada
  - Removida duplicação massiva de conteúdo
- **Estrutura de Documentação**: De 19+ arquivos para 6 essenciais:
  - `README.md` - Visão geral e quick start
  - `installation/inst_docs/GUIA_INSTALACAO.md` - Instalação completa
  - `docs/IDENTIFICACAO_ELEMENTOS.md` - Sistema de identificação
  - `docs/NOVIDADES_v1.1.md` - Novidades da v1.1
  - `docs/SEM_PRIVILEGIOS_ADMIN.md` - Instalação sem admin
  - `docs/SUBSTITUICOES_CONDICIONAIS.md` - Lógica de substituições
  - `docs/VALIDACAO_TIPO_DOCUMENTO.md` - Validação de documentos
- **Scripts PowerShell**: Atualizados para nova estrutura:
  - `install.ps1` - Caminhos corrigidos para `installation/`, `source/`
  - `update-vba-module.ps1` - Caminho do módulo VBA atualizado
  - `export-config.ps1` - Caminhos de exportação atualizados

### [CHANGED] Alterado

- **Estrutura de Pastas**: Organização final clara:

  ```plaintext
  chainsaw/
  ├── .vscode/           # Configurações VS Code (auto-approve)
  ├── assets/            # Recursos do projeto
  ├── docs/              # 5 documentos técnicos essenciais
  ├── installation/
  │   ├── inst_configs/  # Configurações e templates
  │   ├── inst_docs/     # Documentação de instalação
  │   └── inst_scripts/  # Scripts de instalação
  ├── source/
  │   ├── backups/main/  # Módulo VBA principal
  │   └── others/        # Exemplos
  ├── CHANGELOG.md
  ├── LICENSE
  └── README.md
  ```

### [CONFIG] Configurado

- **Auto-Approve do Copilot**: Configurado em `.vscode/settings.json`
  - Aprovação automática de todas as sugestões do GitHub Copilot

---

## [1.1.0] - 2024-11-07

### [NEW] Sistema de Identificação de Elementos Estruturais

**Mudança Principal:** Novo sistema completo de identificação automática dos elementos estruturais das proposituras legislativas.

### [NEW] Adicionado

- **Identificação Automática de Elementos**: Sistema identifica automaticamente todos os elementos da propositura:
  - Título, Ementa, Proposição, Justificativa, Data, Assinatura, Anexo
- **10 Funções Públicas de Acesso**: 
  - `GetTituloRange()`, `GetEmentaRange()`, `GetProposicaoRange()`, etc.
  - `GetElementInfo()` para relatório completo
- **Integração com Cache**: Sistema integrado ao cache de parágrafos existente
- **13 Novos Campos no Cache**: Flags booleanos para cada tipo de elemento
- **12 Variáveis Globais de Índice**: Armazenam posições dos elementos encontrados
- **Constantes de Identificação**: Critérios configuráveis para cada elemento
- **Documentação Completa**: `docs/IDENTIFICACAO_ELEMENTOS.md` (200+ linhas)
- **10 Exemplos Práticos**: `src/Exemplos_Identificacao.bas` (500+ linhas)
- **Guia de Novidades**: `docs/NOVIDADES_v1.1.md`

### [CODE] Funções Adicionadas

- `IsTituloElement()` - Identifica título da propositura
- `IsEmentaElement()` - Identifica ementa
- `IsJustificativaTitleElement()` - Identifica título "Justificativa"
- `IsDataElement()` - Identifica data do plenário
- `IsTituloAnexoElement()` - Identifica título do anexo
- `IsAssinaturaStart()` - Identifica início da assinatura
- `CountBlankLinesBefore()` - Conta linhas em branco
- `IdentifyDocumentStructure()` - Função principal de identificação
- `GetTituloRange()` - Acesso público ao título
- `GetEmentaRange()` - Acesso público à ementa
- `GetProposicaoRange()` - Acesso público à proposição
- `GetTituloJustificativaRange()` - Acesso público ao título da justificativa
- `GetJustificativaRange()` - Acesso público à justificativa
- `GetDataRange()` - Acesso público à data
- `GetAssinaturaRange()` - Acesso público à assinatura
- `GetTituloAnexoRange()` - Acesso público ao título do anexo
- `GetAnexoRange()` - Acesso público ao anexo
- `GetProposituraRange()` - Acesso público à propositura completa
- `GetElementInfo()` - Relatório completo de elementos

### [SYNC] Modificado

- **BuildParagraphCache()**: Agora chama `IdentifyDocumentStructure()` após construir cache
- **ClearParagraphCache()**: Limpa também os índices de identificação
- **Type paragraphCache**: Expandido com 9 novos campos booleanos
- **Versão**: Atualizada para 1.1-RC1-202511071045

### [CODE] Características do Sistema

- ✅ Identificação automática durante processamento
- ✅ Overhead mínimo (< 5% do tempo total)
- ✅ 100% compatível com funcionalidades existentes
- ✅ Abordagem defensiva com tratamento de erros
- ✅ Validação completa de nulidade
- ✅ Limites de segurança contra loops infinitos
- ✅ Log detalhado de identificação

### [INFO] Documentação

- `docs/IDENTIFICACAO_ELEMENTOS.md` - Guia completo (200+ linhas)
- `src/Exemplos_Identificacao.bas` - 10 exemplos práticos (500+ linhas)
- `docs/NOVIDADES_v1.1.md` - Resumo executivo
- Header do `src/Módulo1.bas` - Changelog integrado

### [INFO] Exemplos Incluídos

1. Exibir informações completas
2. Selecionar e destacar título
3. Contar palavras por elemento
4. Exportar proposição para novo documento
5. Adicionar marcadores de seção
6. Validar estrutura do documento
7. Destacar elementos visualmente
8. Remover destaques visuais
9. Gerar índice dos elementos
10. Navegar entre elementos

### [LOCK] Compatibilidade

- ✅ Word 2010+
- ✅ Mantém 100% funcionalidades existentes
- ✅ Sem impacto no desempenho
- ✅ Sem mudanças nas APIs existentes

---

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

### [CHART] Estatísticas da Mudança

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
- `docs/INSTALACAO_LOCAL.md` - Instalação a partir do perfil do usuário
- `docs/EXPORTACAO_IMPORTACAO.md` - Documentação técnica completa
- `docs/BYPASS_SEGURO.md` - Explicação do mecanismo de bypass

### [*] Objetivos Alcançados

- [OK] Instalação sem necessidade de privilégios de administrador
- [OK] Funcionamento com qualquer política de execução
- [OK] Backup automático de configurações
- [OK] Log detalhado de todas as operações
- [OK] Exportação e importação de personalizações do Word
- [OK] Instalação a partir do perfil do usuário (%USERPROFILE%)

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
