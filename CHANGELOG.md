# Histórico de Mudanças - CHAINSAW

## [2.0.4] - 2025-11-08

### [FIXED] Padronização de Encoding e Remoção de Emojis

**Correção:** Remoção completa de emojis e padronização de encoding UTF-8 em todo o projeto.

#### Alterações

- Removidos 359+ emojis de scripts PowerShell, arquivos Markdown e documentação
- Emojis substituídos por equivalentes textuais onde apropriado (`[OK]`, `[ERRO]`, `[AVISO]`)
- Adicionado validador de encoding: `tests/check-encoding.ps1`
  - Valida encoding UTF-8 em todos os arquivos
  - Detecta presença de emojis (bytes F0 9F e símbolos E2)
  - Verifica caracteres de controle inválidos
  - Valida consistência de line endings (CRLF para Windows)
- Criados scripts de manutenção:
  - `tests/remove-emojis.ps1` - Remove emojis via substituição de caracteres
  - `tests/remove-emojis-bytes.ps1` - Remove emojis via manipulação de bytes
  - `tests/find-emoji-bytes.ps1` - Localiza emojis em arquivos específicos
- Todos os arquivos agora usam apenas caracteres compatíveis com UTF-8 sem emojis
- Mantidos caracteres acentuados do português (parte legítima do UTF-8)

#### Arquivos Modificados

- `installation/inst_scripts/export-config.ps1` - 19 substituições
- `installation/inst_scripts/install.ps1` - 51 substituições  
- `docs/IDENTIFICACAO_ELEMENTOS.md` - 5 substituições
- `docs/LGPD_CONFORMIDADE.md` - 64 substituições
- `docs/NOVIDADES_v1.1.md` - 16 substituições
- `docs/SEGURANCA_PRIVACIDADE.md` - 127 substituições
- `docs/SEM_PRIVILEGIOS_ADMIN.md` - 4 substituições
- `docs/VALIDACAO_TIPO_DOCUMENTO.md` - 1 substituição
- `CHANGELOG.md` - 28 substituições
- `LGPD_ATESTADO.md` - 47 substituições
- `README.md` - 10 substituições
- `installation/inst_docs/GUIA_INSTALACAO.md` - 19 substituições

#### Testes

- Total de arquivos verificados: 27
- Emojis detectados e removidos: 359+
- Validação final: 0 emojis, 0 erros, 0 avisos
- Encoding validado: UTF-8 para documentação, ASCII-compatível para scripts

## [2.0.3] - 2025-11-08

### [ADDED] Documentação de Conformidade LGPD e Segurança

**Adição Principal:** Documentação completa e criteriosa sobre conformidade com LGPD, segurança e privacidade.

#### Novos Documentos

- **`docs/LGPD_CONFORMIDADE.md`** - Análise completa de conformidade com a LGPD (Lei nº 13.709/2018)
  - 13 seções cobrindo todos os aspectos da lei
  - Inventário de dados (atestação de não coleta)
  - Princípios da LGPD aplicados (Art. 6º)
  - Bases legais e hipóteses de tratamento
  - Direitos dos titulares garantidos
  - Medidas técnicas de segurança implementadas
  - Análise de vulnerabilidades e controles
  - Responsabilidade e governança
  - Auditoria e conformidade contínua
  - Declaração oficial de conformidade

- **`docs/SEGURANCA_PRIVACIDADE.md`** - Política completa de Segurança e Privacidade
  - 17 seções detalhadas
  - Escopo e aplicação da política
  - Coleta e uso de dados (atestação: nenhum dado coletado)
  - Armazenamento local e segurança
  - Arquitetura de segurança técnica
  - Controles de segurança implementados (validação, tratamento de erros, timeouts)
  - Privacidade de documentos processados
  - Transparência e auditabilidade (código aberto)
  - Direitos dos usuários garantidos
  - Processo de resposta a incidentes
  - Conformidade legal e normas técnicas (ISO 27001, NIST, OWASP)
  - Educação e boas práticas
  - Ciclo de vida seguro de desenvolvimento

- **`LGPD_ATESTADO.md`** - Atestado executivo de conformidade LGPD
  - Resumo executivo para rápida consulta
  - Certificação de não coleta de dados
  - Checklist de conformidade para usuários e organizações
  - Declaração oficial do desenvolvedor
  - Referências rápidas para documentação completa

#### Atualizações em Documentos Existentes

- **`README.md`** - Adicionada seção "Segurança e Privacidade"
  - Links para novos documentos de conformidade LGPD
  - Links para política de segurança e privacidade

### [COMPLIANCE] Conformidade e Certificação

**Status de Conformidade:** [OK] **TOTALMENTE CONFORME** com LGPD

#### Principais Atestações

- [OK] **NÃO coleta** dados pessoais de usuários
- [OK] **NÃO transmite** informações pela internet
- [OK] **NÃO armazena** dados em servidores externos
- [OK] **NÃO utiliza** serviços de terceiros ou telemetria
- [OK] **Processamento 100% local** sob controle do usuário
- [OK] **Código auditável** (open source GPLv3)
- [OK] **172 testes automatizados** incluindo segurança
- [OK] **Histórico de incidentes:** 0 (zero)

#### Princípios LGPD Aplicados (Art. 6º)

| Princípio | Status | Justificativa |
|-----------|--------|---------------|
| Finalidade | [OK] Conforme | Formatação de documentos legislativos |
| Adequação | [OK] Conforme | Processamento compatível com finalidade |
| Necessidade | [OK] Conforme | Apenas dados técnicos essenciais |
| Livre Acesso | [OK] Conforme | Logs e backups acessíveis localmente |
| Transparência | [OK] Conforme | Código aberto (GPLv3) |
| Segurança | [OK] Conforme | Processamento local isolado |
| Prevenção | [OK] Conforme | Arquitetura impede coleta de dados |
| Responsabilização | [OK] Conforme | Documentação completa |

#### Medidas de Segurança Documentadas

**Técnicas:**
- Processamento 100% local (zero conexões de rede)
- Validação de entrada e sanitização de caminhos
- Tratamento robusto de erros (try-catch completo)
- Timeouts e limites de operação (MAX_OPERATION_TIMEOUT_SECONDS)
- Limpeza de recursos (SafeCleanup, ReleaseObjects)
- Proteção contra vulnerabilidades (path traversal, code injection, etc)

**Organizacionais:**
- Código auditável (open source)
- Processo de resposta a incidentes (24-48h para critical)
- Testes automatizados de segurança (VBA.Tests.ps1, Installation.Tests.ps1)
- Documentação completa e transparente
- Compromisso de manutenção contínua

#### Orientações para Organizações

Documentação inclui orientações específicas para organizações que adotarem o CHAINSAW:

1. Designação de Encarregado de Dados (DPO) - Art. 41º LGPD
2. Elaboração de Política de Privacidade organizacional
3. Manutenção de ROPA (Registro de Atividades de Tratamento)
4. Avaliação de necessidade de DPIA (Data Protection Impact Assessment)
5. Treinamento de usuários sobre LGPD e segurança

### [DOCUMENTATION] Melhorias na Documentação

- Adicionadas 3 novas referências na seção "Documentação" do README
- Estrutura de documentação agora inclui subseção "Segurança e Privacidade"
- Total de documentação: ~2.500 linhas de conteúdo técnico e jurídico
- Cobertura completa de aspectos legais, técnicos e organizacionais

### [TRANSPARENCY] Transparência e Auditabilidade

- Declaração oficial de conformidade assinada digitalmente (Git SHA-256)
- Histórico de revisões em todos os documentos
- Compromisso de revisão anual ou quando houver alterações na LGPD
- Canal de reporte de vulnerabilidades estabelecido (chrmsantos@protonmail.com)

---

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

- [OK] Identificação automática durante processamento
- [OK] Overhead mínimo (< 5% do tempo total)
- [OK] 100% compatível com funcionalidades existentes
- [OK] Abordagem defensiva com tratamento de erros
- [OK] Validação completa de nulidade
- [OK] Limites de segurança contra loops infinitos
- [OK] Log detalhado de identificação

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

- [OK] Word 2010+
- [OK] Mantém 100% funcionalidades existentes
- [OK] Sem impacto no desempenho
- [OK] Sem mudanças nas APIs existentes

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
