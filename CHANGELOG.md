# Changelog

Todas as mudanças notáveis neste projeto serão documentadas neste arquivo.

O formato é baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.0.0/),
e este projeto adere ao [Semantic Versioning](https://semver.org/lang/pt-BR/).

## [2.0.2] - 2025-11-25

### Adicionado
- Sistema de logging completo no `chainsaw_installer.cmd`
  - Log salvo no diretório do instalador
  - Log copiado para `installation/inst_docs/inst_logs/`
- Sistema de backup automático da instalação existente
  - Backup criado antes de modificar a instalação
  - Backup nomeado com timestamp: `chainsaw_backup_YYYYMMDD_HHMM`
  - Backup completo com fallback para backup seletivo
- Testes para `chainsaw_installer.cmd` no test suite de instalação
  - Validação de estrutura e conteúdo
  - Validação de ordem de execução
  - Validação de segurança e logging

### Modificado
- `chainsaw_installer.cmd` agora executa operações na ordem correta:
  1. Download do código-fonte
  2. Verificação do download
  3. Criação de backup (se instalação existente)
  4. Remoção da instalação antiga
  5. Extração e instalação da nova versão
- Melhorias na segurança: pasta antiga só é removida após confirmação de download bem-sucedido

### Corrigido
- Prevenção de perda de dados: backup criado antes de qualquer modificação
- Download verificado antes de qualquer alteração no sistema

## [2.0.1] - 2025-11-24

### Adicionado
- Estrutura inicial do projeto CHAINSAW
- Scripts de instalação automatizada
- Módulos VBA monolíticos
- Documentação completa de instalação

### Segurança
- Conformidade com LGPD
- Sistema de backup e restauração
- Validação de integridade dos arquivos

[2.0.2]: https://github.com/chrmsantos/chainsaw/compare/v2.0.1...v2.0.2
[2.0.1]: https://github.com/chrmsantos/chainsaw/releases/tag/v2.0.1
