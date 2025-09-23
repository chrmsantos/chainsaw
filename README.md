# CHAINSAW PROPOSITURAS

**v1.9.1-Alpha-8** - A solução open source em VBA para padronização e automação avançada de documentos legislativos no Microsoft Word, desenvolvida especificamente para Câmaras Municipais e ambientes institucionais.

## 🆕 Novidades da Versão 1.9.1-Alpha-8

### Sistema de Configuração Avançado

- **Arquivo de configuração externo:** `chainsaw-config.ini` com mais de 100 configurações
- **Controle granular:** Habilite/desabilite qualquer funcionalidade do sistema
- **15 categorias de configuração:** Geral, Validações, Backup, Formatação, Limpeza, Performance, etc.
- **Configuração automática:** Sistema carrega valores padrão se arquivo não encontrado

### Otimizações de Performance

- **Processamento em lote:** Parágrafos processados em grupos para melhor performance
- **Operações otimizadas:** Find/Replace em bulk, cache de objetos frequentes
- **Gestão de memória:** Coleta de lixo inteligente e minimização de criação de objetos
- **Compatibilidade preservada:** Todas as otimizações mantêm compatibilidade com Word 2010+

### Sistema de Logging Aprimorado

- **Controle detalhado:** Configure níveis de log (ERROR, WARNING, INFO, DEBUG)
- **Performance tracking:** Medição precisa de tempo de execução
- **Configuração flexível:** Enable/disable logging por categoria

## Principais Funcionalidades

- **Padronização automática de proposituras legislativas:**  
  Formatação específica para INDICAÇÕES, REQUERIMENTOS e MOÇÕES com controle de layout institucional.
- **Validação de conteúdo configurável:**  
  Verificação de consistência entre ementa e teor das proposituras (pode ser desabilitada).
- **Remoção inteligente de elementos visuais:**  
  Limpeza automática de elementos ocultos e formatação inadequada (totalmente configurável).
- **Sistema robusto de backup:**  
  Backup automático antes de modificações, com recuperação de emergência.
- **Formatação institucional:**  
  Cabeçalho com logotipo, numeração de páginas e margens padronizadas.
- **Logging detalhado:**  
  Geração de logs com timestamps, níveis de severidade e rastreamento completo.
- **Interface aprimorada:**  
  Mensagens claras ao usuário e validações interativas.
- **Performance otimizada:**  
  Processamento eficiente mesmo para documentos grandes.
- **Segurança avançada:**  
  Validação de integridade, verificação de versão e proteção contra falhas.

## Instalação

1. Baixe o repositório:  
   [github.com/chrmsantos/chainsaw-proposituras](https://github.com/chrmsantos/chainsaw-proposituras)
2. Execute o script PowerShell de instalação automatizada (recomendado):

   ```powershell
   .\install-chainsaw-proposituras.ps1
   ```

3. **OU** faça a instalação manual:
   - Importe o módulo `Módulo1.bas` no editor VBA do Word (Alt+F11)
   - Configure as permissões de segurança de macro (veja seção **Configurações de Segurança**)

## ⚙️ Sistema de Configuração

### Arquivo de Configuração (`chainsaw-config.ini`)

O sistema utiliza um arquivo de configuração externo que permite controle granular sobre todas as funcionalidades:

```ini
[GERAL]
debug_mode = false
performance_mode = true
compatibility_mode = true

[VALIDACOES]
validate_document_integrity = true
validate_proposition_type = true
check_word_version = true
min_word_version = 14.0

[PERFORMANCE]
disable_screen_updating = true
use_bulk_operations = true
batch_paragraph_operations = true
optimize_find_replace = true
```

### Localização do Arquivo

O sistema procura o arquivo `chainsaw-config.ini` em:

1. **Pasta do documento atual** (se houver documento aberto)
2. **Pasta Documentos do usuário** (fallback)

### Configuração Automática

- Se o arquivo não for encontrado, o sistema **usa valores padrão**
- Todas as funcionalidades principais permanecem **habilitadas por padrão**
- Permite **personalização completa** sem quebrar funcionalidade básica

### Principais Categorias de Configuração

| Categoria | Descrição | Configurações |
|-----------|-----------|---------------|
| **GERAL** | Configurações básicas do sistema | Debug, Performance, Compatibilidade |
| **VALIDACOES** | Controle de validações | Integridade, Versão, Tipo de documento |
| **BACKUP** | Sistema de backup | Auto-backup, Retenção, Tentativas |
| **FORMATACAO** | Controle de formatação | Fonte, Parágrafos, Hifenização |
| **LIMPEZA** | Limpeza de documento | Espaços, Elementos visuais, Formatação |
| **PERFORMANCE** | Otimizações | Processamento em lote, Cache, Loops |
| **INTERFACE** | Mensagens e progresso | Alertas, Status, Confirmações |
| **SEGURANCA** | Validações de segurança | Permissões, Proteção, Sanitização |

## Uso Básico

1. Execute a macro `PadronizarDocumentoMain` em seu documento.

## Configurações de Segurança

### Configuração de Macros no Microsoft Word

Para usar o chainsaw-fprops com segurança, configure o Word da seguinte forma:

1. **Acesse as configurações de segurança:**
   - Arquivo → Opções → Central de Confiabilidade → Configurações da Central de Confiabilidade
   - Clique em "Configurações de Macro"

2. **Configuração recomendada:**
   - Selecione "Desabilitar todas as macros com notificação"
   - Esta opção permite que você escolha quando executar macros

3. **Locais confiáveis (opcional):**
   - Adicione a pasta do chainsaw-fprops aos "Locais Confiáveis"
   - Isso permitirá execução automática apenas desta pasta específica

### Verificação de Segurança

Antes de executar a macro:

- ✅ Verifique se o arquivo foi baixado de fonte confiável
- ✅ Execute em documentos com backup disponível
- ✅ Teste primeiro em documentos não-críticos
- ✅ Mantenha o antivírus atualizado

**Importante:** O CHAINSAW PROPOSITURAS é open source e não se conecta à internet. Todo o código pode ser inspecionado no arquivo VBA.

Para ambientes corporativos, consulte também a [Política de Segurança para Macros](MACRO_SECURITY_POLICY.md).

## Requisitos

- Microsoft Word 2010 ou superior (Windows)
- Permissão para executar macros VBA

## Licença

Código sob licença [Apache 2.0 modificada com cláusula 10 (restrição comercial), conforme LICENSE](LICENSE).  
O Microsoft Word é software proprietário e requer licença própria.

## Autor

Christian Martin dos Santos

## Contribuição

Colaborações são bem-vindas! Consulte o arquivo [CONTRIBUTORS.md](CONTRIBUTORS.md) para detalhes.

---
