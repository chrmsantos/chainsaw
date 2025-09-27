# CHAINSAW PROPOSITURAS

## v1.0.0-Beta1

*A solução open source em VBA para padronização e automação avançada de documentos legislativos no Microsoft Word, desenvolvida especificamente para Câmaras Municipais e ambientes institucionais.*

[![License](https://img.shields.io/badge/License-Apache%202.0%20Modified-blue.svg)](LICENSE)
![Word Version](https://img.shields.io/badge/Word-2010+-green.svg)
![Language](https://img.shields.io/badge/Language-VBA-orange.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)

## 📋 Índice

- [Novidades da Versão](#-novidades-da-versão-100-beta1)
- [Principais Funcionalidades](#-principais-funcionalidades)
- [Estrutura do Projeto](#-estrutura-do-projeto)
- [Instalação](#-instalação)
- [Configuração](#️-configuração)
- [Uso](#-uso)
- [Segurança](#-segurança)
- [Requisitos](#-requisitos)
- [Documentação](#-documentação)
- [Contribuição](#-contribuição)
- [Licença](#-licença)

## 🆕 Novidades da Versão 1.0.0-Beta1

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

## 🚀 Principais Funcionalidades

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

## 📁 Estrutura do Projeto

```text
chainsaw/
├── 📁 assets/              # Recursos (imagens, ícones)
│   └── stamp.png          # Logo institucional
├── 📁 config/             # Arquivos de configuração
│   ├── chainsaw-config.ini # Configuração principal
│   └── word/              # Configurações específicas do Word
├── 📁 docs/               # Documentação
│   ├── CONTRIBUTORS.md    # Lista de contribuidores
│   └── SECURITY.md        # Políticas de segurança
├── 📁 examples/           # Documentos de exemplo
│   └── prop-de-testes-01.docx
├── 📁 scripts/            # Scripts de instalação
│   ├── install-chainsaw.ps1  # Instalador automatizado
│   ├── install-config.ini    # Configuração do instalador
│   └── INSTALL.md           # Guia de instalação
├── 📁 src/                # Código fonte
│   └── chainsaw0.bas      # Módulo VBA principal
├── LICENSE                # Licença do projeto
└── README.md             # Este arquivo
```

## 🔧 Instalação

### Instalação Rápida (Recomendada)

1. **Download do projeto:**
   ```bash
   git clone https://github.com/chrmsantos/chainsaw-proposituras.git
   ```

2. **Execute o instalador automatizado:**

   ```powershell
   cd chainsaw-proposituras
   .\scripts\install-chainsaw.ps1
   ```

### Instalação Manual

Consulte o guia detalhado em [`docs/INSTALL.md`](scripts/INSTALL.md) para instruções completas de instalação manual.

## ⚙️ Configuração

O sistema utiliza um arquivo de configuração externo (`config/chainsaw-config.ini`) que permite controle granular sobre todas as funcionalidades.

### Configuração Rápida

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
```

Para configuração completa, consulte [`config/chainsaw-config.ini`](config/chainsaw-config.ini).

### Localização do Arquivo

O sistema procura o arquivo `chainsaw-config.ini` em:

1. **Pasta do documento atual** (se houver documento aberto)
2. **Pasta Documentos do usuário** (fallback)

## 📖 Uso

### Uso Básico

1. Abra um documento no Microsoft Word
2. Execute a macro `PadronizarDocumentoMain`
3. O sistema processará automaticamente o documento seguindo as configurações

### Funcionalidades Principais

- **Alt + F8**: Abrir lista de macros
- **Ctrl + Shift + P**: Atalho personalizado (configurável)

## 🔒 Segurança

### Configuração de Macros no Microsoft Word

Para usar o CHAINSAW PROPOSITURAS com segurança:

1. **Configurações de Segurança:**
   - Arquivo → Opções → Central de Confiabilidade
   - Configurações de Macro → "Desabilitar todas as macros com notificação"

2. **Verificações de Segurança:**
   - ✅ Código fonte aberto e auditável
   - ✅ Não requer conexão com internet
   - ✅ Backup automático antes de modificações
   - ✅ Tratamento robusto de erros

Para políticas corporativas, consulte [`docs/SECURITY.md`](docs/SECURITY.md).

## 📋 Requisitos

### Mínimos

- **Sistema Operacional:** Windows 7 ou superior
- **Microsoft Word:** 2010 ou superior
- **Permissões:** Execução de macros VBA habilitada
- **Espaço em Disco:** 50MB livres

### Recomendados

- **Microsoft Word:** 2016 ou superior
- **RAM:** 4GB ou superior
- **Processador:** Intel/AMD 64-bit

## 📚 Documentação

### Documentos Disponíveis

- [`docs/SECURITY.md`](docs/SECURITY.md) - Políticas de segurança
- [`docs/CONTRIBUTORS.md`](docs/CONTRIBUTORS.md) - Lista de contribuidores
- [`scripts/INSTALL.md`](scripts/INSTALL.md) - Guia de instalação detalhado

### Exemplos

Consulte a pasta [`examples/`](examples/) para documentos de exemplo e casos de uso.

## 🤝 Contribuição

Colaborações são bem-vindas! Para contribuir:

1. Fork o repositório
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

Consulte [`docs/CONTRIBUTORS.md`](docs/CONTRIBUTORS.md) para detalhes sobre o processo de contribuição.

## 📄 Licença

Este projeto está licenciado sob a **Apache 2.0 License modificada com cláusula 10 (restrição comercial)** - consulte o arquivo [LICENSE](LICENSE) para detalhes.

**Nota:** O Microsoft Word é software proprietário e requer licença própria.

## 👨‍💻 Autor

**Christian Martin dos Santos** - [chrmsantos](https://github.com/chrmsantos)

---

---

Desenvolvido com ❤️ para a comunidade legislativa brasileira
