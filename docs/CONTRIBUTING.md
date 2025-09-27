# Contributing Guidelines

Agradecemos o seu interesse em contribuir com o CHAINSAW PROPOSITURAS! Este documento fornece diretrizes e informações para ajudá-lo a contribuir de forma eficaz.

## Índice

- [Como Contribuir](#como-contribuir)
- [Configuração do Ambiente de Desenvolvimento](#configuração-do-ambiente-de-desenvolvimento)
- [Padrões de Código](#padrões-de-código)
- [Processo de Pull Request](#processo-de-pull-request)
- [Reportar Problemas](#reportar-problemas)
- [Comunidade e Conduta](#comunidade-e-conduta)

## Como Contribuir

Existem várias formas de contribuir com o projeto:

### 🐛 Reportar Bugs
- Use o template de issue para bugs
- Forneça informações detalhadas sobre o ambiente
- Inclua passos para reproduzir o problema

### 💡 Sugerir Melhorias
- Use o template de issue para features
- Descreva claramente o problema que a feature resolve
- Explique como a implementação funcionaria

### 📝 Melhorar Documentação
- Corrija erros de digitação ou gramática
- Adicione exemplos ou esclareça instruções
- Traduza documentação para outros idiomas

### 🔧 Contribuir com Código
- Corrija bugs existentes
- Implemente novas funcionalidades
- Otimize performance e qualidade do código

## Configuração do Ambiente de Desenvolvimento

### Pré-requisitos

- Windows 7 ou superior
- Microsoft Word 2010 ou superior
- Git instalado e configurado
- Editor de texto/IDE (recomendado: VS Code)

### Setup do Projeto

1. **Fork o repositório**
   ```bash
   # Clique em "Fork" no GitHub ou use a CLI do GitHub
   gh repo fork chrmsantos/chainsaw-proposituras
   ```

2. **Clone seu fork**
   ```bash
   git clone https://github.com/SEU_USUARIO/chainsaw-proposituras.git
   cd chainsaw-proposituras
   ```

3. **Configure o upstream**
   ```bash
   git remote add upstream https://github.com/chrmsantos/chainsaw-proposituras.git
   ```

4. **Instale o projeto localmente**
   ```powershell
   .\scripts\install-chainsaw.ps1 -CheckOnly
   ```

## Padrões de Código

### VBA Guidelines

#### Nomenclatura
- **Variáveis**: camelCase (`minhaVariavel`)
- **Constantes**: UPPER_SNAKE_CASE (`MINHA_CONSTANTE`)
- **Procedimentos**: PascalCase (`MinhaProcedure`)
- **Prefixos**: Use prefixos descritivos (`str` para String, `int` para Integer)

#### Estrutura de Código
```vba
' Header obrigatório
' =============================================================================
' NOME DA FUNÇÃO/PROCEDIMENTO - Breve descrição
' =============================================================================
' Descrição detalhada da funcionalidade
' Parâmetros: param1 (tipo) - descrição
' Retorna: tipo - descrição
' Autor: Nome do Contribuidor
' Data: YYYY-MM-DD
' =============================================================================

Public Function MinhaFuncao(param1 As String) As Boolean
    ' Declaração de variáveis locais
    Dim resultado As Boolean
    Dim mensagem As String
    
    ' Validação de entrada
    If Len(param1) = 0 Then
        LogMessage "Parâmetro inválido", LOG_LEVEL_ERROR
        MinhaFuncao = False
        Exit Function
    End If
    
    ' Lógica principal
    On Error GoTo ErrorHandler
    
    ' ... código ...
    
    MinhaFuncao = True
    Exit Function
    
ErrorHandler:
    LogMessage "Erro em MinhaFuncao: " & Err.Description, LOG_LEVEL_ERROR
    MinhaFuncao = False
End Function
```

#### Boas Práticas
- **Sempre use `Option Explicit`**
- **Trate erros apropriadamente**
- **Use logging consistente**
- **Comente código complexo**
- **Evite procedimentos muito longos** (máximo 50 linhas)
- **Use nomes descritivos** para variáveis e procedimentos

### Documentação

#### Comentários
```vba
' Comentário de linha única

' =============================================================================
' COMENTÁRIO DE SEÇÃO - Para separar grandes blocos de código
' =============================================================================

' TODO: Implementar validação adicional
' FIXME: Corrigir problema de performance
' NOTE: Esta função será descontinuada na v2.0
```

#### Arquivos de Configuração
- Use formato INI para configurações
- Comente todas as opções
- Forneça valores padrão sensatos
- Agrupe configurações logicamente

## Processo de Pull Request

### 1. Prepare sua Contribuição

```bash
# Crie uma branch para sua feature/fix
git checkout -b feature/minha-nova-feature

# Ou para correções
git checkout -b fix/corrigir-problema-especifico
```

### 2. Faça suas Mudanças

- Siga os padrões de código estabelecidos
- Adicione/atualize testes quando aplicável
- Atualize documentação relevante
- Teste suas mudanças em diferentes versões do Word

### 3. Commit suas Mudanças

```bash
# Use mensagens de commit descritivas
git add .
git commit -m "feat: adiciona funcionalidade X para melhorar Y

- Implementa algoritmo otimizado para processamento
- Adiciona validação de entrada robusta  
- Atualiza documentação com exemplos de uso"
```

#### Formato de Mensagens de Commit

Use o formato [Conventional Commits](https://www.conventionalcommits.org/):

- `feat:` nova funcionalidade
- `fix:` correção de bug
- `docs:` mudanças na documentação
- `style:` formatação, espaços em branco, etc.
- `refactor:` refatoração de código
- `perf:` melhorias de performance
- `test:` adição ou correção de testes
- `chore:` mudanças no processo de build, auxiliares, etc.

### 4. Abra o Pull Request

1. **Push sua branch**
   ```bash
   git push origin feature/minha-nova-feature
   ```

2. **Crie o PR no GitHub**
   - Use o template de PR disponível
   - Descreva claramente as mudanças
   - Referencie issues relacionadas
   - Adicione screenshots quando aplicável

3. **Aguarde Review**
   - Responda feedback construtivamente
   - Faça ajustes solicitados
   - Mantenha o PR atualizado com a branch main

## Reportar Problemas

### Informações Necessárias

Ao reportar um bug, inclua:

- **Versão do Word**: (ex: 2016, 2019, 365)
- **Versão do Windows**: (ex: Windows 10 21H2)
- **Versão do CHAINSAW**: (ex: 1.9.1-Alpha-8)
- **Passos para reproduzir**
- **Comportamento esperado vs atual**
- **Screenshots/logs** quando aplicável
- **Documento de teste** (sem dados sensíveis)

### Template de Bug Report

```markdown
**Versão do Ambiente:**
- Word: [versão]
- Windows: [versão]
- CHAINSAW: [versão]

**Descrição do Problema:**
[Descreva claramente o problema]

**Passos para Reproduzir:**
1. [Primeiro passo]
2. [Segundo passo]
3. [Terceiro passo]

**Comportamento Esperado:**
[O que deveria acontecer]

**Comportamento Atual:**
[O que está acontecendo]

**Logs/Screenshots:**
[Adicione informações adicionais]
```

## Comunidade e Conduta

### Código de Conduta

- **Seja respeitoso** com todos os participantes
- **Seja construtivo** em feedback e críticas
- **Seja paciente** com iniciantes
- **Seja colaborativo** e ajude outros contribuidores

### Canais de Comunicação

- **Issues**: Para bugs, features e discussões técnicas
- **Discussions**: Para perguntas gerais e ideias
- **Email**: Para questões sensíveis ou privadas

### Reconhecimento

Todos os contribuidores são reconhecidos no arquivo [`docs/CONTRIBUTORS.md`](CONTRIBUTORS.md). Sua contribuição, por menor que seja, é valorizada e registrada.

## Licença

Ao contribuir com este projeto, você concorda que suas contribuições serão licenciadas sob a mesma licença do projeto (Apache 2.0 modificada).

---

**Obrigado por contribuir com o CHAINSAW PROPOSITURAS!** 🎉

Sua contribuição ajuda a melhorar ferramentas para a comunidade legislativa brasileira.