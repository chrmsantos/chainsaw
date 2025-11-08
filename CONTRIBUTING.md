# Guia de Contribuição - CHAINSAW

## Bem-vindo!

Obrigado por considerar contribuir para o projeto CHAINSAW! Este documento fornece diretrizes para contribuições.

## Índice

- [Código de Conduta](#código-de-conduta)
- [Como Contribuir](#como-contribuir)
- [Padrões de Código](#padrões-de-código)
- [Processo de Pull Request](#processo-de-pull-request)
- [Reportar Bugs](#reportar-bugs)
- [Sugerir Melhorias](#sugerir-melhorias)

## Código de Conduta

Este projeto segue um código de conduta básico:

- Seja respeitoso e profissional
- Aceite críticas construtivas
- Foque no que é melhor para o projeto
- Seja colaborativo

## Como Contribuir

### Configuração do Ambiente

1. **Fork o repositório** no GitHub
2. **Clone seu fork**:
   ```bash
   git clone https://github.com/seu-usuario/chainsaw.git
   cd chainsaw
   ```
3. **Configure o upstream**:
   ```bash
   git remote add upstream https://github.com/chrmsantos/chainsaw.git
   ```

### Fluxo de Trabalho

1. **Crie uma branch** para sua contribuição:
   ```bash
   git checkout -b feature/minha-contribuicao
   ```
2. **Faça suas alterações** seguindo os padrões de código
3. **Execute os testes**:
   ```powershell
   cd tests
   .\Run-Tests.ps1
   ```
4. **Valide encoding** (sem emojis):
   ```powershell
   powershell -ExecutionPolicy Bypass -File tests\check-encoding.ps1
   ```
5. **Commit suas alterações**:
   ```bash
   git add .
   git commit -m "feat: adiciona nova funcionalidade"
   ```
6. **Push para seu fork**:
   ```bash
   git push origin feature/minha-contribuicao
   ```
7. **Abra um Pull Request** no GitHub

## Padrões de Código

### Encoding e Caracteres

**IMPORTANTE**: Este projeto NÃO permite emojis no código-fonte ou documentação.

#### Regras de Encoding

- **Encoding**: UTF-8 (com ou sem BOM)
- **Line Endings**: CRLF (Windows)
- **Indentação**: 4 espaços (não tabs)
- **Caracteres Permitidos**:
  - ASCII básico (0x00-0x7F)
  - Caracteres acentuados portugueses (á, ç, ê, õ, etc.)
  - Pontuação padrão
  
#### Caracteres PROIBIDOS

- [PROIBIDO] Emojis (     ⚠ etc.)
- [PROIBIDO] Símbolos decorativos Unicode
- [PROIBIDO] Caracteres de controle inválidos

**Use alternativas textuais**:
- Em vez de  use `[OK]`
- Em vez de  ou  use `[ERRO]`
- Em vez de ⚠ use `[AVISO]`
- Em vez de   etc. não use nada ou descreva em texto

#### Validação Automática

O projeto inclui validação automática via git hooks:

```powershell
# Validar manualmente antes de commitar
powershell -ExecutionPolicy Bypass -File tests\check-encoding.ps1

# Remover emojis se detectados
powershell -ExecutionPolicy Bypass -File tests\remove-emojis-bytes.ps1
```

### PowerShell

#### Estilo de Código

```powershell
# Use nomes descritivos
$nomeVariavel = "valor"  # Bom
$x = "valor"             # Ruim

# Funcoes com verbos aprovados
function Get-Configuracao { }  # Bom
function Pegar-Config { }      # Ruim

# Comentarios claros
# Faz X porque Y
$resultado = Get-Something

# Tratamento de erros
try {
    # Codigo que pode falhar
} catch {
    Write-Error "Mensagem clara: $_"
}
```

#### Padrões PowerShell

- Use `$ErrorActionPreference = "Stop"` no início dos scripts
- Prefira cmdlets nativos sobre comandos externos
- Use `Write-Host` para output do usuário, `Write-Verbose` para debug
- Documente funções com comentários `# === ===`
- Use `-WhatIf` e `-Confirm` para operações destrutivas

### Markdown

#### Formatação

```markdown
# Título Nível 1

## Título Nível 2

### Título Nível 3

- Lista com hífen
- Use espaço após o hífen

1. Listas numeradas
2. Com ponto após número

**Negrito** para ênfase forte
*Itálico* para ênfase leve

`código inline` com crases
```

#### Blocos de Código

Sempre especifique a linguagem:

````markdown
```powershell
# Código PowerShell
Get-Process
```

```bash
# Código Bash
ls -la
```
````

### VBA

- Use indentação de 4 espaços
- Comente blocos lógicos
- Use nomes descritivos para variáveis
- Prefixe variáveis com tipo: `str`, `int`, `obj`, etc.

```vba
' Comentario descritivo
Public Function MinhaFuncao(ByVal strParametro As String) As Boolean
    ' Implementacao
    On Error GoTo ErrorHandler
    
    MinhaFuncao = True
    Exit Function
    
ErrorHandler:
    MinhaFuncao = False
End Function
```

## Processo de Pull Request

### Antes de Abrir PR

- [ ] Todos os testes passam (`.\Run-Tests.ps1`)
- [ ] Validação de encoding OK (`.\check-encoding.ps1`)
- [ ] Código segue os padrões deste guia
- [ ] Documentação atualizada (se aplicável)
- [ ] CHANGELOG.md atualizado

### Template de PR

```markdown
## Descrição

Breve descrição das alterações.

## Tipo de Mudança

- [ ] Correção de bug
- [ ] Nova funcionalidade
- [ ] Mudança que quebra compatibilidade
- [ ] Documentação

## Checklist

- [ ] Testes passam
- [ ] Sem emojis (validação OK)
- [ ] Documentação atualizada
- [ ] CHANGELOG.md atualizado

## Testes

Descreva os testes realizados.
```

### Revisão

- PRs serão revisados dentro de 1-3 dias úteis
- Responda aos comentários prontamente
- Seja aberto a sugestões
- Faça alterações solicitadas em commits separados

## Reportar Bugs

### Antes de Reportar

1. **Verifique issues existentes** - Seu bug pode já ter sido reportado
2. **Tente a versão mais recente** - O bug pode já estar corrigido
3. **Colete informações** - Versão, SO, passos para reproduzir

### Template de Bug Report

```markdown
**Descrição do Bug**
Descrição clara e concisa.

**Passos para Reproduzir**
1. Vá para '...'
2. Execute '...'
3. Veja erro

**Comportamento Esperado**
O que deveria acontecer.

**Comportamento Atual**
O que realmente acontece.

**Ambiente**
- SO: Windows 10/11
- PowerShell: 5.1 / 7.x
- Word: 2016 / 2019 / 365
- Versão CHAINSAW: 2.0.x

**Logs**
Cole logs relevantes aqui.

**Screenshots**
Se aplicável.
```

## Sugerir Melhorias

### Template de Feature Request

```markdown
**O Problema**
Descreva o problema que esta funcionalidade resolveria.

**Solução Proposta**
Descreva como você imagina a solução.

**Alternativas Consideradas**
Outras abordagens que você considerou.

**Contexto Adicional**
Qualquer informação relevante.
```

## Testes

### Executar Testes

```powershell
# Todos os testes
cd tests
.\Run-Tests.ps1

# Testes específicos
Invoke-Pester -Script .\VBA.Tests.ps1
Invoke-Pester -Script .\Installation.Tests.ps1
Invoke-Pester -Script .\Encoding.Tests.ps1
```

### Escrever Testes

Use Pester 3.4.0:

```powershell
Describe 'Meu Componente' {
    Context 'Quando faz X' {
        It 'Deve retornar Y' {
            $resultado = Get-Algo
            $resultado | Should Be "Y"
        }
    }
}
```

## Documentação

### Atualizar Documentação

- Mantenha README.md atualizado
- Atualize docs/ para novos recursos
- Adicione comentários no código
- Atualize CHANGELOG.md

### Estilo de Documentação

- Use português brasileiro
- Seja claro e conciso
- Forneça exemplos
- Use formatação Markdown correta
- **NÃO use emojis**

## Versionamento

Este projeto segue [Semantic Versioning](https://semver.org/):

- **MAJOR** (X.0.0): Mudanças que quebram compatibilidade
- **MINOR** (0.X.0): Novas funcionalidades (compatíveis)
- **PATCH** (0.0.X): Correções de bugs

## Convenções de Commit

Use mensagens descritivas:

```
feat: adiciona suporte para Word 2024
fix: corrige erro de encoding em UTF-8
docs: atualiza guia de instalacao
test: adiciona testes para modulo X
refactor: melhora performance de Y
style: remove emojis e padroniza encoding
```

## Licença

Ao contribuir, você concorda que suas contribuições serão licenciadas sob a mesma licença do projeto.

## Perguntas?

- Abra uma issue com a tag `question`
- Entre em contato: chrmsantos@protonmail.com

## Agradecimentos

Obrigado por contribuir para o CHAINSAW! Sua ajuda é muito apreciada.

---

**Lembre-se**: Mantenha o código limpo, testado e **sem emojis**!  <- Ops, até aqui não! Use texto simples.
