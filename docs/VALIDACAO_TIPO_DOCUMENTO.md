# Validação de Tipo de Documento

## Resumo

Implementação de validação automática do tipo de documento antes do processamento, garantindo que apenas **Indicações**, **Requerimentos** ou **Moções** sejam processados (ou processados com confirmação explícita do usuário).

---

## Funcionalidade

### Quando a Validação Ocorre

A validação é executada **logo no início** do processamento, antes de qualquer modificação no documento:

```
1. [OK] Verifica versão do Word
2. [OK] Verifica integridade do documento
3. [OK] Inicializa sistema de logs
4. [NEW] VALIDA TIPO DE DOCUMENTO ← NOVA VALIDAÇÃO
5. [OK] Cria backup
6. [OK] Inicia formatação...
```

### Tipos de Documento Válidos

A validação aceita automaticamente documentos que iniciam com (case insensitive):

- [OK] **INDICAÇÃO**
- [OK] **REQUERIMENTO**
- [OK] **MOÇÃO**

---

## Cenários de Uso

### Cenário 1: Documento Válido (INDICAÇÃO)

**Primeira linha do documento:**
```
INDICAÇÃO N.º 123/2024
```

**Resultado:**
- [OK] Validação passa automaticamente
- [OK] Log: "Documento identificado como: INDICAÇÃO"
- [OK] Processamento continua normalmente

---

### Cenário 2: Documento Válido (REQUERIMENTO)

**Primeira linha do documento:**
```
REQUERIMENTO
```

**Resultado:**
- [OK] Validação passa automaticamente
- [OK] Log: "Documento identificado como: REQUERIMENTO"
- [OK] Processamento continua normalmente

---

### Cenário 3: Documento Válido (MOÇÃO)

**Primeira linha do documento:**
```
Moção n.º 45/2024
```

**Resultado:**
- [OK] Validação passa automaticamente (case insensitive)
- [OK] Log: "Documento identificado como: MOÇÃO"
- [OK] Processamento continua normalmente

---

### Cenário 4: Tipo Não Reconhecido

**Primeira linha do documento:**
```
PROJETO DE LEI N.º 789/2024
```

**Resultado:**
- [!] Exibe mensagem ao usuário:

```
╔═══════════════════════════════════════════════════════════╗
║   Tipo de Documento Não Reconhecido                       ║
╠═══════════════════════════════════════════════════════════╣
║                                                           ║
║   O documento parece não ser uma Indicação,               ║
║   Requerimento ou Moção.                                  ║
║                                                           ║
║   Primeira palavra identificada: "PROJETO"                ║
║                                                           ║
║   Tipos válidos esperados:                                ║
║   • INDICAÇÃO                                             ║
║   • REQUERIMENTO                                          ║
║   • MOÇÃO                                                 ║
║                                                           ║
║   Possíveis causas:                                       ║
║   • Erro de grafia no título da propositura               ║
║   • Documento de tipo diferente                           ║
║   • Formatação incorreta do título                        ║
║                                                           ║
║   Deseja cancelar ou prosseguir mesmo assim?              ║
║                                                           ║
║              [ Sim ]         [ Não ]                      ║
╚═══════════════════════════════════════════════════════════╝
```

**Se o usuário clicar em "Sim" (prosseguir):**
- [!] Log: "Usuário optou por prosseguir com documento tipo: PROJETO"
- [OK] Processamento continua

**Se o usuário clicar em "Não" (cancelar):**
- [X] Log: "Usuário cancelou processamento - tipo de documento não reconhecido: PROJETO"
- [X] Status bar: "Cancelado: tipo de documento não reconhecido"
- [X] Processamento é interrompido

---

### Cenário 5: Erro de Grafia

**Primeira linha do documento:**
```
INDCAÇÃO N.º 123/2024
```
(faltou a letra "I" em INDICAÇÃO)

**Resultado:**
- [!] Primeira palavra identificada: "INDCAÇÃO"
- [!] Exibe mensagem ao usuário (similar ao Cenário 4)
-  Usuário pode perceber o erro e cancelar para corrigir
- [OK] Ou pode prosseguir se for intencional

---

### Cenário 6: Documento Vazio

**Documento sem conteúdo ou apenas parágrafos vazios**

**Resultado:**
- [!] Exibe mensagem ao usuário:

```
╔═══════════════════════════════════════════════════════════╗
║   Tipo de Documento Não Identificado                      ║
╠═══════════════════════════════════════════════════════════╣
║                                                           ║
║   Não foi possível identificar o tipo do documento.       ║
║                                                           ║
║   O documento parece estar vazio ou sem texto válido.     ║
║                                                           ║
║   Deseja cancelar ou prosseguir mesmo assim?              ║
║                                                           ║
║              [ Sim ]         [ Não ]                      ║
╚═══════════════════════════════════════════════════════════╝
```

---

## Detalhes Técnicos

### Função Principal: `ValidateDocumentType`

**Localização**: `src/chainsaw.bas` (após `GetFirstWordOfDocument`)

**Lógica de Validação**:

```vba
1. Obtém primeira palavra do documento via GetFirstWordOfDocument()
2. Se primeira palavra vazia:
   → Alerta "documento vazio"
   → Usuário decide: prosseguir ou cancelar
3. Se primeira palavra = "INDICAÇÃO" OU "REQUERIMENTO" OU "MOÇÃO":
   → Validação OK (retorna True)
   → Log informativo
4. Se primeira palavra diferente:
   → Alerta "tipo não reconhecido"
   → Mostra primeira palavra identificada
   → Lista tipos válidos
   → Explica possíveis causas
   → Usuário decide: prosseguir ou cancelar
```

### Chamada na Função Principal

**Localização**: `PadronizarDocumentoMain()` (linha ~220)

**Momento**: Logo após validação de integridade e inicialização de logs

```vba
' Valida o tipo de documento (INDICAÇÃO, REQUERIMENTO ou MOÇÃO)
If Not ValidateDocumentType(doc) Then
    Application.StatusBar = "Cancelado: tipo de documento não reconhecido"
    LogMessage "Processamento cancelado pelo usuário após validação de tipo", LOG_LEVEL_INFO
    Exit Sub
End If
```

---

## Características

### [OK] Segurança

- Comparação **case insensitive** (aceita "INDICAÇÃO", "Indicação", "indicação")
- Reutiliza função `GetFirstWordOfDocument` já validada
- Remove pontuação automaticamente (`:`, `,`, `.`, `;`)
- Tratamento completo de erros

### [OK] User Experience

- Mensagens claras e informativas
- Explica possíveis causas do problema
- Lista tipos válidos esperados
- Mostra a palavra identificada
- Permite ao usuário decidir (não é bloqueante por padrão)
- Status bar atualizada em caso de cancelamento

### [OK] Logging

- Log de tipo identificado (INFO)
- Log de decisão do usuário (WARNING/INFO)
- Log de erros (ERROR)
- Rastreabilidade completa das decisões

---

## Exemplos de Log

### Tipo Válido Identificado
```
[INFO] Documento identificado como: INDICAÇÃO
```

### Tipo Não Reconhecido - Usuário Prossegue
```
[WARNING] Usuário optou por prosseguir com documento tipo: PROJETO
```

### Tipo Não Reconhecido - Usuário Cancela
```
[INFO] Usuário cancelou processamento - tipo de documento não reconhecido: PROJETO
[INFO] Processamento cancelado pelo usuário após validação de tipo
```

### Documento Vazio - Usuário Cancela
```
[INFO] Usuário cancelou - documento de tipo não identificado
[INFO] Processamento cancelado pelo usuário após validação de tipo
```

---

## Benefícios

### 1. Prevenção de Erros
- Detecta erros de grafia no título antes do processamento
- Evita processar documentos do tipo errado
- Reduz retrabalho

### 2. Flexibilidade
- Não bloqueia completamente outros tipos de documento
- Usuário pode prosseguir se necessário
- Útil para testes ou casos especiais

### 3. Rastreabilidade
- Todos os logs registram o tipo de documento
- Decisões do usuário são registradas
- Facilita auditoria e troubleshooting

### 4. Educação do Usuário
- Mensagens explicam os tipos válidos
- Alerta sobre possíveis causas de problema
- Ajuda a manter padrões de nomenclatura

---

## Manutenção

### Para Adicionar Novos Tipos Válidos

Editar a função `ValidateDocumentType`:

```vba
' Linha atual:
If firstWord = "INDICAÇÃO" Or firstWord = "REQUERIMENTO" Or firstWord = "MOÇÃO" Then

' Adicionar novo tipo (ex: "OFÍCIO"):
If firstWord = "INDICAÇÃO" Or firstWord = "REQUERIMENTO" Or firstWord = "MOÇÃO" Or firstWord = "OFÍCIO" Then
```

E atualizar a mensagem de lista de tipos válidos:

```vba
validTypes = "• INDICAÇÃO" & vbCrLf & "• REQUERIMENTO" & vbCrLf & "• MOÇÃO" & vbCrLf & "• OFÍCIO"
```

---

## Integração com Substituições Condicionais

Esta validação trabalha em conjunto com as **substituições condicionais de texto**:

1. **ValidateDocumentType** → Valida tipo no início (pode cancelar)
2. **GetFirstWordOfDocument** → Usado para determinar substituições
3. **FormatSecondParagraph** → Aplica substituições baseadas no tipo

---

## Status

- [OK] Implementado
- [OK] Testado (sem erros de sintaxe)
- [OK] Documentado
- [OK] Integrado com fluxo principal
- [OK] Pronto para uso

---

## Histórico

| Data | Versão | Descrição |
|------|--------|-----------|
| 2024 | 1.0 | Implementação inicial da validação de tipo de documento |
