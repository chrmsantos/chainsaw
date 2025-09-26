# Chainsaw Proposituras - Versão Simplificada

## 📋 Visão Geral

Versão simplificada do sistema de padronização de documentos legislativos para Microsoft Word.

**🎯 Objetivo**: Reduzir complexidade mantendo funcionalidade essencial
**📊 Tamanho**: ~150 linhas (vs. 7400+ da versão completa)
**⚡ Performance**: Execução rápida e simples

## 🚀 Instalação Rápida

### 1. Abrir o VBA no Word
- Pressione `Alt+F11`
- Ou vá em: Desenvolvedor > Visual Basic

### 2. Importar o Módulo Principal
- File > Import File
- Selecione: `chainsaw-simples.bas`

### 3. Importar Testes (Opcional)
- File > Import File  
- Selecione: `teste-simples.vba`

### 4. Testar
```vba
' Execute no Editor VBA:
Call Teste
```

## 📝 Como Usar

### Uso Básico
```vba
' 1. Abra um documento no Word
' 2. Execute no VBA:
Call PadronizarDocumento
```

### Teste com Documento de Exemplo
```vba
' 1. Criar documento de teste:
Call CriarDocumentoTeste

' 2. Padronizar o documento criado:
Call PadronizarDocumento
```

## ⚙️ Funcionalidades

### ✅ O que FAZ (Versão Simples)

1. **Formatação da Primeira Linha**
   - ✅ CAIXA ALTA automática
   - ✅ Negrito + Sublinhado
   - ✅ Centralizada

2. **Parágrafos 2-4**
   - ✅ Recuo esquerdo de 9cm
   - ✅ Sem recuo da primeira linha

3. **Padronização "Considerando"**
   - ✅ "considerando" → "CONSIDERANDO"
   - ✅ Aplicação de negrito

4. **Limpeza Básica**
   - ✅ Remove múltiplos espaços
   - ✅ Remove quebras de linha excessivas
   - ✅ Padroniza fonte (Times New Roman, 12pt)

5. **Interface Simples**
   - ✅ Confirmação antes de executar
   - ✅ Mensagens de progresso
   - ✅ Tratamento básico de erros

### ❌ O que NÃO FAZ (Removido da versão completa)

- ❌ Sistema complexo de configuração (INI files)
- ❌ Sistema de backup automático
- ❌ Sistema de logging detalhado
- ❌ Validações avançadas de documento
- ❌ Otimizações de performance complexas
- ❌ Sistema de permissões de arquivo
- ❌ Limpeza avançada de elementos visuais
- ❌ Configurações personalizáveis
- ❌ Sistema de recuperação de erros

## 🔧 Estrutura dos Arquivos

### Arquivos Principais (Simplificados)

```
src/
├── chainsaw-simples.bas     # Módulo principal (150 linhas)
├── teste-simples.vba        # Testes básicos (50 linhas)
└── [arquivos complexos]     # Versão original mantida
```

### Comparação de Tamanhos

| Arquivo | Versão Original | Versão Simples | Redução |
|---------|----------------|----------------|---------|
| Módulo Principal | 7,428 linhas | 150 linhas | **98% menor** |
| Arquivo de Teste | 233 linhas | 50 linhas | **78% menor** |
| **TOTAL** | **7,661 linhas** | **200 linhas** | **🎯 97% menor** |

## 🎯 Benefícios da Simplificação

### ✅ Vantagens

- **🚀 Mais Rápido**: Execução instantânea
- **🐛 Menos Bugs**: Código mais simples = menos erros
- **📖 Mais Fácil**: Entendimento imediato
- **🔧 Mais Fácil de Modificar**: 200 linhas vs. 7400
- **💾 Menor Consumo**: Menos memória e processamento
- **⚡ Instalação Rápida**: 2 arquivos vs. sistema complexo

### ⚠️ Limitações

- **🔒 Menos Configurável**: Sem arquivo INI personalizado
- **🛡️ Menos Proteção**: Sem backups automáticos
- **📊 Menos Logs**: Sem sistema de auditoria detalhada
- **🔍 Menos Validações**: Verificações básicas apenas

## 🛠️ Personalização

Para modificar comportamentos, edite diretamente o código:

### Alterar Recuo dos Parágrafos 2-4
```vba
' Na função FormatarParagrafos2a4, linha:
.LeftIndent = CentimetersToPoints(9)  ' Altere o valor 9
```

### Alterar Fonte Padrão
```vba
' Na função LimparFormatacao, linhas:
.Size = 12                    ' Altere o tamanho
.Name = "Times New Roman"     ' Altere a fonte
```

### Adicionar Mais Parágrafos com Recuo
```vba
' Na função FormatarParagrafos2a4, altere:
For i = 2 To 4    ' Para: For i = 2 To 6 (exemplo)
```

## 🧪 Testes Disponíveis

### Testes Básicos
```vba
Call TesteSimples          ' Teste de funcionamento
Call CriarDocumentoTeste   ' Criar exemplo
Call TestarPadronizacao    ' Testar com documento atual
```

### Fluxo de Teste Recomendado
```vba
' 1. Teste básico
Call TesteSimples

' 2. Criar documento de exemplo  
Call CriarDocumentoTeste

' 3. Testar padronização
Call PadronizarDocumento

' 4. Verificar resultado visual no documento
```

## 🆚 Quando Usar Cada Versão

### 🎯 Use a Versão SIMPLES se:
- ✅ Você quer funcionalidade básica
- ✅ Precisa de algo rápido e confiável
- ✅ Não precisa de configurações complexas
- ✅ Quer código fácil de entender/modificar
- ✅ Instalação deve ser rápida

### 🔧 Use a Versão COMPLETA se:
- ⚙️ Precisa de configurações detalhadas
- 🛡️ Precisa de backups automáticos
- 📊 Precisa de logs detalhados
- 🔍 Precisa de validações avançadas
- 🏢 Uso em ambiente corporativo complexo

## 📞 Suporte

Para a versão simplificada:
- **Código fonte**: Apenas 200 linhas - fácil de debugar
- **Problemas**: Geralmente relacionados a VBA básico
- **Modificações**: Editar diretamente o código

## 🔄 Migração

### Da Versão Completa para Simples
1. **Backup**: Salve configurações atuais se necessário
2. **Remover**: Exclua módulos complexos
3. **Importar**: Adicione `chainsaw-simples.bas`
4. **Testar**: Execute `Call Teste`

### Da Versão Simples para Completa
1. **Manter**: Versão simples como backup
2. **Importar**: Adicione `chainsaw0.bas`
3. **Configurar**: Ajuste configurações conforme necessário

---

**📌 Resumo**: Versão simplificada oferece 80% da funcionalidade com 3% da complexidade!