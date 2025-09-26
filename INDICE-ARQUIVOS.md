# 📁 Índice dos Arquivos - Chainsaw Proposituras

## 🎯 VERSÕES DISPONÍVEIS

### ⚡ VERSÃO SIMPLES (Recomendada - 200 linhas)
```
src/chainsaw-simples.bas     # Módulo principal (150 linhas)
src/teste-simples.vba        # Testes básicos (50 linhas)
```
**Documentação:**
```
docs/README-VERSAO-SIMPLES.md    # Guia completo da versão simples
README-SIMPLIFICADO.md           # Overview das duas versões
```

### 🔧 VERSÃO COMPLETA (Avançada - 7.400+ linhas)
```
src/chainsaw0.bas            # Módulo principal (7.428 linhas)
src/teste-execucao.vba       # Testes completos (233 linhas)
```
**Documentação:**
```
docs/SOLUCAO_PERMISSAO_CONFIG.md           # Sistema de permissões
docs/CORRECAO_TESTE_LOADCONFIGURATION.md   # Correções de teste
docs/SOLUCAO_TESTE_LOADCONFIGURATION_NOT_DEFINED.md  # Troubleshooting
```

## 🚀 QUAL ARQUIVO USAR?

### Para Início Rápido (90% dos usuários):
1. **Baixe**: `src/chainsaw-simples.bas`
2. **Teste**: `src/teste-simples.vba`
3. **Leia**: `docs/README-VERSAO-SIMPLES.md`

### Para Uso Corporativo Avançado:
1. **Baixe**: `src/chainsaw0.bas`  
2. **Configure**: `config/chainsaw-config.ini`
3. **Teste**: `src/teste-execucao.vba`
4. **Leia**: Documentação na pasta `docs/`

## 📋 GUIA DE INSTALAÇÃO RÁPIDA

### Versão Simples (2 minutos):
```
1. Alt+F11 (Editor VBA)
2. File > Import File > chainsaw-simples.bas
3. Execute: Call Teste
4. Use: Call PadronizarDocumento
```

### Versão Completa (30 minutos):
```
1. Alt+F11 (Editor VBA)  
2. File > Import File > chainsaw0.bas
3. Configure chainsaw-config.ini (se necessário)
4. Execute: Call TesteCompleto
5. Use: Call PadronizarDocumentoMain
```

## 🔍 ESTRUTURA COMPLETA DO PROJETO

```
chainsaw-proposituras/
│
├── 📁 src/                          # Código fonte
│   ├── ⚡ chainsaw-simples.bas      # Versão simples (150 linhas)
│   ├── ⚡ teste-simples.vba         # Testes simples (50 linhas)
│   ├── 🔧 chainsaw0.bas             # Versão completa (7.428 linhas)
│   └── 🔧 teste-execucao.vba        # Testes completos (233 linhas)
│
├── 📁 docs/                         # Documentação
│   ├── 📖 README-VERSAO-SIMPLES.md
│   ├── 📖 GUIA-SIMPLIFICACAO.md
│   ├── 🔧 SOLUCAO_PERMISSAO_CONFIG.md
│   ├── 🔧 CORRECAO_TESTE_LOADCONFIGURATION.md
│   └── 🔧 SOLUCAO_TESTE_LOADCONFIGURATION_NOT_DEFINED.md
│
├── 📁 config/                       # Configurações (apenas versão completa)
│   └── chainsaw-config.ini
│
├── 📄 README-SIMPLIFICADO.md        # Overview principal
├── 📄 README.md                     # README original
└── 📄 INDICE-ARQUIVOS.md           # Este arquivo
```

## 🎯 DECISÃO RÁPIDA

### "Qual arquivo devo usar?"

#### Se você quer:
- ✅ **Funcionar imediatamente** → `chainsaw-simples.bas`
- ✅ **Código simples de entender** → `chainsaw-simples.bas`  
- ✅ **Modificar facilmente** → `chainsaw-simples.bas`
- ✅ **Performance máxima** → `chainsaw-simples.bas`

#### Se você precisa:
- 🔧 **Configurações avançadas** → `chainsaw0.bas`
- 🔧 **Sistema de backup automático** → `chainsaw0.bas`
- 🔧 **Logs detalhados** → `chainsaw0.bas`
- 🔧 **Validações complexas** → `chainsaw0.bas`

## 🧪 TESTE RÁPIDO

### Para testar a versão simples:
```vba
' 1. Importe chainsaw-simples.bas
' 2. Execute:
Call Teste                    # Verificar funcionamento
Call CriarDocumentoTeste      # Criar exemplo
Call PadronizarDocumento      # Testar padronização
```

### Para testar a versão completa:
```vba
' 1. Importe chainsaw0.bas
' 2. Execute:
Call TesteCompleto           # Bateria completa de testes
Call PadronizarDocumentoMain # Execução principal
```

## 📊 COMPARAÇÃO RÁPIDA

| Característica | Simples | Completa |
|----------------|---------|----------|
| **Linhas** | 200 | 7.400+ |
| **Arquivos** | 2 | 15+ |
| **Instalação** | 2 min | 30 min |
| **Configuração** | No código | Arquivo INI |
| **Performance** | ⚡ Instantânea | 🔧 5-30 seg |
| **Manutenção** | 🎯 Trivial | 🔧 Complexa |
| **Uso ideal** | Pessoal/Simples | Corporativo |

---

**🎯 RECOMENDAÇÃO**: 90% dos usuários devem começar com `chainsaw-simples.bas`

**📞 SUPORTE**: Para versão simples, o código é auto-explicativo (200 linhas)