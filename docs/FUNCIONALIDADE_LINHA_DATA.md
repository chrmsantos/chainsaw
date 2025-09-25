# 📅 Funcionalidade de Substituição de Linha de Data

## 🎯 Objetivo

Detecta automaticamente a linha de data localizada **3 linhas acima** do parágrafo "- Vereador -" e a substitui por uma linha padronizada com a data atual por extenso.

## 🔍 Critérios de Detecção

Para que uma linha seja considerada válida para substituição, ela deve atender **TODOS** os critérios:

### ✅ **Critério 1: Início da Linha**
- Deve iniciar com "**Palácio**" ou "**Plenário**" (não diferencia maiúsculas/minúsculas)

### ✅ **Critério 2: Presença de Mês**
- Deve conter nome de mês por extenso em português:
  - janeiro, fevereiro, março, abril, maio, junho
  - julho, agosto, setembro, outubro, novembro, dezembro

### ✅ **Critério 3: Tamanho do Texto**
- Deve ter **menos de 20 palavras**

### ✅ **Critério 4: Final da Linha**
- Deve terminar com **número** seguido ou não por **ponto final**
  - Exemplos válidos: `2025`, `2025.`

## 📍 Posicionamento

A funcionalidade busca especificamente na **3ª linha acima** do parágrafo que contém "- Vereador -".

**Estrutura do documento:**
```
[Linha qualquer]          ← 4ª linha acima
[Linha qualquer]          ← 3ª linha acima (AQUI É VERIFICADO)
[Linha qualquer]          ← 2ª linha acima  
[Linha qualquer]          ← 1ª linha acima
- Vereador -              ← Referência para busca
```

## 🔄 Substituição

### **Texto Original (Exemplos Válidos):**
- `Palácio Municipal "Antônio Carlos Magalhães", 15 de setembro de 2025.`
- `Plenário Dr. Tancredo Neves, 23 de dezembro de 2024`
- `PALÁCIO MUNICIPAL, 05 de janeiro de 2025.`

### **Texto Substituído:**
```
Plenário "Dr. Tancredo Neves", [DIA] de [MÊS] de [ANO].
```

**Exemplo com data atual (25/09/2025):**
```
Plenário "Dr. Tancredo Neves", 25 de setembro de 2025.
```

## ⚙️ Configuração

### **Arquivo de Configuração**
No arquivo `config/chainsaw-config.ini`, seção `[LIMPEZA]`:

```ini
replace_date_line_before_vereador=true    # Ativa/desativa a funcionalidade
```

### **Valores Aceitos:**
- `true` - Funcionalidade **ativada** (padrão)
- `false` - Funcionalidade **desativada**

## 📝 Exemplos Práticos

### ✅ **Exemplo 1: Substituição Bem-sucedida**

**Antes:**
```
Projeto de Lei nº 123/2025

AUTOR: Fulano de Tal

Palácio Municipal de Nova Iguaçu, 15 de setembro de 2025.

- Vereador -
```

**Depois:**
```
Projeto de Lei nº 123/2025

AUTOR: Fulano de Tal

Plenário "Dr. Tancredo Neves", 25 de setembro de 2025.

- Vereador -
```

### ❌ **Exemplo 2: Não Substituído (Critérios Não Atendidos)**

**Caso 1 - Não inicia com "Palácio" ou "Plenário":**
```
Casa de Leis de Nova Iguaçu, 15 de setembro de 2025.  ← Não será substituído
```

**Caso 2 - Não contém mês por extenso:**
```
Palácio Municipal de Nova Iguaçu, 15/09/2025.  ← Não será substituído
```

**Caso 3 - Muito longo (>20 palavras):**
```
Palácio Municipal "Prefeito Antônio Carlos Magalhães" da cidade de Nova Iguaçu do Estado do Rio de Janeiro, 15 de setembro de 2025.  ← Não será substituído
```

**Caso 4 - Não termina com número:**
```
Palácio Municipal de Nova Iguaçu, 15 de setembro.  ← Não será substituído
```

## 🚨 Comportamento de Erro

### **Linha Não Encontrada**
Se nenhuma linha atender aos critérios, o sistema:

1. **Registra no log**: "Nenhuma linha de data foi encontrada que atenda aos critérios"
2. **Exibe mensagem ao usuário**:
   ```
   A linha da data não foi encontrada.

   Critérios de busca:
   • Deve estar na 3ª linha acima de '- Vereador -'
   • Deve iniciar com 'Palácio' ou 'Plenário'  
   • Deve conter nome de mês por extenso
   • Deve ter menos de 20 palavras
   • Deve terminar com número seguido ou não por ponto
   ```

### **Parágrafos Insuficientes**
Se não houver 3 parágrafos acima de "- Vereador -":
- **Log**: "Não foi possível encontrar a 3ª linha acima de '- Vereador -'"
- **Não exibe erro ao usuário** (situação normal em documentos curtos)

## 🔧 Funções Técnicas

### **Função Principal**
- `ProcessDateLineReplacement()` - Processa toda a lógica de busca e substituição

### **Funções Auxiliares**
- `IsValidDateLine()` - Valida se a linha atende aos critérios
- `ContainsMonthName()` - Verifica presença de mês por extenso  
- `EndsWithNumberAndOptionalPeriod()` - Valida final numérico
- `GenerateStandardDateLine()` - Gera linha padronizada
- `GetCurrentDateExtended()` - Retorna data atual por extenso

## 📊 Log de Atividade

### **Mensagens de Log Geradas:**

**Sucesso:**
```
INFO: Encontrado parágrafo '- Vereador -' no índice: 1245
INFO: Substituindo linha de data: 'Palácio Municipal, 15 de setembro de 2025.' por 'Plenário "Dr. Tancredo Neves", 25 de setembro de 2025.'
INFO: Processamento de linha de data concluído: 1 substituições realizadas
```

**Critérios não atendidos:**
```
INFO: Linha 3 acima de '- Vereador -' não atende aos critérios: 'Casa de Leis, 15/09/2025'
INFO: Nenhuma linha de data foi encontrada que atenda aos critérios especificados
```

**Erros:**
```
ERROR: Erro no processamento de linha de data: [Descrição do erro]
```

## 🎛️ Ativação/Desativação

### **Via Código VBA:**
```vba
' Ativar funcionalidade
Config.replaceDateLineBeforeVereador = True

' Desativar funcionalidade  
Config.replaceDateLineBeforeVereador = False
```

### **Via Arquivo de Configuração:**
```ini
[LIMPEZA]
replace_date_line_before_vereador=false  # Desativa
replace_date_line_before_vereador=true   # Ativa (padrão)
```

---

## 💡 **Dicas de Uso**

1. **Backup**: Sempre faça backup antes de usar em documentos importantes
2. **Teste**: Teste em documento de exemplo primeiro
3. **Verificação**: Confira o resultado após o processamento
4. **Configuração**: Desative se não precisar desta funcionalidade

---

**Versão da Documentação:** 1.0  
**Data:** 25 de setembro de 2025  
**Funcionalidade:** Chainsaw Proposituras v1.9.1-Alpha-8