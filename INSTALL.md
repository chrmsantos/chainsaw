# CHAINSAW PROPOSITURAS - Instalação Rápida

## 🚀 Instalação Automatizada (Recomendada)

### 1. Download
Baixe todos os arquivos do projeto em: <https://github.com/chrmsantos/chainsaw-proposituras>

### 2. Execução do Instalador
Abra o PowerShell como Administrador e execute:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\install-chainsaw-proposituras.ps1
```

### 3. Parâmetros do Instalador

```powershell
# Instalação padrão
.\install-chainsaw-proposituras.ps1

# Instalação customizada
.\install-chainsaw-proposituras.ps1 -InstallPath "C:\MinhaPasta" -AllUsers -Silent

# Apenas verificar compatibilidade
.\install-chainsaw-proposituras.ps1 -CheckOnly
```

## 🛠️ Instalação Manual

### Pré-requisitos
- Microsoft Word 2010 ou superior
- Windows 7/8/10/11
- Macros habilitadas no Word

### Passos

1. **Criar estrutura de pastas:**
   ```
   CHAINSAW-PROPOSITURAS/
   ├── src/
   ├── private/
   │   ├── header/
   │   ├── backups/
   │   └── logs/
   └── docs/
   ```

2. **Copiar arquivos:**
   - `src/Módulo1.bas` → Pasta src
   - `private/header/stamp.png` → Pasta header
   - Documentação → Pasta docs

3. **Instalar módulo VBA:**
   - Abrir Word
   - Alt+F11 (Editor VBA)
   - Arquivo → Importar
   - Selecionar `Módulo1.bas`

4. **Configurar segurança:**
   - Arquivo → Opções → Central de Confiabilidade
   - Configurações de Macro → "Desabilitar todas as macros com notificação"

## ⚡ Uso Rápido

### Executar Padronização
1. Abrir documento no Word
2. Alt+F8 → Executar Macro
3. Selecionar: `PadronizarDocumentoMain`
4. Confirmar execução

### Atalhos Disponíveis (após instalação automatizada)
- **Área de Trabalho:** "Chainsaw Proposituras"
- **Menu Iniciar:** Programas → Chainsaw Proposituras

## 🔧 Configurações Principais

### Tipos de Documento Suportados
- ✅ INDICAÇÃO
- ✅ REQUERIMENTO  
- ✅ MOÇÃO

### Formatações Aplicadas
- ✅ Margens institucionais (4.6/2/3/3 cm)
- ✅ Fonte Arial 12pt, espaçamento 1.4
- ✅ Cabeçalho com logotipo
- ✅ Numeração de páginas
- ✅ Formatação de parágrafos especiais
- ✅ Limpeza de elementos visuais desnecessários

## 📋 Autotexto Instalado

| Código | Resultado |
|--------|-----------|
| `indicacao` | INDICAÇÃO Nº $NUMERO$/$ANO$ |
| `requerimento` | REQUERIMENTO Nº $NUMERO$/$ANO$ |
| `mocao` | MOÇÃO Nº $NUMERO$/$ANO$ |
| `considerando` | CONSIDERANDO que |
| `justificativa` | JUSTIFICATIVA |
| `vereador` | - VEREADOR - |

## 🔒 Segurança

### Configurações Recomendadas
- Macros com notificação habilitada
- Pasta do projeto como local confiável
- Antivírus atualizado
- Backups automáticos ativos

### Validações do Sistema
- ✅ Verificação de versão do Word
- ✅ Validação de integridade do documento
- ✅ Backup automático antes de modificações
- ✅ Log detalhado de operações
- ✅ Recuperação de emergência

## 📁 Estrutura de Arquivos

```
CHAINSAW-PROPOSITURAS/
├── src/
│   └── Módulo1.bas              # Código VBA principal
├── private/
│   ├── header/
│   │   └── stamp.png            # Logotipo institucional
│   ├── backups/                 # Backups automáticos
│   └── logs/                    # Arquivos de log
├── templates/                   # Templates personalizados
├── docs/                        # Documentação adicional
├── README.md                    # Documentação principal
├── SECURITY.md                  # Política de segurança
├── MACRO_SECURITY_POLICY.md     # Política corporativa
├── LICENSE                      # Licença Apache 2.0
├── install-chainsaw-proposituras.ps1  # Instalador
└── install-config.ini           # Configurações
```

## 🆘 Solução de Problemas

### Erro: "Macro não encontrada"
1. Verificar se módulo foi importado corretamente
2. Reabrir Word
3. Verificar nome da macro: `PadronizarDocumentoMain`

### Erro: "Acesso negado ao VBA"
1. Word → Opções → Central de Confiabilidade
2. Configurações de Macro
3. Habilitar "Acesso ao modelo de objeto do projeto VBA"

### Erro: "Documento protegido"
1. Remover proteção do documento
2. Salvar documento
3. Executar macro novamente

### Performance Lenta
1. Fechar outros documentos do Word
2. Verificar tamanho do documento (máx. 500KB recomendado)
3. Aguardar conclusão completa

## 📞 Suporte

- **Repositório:** <https://github.com/chrmsantos/chainsaw-proposituras>
- **Issues:** <https://github.com/chrmsantos/chainsaw-proposituras/issues>
- **Email:** chrmsantos@gmail.com

## 📄 Licença

Apache 2.0 modificada - Ver arquivo LICENSE para detalhes completos.

---

**CHAINSAW PROPOSITURAS v2.0.0** - Sistema de padronização de documentos legislativos  
© 2025 Christian Martin dos Santos