# CHAINSAW PROPOSITURAS - Quick Installation

## 🚀 Automated Installation (Recommended)

### 1. Download
Download all project files from: <https://github.com/chrmsantos/chainsaw-proposituras>

### 2. Run the Installer
Open PowerShell as Administrator and run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\install-chainsaw.ps1
```

### 3. Installer Parameters

```powershell
# Default installation
.\install-chainsaw.ps1

# Custom installation
.\install-chainsaw.ps1 -InstallPath "C:\MyFolder" -AllUsers -Silent

# Check compatibility only
.\install-chainsaw.ps1 -CheckOnly
```

## 🛠️ Manual Installation

### Prerequisites

### Steps

1. **Create folder structure:**

   ```text
   CHAINSAW-PROPOSITURAS/
   ├── src/
   ├── private/
   │   ├── header/
   │   ├── backups/
   │   └── logs/
   └── docs/
   ```

2. **Copy files:**

3. **Install VBA module:**

4. **Configure security:**

## ⚡ Quick Use

### Run Standardization

1. Open a document in Word
2. Alt+F8 → Run Macro
3. Select: `StandardizeDocumentMain`
4. Confirm execution

### Shortcuts (after automated installation)

## 🔧 Main Settings

### Supported Document Types

### Applied Formatting

## 📋 Installed Autotext

| Código | Resultado |
|--------|-----------|
| `indicacao` | INDICAÇÃO Nº $NUMERO$/$ANO$ |
| `requerimento` | REQUERIMENTO Nº $NUMERO$/$ANO$ |
| `mocao` | MOÇÃO Nº $NUMERO$/$ANO$ |
| `considerando` | CONSIDERANDO que |
| `justificativa` | JUSTIFICATIVA |
| `vereador` | - VEREADOR - |

## 🔒 Security

### Recommended Settings

### System Validations

## 📁 File Structure

```text
CHAINSAW-PROPOSITURAS/
├── src/
│   └── chainsaw0.bas            # Main VBA code
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

## 🆘 Troubleshooting

### Error: "Macro not found"

1. Check if the module was imported correctly
2. Reopen Word
3. Check macro name: `StandardizeDocumentMain`

### Error: "Access to VBA denied"

1. Word → Options → Trust Center
2. Macro Settings
3. Enable "Trust access to the VBA project object model"

### Error: "Document protected"

1. Remove document protection
2. Save the document
3. Run the macro again

### Slow Performance

1. Close other Word documents
2. Check document size (max. 500KB recommended)
3. Wait for completion

## 📞 Support


## 📄 License

Apache 2.0 modified - See LICENSE for details.


CHAINSAW PROPOSITURAS v2.0.0 - Legislative document standardization system  
© 2025 Christian Martin dos Santos
