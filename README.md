# CHAINSAW

Ferramenta para padronizar proposituras legislativas no Microsoft Word.

## Exportar configuracoes do Word

- Abrir PowerShell na raiz do projeto
- Executar:
  - `cd tools\export`
  - `exportar_configs.cmd`
- Para caminho especifico:
  - `powershell -ExecutionPolicy Bypass -NoProfile -File exportar_configs.ps1 -ExportPath "C:\\Backup\\WordConfig"`

## Requisitos

- Windows 10+
- PowerShell 5.1+
- Microsoft Word 2010+

## Estrutura

```text
chainsaw/
├── assets/
├── props/
│   ├── backups/
│   └── recovery_tmp/
├── source/main/
├── tools/export/
└── tests/
```

## Testes

```powershell
cd tests
.\Run-Tests.ps1 -TestSuite Export
```

## Licenca

GPLv3 - ver LICENSE

