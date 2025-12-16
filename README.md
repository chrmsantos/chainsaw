# CHAINSAW

Ferramenta para padronizar proposituras legislativas no Microsoft Word.

## Requisitos

- Windows 10+
- PowerShell 5.1+
- Microsoft Word 2010+

## Estrutura

```text
chainsaw/
|-- assets/
|-- props/
|   |-- backups/
|-- source/
|   |-- logs/
|   `-- main/
`-- tests/
```

## Testes

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1

# Suites:
# powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1 -TestSuite VBA
# powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1 -TestSuite Encoding
```

## Licenca

GPLv3 - ver LICENSE

