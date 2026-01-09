# CHAINSAW

Ferramenta para padronizar proposituras legislativas no Microsoft Word.

![Chainsaw](assets/chainsaw.png)

## Requisitos

- Windows 10+
- PowerShell 5.1+
- Microsoft Word 2010+

## Estrutura

```text
chainsaw/
|-- assets/
|-- source/
|   `-- main/
`-- tests/

Logs: %USERPROFILE%\chainsaw\source\logs
Backups: %TEMP%\.chainsaw\props\backups
```

## Testes

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1

# Suites:
# powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1 -TestSuite VBA
# powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1 -TestSuite Encoding
```

## Licenca

- Texto original (ingles): LICENSE
- Traducao (pt-BR, nao-oficial): LICENSE.pt-BR.md

