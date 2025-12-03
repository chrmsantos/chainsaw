# CHAINSAW

Sistema de padronização automática de proposituras legislativas para Microsoft Word.

## Instalação

```cmd
cd installation\inst_scripts
chainsaw_installer.cmd
```

Instalação guiada em modo texto com confirmações simples (S/N ou menus numerados).

## Requisitos

- Windows 10+ | PowerShell 5.1+ | Word 2010+
- Permissões de usuário (sem admin)
- Internet (primeira instalação)

## Uso

```cmd
# Instalar ou atualizar Word Templates e VBA
cd %USERPROFILE%\chainsaw\installation\inst_scripts
chainsaw_installer.cmd

# Exportar configurações atuais do Word
exportar_configs.cmd
```

## Estrutura

```text
chainsaw/
├── installation/
│   ├── inst_scripts/     # Pipelines (chainsaw_installer/exportar_configs) e scripts PowerShell
│   ├── inst_configs/     # Templates do Word
│   └── inst_docs/        # Documentação e logs
├── source/main/          # Módulo VBA (Módulo1.bas)
├── tests/                # Testes automatizados (Pester)
└── assets/               # Recursos (stamp.png)
```

## Troubleshooting

**Erro de execução:**

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**Word não abre:**

- Feche todas as instâncias
- Execute novamente

## Segurança

- Backups automáticos antes de modificações
- Código fonte aberto (VBA + PowerShell)
- Instalação 100% local (sem envio de dados)
- Sem privilégios de administrador
- Validação automática de qualidade

## Testes

```cmd
cd chainsaw\tests
run-tests.cmd
```

## Licença

GPLv3 - Ver [LICENSE](LICENSE)

---

**Versão:** 2.0.4 | **Desenvolvido por:** chrmsantos | **Atualizado:** Nov 2025

