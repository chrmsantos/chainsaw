# CHAINSAW - Sistema de Padronização de Proposituras Legislativas

Sistema automatizado para padronização de documentos legislativos no Microsoft Word.

##  Requisitos

- Windows 10+ | PowerShell 5.1+
- Microsoft Word 2010+
- Permissões de usuário normal (não requer admin)

##  Instalação Rápida

1. Copie a pasta `chainsaw` para: `C:\Users\[seu_usuario]\chainsaw`
2. Navegue até: `chainsaw\installation\inst_scripts\`
3. Dê duplo-clique em: `install.cmd`

**Pronto!** O instalador fará tudo automaticamente.

### Alternativa (PowerShell):

```powershell
cd "$env:USERPROFILE\chainsaw\installation\inst_scripts"
.\install.ps1
```

##  Documentação

### Instalação e Uso

-  **[Guia de Instalação](installation/inst_docs/GUIA_INSTALACAO.md)** - Instalação detalhada e troubleshooting
-  **[Novidades v1.1](docs/NOVIDADES_v1.1.md)** - Sistema de identificação de elementos
-  **[Identificação](docs/IDENTIFICACAO_ELEMENTOS.md)** - API de identificação automática
-  **[Sem Admin](docs/SEM_PRIVILEGIOS_ADMIN.md)** - Instalação em ambientes restritos
-  **[Substituições](docs/SUBSTITUICOES_CONDICIONAIS.md)** - Lógica de substituições
-  **[Validação](docs/VALIDACAO_TIPO_DOCUMENTO.md)** - Validação de tipos de documento

### Segurança e Privacidade

-  **[Conformidade LGPD](docs/LGPD_CONFORMIDADE.md)** - Conformidade com Lei Geral de Proteção de Dados
-  **[Segurança e Privacidade](docs/SEGURANCA_PRIVACIDADE.md)** - Política completa de segurança e privacidade

##  Estrutura

```
chainsaw/
 installation/          # Scripts e configurações de instalação
    inst_configs/     # Templates do Word
    inst_scripts/     # Scripts (.ps1, .cmd)
    inst_docs/        # Documentação e logs
 source/
    main/             # Módulo VBA principal (monolithicMod.bas)
    backups/          # Backups timestamped do módulo VBA
    others/           # Exemplos e código auxiliar
 docs/                 # Documentação técnica
 assets/               # Recursos (imagens)
 README.md
 CHANGELOG.md
 LICENSE
```

##  Atualização VBA

```powershell
cd "$env:USERPROFILE\chainsaw\installation\inst_scripts"
.\update-vba-module.ps1
```

Ou dê duplo-clique em: `update-vba-module.cmd`

##  Exportar Configurações

```powershell
cd "$env:USERPROFILE\chainsaw\installation\inst_scripts"
.\export-config.ps1
```

##  Solução de Problemas

### Erro: "Não é possível executar scripts"

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Ou use o `install.cmd` (duplo-clique) que contorna automaticamente.

### Ribbon não aparece

1. Feche completamente o Word
2. Execute o instalador novamente
3. Verifique logs em `installation/inst_docs/inst_logs/`

### "Normal.dotm está em uso"

1. Feche Word
2. Gerenciador de Tarefas → Finalize `WINWORD.EXE`
3. Execute instalador novamente

##  Testes Automatizados

O projeto inclui um sistema completo de testes unitários usando **Pester** para garantir a qualidade do código.

### Executar Testes

**Opção 1 - Script CMD (recomendado):**
```cmd
cd chainsaw\tests
run-tests.cmd
```

**Opção 2 - PowerShell:**
```powershell
cd "$env:USERPROFILE\chainsaw\tests"
powershell -NoProfile -ExecutionPolicy Bypass -File .\Run-Tests.ps1
```

### O que é testado

- [OK] **Scripts PowerShell** - Validação de sintaxe (export-config.ps1, install.ps1, update-vba-module.ps1)
- [OK] **Módulos VBA** - Verificação de existência e duplicatas (monolithicMod.bas)
- [OK] **Documentação** - Integridade de arquivos essenciais
- [OK] **CHANGELOG** - Verificação de versão atual

### Pré-requisitos

- PowerShell 5.1+
- Pester 3.4.0+ (instalado automaticamente se necessário)

### Ver Resultados Detalhados

```powershell
cd chainsaw\tests
powershell -NoProfile -ExecutionPolicy Bypass -File .\Run-Tests.ps1 -Detailed
```

##  Segurança

- [OK] Backups automáticos antes de qualquer modificação
- [OK] Código fonte aberto (VBA + PowerShell)
- [OK] Instalação 100% local (sem envio de dados)
- [OK] Sem privilégios de administrador
- [OK] Encoding UTF-8 padronizado, sem emojis
- [OK] Validação automática de qualidade

##  Contribuindo

Contribuições são bem-vindas! Veja [CONTRIBUTING.md](CONTRIBUTING.md) para diretrizes:

- Padrões de código e encoding
- Processo de Pull Request
- Como reportar bugs
- **IMPORTANTE**: Projeto não permite emojis no código

##  Licença

MIT License - Veja [LICENSE](LICENSE)

---

**Versão:** 2.0.4 | **Desenvolvido por:** chrmsantos | **Atualizado:** Nov 2025

