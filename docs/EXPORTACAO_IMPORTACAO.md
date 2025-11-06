# Exportação e Importação de Personalizações do Word

## [INFO] Visão Geral

O Chainsaw agora inclui scripts completos para exportar e importar todas as personalizações da interface do Microsoft Word, incluindo:

- [OK] **Faixa de Opções Personalizada** (abas customizadas)
- [OK] **Blocos de Construção** (Building Blocks)
- [OK] **Partes Rápidas** (Quick Parts)
- [OK] **Temas e Estilos**
- [OK] **Barra de Ferramentas de Acesso Rápido**
- [OK] **Normal.dotm** (template global com macros)
- [OK] **Configurações do Registro** (opcional)

## [*] Fluxo de Trabalho

### 1️⃣ Exportar Personalizações (Máquina de Origem)

Execute na máquina que possui as personalizações que você deseja copiar:

```powershell
cd "$env:USERPROFILE\Documents\chainsaw"
.\export-config.ps1
```

Ou usando o launcher seguro:

```cmd
cd "%USERPROFILE%\Documents\chainsaw"
powershell.exe -ExecutionPolicy Bypass -File ".\export-config.ps1"
```

#### O que o Script Faz

1. **Verifica se o Word está aberto** - Recomenda fechar para garantir export completo
2. **Exporta Normal.dotm** - Template global com todas as personalizações
3. **Exporta Building Blocks** - Blocos de construção e partes rápidas
4. **Exporta Ribbon** - Personalizações da Faixa de Opções
5. **Exporta Temas** - Temas e estilos customizados
6. **Exporta UI** - Configurações da interface
7. **Cria Manifesto** - Arquivo JSON com metadata

#### Resultado

```
exported-config/
├── Templates/
│   ├── Normal.dotm
│   └── LiveContent/
│       └── 16/
│           ├── Managed/
│           │   ├── Document Themes/
│           │   └── Word Document Building Blocks/
│           └── User/
│               ├── Document Themes/
│               └── Word Document Building Blocks/
├── RibbonCustomization/
│   └── Word.officeUI
├── OfficeCustomUI/
│   └── [arquivos .officeUI]
├── Registry/
│   └── [arquivos .reg]
├── MANIFEST.json
├── README.txt
└── logs/
    └── export_YYYYMMDD_HHMMSS.log
```

### 2️⃣ Transferir Arquivos

Copie a pasta `exported-config` para a máquina de destino:

**Opção 1: Substituir no pacote Chainsaw**

```cmd
# Na máquina de destino
robocopy "C:\Temp\exported-config" "%USERPROFILE%\Documents\chainsaw\exported-config" /E /IS
```

**Opção 2: USB/Email**

1. Compacte a pasta `exported-config`
2. Transfira por USB, email ou rede
3. Extraia na máquina de destino

### 3️⃣ Importar Personalizações (Máquina de Destino)

**IMPORTANTE: Feche o Microsoft Word antes de importar!**

```powershell
cd "$env:USERPROFILE\Documents\chainsaw"
.\import-config.ps1
```

Ou usando o launcher seguro:

```cmd
cd "%USERPROFILE%\Documents\chainsaw"
powershell.exe -ExecutionPolicy Bypass -File ".\import-config.ps1"
```

#### O que o Script Faz

1. **Verifica se o Word está fechado** - Aborta se estiver aberto
2. **Cria backup automático** - Salva configurações atuais
3. **Importa Normal.dotm** - Substitui template global
4. **Importa Building Blocks** - Copia blocos de construção
5. **Importa Ribbon** - Restaura Faixa de Opções
6. **Importa Temas** - Restaura temas personalizados
7. **Importa UI** - Restaura configurações de interface
8. **Registra tudo em log**

## [>>] Uso Avançado

### Exportar para Caminho Específico

```powershell
.\export-config.ps1 -ExportPath "C:\Backup\MinhasPersonalizacoes"
```

### Incluir Configurações do Registro

```powershell
.\export-config.ps1 -IncludeRegistry
```

### Importar sem Backup

[!] **Não recomendado** - Use apenas se tiver certeza:

```powershell
.\import-config.ps1 -NoBackup
```

### Importar sem Confirmação

```powershell
.\import-config.ps1 -Force
```

### Importar de Caminho Específico

```powershell
.\import-config.ps1 -ImportPath "C:\Backup\MinhasPersonalizacoes"
```

## [PKG] Integração com Instalador Principal

O instalador principal (`install.ps1`) pode automaticamente importar as personalizações se encontrar a pasta `exported-config`:

```cmd
cd "%USERPROFILE%\Documents\chainsaw"
install.cmd
```

Isso irá:
1. Copiar `stamp.png`
2. Instalar Templates
3. **Importar personalizações** (se `exported-config` existir)

## 🔍 Estrutura Detalhada

### Normal.dotm

Contém:
- Macros personalizadas
- Estilos customizados
- Configurações globais
- AutoTexto
- Atalhos de teclado

### Building Blocks

Incluem:
- Partes Rápidas
- Cabeçalhos e Rodapés
- Páginas de Capa
- Marcas d'água
- Equações
- Tabelas

### Ribbon Customization

Personalizações da Faixa de Opções:
- Abas customizadas
- Grupos personalizados
- Botões adicionados/removidos
- Ordem das abas

### Office Custom UI

Configurações gerais:
- Barra de Ferramentas de Acesso Rápido (QAT)
- Temas do Office
- Preferências de interface

## [!] Avisos Importantes

### [X] NÃO Execute com Word Aberto

A importação **REQUER** que o Word esteja fechado. Se detectar o Word em execução, o script abortará automaticamente.

### [OK] Sempre Crie Backup

Por padrão, o script de importação cria backup automático. Não desabilite isso a menos que tenha outro backup.

### [SYNC] Compatibilidade de Versões

As personalizações são compatíveis entre:
- [OK] Mesma versão do Office
- [!] Versões próximas (ex: Office 2019 → Office 2021)
- [X] Versões muito diferentes (ex: Office 2010 → Office 365)

## [CHART] Exemplo de Uso Completo

### Cenário: Configurar 5 máquinas iguais

**Passo 1: Preparar máquina master**

```powershell
# Configurar o Word com todas as personalizações desejadas
# Testar e validar

# Exportar configurações
cd "$env:USERPROFILE\Documents\chainsaw"
.\export-config.ps1 -IncludeRegistry

# Resultado: exported-config/ criado
```

**Passo 2: Distribuir**

```cmd
# Copiar toda a pasta chainsaw incluindo exported-config
robocopy "C:\Master\chainsaw" "\\FileServer\Share\chainsaw" /E /IS

# Ou criar um ZIP
Compress-Archive -Path "C:\Master\chainsaw" -DestinationPath "Chainsaw-Complete.zip"
```

**Passo 3: Instalar em cada máquina**

```cmd
# Em cada máquina de destino:

# 1. Copiar pasta chainsaw para Documentos
robocopy "\\FileServer\Share\chainsaw" "%USERPROFILE%\Documents\chainsaw" /E /IS

# 2. Executar instalador
cd "%USERPROFILE%\Documents\chainsaw"
install.cmd

# 3. Abrir Word e verificar
```

## [SEC] Segurança e Privacidade

### O que é Exportado

- [OK] Personalizações de UI
- [OK] Blocos de construção
- [OK] Temas
- [OK] Configurações visuais

### O que NÃO é Exportado

- [X] Documentos pessoais
- [X] Histórico de uso
- [X] Senhas
- [X] Dados de conta Microsoft

### Registro (Opcional)

Se usar `-IncludeRegistry`, serão exportadas:
- Preferências do Word
- Configurações de interface
- Nenhuma informação sensível

## [CFG] Solução de Problemas

### Erro: "Word está em execução"

**Solução:**
1. Feche completamente o Word
2. Verifique no Gerenciador de Tarefas se `WINWORD.EXE` ainda está aberto
3. Termine o processo se necessário
4. Execute o script novamente

### Erro: "Fonte de importação não encontrada"

**Solução:**
1. Verifique se a pasta `exported-config` existe
2. Certifique-se que está no diretório correto
3. Use `-ImportPath` para especificar o caminho correto

### Personalizações não aparecem no Word

**Causas possíveis:**
1. Word não foi reiniciado após importação
2. Versões incompatíveis do Office
3. Políticas de grupo corporativas bloqueando personalizações

**Solução:**
1. Reinicie o Word completamente
2. Verifique se a versão do Office é compatível
3. Consulte o administrador de TI sobre políticas

### Normal.dotm corrompido

**Sintomas:**
- Word trava ao abrir
- Personalizações desaparecem
- Erros de macro

**Solução:**
1. Feche o Word
2. Renomeie `Normal.dotm` para `Normal.old`
3. O Word criará um novo Normal.dotm automático
4. Re-importe as personalizações

## [LOG] Logs

Todos os logs são salvos em:
- **Exportação**: `exported-config/logs/export_YYYYMMDD_HHMMSS.log`
- **Importação**: `%USERPROFILE%\chainsaw\logs\import_YYYYMMDD_HHMMSS.log`

## [SYNC] Atualização de Personalizações

Para atualizar personalizações existentes:

1. **Exportar novas personalizações**
   ```powershell
   .\export-config.ps1 -ExportPath ".\exported-config-v2"
   ```

2. **Distribuir atualização**
   - Substitua `exported-config` antiga pela nova

3. **Re-importar**
   ```powershell
   .\import-config.ps1 -Force
   ```

## 📞 Suporte

Para problemas ou dúvidas:

1. Consulte os logs em `chainsaw\logs\`
2. Verifique `INSTALL.md` para documentação geral
3. Entre em contato: chrmsantos@protonmail.com

---

**Versão:** 1.0.0  
**Última Atualização:** 05/11/2025  
**Autor:** Christian Martin dos Santos
