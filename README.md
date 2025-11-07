# CHAINSAW - Sistema de Padronização de Proposituras Legislativas

Sistema automatizado para padronização de documentos legislativos no Microsoft Word, garantindo conformidade com normas de formatação institucional.

## Requisitos

### Sistema

- Windows 10 ou superior
- PowerShell 5.1 ou superior
- Acesso à rede corporativa (para instalação inicial)

### Aplicações

- Microsoft Word 2010 ou superior
- Permissões de leitura/escrita no perfil do usuário

## Instalação

### Instalação Automática (Recomendado)

O sistema CHAINSAW inclui um script automatizado de instalação que configura todos os componentes necessários.

#### Pré-requisitos

- Pasta `chainsaw` no perfil do usuário com todos os arquivos necessários
- Permissões de escrita no perfil do usuário (`%USERPROFILE%`)
- Word deve estar fechado durante a instalação

#### Como Executar

1. **Copie a pasta `chainsaw` para o seu perfil de usuário**
   - Caminho típico: `C:\Users\[seu_usuario]\chainsaw`

2. **Abra o PowerShell** (não é necessário executar como Administrador)
   - Pressione `Win + X` e selecione "Windows PowerShell"

3. **Navegue até a pasta do script**

   ```powershell
   cd "$env:USERPROFILE\chainsaw"
   ```

4. **Execute o script de instalação**

   [LOCK] **Método Recomendado - Bypass Automático Seguro:**

   ```cmd
   install.cmd
   ```

   Este launcher automático:
   - [OK] Funciona em QUALQUER política de execução
   - [OK] Não requer configuração manual
   - [OK] Usa bypass temporário apenas para este script
   - [OK] Não altera configurações permanentes do sistema
   - [OK] Totalmente seguro e transparente

   **Alternativa - Executar diretamente (requer política adequada):**

   ```powershell
   .\install.ps1
   ```

   **Com opções:**

   ```cmd
   install.cmd -Force          # Modo automático (sem confirmação)
   install.cmd -NoBackup       # Sem criar backup (não recomendado)
   ```

5. **Aguarde a conclusão**
   - O script exibirá o progresso de cada etapa
   - Se necessário, o script se relançará automaticamente (você verá uma mensagem explicativa)
   - Um arquivo de log será criado em `%USERPROFILE%\chainsaw\logs\`

#### O que o Script Faz

O script de instalação realiza automaticamente as seguintes operações:

1. **Verificação de Pré-requisitos**
   - Valida versão do Windows (10+)
   - Valida versão do PowerShell (5.1+)
   - Verifica existência dos arquivos necessários
   - Confirma permissões de escrita

2. **Cópia do Arquivo de Imagem**
   - Copia `stamp.png` para `%USERPROFILE%\chainsaw\assets\`
   - Verifica integridade do arquivo copiado

3. **Backup Automático**
   - Renomeia a pasta `%APPDATA%\Microsoft\Templates` existente
   - Formato do backup: `Templates_backup_AAAAMMDD_HHMMSS`
   - Remove backups antigos (mantém os 5 mais recentes)

4. **Instalação dos Templates**
   - Copia todos os templates para `%APPDATA%\Microsoft\Templates`
   - Preserva estrutura de pastas e arquivos

5. **Importação Automática de Personalizações** [NEW] **NOVO**
   - Detecta automaticamente a pasta `exported-config` (se existir)
   - Importa personalizações da interface do Word:
     - Faixa de Opções Personalizada (Ribbon)
     - Partes Rápidas (Quick Parts)
     - Blocos de Construção (Building Blocks)
     - Temas de Documentos
     - Template Normal.dotm
   - Solicita confirmação antes de importar (modo interativo)
   - Cria backup automático das personalizações existentes

6. **Registro de Log**
   - Cria log detalhado em `%USERPROFILE%\chainsaw\logs\`
   - Registra todas as operações, avisos e erros
   - Formato do log: `install_AAAAMMDD_HHMMSS.log`

#### Tratamento de Erros

O script inclui mecanismos robustos de tratamento de erros:

- **Validação prévia**: Verifica todos os requisitos antes de iniciar
- **Backup automático**: Sempre cria backup antes de modificar arquivos
- **Rollback**: Em caso de erro, tenta restaurar o backup automaticamente
- **Log detalhado**: Registra todas as operações para diagnóstico

#### Recuperação de Backup

Se precisar restaurar uma configuração anterior:

1. Navegue até `%APPDATA%\Microsoft\`
2. Renomeie a pasta `Templates` atual
3. Renomeie o backup desejado (ex: `Templates_backup_20251105_143022`) para `Templates`

#### Solução de Problemas

##### Erro: "Não foi possível acessar o caminho de rede"

- Verifique conexão com a rede corporativa
- Confirme que o caminho `\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw` está acessível
- Verifique suas credenciais de rede

##### Erro: "Permissões insuficientes"

- Não execute como Administrador (pode causar problemas de permissões)
- Verifique se você tem permissões de escrita no seu perfil

##### Erro: "Arquivo em uso"

- Feche o Microsoft Word completamente
- Feche todos os documentos do Office
- Se persistir, reinicie o computador

##### Consultar logs

```powershell
notepad "$env:USERPROFILE\chainsaw\logs\install_*.log"
```

### Instalação Manual

Caso não seja possível executar o script automatizado:

1. **Copiar arquivo de imagem**
   - Copie `\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw\assets\stamp.png`
   - Para `%USERPROFILE%\chainsaw\assets\stamp.png`

2. **Fazer backup dos Templates**
   - Renomeie `%APPDATA%\Microsoft\Templates`
   - Para `Templates_backup_AAAAMMDD`

3. **Copiar Templates**
   - Copie `\\strqnapmain\Dir. Legislativa\_Christian261\chainsaw\configs\Templates`
   - Para `%APPDATA%\Microsoft\Templates`

4. **Importar macro VBA**
   - Abra o Microsoft Word
   - Pressione `Alt + F11` para abrir o Editor VBA
   - Vá em `Arquivo > Importar Arquivo`
   - Selecione `CHAINSAW.bas` do caminho de rede
   - Feche o Editor VBA

## Execução

1. Abra o documento que deseja padronizar no Word
3. **Salve o documento** antes de executar (requisito obrigatório)
3. Pressione `Alt + F8` para abrir a lista de macros
4. Selecione `PadronizarDocumentoMain` e clique em `Executar`
5. Aguarde a conclusão do processamento

## [NEW] Novo: Exportação e Importação de Personalizações

O CHAINSAW agora permite **exportar e importar** todas as suas personalizações do Word:

- [ART] **Faixa de Opções** - Abas customizadas
- [PKG] **Blocos de Construção** - Building Blocks e Partes Rápidas
- [THEME] **Temas** - Temas e estilos personalizados
- [FAST] **Barra de Acesso Rápido** - Botões customizados
- [LOG] **Normal.dotm** - Template global com macros

### Como Usar

**Exportar (máquina de origem):**
```cmd
export-config.cmd
```

**Importar (máquina de destino):**
```cmd
import-config.cmd
```

📖 **Documentação completa:** `docs\EXPORTACAO_IMPORTACAO.md`

## Funcionalidades

### 1. **Configuração de Página**

- Margens: Superior (4,6 cm), Inferior (2 cm), Esquerda/Direita (3 cm)
- Orientação retrato
- Distância do cabeçalho: 0,3 cm
- Distância do rodapé: 0,9 cm

### 2. **Formatação de Fonte**

- Fonte padrão: Arial 12pt
- Cor automática do texto
- Remove formatações inconsistentes
- Preserva imagens inline durante formatação

### 3. **Formatação de Parágrafos**

- Alinhamento justificado
- Espaçamento entrelinhas: 14pt
- Espaçamento antes/depois: 0pt
- Recuo primeira linha: 0cm (exceto parágrafos especiais)

### 4. **Cabeçalho e Rodapé**

- Inserção automática de imagem institucional no cabeçalho
- Numeração de páginas no rodapé (formato: Página X de Y)
- Fonte do rodapé: Arial 9pt
- Centralização automática

### 5. **Remoção de Elementos**

- Marcas d'água
- Quebras de página manuais
- Espaços múltiplos
- Tabulações excessivas
- Linhas vazias sequenciais (máximo 1)

### 6. **Formatação de Título**

- Primeira linha: caixa alta, negrito, sublinhado, centralizado
- Detecção automática de proposituras (Indicação, Requerimento, Moção)
- Substituição automática por `$NUMERO$/$ANO$` quando aplicável

### 7. **Formatação de Parágrafos Especiais**

#### "CONSIDERANDO"

- Detecção automática
- Formatação: caixa alta, negrito
- Preserva espaçamento após palavra-chave

#### "Justificativa"

- Centralizado, negrito
- Inserção automática de 2 linhas em branco antes e depois

#### "Vereador/Vereadora"

- Parágrafo centralizado sem negrito
- Linha anterior: caixa alta, negrito, centralizado
- Linha posterior: centralizada

#### "Anexo/Anexos"

- Alinhamento à esquerda, negrito

#### "Diante do exposto"

- Primeiros 17 caracteres: caixa alta, negrito

#### "REQUEIRO"

- Parágrafo completo: caixa alta, negrito

### 8. **Substituições de Texto**

- Normalização de "d'Oeste" (16 variantes de aspas/acentos)
- Remoção de caracteres especiais inconsistentes

### 9. **Formatação de Local e Data**

- Padrão: "Plenário Dr. Tancredo Neves, em $DATAATUALEXTENSO$"
- Inserção automática de 2 linhas em branco antes e depois

### 10. **Formatação "Excelentíssimo Senhor Prefeito Municipal"**

- Inserção de 2 linhas em branco após o parágrafo
- Remoção de linhas vazias excedentes

### 11. **Formatação de Listas**

- Backup e restauração de listas numeradas
- Backup e restauração de listas com marcadores
- Aplicação de recuo padrão (36pt) para parágrafos iniciados com número ou marcador

### 12. **Proteção de Imagens**

- Backup de propriedades de todas as imagens
- Verificação de integridade após processamento
- Correção automática de dimensões alteradas
- Centralização de imagens entre 5ª e 7ª linha após "Plenário"
- Remoção de recuos em parágrafos com imagens

### 13. **Validações de Segurança**

- Verificação de integridade estrutural do documento
- Verificação de espaço em disco suficiente
- Detecção de timeout em operações longas (300s)
- Proteção contra loops infinitos (limite: 1000 iterações)

### 14. **Validação de Endereços**

- Verifica consistência entre endereço na ementa (2º parágrafo) e texto (4º parágrafo)
- Compara 2 palavras após "Rua" em contexto de 100 caracteres
- Normalização de "n.º", "nº", "número"
- Recomendação em caso de inconsistência

### 15. **Sistema de Backup Automático**

- Backup criado antes de qualquer modificação
- Localização: **mesma pasta do documento sendo editado**
- Formato: `nomedocumento_backup_AAAA-MM-DD_HHMMSS.docx`
- Limpeza automática com aviso após 15 arquivos

### 16. **Sistema de Logs**

- Registro detalhado de todas as operações
- Localização: **mesma pasta do documento sendo editado**
- Formato: `CHAINSAW_AAAAMMDD_HHMMSS_nomedocumento.log`
- Níveis: INFO, WARNING, ERROR

> **📍 Nota Importante sobre Localização:**  
> Tanto os **backups** quanto os **logs** são salvos na **mesma pasta do documento sendo editado**. Isso facilita o gerenciamento e garante que os arquivos relacionados fiquem juntos. Por exemplo:
> - Documento: `C:\Users\usuario\Meus Arquivos\MinhaProposicao.docx`
> - Backup: `C:\Users\usuario\Meus Arquivos\MinhaProposicao_backup_2025-11-05_143022.docx`
> - Log: `C:\Users\usuario\Meus Arquivos\chainsaw_20251105_143022_MinhaProposicao.log`

### 17. **Recuperação de Erros**

- Tratamento amigável de erros comuns
- Recuperação automática de estado da aplicação
- Mensagens descritivas para o usuário

## Fluxo de Processamento

### Fase 1: Verificações Iniciais

1. Verificação de versão do Word (mínimo: 2010)
2. Validação de integridade estrutural do documento
3. Verificação de documento salvo
4. Verificação de espaço em disco
5. Inicialização do sistema de logs

### Fase 2: Preparação

1. Criação de backup automático
2. Backup de configurações de visualização
3. Backup de propriedades de todas as imagens
4. Backup de formatações de listas
5. Desativação de alertas e atualização de tela

### Fase 3: Limpeza Estrutural

1. Remoção de formatações inconsistentes
2. Substituição de quebras de linha por quebras de parágrafo
3. Remoção de quebras de página manuais
4. Limpeza de espaços múltiplos
5. Remoção de tabulações excessivas
6. Limitação de linhas vazias sequenciais (máximo 1)

### Fase 4: Configuração Base

1. Aplicação de configurações de página (margens, orientação)
2. Remoção de marcas d'água
3. Formatação padrão de fonte (Arial 12pt)
4. Formatação padrão de parágrafos (justificado, 14pt)

### Fase 5: Formatações Especiais

1. Formatação do título (primeira linha)
2. Formatação de parágrafos "CONSIDERANDO"
3. Formatação de "Justificativa" com linhas em branco
4. Formatação de parágrafos "Vereador/Vereadora"
5. Formatação "Anexo/Anexos"
6. Formatação "Diante do exposto"
7. Formatação de parágrafos "REQUEIRO"
8. Substituição de parágrafo "Plenário" com data
9. Formatação "Excelentíssimo Senhor Prefeito Municipal"

### Fase 6: Aplicação de Regras de Texto

1. Substituições de texto (d'Oeste, etc.)
2. Limpeza final de espaços múltiplos
3. Controle final de linhas vazias

### Fase 7: Cabeçalho e Rodapé

1. Inserção de imagem no cabeçalho
2. Inserção de numeração de páginas no rodapé

### Fase 8: Formatações de Listas e Imagens

1. Formatação de recuos para parágrafos numerados
2. Formatação de recuos para parágrafos com marcadores
3. Restauração de formatações de listas originais
4. Formatação de recuos de imagens (zerado)
5. Centralização de imagem após "Plenário"

### Fase 9: Validações Finais

1. Validação de estrutura do documento
2. Validação de consistência de endereços
3. Verificação de dados sensíveis (CPF, RG, CNH)
4. Verificação de integridade das imagens
5. Correção de propriedades de imagens se necessário

### Fase 10: Finalização

1. Restauração de configurações de visualização (exceto zoom 110%)
2. Restauração de alertas e atualização de tela
3. Limpeza de variáveis globais
4. Finalização do sistema de logs
5. Exibição de mensagem de sucesso

## Utilitários Adicionais

### Abertura de Pasta de Logs/Backups

Execute a macro `AbrirPastaLogsEBackups` para abrir automaticamente:

- Pasta de backups (se existir)
- Pasta do documento (onde ficam os logs)

## Tratamento de Erros

O sistema implementa múltiplas camadas de proteção:

- **Erro 91** (Objeto não inicializado): Recomenda reiniciar o Word
- **Erro 5** (Chamada inválida): Verifica formato do documento
- **Erro 70** (Permissão negada): Indica documento protegido
- **Erro 53** (Arquivo não encontrado): Verifica salvamento do documento

Em caso de erro crítico, o sistema:

1. Registra o erro detalhado no log
2. Executa recuperação de emergência
3. Restaura estado da aplicação
4. Exibe mensagem amigável ao usuário

## Limitações e Considerações

- O documento **deve estar salvo** antes da execução
- Arquivos muito grandes (>50.000 parágrafos) podem ter processamento limitado
- Operações com timeout de 300 segundos
- A macro não cria cópias automáticas em rede - apenas local
- Imagens flutuantes podem ter comportamento diferente de imagens inline

## Licença

GNU General Public License v3.0 ou superior

## Autor

**Christian Martin dos Santos**  
Email: <chrmsantos@protonmail.com>  
GitHub: <https://github.com/chrmsantos>

## Versão

1.0-RC1 (Release Candidate 1)  
Data: 05/11/2025
