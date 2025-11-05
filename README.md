# Chainsaw - Sistema de Padronização de Proposituras Legislativas

Sistema automatizado para padronização de documentos legislativos no Microsoft Word, garantindo conformidade com normas de formatação institucional.

## Requisitos

- Microsoft Word 2010 ou superior
- Sistema operacional Windows
- Permissões de leitura/escrita no diretório do documento

## Instalação

1. Copie o arquivo `chainsaw.bas` para a pasta desejada
2. No Microsoft Word, pressione `Alt + F11` para abrir o Editor VBA
3. Vá em `Arquivo > Importar Arquivo` e selecione `chainsaw.bas`
4. Feche o Editor VBA

## Execução

1. Abra o documento que deseja padronizar no Word
2. **Salve o documento** antes de executar (requisito obrigatório)
3. Pressione `Alt + F8` para abrir a lista de macros
4. Selecione `PadronizarDocumentoMain` e clique em `Executar`
5. Aguarde a conclusão do processamento

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
- Localização: pasta `backups\` no mesmo diretório do documento
- Formato: `nomedocumento_backup_AAAA-MM-DD_HHMMSS.docx`
- Limpeza automática com aviso após 15 arquivos

### 16. **Sistema de Logs**

- Registro detalhado de todas as operações
- Localização: mesmo diretório do documento
- Formato: `chainsaw_log_AAAA-MM-DD.txt`
- Níveis: INFO, WARNING, ERROR

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
