# Atualização do Módulo VBA - CHAINSAW v1.1

## 📋 Mudança Importante

O módulo principal do CHAINSAW foi renomeado de `Módulo1.bas` para `monolithicMod.bas`. Esta mudança melhora a organização e clareza do código.

## 🔄 Como Atualizar

### Opção 1: Instalação Completa (Recomendado)

Execute o script de instalação normal. Ele agora inclui automaticamente a atualização do módulo VBA:

```cmd
install.cmd
```

O script irá:
1. Copiar o arquivo `stamp.png`
2. Fazer backup do `Templates` atual
3. Copiar o novo `Templates`
4. **✨ NOVO: Importar o módulo VBA mais recente (`monolithicMod.bas`)**
5. Importar personalizações (se disponíveis)

### Opção 2: Atualizar Apenas o Módulo VBA

Se você já tem tudo instalado e só precisa atualizar o módulo VBA:

```cmd
update-vba-module.cmd
```

Este script:
- ✅ Fecha o Word automaticamente (com sua confirmação)
- ✅ Faz backup do módulo antigo
- ✅ Importa o novo módulo `monolithicMod.bas`
- ✅ Salva automaticamente

## 🎯 Para Novos Usuários

Se é sua primeira instalação:

```cmd
install.cmd
```

Tudo será configurado automaticamente, incluindo o módulo VBA mais recente.

## 🔧 Importação Manual (Se Necessário)

Caso os scripts automáticos não funcionem (devido a configurações de segurança do VBA):

1. Abra o Microsoft Word
2. Pressione `Alt + F11` (abre o editor VBA)
3. Clique em **Arquivo** > **Importar Arquivo**
4. Navegue até: `C:\Users\[seu_usuario]\chainsaw\src\monolithicMod.bas`
5. Selecione o arquivo e clique em **Abrir**
6. O módulo será importado para o `Normal.dotm`
7. Feche o editor VBA (`Alt + Q`)
8. Salve quando solicitado

## ⚠️ Possíveis Problemas

### Erro: "Acesso programático ao projeto VBA negado"

**Causa:** Configuração de segurança do Word bloqueia acesso ao VBA.

**Solução:**
1. Abra o Word
2. Vá em **Arquivo** > **Opções**
3. Selecione **Central de Confiabilidade**
4. Clique em **Configurações da Central de Confiabilidade**
5. Vá em **Configurações de Macro**
6. Marque: **"Confiar no acesso ao modelo de objeto do projeto VBA"**
7. Clique em **OK** e feche o Word
8. Execute o script novamente

### Word não fecha automaticamente

**Solução:** Feche o Word manualmente antes de executar o script.

## 📁 Estrutura de Arquivos

```
chainsaw/
├── install.cmd              # Instalação completa (RECOMENDADO)
├── install.ps1              # Script PowerShell de instalação
├── update-vba-module.cmd    # Atualizar apenas módulo VBA
├── update-vba-module.ps1    # Script PowerShell de atualização
└── src/
    ├── monolithicMod.bas    # ⭐ MÓDULO VBA PRINCIPAL (v1.1)
    └── Exemplos_Identificacao.bas  # Exemplos de uso
```

## ✅ Verificação

Para verificar se o módulo foi importado corretamente:

1. Abra o Word
2. Pressione `Alt + F11`
3. Na janela do editor VBA, procure por **"monolithicMod"** na árvore de projetos
4. Se estiver lá, a importação foi bem-sucedida! ✓

## 🆕 O Que Mudou na v1.1

- ✨ Novo sistema de identificação de elementos estruturais
- ✨ 11 novas funções públicas de acesso
- ✨ Integração com cache de parágrafos
- ✨ Documentação completa e exemplos
- ✨ Instalação automática do módulo VBA

## 📚 Documentação

- **Guia Rápido:** `GUIA_RAPIDO_IDENTIFICACAO.md`
- **Documentação Completa:** `docs/IDENTIFICACAO_ELEMENTOS.md`
- **Exemplos Práticos:** `src/Exemplos_Identificacao.bas`
- **Novidades v1.1:** `docs/NOVIDADES_v1.1.md`

## 🆘 Suporte

**Email:** chrmsantos@protonmail.com  
**Versão:** CHAINSAW v1.1-RC1-202511071045  
**Licença:** GNU GPLv3

---

**Última atualização:** 07/11/2024
