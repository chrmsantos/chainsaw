# 🎯 Guia Rápido: Exportação e Importação de Personalizações

## 📦 O que é exportado?

```
┌─────────────────────────────────────────────────────────────────┐
│                    PERSONALIZAÇÕES DO WORD                      │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  🎨 Faixa de Opções (Ribbon)                                   │
│     └─ Abas customizadas                                       │
│     └─ Grupos personalizados                                   │
│     └─ Botões adicionados/removidos                           │
│                                                                 │
│  📦 Blocos de Construção (Building Blocks)                     │
│     └─ Partes Rápidas                                          │
│     └─ Cabeçalhos e Rodapés                                    │
│     └─ Páginas de Capa                                         │
│     └─ Marcas d'água                                           │
│                                                                 │
│  🎭 Temas e Estilos                                            │
│     └─ Temas personalizados                                    │
│     └─ Estilos customizados                                    │
│     └─ Cores do documento                                      │
│                                                                 │
│  ⚡ Barra de Ferramentas de Acesso Rápido                      │
│     └─ Botões personalizados                                   │
│     └─ Ordem dos comandos                                      │
│                                                                 │
│  📝 Normal.dotm                                                 │
│     └─ Template global                                         │
│     └─ Macros                                                  │
│     └─ AutoTexto                                               │
│     └─ Atalhos de teclado                                      │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

## 🚀 Fluxo Completo

### Máquina A (Origem) → Máquina B (Destino)

```
┌─────────────┐
│  Máquina A  │  (Origem - com personalizações)
│             │
│   1. ✅     │  Configure o Word como desejar
│   2. 📤     │  Execute: export-config.cmd
│   3. 📦     │  Pasta 'exported-config' criada
│   4. 💾     │  Copie para USB/rede/email
│             │
└─────────────┘
       │
       │ Transferir arquivos
       ↓
┌─────────────┐
│  Máquina B  │  (Destino - receberá personalizações)
│             │
│   1. 📁     │  Copie pasta 'CHAINSAW' para Documentos
│   2. ❌     │  FECHE o Word completamente
│   3. 📥     │  Execute: import-config.cmd
│   4. ✅     │  Abra o Word
│   5. 🎉     │  Personalizações aplicadas!
│             │
└─────────────┘
```

## ⚡ Comandos Rápidos

### Exportar

```cmd
cd "%USERPROFILE%\Documents\CHAINSAW"
export-config.cmd
```

**Resultado:**
- ✅ Cria pasta `exported-config`
- ✅ Manifesto JSON com metadata
- ✅ README com instruções
- ✅ Log detalhado

### Importar

```cmd
cd "%USERPROFILE%\Documents\CHAINSAW"
import-config.cmd
```

**Requisitos:**
- ❌ Word DEVE estar fechado
- ✅ Pasta `exported-config` deve existir
- ✅ Backup automático criado

## 🎬 Cenários de Uso

### Cenário 1: Configurar Nova Máquina

```
Você → Nova Máquina
1. export-config.cmd na sua máquina
2. Copiar exported-config
3. import-config.cmd na nova máquina
```

### Cenário 2: Padronizar Equipe

```
Master → Várias Máquinas
1. Configurar Word master
2. export-config.cmd
3. Distribuir CHAINSAW completo
4. Cada usuário: install.cmd ou import-config.cmd
```

### Cenário 3: Backup de Segurança

```
Periódico
1. export-config.cmd mensalmente
2. Guardar exported-config em backup
3. Restaurar quando necessário
```

### Cenário 4: Teste de Configurações

```
Sandbox
1. Exportar configurações atuais
2. Testar novas configurações
3. Se não gostar: importar backup
4. Se gostar: exportar novo padrão
```

## 📋 Checklist de Exportação

- [ ] Word está aberto? (pode exportar, mas recomenda fechar)
- [ ] Todas personalizações configuradas?
- [ ] Macros testadas?
- [ ] Blocos de construção criados?
- [ ] Execute: `export-config.cmd`
- [ ] Verifique pasta `exported-config` criada
- [ ] Copie para destino

## 📋 Checklist de Importação

- [ ] Pasta `exported-config` copiada?
- [ ] Word está **COMPLETAMENTE FECHADO**? ⚠️
- [ ] Feche outros documentos Office
- [ ] Execute: `import-config.cmd`
- [ ] Aguarde conclusão
- [ ] Abra Word
- [ ] Verifique personalizações

## ⚠️ Avisos Importantes

### ✅ PODE fazer

- ✅ Exportar com Word aberto (não recomendado)
- ✅ Exportar múltiplas vezes
- ✅ Importar várias vezes
- ✅ Ter backups de exports

### ❌ NÃO PODE fazer

- ❌ Importar com Word aberto → Script aborta!
- ❌ Importar sem `exported-config`
- ❌ Pular o backup (use `-NoBackup` com cautela)

## 🔧 Opções Avançadas

### Exportar com Registro

```powershell
.\export-config.ps1 -IncludeRegistry
```

### Importar sem Confirmação

```powershell
.\import-config.ps1 -Force
```

### Importar sem Backup

```powershell
.\import-config.ps1 -NoBackup
```

⚠️ **Cuidado:** Não recomendado!

### Caminho Customizado

```powershell
# Exportar para local específico
.\export-config.ps1 -ExportPath "C:\Backup\Config2025"

# Importar de local específico
.\import-config.ps1 -ImportPath "C:\Backup\Config2025"
```

## 📊 Tamanho Típico

```
exported-config/
├── Templates/         (~100 KB)
├── RibbonCustomization/  (~10 KB)
├── OfficeCustomUI/    (~5 KB)
├── Building Blocks/   (~50 KB)
└── Registry/          (~20 KB)
────────────────────────────────
Total:                 ~200 KB

Compactado (ZIP):      ~50 KB
```

💡 **Dica:** Facilmente enviável por email!

## 🆘 Solução Rápida de Problemas

| Problema | Solução |
|----------|---------|
| "Word em execução" | Feche COMPLETAMENTE o Word |
| "exported-config não encontrado" | Execute export-config.cmd primeiro |
| "Ribbon não aparece" | Reinicie Word, verifique versão |
| "Macros não funcionam" | Habilite macros nas configurações |
| "Erro de permissão" | NÃO execute como Administrador |

## 📞 Precisa de Ajuda?

1. **Logs**: `CHAINSAW\logs\`
2. **Documentação**: `docs\EXPORTACAO_IMPORTACAO.md`
3. **Email**: chrmsantos@protonmail.com

---

**Versão:** 1.0.0  
**Última Atualização:** 05/11/2025
