# POLÍTICA DE PRIVACIDADE

**Projeto:** CHAINSAW - Sistema de Padronização de Proposituras Legislativas  
**Versão:** 2.0.3  
**Data de Vigência:** 04 de dezembro de 2025  
**Última Atualização:** 04 de dezembro de 2025

---

## 1. INTRODUÇÃO

Esta Política de Privacidade descreve como o CHAINSAW trata informações em conformidade com:

- **LGPD** - Lei Geral de Proteção de Dados Pessoais (Lei nº 13.709/2018)
- **GPLv3** - GNU General Public License versão 3

O CHAINSAW é um software livre e de código aberto que processa documentos **exclusivamente no computador do usuário**, sem coletar, transmitir ou armazenar dados pessoais em servidores externos.

---

## 2. CONTROLADOR E ENCARREGADO DE DADOS

### 2.1 Controlador de Dados

**Nome:** Projeto CHAINSAW (Open Source)  
**Repositório:** https://github.com/chrmsantos/chainsaw  
**Natureza:** Software Livre (GPLv3)  
**Tipo:** Ferramenta de Processamento Local

### 2.2 Encarregado de Dados (DPO)

Em conformidade com o Art. 41 da LGPD, informamos que:

- O CHAINSAW **não coleta dados pessoais**
- Não há necessidade de nomeação de Encarregado de Dados
- Questões sobre privacidade podem ser direcionadas ao repositório do projeto

---

## 3. COLETA E PROCESSAMENTO DE DADOS

### 3.1 Declaração de Não Coleta

**O CHAINSAW NÃO coleta dados pessoais.**

O sistema opera de acordo com os princípios de **Privacy by Design** e **Privacy by Default**:

- ✅ **NÃO coleta** dados pessoais de usuários
- ✅ **NÃO transmite** informações pela internet
- ✅ **NÃO armazena** dados em servidores externos
- ✅ **NÃO rastreia** comportamento de usuários
- ✅ **NÃO utiliza** cookies, telemetria ou analytics
- ✅ **NÃO possui** sistema de autenticação ou cadastro
- ✅ **NÃO compartilha** dados com terceiros

### 3.2 Dados Técnicos Processados Localmente

O CHAINSAW processa **exclusivamente em seu computador**:

| Tipo de Dado | Finalidade | Localização | Base Legal (LGPD) |
|--------------|-----------|-------------|-------------------|
| Configurações do Word | Aplicar formatação padrão | Registro do Windows | Art. 7º, V (Execução de contrato) |
| Backups de documentos | Recuperação em caso de erro | `%USERPROFILE%\chainsaw\props\backups` | Art. 7º, V (Execução de contrato) |
| Logs de operação | Debugging e diagnóstico | `installation\inst_docs\inst_logs` | Art. 7º, V (Execução de contrato) |
| Templates | Padronização de documentos | `%APPDATA%\Microsoft\Templates` | Art. 7º, V (Execução de contrato) |

**Importante:** Nenhum dos dados acima constitui "dado pessoal" conforme Art. 5º, I da LGPD.

### 3.3 Documentos do Usuário

- Os documentos criados e editados **permanecem no seu computador**
- O CHAINSAW **não acessa** o conteúdo dos documentos para fins além da formatação
- O CHAINSAW **não envia** documentos para servidores externos
- Você mantém **controle total** sobre seus arquivos

---

## 4. CONFORMIDADE COM A LGPD

### 4.1 Princípios da LGPD (Art. 6º)

| Princípio | Status | Implementação |
|-----------|--------|---------------|
| **Finalidade** | ✅ Conforme | Formatação de documentos legislativos |
| **Adequação** | ✅ Conforme | Processamento compatível com finalidade declarada |
| **Necessidade** | ✅ Conforme | Apenas dados técnicos essenciais ao funcionamento |
| **Livre Acesso** | ✅ Conforme | Logs e backups acessíveis localmente pelo usuário |
| **Qualidade dos Dados** | ✅ Conforme | Dados técnicos de sistema, sem dados pessoais |
| **Transparência** | ✅ Conforme | Código aberto (GPLv3), documentação completa |
| **Segurança** | ✅ Conforme | Processamento local isolado, sem transmissão |
| **Prevenção** | ✅ Conforme | Arquitetura projetada para não coletar dados |
| **Não Discriminação** | ✅ Conforme | Software acessível a todos os usuários |
| **Responsabilização** | ✅ Conforme | Documentação e código auditável |

### 4.2 Direitos dos Titulares (Art. 18)

Como o CHAINSAW **não coleta dados pessoais**, os seguintes direitos não são aplicáveis:

- Confirmação e acesso aos dados
- Correção de dados incompletos ou desatualizados
- Anonimização, bloqueio ou eliminação
- Portabilidade dos dados
- Revogação do consentimento

**Seus documentos locais:** Você possui controle total sobre arquivos armazenados em seu computador e pode excluí-los a qualquer momento.

### 4.3 Base Legal

O processamento técnico local é realizado com base em:

- **Art. 7º, V da LGPD**: Execução de contrato (instalação e uso do software)
- **Legítimo interesse**: Funcionalidades essenciais ao software (Art. 10)

---

## 5. CONFORMIDADE COM GPLv3

### 5.1 Liberdade e Transparência

O CHAINSAW é licenciado sob **GNU General Public License v3.0**:

- ✅ Código-fonte **100% aberto** e auditável
- ✅ Liberdade para **executar, estudar, modificar e distribuir**
- ✅ **Transparência total** nas operações do software
- ✅ **Sem componentes proprietários** ocultos
- ✅ **Copyleft**: Derivações devem manter a mesma licença

### 5.2 Garantias de Privacidade pela GPLv3

A licença GPLv3 garante:

1. **Transparência**: Você pode auditar o código para verificar práticas de privacidade
2. **Controle**: Você pode modificar o software para atender suas necessidades
3. **Ausência de backdoors**: Código aberto impede coleta oculta de dados
4. **Responsabilidade**: A comunidade pode identificar e corrigir problemas

### 5.3 Ausência de Garantias

Conforme GPLv3, Seção 15 e 16:

> O CHAINSAW é fornecido "COMO ESTÁ", sem garantias de qualquer tipo. O desenvolvedor não se responsabiliza por danos decorrentes do uso do software.

---

## 6. SEGURANÇA E PROTEÇÃO

### 6.1 Medidas de Segurança Técnicas

O CHAINSAW implementa:

- **Processamento Local**: Zero transmissão de dados pela rede
- **Isolamento**: Operação restrita ao ambiente do usuário
- **Backups Locais**: Sistema de proteção em diretório do usuário
- **Logs Locais**: Registros armazenados apenas localmente
- **Sem Autenticação Externa**: Não há credenciais a serem comprometidas

### 6.2 Responsabilidade do Usuário

Você é responsável por:

- Proteger seu computador com antivírus e firewall
- Manter backups dos seus documentos importantes
- Proteger o acesso físico ao seu computador
- Atualizar o sistema operacional e o Microsoft Word

### 6.3 Incidentes de Segurança

Para reportar vulnerabilidades de segurança, consulte `SECURITY.md` no repositório do projeto.

---

## 7. ARMAZENAMENTO E RETENÇÃO

### 7.1 Localização dos Dados

Todos os dados são armazenados **localmente** em:

```
%USERPROFILE%\chainsaw\
├── installation\inst_docs\inst_logs\     # Logs de instalação
├── props\backups\                        # Backups de documentos
└── props\recovery_tmp\                   # Arquivos temporários de recuperação

%APPDATA%\Microsoft\Templates\
└── Normal.dotm                           # Template do Word
```

### 7.2 Retenção de Dados

- **Logs**: Mantidos indefinidamente (podem ser excluídos manualmente)
- **Backups**: Mantidos indefinidamente (podem ser excluídos manualmente)
- **Configurações**: Mantidas enquanto o software estiver instalado

### 7.3 Exclusão de Dados

Para remover todos os dados:

1. Desinstale o CHAINSAW pelo painel de controle do Word
2. Exclua manualmente a pasta `%USERPROFILE%\chainsaw`
3. Remova o template `Normal.dotm` se desejar

---

## 8. TRANSFERÊNCIA INTERNACIONAL DE DADOS

**Não aplicável**: O CHAINSAW não transfere dados para fora do seu computador.

---

## 9. COOKIES E RASTREAMENTO

**O CHAINSAW não utiliza:**

- ❌ Cookies
- ❌ Web beacons
- ❌ Pixels de rastreamento
- ❌ Analytics (Google Analytics, etc.)
- ❌ Telemetria
- ❌ Fingerprinting
- ❌ Rastreamento de comportamento

---

## 10. INTEGRAÇÃO COM MICROSOFT WORD

### 10.1 VBA e Macros

O CHAINSAW utiliza **VBA (Visual Basic for Applications)** para:

- Aplicar formatação a documentos
- Criar backups de segurança
- Gerenciar templates

### 10.2 Acesso a Documentos

- O VBA acessa documentos **apenas para formatação**
- Não há **leitura de conteúdo** para fins de análise ou coleta
- Não há **envio de dados** para servidores externos

### 10.3 Permissões do Word

O CHAINSAW requer permissões padrão do Word para:

- Modificar configurações de formatação
- Aplicar estilos e templates
- Criar backups locais

---

## 11. ATUALIZAÇÕES DO SOFTWARE

### 11.1 Verificação de Atualizações

O CHAINSAW pode verificar atualizações via:

- **Manual**: Comando `update-from-github.cmd`
- **Conexão**: Acesso ao GitHub para baixar versões mais recentes

### 11.2 Dados Transmitidos

Durante atualização:

- ✅ Acesso ao repositório público do GitHub
- ✅ Download de arquivos do projeto
- ❌ **Nenhum dado pessoal é enviado**
- ❌ **Nenhum dado de uso é coletado**

---

## 12. SOFTWARE DE TERCEIROS

### 12.1 Dependências

O CHAINSAW utiliza apenas:

- **Microsoft Word**: Aplicativo local do usuário
- **PowerShell**: Componente nativo do Windows
- **Windows Registry**: APIs padrão do sistema operacional

### 12.2 Nenhum Serviço Externo

- ❌ Sem APIs de terceiros
- ❌ Sem bibliotecas remotas
- ❌ Sem serviços em nuvem
- ❌ Sem SDKs de analytics

---

## 13. DIREITOS DO USUÁRIO

### 13.1 Controle Total

Você possui **controle total** sobre:

- Instalação e desinstalação do software
- Documentos criados e editados
- Backups e logs armazenados localmente
- Configurações e preferências

### 13.2 Código Aberto

Como software GPLv3, você tem direito de:

- ✅ Inspecionar o código-fonte completo
- ✅ Modificar o software para suas necessidades
- ✅ Redistribuir cópias (mantendo a licença GPLv3)
- ✅ Contribuir com melhorias ao projeto

### 13.3 Sem Vendor Lock-in

- Você pode **cessar o uso** a qualquer momento
- Documentos criados são **compatíveis com Word padrão**
- **Nenhuma dependência** de serviços proprietários

---

## 14. MENORES DE IDADE

- O CHAINSAW **não coleta dados** de menores ou adultos
- Não há restrições de idade para uso do software
- Pais/responsáveis devem supervisionar o uso de computadores por menores

---

## 15. ALTERAÇÕES NESTA POLÍTICA

### 15.1 Notificação de Mudanças

Alterações nesta Política de Privacidade serão:

- Publicadas no repositório GitHub do projeto
- Incluídas em notas de versão (`VERSION`)
- Datadas e versionadas

### 15.2 Aceitação Continuada

O uso continuado do CHAINSAW após alterações constitui aceitação da nova política.

---

## 16. LEGISLAÇÃO APLICÁVEL

### 16.1 Brasil

- **LGPD** - Lei nº 13.709/2018
- **Marco Civil da Internet** - Lei nº 12.965/2014
- **Código de Defesa do Consumidor** - Lei nº 8.078/1990

### 16.2 Internacional

- **GPLv3** - GNU General Public License v3.0
- Conformidade com princípios internacionais de privacidade

### 16.3 Foro

Para questões relacionadas ao software, utilize o repositório GitHub do projeto.

---

## 17. CONTATO E SUPORTE

### 17.1 Questões sobre Privacidade

- **Repositório GitHub**: https://github.com/chrmsantos/chainsaw
- **Issues**: Para reportar problemas ou fazer perguntas
- **Documentação**: Consulte os arquivos .md no repositório

### 17.2 Relatórios de Segurança

Consulte `SECURITY.md` para instruções sobre como reportar vulnerabilidades.

### 17.3 Conformidade LGPD

Para dúvidas sobre conformidade com LGPD, consulte:

- `LGPD_ATESTADO.md`
- `docs/LGPD_CONFORMIDADE.md`

---

## 18. CERTIFICAÇÃO E ATESTADOS

### 18.1 Atestado de Conformidade

Este software possui **Atestado de Conformidade LGPD** disponível em `LGPD_ATESTADO.md`.

### 18.2 Auditoria

O código-fonte está disponível para auditoria em:
https://github.com/chrmsantos/chainsaw

### 18.3 Certificação de Não Coleta

**CERTIFICAMOS** que o CHAINSAW:

- ✅ **NÃO coleta** dados pessoais
- ✅ **NÃO transmite** dados pela internet (exceto para atualizações opcionais)
- ✅ **NÃO armazena** dados em servidores externos
- ✅ **É totalmente auditável** por ser código aberto

---

## 19. GLOSSÁRIO

**Dados Pessoais**: Informação relacionada a pessoa natural identificada ou identificável (Art. 5º, I da LGPD).

**Processamento Local**: Operações realizadas exclusivamente no computador do usuário, sem transmissão externa.

**GPLv3**: Licença de software livre que garante liberdades de uso, modificação e distribuição.

**Privacy by Design**: Princípio de projeto que incorpora privacidade desde a concepção do sistema.

**VBA**: Visual Basic for Applications - linguagem de macros do Microsoft Office.

---

## 20. DECLARAÇÃO FINAL

O CHAINSAW foi desenvolvido com **máximo respeito à privacidade** dos usuários. Por ser um software de processamento local e código aberto, oferece:

- ✅ **Transparência total** através do código-fonte aberto
- ✅ **Controle total** do usuário sobre seus dados
- ✅ **Privacidade por design** desde a concepção
- ✅ **Conformidade plena** com LGPD e princípios internacionais
- ✅ **Liberdade de software** garantida pela GPLv3

**Esta política reflete nosso compromisso com sua privacidade e liberdade digital.**

---

**Data de Vigência:** 04 de dezembro de 2025  
**Versão da Política:** 1.0  
**Versão do Software:** 2.0.3  

**Última Atualização:** 04 de dezembro de 2025

---

© 2025 CHAINSAW Project  
Licenciado sob GNU General Public License v3.0  
Software Livre - Código Aberto - Privacidade Garantida
