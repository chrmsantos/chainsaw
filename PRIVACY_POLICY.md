# POLITICA DE PRIVACIDADE

**Projeto:** CHAINSAW - Sistema de Padronizacao de Proposituras Legislativas  
**Versao:** 2.8.8  
**Data de Vigencia:** 04 de dezembro de 2025  
**Ultima Atualizacao:** 04 de dezembro de 2025

---

## 1. INTRODUCAO

Esta Politica de Privacidade descreve como o CHAINSAW trata informacoes em conformidade com:

- **LGPD** - Lei Geral de Protecao de Dados Pessoais (Lei n 13.709/2018)
- **GPLv3** - GNU General Public License versao 3

O CHAINSAW e um software livre e de codigo aberto que processa documentos **exclusivamente no computador do usuario**, sem coletar, transmitir ou armazenar dados pessoais em servidores externos.

---

## 2. CONTROLADOR E ENCARREGADO DE DADOS

### 2.1 Controlador de Dados

**Nome:** Projeto CHAINSAW (Open Source)  
**Repositorio:** <https://github.com/chrmsantos/chainsaw>  
**Natureza:** Software Livre (GPLv3)  
**Tipo:** Ferramenta de Processamento Local

### 2.2 Encarregado de Dados (DPO)

Em conformidade com o Art. 41 da LGPD, informamos que:

- O CHAINSAW **nao coleta dados pessoais**
- Nao ha necessidade de nomeacao de Encarregado de Dados
- Questoes sobre privacidade podem ser direcionadas ao repositorio do projeto

---

## 3. COLETA E PROCESSAMENTO DE DADOS

### 3.1 Declaracao de Nao Coleta

**O CHAINSAW NAO coleta dados pessoais.**

O sistema opera de acordo com os principios de **Privacy by Design** e **Privacy by Default**:

- **NAO coleta** dados pessoais de usuarios
- **NAO transmite** informacoes pela internet
- **NAO armazena** dados em servidores externos
- **NAO rastreia** comportamento de usuarios
- **NAO utiliza** cookies, telemetria ou analytics
- **NAO possui** sistema de autenticacao ou cadastro
- **NAO compartilha** dados com terceiros

### 3.2 Dados Tecnicos Processados Localmente

O CHAINSAW processa **exclusivamente em seu computador**:

| Tipo de Dado | Finalidade | Localizacao | Base Legal (LGPD) |
| --- | --- | --- | --- |
| Configuracoes do Word | Aplicar formatacao padrao | Registro do Windows | Art. 7, V (Execucao de contrato) |
| Backups de documentos | Recuperacao em caso de erro | `%TEMP%\.chainsaw\props\backups` | Art. 7, V (Execucao de contrato) |
| Logs de operacao | Debugging e diagnostico | `%USERPROFILE%\chainsaw\source\logs` | Art. 7, V (Execucao de contrato) |
| Templates | Padronizacao de documentos | `%APPDATA%\Microsoft\Templates` | Art. 7, V (Execucao de contrato) |

**Importante:** Nenhum dos dados acima constitui "dado pessoal" conforme Art. 5, I da LGPD.

### 3.3 Documentos do Usuario

- Os documentos criados e editados **permanecem no seu computador**
- O CHAINSAW **nao acessa** o conteudo dos documentos para fins alem da formatacao
- O CHAINSAW **nao envia** documentos para servidores externos
- Voce mantem **controle total** sobre seus arquivos

---

## 4. CONFORMIDADE COM A LGPD

### 4.1 Principios da LGPD (Art. 6)

| Principio | Status | Implementacao |
| --- | --- | --- |
| **Finalidade** | Conforme | Formatacao de documentos legislativos |
| **Adequacao** | Conforme | Processamento compativel com finalidade declarada |
| **Necessidade** | Conforme | Apenas dados tecnicos essenciais ao funcionamento |
| **Livre Acesso** | Conforme | Logs e backups acessiveis localmente pelo usuario |
| **Qualidade dos Dados** | Conforme | Dados tecnicos de sistema, sem dados pessoais |
| **Transparencia** | Conforme | Codigo aberto (GPLv3), documentacao completa |
| **Seguranca** | Conforme | Processamento local isolado, sem transmissao |
| **Prevencao** | Conforme | Arquitetura projetada para nao coletar dados |
| **Nao Discriminacao** | Conforme | Software acessivel a todos os usuarios |
| **Responsabilizacao** | Conforme | Documentacao e codigo auditavel |

### 4.2 Direitos dos Titulares (Art. 18)

Como o CHAINSAW **nao coleta dados pessoais**, os seguintes direitos nao sao aplicaveis:

- Confirmacao e acesso aos dados
- Correcao de dados incompletos ou desatualizados
- Anonimizacao, bloqueio ou eliminacao
- Portabilidade dos dados
- Revogacao do consentimento

**Seus documentos locais:** Voce possui controle total sobre arquivos armazenados em seu computador e pode exclui-los a qualquer momento.

### 4.3 Base Legal

O processamento tecnico local e realizado com base em:

- **Art. 7, V da LGPD**: Execucao de contrato (uso do software)
- **Legitimo interesse**: Funcionalidades essenciais ao software (Art. 10)

---

## 5. CONFORMIDADE COM GPLv3

### 5.1 Liberdade e Transparencia

O CHAINSAW e licenciado sob **GNU General Public License v3.0**:

- Codigo-fonte **100% aberto** e auditavel
- Liberdade para **executar, estudar, modificar e distribuir**
- **Transparencia total** nas operacoes do software
- **Sem componentes proprietarios** ocultos
- **Copyleft**: Derivacoes devem manter a mesma licenca

### 5.2 Garantias de Privacidade pela GPLv3

A licenca GPLv3 garante:

1. **Transparencia**: Voce pode auditar o codigo para verificar praticas de privacidade
2. **Controle**: Voce pode modificar o software para atender suas necessidades
3. **Ausencia de backdoors**: Codigo aberto impede coleta oculta de dados
4. **Responsabilidade**: A comunidade pode identificar e corrigir problemas

### 5.3 Ausencia de Garantias

Conforme GPLv3, Secao 15 e 16:

> O CHAINSAW e fornecido "COMO ESTA", sem garantias de qualquer tipo. O desenvolvedor nao se responsabiliza por danos decorrentes do uso do software.

---

## 6. SEGURANCA E PROTECAO

### 6.1 Medidas de Seguranca Tecnicas

O CHAINSAW implementa:

- **Processamento Local**: Zero transmissao de dados pela rede
- **Isolamento**: Operacao restrita ao ambiente do usuario
- **Backups Locais**: Sistema de protecao em diretorio do usuario
- **Logs Locais**: Registros armazenados apenas localmente
- **Sem Autenticacao Externa**: Nao ha credenciais a serem comprometidas

### 6.2 Responsabilidade do Usuario

Voce e responsavel por:

- Proteger seu computador com antivirus e firewall
- Manter backups dos seus documentos importantes
- Proteger o acesso fisico ao seu computador
- Atualizar o sistema operacional e o Microsoft Word

### 6.3 Incidentes de Seguranca

Para reportar vulnerabilidades de seguranca, consulte `SECURITY.md` no repositorio do projeto.

---

## 7. ARMAZENAMENTO E RETENCAO

### 7.1 Localizacao dos Dados

Todos os dados sao armazenados **localmente** em:

```text
%USERPROFILE%\chainsaw\
 source\logs\                          # Logs de operacao
 %TEMP%\.chainsaw\props\backups\       # Backups de documentos (runtime)
 props\recovery_tmp\                   # Arquivos temporarios de recuperacao

%APPDATA%\Microsoft\Templates\
 Normal.dotm                           # Template do Word
```

### 7.2 Retencao de Dados

- **Logs**: Retencao limitada (mantem os 5 arquivos mais recentes por padrao)
- **Backups**: Retencao limitada (mantem os 10 backups mais recentes por documento, por padrao)
- **Configuracoes**: Mantidas enquanto o software for utilizado

### 7.3 Exclusao de Dados

Para remover todos os dados:

1. Feche o Word
2. Exclua manualmente a pasta `%USERPROFILE%\chainsaw`
3. Remova o template `Normal.dotm` se desejar

---

## 8. TRANSFERENCIA INTERNACIONAL DE DADOS

**Nao aplicavel**: O CHAINSAW nao transfere dados para fora do seu computador.

---

## 9. COOKIES E RASTREAMENTO

**O CHAINSAW nao utiliza:**

- Cookies
- Web beacons
- Pixels de rastreamento
- Analytics (Google Analytics, etc.)
- Telemetria
- Fingerprinting
- Rastreamento de comportamento

---

## 10. INTEGRACAO COM MICROSOFT WORD

### 10.1 VBA e Macros

O CHAINSAW utiliza **VBA (Visual Basic for Applications)** para:

- Aplicar formatacao a documentos
- Criar backups de seguranca
- Gerenciar templates

### 10.2 Acesso a Documentos

- O VBA acessa documentos **apenas para formatacao**
- Nao ha **leitura de conteudo** para fins de analise ou coleta
- Nao ha **envio de dados** para servidores externos

### 10.3 Permissoes do Word

O CHAINSAW requer permissoes padrao do Word para:

- Modificar configuracoes de formatacao
- Aplicar estilos e templates
- Criar backups locais

---

## 11. ATUALIZACOES DO SOFTWARE

### 11.1 Verificacao de Atualizacoes

O CHAINSAW pode verificar atualizacoes via:

- **Manual**: Comando `update-from-github.cmd`
- **Conexao**: Acesso ao GitHub para baixar versoes mais recentes

### 11.2 Dados Transmitidos

Durante atualizacao:

- Acesso ao repositorio publico do GitHub
- Download de arquivos do projeto
- **Nenhum dado pessoal e enviado**
- **Nenhum dado de uso e coletado**

---

## 12. SOFTWARE DE TERCEIROS

### 12.1 Dependencias

O CHAINSAW utiliza apenas:

- **Microsoft Word**: Aplicativo local do usuario
- **PowerShell**: Componente nativo do Windows
- **Windows Registry**: APIs padrao do sistema operacional

### 12.2 Nenhum Servico Externo

- Sem APIs de terceiros
- Sem bibliotecas remotas
- Sem servicos em nuvem
- Sem SDKs de analytics


---

## 13. DIREITOS DO USUARIO

### 13.1 Controle Total

Voce possui **controle total** sobre:

- Uso e remocao do software
- Documentos criados e editados
- Backups e logs armazenados localmente
- Configuracoes e preferencias

### 13.2 Codigo Aberto

Como software GPLv3, voce tem direito de:

- Inspecionar o codigo-fonte completo
- Modificar o software para suas necessidades
- Redistribuir copias (mantendo a licenca GPLv3)
- Contribuir com melhorias ao projeto

### 13.3 Sem Vendor Lock-in

- Voce pode **cessar o uso** a qualquer momento
- Documentos criados sao **compativeis com Word padrao**
- **Nenhuma dependencia** de servicos proprietarios

---

## 14. MENORES DE IDADE

- O CHAINSAW **nao coleta dados** de menores ou adultos
- Nao ha restricoes de idade para uso do software
- Pais/responsaveis devem supervisionar o uso de computadores por menores

---

## 15. ALTERACOES NESTA POLITICA

### 15.1 Notificacao de Mudancas

Alteracoes nesta Politica de Privacidade serao:

- Publicadas no repositorio GitHub do projeto
- Incluidas em notas de versao (`VERSION`)
- Datadas e versionadas

### 15.2 Aceitacao Continuada

O uso continuado do CHAINSAW apos alteracoes constitui aceitacao da nova politica.

---

## 16. LEGISLACAO APLICAVEL

### 16.1 Brasil

- **LGPD** - Lei n 13.709/2018
- **Marco Civil da Internet** - Lei n 12.965/2014
- **Codigo de Defesa do Consumidor** - Lei n 8.078/1990

### 16.2 Internacional

- **GPLv3** - GNU General Public License v3.0
- Conformidade com principios internacionais de privacidade

### 16.3 Foro

Para questoes relacionadas ao software, utilize o repositorio GitHub do projeto.

---

## 17. CONTATO E SUPORTE

### 17.1 Questoes sobre Privacidade

- **Repositorio GitHub**: <https://github.com/chrmsantos/chainsaw>
- **Issues**: Para reportar problemas ou fazer perguntas
- **Documentacao**: Consulte os arquivos .md no repositorio

### 17.2 Relatorios de Seguranca

Consulte `SECURITY.md` para instrucoes sobre como reportar vulnerabilidades.

### 17.3 Conformidade LGPD

Para duvidas sobre conformidade com LGPD, consulte:

- `LGPD_ATESTADO.md`

---

## 18. CERTIFICACAO E ATESTADOS

### 18.1 Atestado de Conformidade

Este software possui **Atestado de Conformidade LGPD** disponivel em `LGPD_ATESTADO.md`.

### 18.2 Auditoria

O codigo-fonte esta disponivel para auditoria em:
<https://github.com/chrmsantos/chainsaw>

### 18.3 Certificacao de Nao Coleta

**CERTIFICAMOS** que o CHAINSAW:

- **NAO coleta** dados pessoais
- **NAO transmite** dados pela internet (exceto para atualizacoes opcionais)
- **NAO armazena** dados em servidores externos
- **E totalmente auditavel** por ser codigo aberto

---

## 19. GLOSSARIO

**Dados Pessoais**: Informacao relacionada a pessoa natural identificada ou identificavel (Art. 5, I da LGPD).

**Processamento Local**: Operacoes realizadas exclusivamente no computador do usuario, sem transmissao externa.

**GPLv3**: Licenca de software livre que garante liberdades de uso, modificacao e distribuicao.

**Privacy by Design**: Principio de projeto que incorpora privacidade desde a concepcao do sistema.

**VBA**: Visual Basic for Applications - linguagem de macros do Microsoft Office.

---

## 20. DECLARACAO FINAL

O CHAINSAW foi desenvolvido com **maximo respeito a privacidade** dos usuarios. Por ser um software de processamento local e codigo aberto, oferece:

- **Transparencia total** atraves do codigo-fonte aberto
- **Controle total** do usuario sobre seus dados
- **Privacidade por design** desde a concepcao
- **Conformidade plena** com LGPD e principios internacionais
- **Liberdade de software** garantida pela GPLv3

**Esta politica reflete nosso compromisso com sua privacidade e liberdade digital.**

---

**Data de Vigencia:** 04 de dezembro de 2025  
**Versao da Politica:** 1.0  
**Versao do Software:** 2.8.8  

**Ultima Atualizacao:** 04 de dezembro de 2025

---

 2025 CHAINSAW Project  
Licenciado sob GNU General Public License v3.0  
Software Livre - Codigo Aberto - Privacidade Garantida
