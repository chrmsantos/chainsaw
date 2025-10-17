# Security Policy

Chainsaw Proposituras é um módulo VBA local para Microsoft Word. Ele não inicia requisições de rede nem telemetria, mas aceitamos relatórios de segurança confidenciais.

## Supported Versions

- Suporte ativo: `main` e a release etiquetada mais recente.
- Ambiente-alvo: Microsoft Word 2010 ou superior no Windows. Versões anteriores são atendidas por melhor esforço.

## Reporting

- Canal preferencial: [chrmsantos@gmail.com](mailto:chrmsantos@gmail.com) com assunto "[Security] chainsaw-proposituras".
- Inclua descrição, passos de reprodução ou documento exemplo, versões de Word e Windows e o commit hash quando possível.
- Se email não for opção, abra uma issue no GitHub sem divulgar dados sensíveis e rotule como segurança.
- Metas de resposta (melhor esforço): acusar recebimento em até 7 dias, compartilhar triagem em 14 dias, buscar mitigação de falhas críticas em 30 dias.

## Scope

In scope:

- Módulos VBA e fluxos de macro documentados.
- Manipulação de arquivos, caminhos ou diálogos realizada pelo projeto.

Out of scope:

- Defeitos no Microsoft Word, Office ou Windows.
- Modelos ou complementos de terceiros não incluídos aqui.

## Macro Hygiene

- Habilite macros apenas de fontes confiáveis e mantenha este projeto em um Local Confiável.
- Configuração recomendada no Trust Center: "Desabilitar todas as macros com notificação".
- O macro lê `assets\stamp.png` relativo ao documento ativo; se o arquivo não existir, a etapa é pulada sem I/O alternativo.

## Disclosure

Os relatores podem optar por receber crédito ou permanecer anônimos. O reconhecimento é concedido após a correção ser publicada, salvo solicitação em contrário.
