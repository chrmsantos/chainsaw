# Contribution Guide

## Scope

- Foque na padronização de documentos do Microsoft Word escrita em VBA.
- Mantenha a abordagem de módulo único salvo quando a manutenção exigir divisão.
- Evite novas dependências, telemetria ou recursos que não envolvam automação do Word.

## Quality Principles

- Prefira padrões defensivos compatíveis com Word 2010 ou superior.
- Reconfigure o tratamento de erros explicitamente após blocos com `On Error Resume Next`.
- Utilize helpers (`SafeSetFont`, `ReplacePlaceholders` etc.) em vez de lógica ad-hoc.
- Agrupe operações para reduzir flicker de interface e mudanças de seleção desnecessárias.

## Workflow

1. Faça fork do repositório e crie uma branch temática (por exemplo `fix/paragraph-spacing`).
2. Implemente a mudança com commits focados; evite ajustes de formatação não relacionados.
3. Atualize a documentação sempre que houver alteração de comportamento ou uso.
4. Teste o macro em documentos exemplo cobrindo título, seções numeradas, carimbos e fluxos de fallback.
5. Abra um pull request descrevendo o problema, a solução e eventuais efeitos colaterais.

## Review Checklist

- Código compila no Word VBA (2010+).
- Espaçamento e nomenclatura seguem o estilo atual do módulo.
- Nenhuma nova dependência externa ou edição de registro é adicionada.
- Notas de validação manual (documentos usados, versão do Word) aparecem no corpo do PR.

## Conduct

Trate todos os participantes com respeito. Relate comportamentos abusivos ao mantenedor imediatamente.

## Licensing

As contribuições são aceitas sob GPL-3.0-or-later. Confirme que você tem direito de enviar o código.

## Contact

Questões sobre o processo podem ser registradas em issues do GitHub ou enviadas para [chrmsantos@gmail.com](mailto:chrmsantos@gmail.com).
