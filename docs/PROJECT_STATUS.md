# VBAsh — Estado atual do projeto

**Snapshot:** 2026-11-09  
**Branch de referência consultada:** feature/phase-3-integration  
**Branch principal consultada:** main  
**HEAD de main na auditoria:** 0d4115f4ced6f3772fa38182a71a34fdfee387ce

Este documento é uma fotografia de estado, não uma promessa de compatibilidade geral com Bash. A implementação descrita deve sempre ser lida junto do branch/commit consultado.

## Resumo executivo

- O desenvolvimento ativo está em feature/phase-3-integration. Na comparação com main, está 155 commits à frente e 0 atrás. A API de comparação não expôs o SHA do HEAD dessa branch nesta consulta; não inventar ou inferir esse SHA.
- O runner VBAshTests-r14 reportou 29/29 suítes PASS, 969 asserções, 0 FAIL e 0 ERROR em execução real no Excel/VBA informada pelo usuário.
- O relatório r14 identificou as revisões dos módulos carregados, mas checkout/branch/SHA do workbook estavam UNKNOWN e a correspondência checkout/módulos UNVERIFIED.
- A Fase 4 (Command Registry) pode ser considerada funcionalmente validada no ambiente testado, com a ressalva de rastreabilidade.
- Lexer e parser já existem e têm cobertura no r14. A Fase 5 deve ser tratada como auditoria e endurecimento da implementação existente, não como criação do lexer do zero.
- A resolução de funções e executáveis externos está implementada; sua execução não está coberta nem faz parte da conclusão da Fase 4.

## Evidência de execução mais recente

Fonte: [phase-4-command-registry-r14.md](test-reports/phase-4-command-registry-r14.md).

| Suíte | Asserções | Resultado informado |
|---|---:|---|
| Tokenizer | 25 | PASS |
| TokenizerWordSegments | 41 | PASS |
| ParserAst | 98 | PASS |
| AstExecution | 46 | PASS |
| StreamContract | 41 | PASS |
| MemoryStream | 112 | PASS |
| NullStream | 42 | PASS |
| FileStream | 74 | PASS |
| FileDescriptorTable | 62 | PASS |
| FileDescriptorTableCloseAll | 41 | PASS |
| DescriptorScopes | 7 | PASS |
| DescriptorRedirection | 58 | PASS |
| ConsoleStream | 6 | PASS |
| PipeStream | 8 | PASS |
| Command/alias/function registry e demais suítes | demais | PASS |
| **Total** | **969** | **29/29 PASS** |

O resultado acima é o relatório informado pelo usuário, preservado no repositório. Não deve ser descrito como execução reproduzida pelo GitHub ou por este assistente. Como o SHA exato do checkout do workbook não foi identificado, não associar a execução a um commit específico.

## Estado por área

| Área | Estado no snapshot | Evidência / ressalva |
|---|---|---|
| Terminal e host Excel | Implementados | VBAsh e instalador; validar no workbook quando houver mudança |
| Session e ambiente | Implementados | clsSession e suítes de execução |
| Streams e descritores | Integrados em parte | r14 cobre streams, tabela, escopos, redirecionamento e adaptadores; stdin interativo continua limitado |
| Command Registry | Funcionalmente validado | r14; funções e externos são classificados/resolvidos, não executados |
| Lexer | Implementado; auditoria pendente | clsTokenizer-r2, clsToken, clsWordSegment; 66 asserções lexicais PASS no r14 informado |
| Parser / AST | Implementado; evolução incremental | clsParser-r3; ParserAst e AstExecution PASS no r14 |
| Builtins | Implementados parcialmente | README lista os comandos disponíveis; não implica compatibilidade Bash completa |
| Process Manager | Pendente | execução de programas externos e gestão de processos não concluídas |
| Pipes | Buffer e sintaxe existentes; execução de pipeline pendente | clsPipeStream é buffer textual; não há execução real de pipeline entre comandos |
| POSIX Layer | Pendente | fase posterior do roadmap |

## Limitações que não devem ser ocultadas

- O shell não é compatível em geral com Bash 5.x.
- O resolver identifica funções e executáveis externos, mas não os executa.
- O pipe atual é um buffer textual; não equivale a scheduler, processos independentes ou pipeline funcional.
- A entrada interativa pelo FD 0 ainda é limitada; redirecionamento de stdin não implica que os builtins consumam esse stream.
- Os resultados do r14 não certificam o SHA do workbook testado.

## Próximo marco

Fazer a auditoria lexical da Fase 5 conforme IMPLEMENTATION_PLAN.md: contrato atual, matriz de casos, lacunas reproduzíveis, correções pequenas e regressão completa no Excel. Não iniciar refatoração estrutural antes dessa auditoria.

