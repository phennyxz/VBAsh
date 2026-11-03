# Pipelines — Referência de testes

**Escopo desta entrega:** reconhecimento sintático de pipelines e rejeição segura pelo executor. É uma subetapa da fase “Pipes” (Fase 9 em `IMPLEMENTATION_PLAN.md`); não implementa execução real de pipelines.

Os testes abaixo foram localizados em `feature/vbash-pipelines`, arquivo `VBAshTests.bas`. As linhas são referências aproximadas do snapshot consultado e podem mudar em commits futuros.

| ID | Suíte | Procedimento | Arquivo / localização | Requisito | Categoria | Resultado esperado |
|---|---|---|---|---|---|---|
| PIPE-SYN-01 | ParserAst | TestParserAst | `VBAshTests.bas:344-355` | Reconhecer `a \| b` como nó PIPELINE com dois estágios ordenados e intervalo de origem correto | Positivo | AST contém PIPELINE e os dois comandos simples na ordem original |
| PIPE-SYN-02 | ParserAst | TestParserAst | `VBAshTests.bas:358-362` | Reconhecer pipeline de três estágios | Positivo | AST contém três estágios na ordem original |
| PIPE-SYN-03 | ParserAst | TestParserAst | `VBAshTests.bas:365-368` | Respeitar precedência entre pipeline, AND/OR e sequência | Integração | Pipeline é agrupado antes de AND/OR, e sequência fica no nível externo |
| PIPE-SYN-04 | ParserAst | TestParserAst | `VBAshTests.bas` (bloco de testes de aspas) | Tratar `|` dentro de aspas simples e duplas como texto literal | Regressão | Pipe citado não cria estágio de pipeline |
| PIPE-SYN-05 | ParserAst | TestParserAst | `VBAshTests.bas:380-402` | Manter redirecionamentos associados ao estágio correspondente e validar casos relacionados | Integração | Redirecionamento permanece no comando simples correto |
| PIPE-SYN-06 | ParserAst | TestParserAst | `VBAshTests.bas:403-406` | Rejeitar pipe inicial, final, estágios vazios e combinação inválida de operadores | Negativo | Parser reporta erro de sintaxe |
| PIPE-SYN-07 | ParserAst | TestParserAst | `VBAshTests.bas:409` | Preservar o comportamento da API legada Parse para pipelines | Regressão | API legada continua rejeitando pipeline |
| PIPE-SYN-08 | AstExecution | TestAstExecution | `VBAshTests.bas:765-773` | Rejeitar PIPELINE antes de executar qualquer estágio | Negativo / segurança | Erro explícito, status 2, sem stdout e sem efeito colateral do primeiro estágio |

## Cobertura ainda não implementada

Os requisitos abaixo pertencem à execução real de pipelines, que permanece fora do escopo desta subetapa. São itens adiados, não falhas dos testes de reconhecimento sintático acima.

| ID | Requisito futuro | Categoria | Resultado esperado |
|---|---|---|---|
| PIPE-RUN-01 | Conectar stdout de um estágio ao stdin do seguinte | Integração | Dados produzidos por uma etapa são recebidos pela seguinte |
| PIPE-RUN-02 | Gerenciar estágios e ciclo de vida de processos como um job | Integração | O pipeline é acompanhado como uma unidade |
| PIPE-RUN-03 | Definir e manter semântica do status final do pipeline | Regressão | Status segue a política documentada |
| PIPE-RUN-04 | Avaliar suporte a `\|&` | Positivo / compatibilidade | stdout e stderr são encaminhados conforme a semântica definida, se suportado |
| PIPE-RUN-05 | Executar pipeline em background, se o gerenciador de jobs suportar | Integração | O shell permanece responsivo e o job é acompanhado |

## Referências

- `IMPLEMENTATION_PLAN.md`, Fase 9 — Pipes.
- `clsParser.cls` — parsing e construção da AST de pipelines.
- `clsAstExecutor.cls` — rejeição explícita do nó PIPELINE sem executar filhos.
- `VBAshTests.bas` — procedimentos `TestParserAst` e `TestAstExecution` na branch `feature/vbash-pipelines`.
