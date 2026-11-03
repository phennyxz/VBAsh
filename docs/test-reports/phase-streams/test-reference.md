# Streams — Referência de testes

**Fase do projeto:** Streams (Fase 3 em `IMPLEMENTATION_PLAN.md`).

Os requisitos abaixo são derivados dos critérios de conclusão da fase no plano. `STR-CONTRACT-01` verifica a interface VBA e seu dispatch por `Implements`; `STR-NULL-01` cobre a implementação de produção `clsNullStream`. Nenhum dos testes listados foi executado neste ambiente.

| ID | Suíte | Procedimento | Arquivo | Requisito | Categoria | Resultado esperado |
|---|---|---|---|---|---|---|
| STR-CONTRACT-01 | StreamContract | TestStreamContract | `VBAshTests.bas` | Contrato textual `IStream`: leitura, escrita, fechamento idempotente, capacidades, EOF significativo somente em stream aberto e legível, propriedades independentes `CanRead`/`IsClosed`/`HasError` e códigos distintos para falha não suportada, fechada, leitura, escrita e abertura | Teste escrito — NOT RUN | Probe implementa o contrato por `Implements`; `AtEndOfStream` é False se não legível ou fechado; `CanRead`, `IsClosed` e `HasError` representam apenas capacidade, estado de fechamento e ocorrência de falha; operações não suportadas/em stream fechado preenchem erro sem alterar as outras propriedades; fechamentos repetidos não falham. |
| STR-NULL-01 | NullStream | TestNullStream | `VBAshTests.bas` | Implementação `clsNullStream` por `IStream`: capacidades, criação aberta, leitura vazia e EOF imediato, escritas descartadas com contagem de caracteres, ausência/limpeza de erro em operações válidas, fechamento idempotente e erros de leitura/escrita após fechar | Teste escrito — NOT RUN | Confirma que escritas comuns, vazias e sucessivas nunca produzem dados de leitura; após fechar `CanRead`/`CanWrite` permanecem True e EOF é False; operações após fechar geram `VBAsh_STREAM_ERROR_CLOSED` e registram estado de erro consistente. |
| STR-GAP-01 | Não localizado | Não localizado | Não localizado neste snapshot | Manter uma tabela de descritores com 0, 1 e 2 definidos (plano, Fase 3) | Lacuna — teste não localizado | Os descritores padrão para stdin, stdout e stderr estão definidos. |
| STR-GAP-02 | Não localizado | Não localizado | Não localizado neste snapshot | Ler stdin do console (plano, Fase 3) | Lacuna — teste não localizado | A entrada lida corresponde ao conteúdo fornecido pelo console. |
| STR-GAP-03 | Não localizado | Não localizado | Não localizado neste snapshot | Escrever stdout no console (plano, Fase 3) | Lacuna — teste não localizado | O texto escrito em stdout chega ao console. |
| STR-GAP-04 | Não localizado | Não localizado | Não localizado neste snapshot | Separar stderr de stdout e permitir redirecionar stderr para arquivo ou buffer (plano, Fase 3) | Lacuna — teste não localizado | stderr pode ser capturado/redirecionado sem se misturar indevidamente com stdout. |
| STR-GAP-05 | Não localizado | Não localizado | Não localizado neste snapshot | Redirecionar stdout com `> file` (teste mínimo planejado, Fase 3) | Lacuna — teste não localizado | O arquivo de saída é criado ou sobrescrito com o conteúdo esperado. |
| STR-GAP-06 | Não localizado | Não localizado | Não localizado neste snapshot | Redirecionar stdin com `< file` (teste mínimo planejado, Fase 3) | Lacuna — teste não localizado | O conteúdo do arquivo é lido como entrada. |
| STR-GAP-07 | Não localizado | Não localizado | Não localizado neste snapshot | Abstrair duplicação de descritores (`dup`, `dup2`) e combinação equivalente a `2>&1` (plano, Fase 3) | Lacuna — teste não localizado | O descritor duplicado referencia o fluxo esperado e a saída combinada é produzida corretamente. |
| STR-GAP-08 | Não localizado | Não localizado | Não localizado neste snapshot | Integrar o stream nulo a `/dev/null` ou redirecionamentos (fora do escopo deste marco) | Lacuna — teste não localizado | Qualquer integração funcional depende de marcos posteriores. |

As linhas `STR-GAP-*` indicam somente que não foi possível localizar testes correspondentes neste snapshot. `clsNullStream` implementa o contrato `IStream`, mas não é integrado a pipes, redirecionamentos ou executor neste marco. Ao implementar ou localizar cada teste, registre o ID real, o nome exato da suíte e do procedimento, o arquivo e a categoria, mantendo explícitas as lacunas ainda abertas.

## Referências

- `IMPLEMENTATION_PLAN.md`, Fase 3 — Streams, critérios de conclusão e testes mínimos.
- `docs/test-reports/phase-streams/execution-log.md` — status atual das evidências de execução.
