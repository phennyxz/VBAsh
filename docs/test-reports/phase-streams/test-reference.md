# Streams — Referência de testes

**Fase do projeto:** Streams (Fase 3 em `IMPLEMENTATION_PLAN.md`).

Os requisitos abaixo são derivados dos critérios de conclusão da fase no plano. `STR-CONTRACT-01` foi escrito nesta branch para verificar a interface VBA e seu dispatch por `Implements`; ele usa um probe de teste, não implementações de stream de produção. Nenhum dos testes listados foi executado neste ambiente.

| ID | Suíte | Procedimento | Arquivo | Requisito | Categoria | Resultado esperado |
|---|---|---|---|---|---|---|
| STR-CONTRACT-01 | StreamContract | TestStreamContract | `VBAshTests.bas` | Contrato `IStream`: leitura, escrita, fechamento, capacidades, EOF, estado de erro e rejeição de operações não suportadas (contrato desta implementação) | Teste escrito — NOT RUN | Implementação VBA dispatcha os membros; leitura/escrita retornam os resultados definidos, EOF/fechamento são observáveis, operações não suportadas lançam `VBAsh_STREAM_ERROR_NOT_SUPPORTED` e preenchem estado de erro, sucesso limpa o erro. |
| STR-GAP-01 | Não localizado | Não localizado | Não localizado neste snapshot | Manter uma tabela de descritores com 0, 1 e 2 definidos (plano, Fase 3) | Lacuna — teste não localizado | Os descritores padrão para stdin, stdout e stderr estão definidos. |
| STR-GAP-02 | Não localizado | Não localizado | Não localizado neste snapshot | Ler stdin do console (plano, Fase 3) | Lacuna — teste não localizado | A entrada lida corresponde ao conteúdo fornecido pelo console. |
| STR-GAP-03 | Não localizado | Não localizado | Não localizado neste snapshot | Escrever stdout no console (plano, Fase 3) | Lacuna — teste não localizado | O texto escrito em stdout chega ao console. |
| STR-GAP-04 | Não localizado | Não localizado | Não localizado neste snapshot | Separar stderr de stdout e permitir redirecionar stderr para arquivo ou buffer (plano, Fase 3) | Lacuna — teste não localizado | stderr pode ser capturado/redirecionado sem se misturar indevidamente com stdout. |
| STR-GAP-05 | Não localizado | Não localizado | Não localizado neste snapshot | Redirecionar stdout com `> file` (teste mínimo planejado, Fase 3) | Lacuna — teste não localizado | O arquivo de saída é criado ou sobrescrito com o conteúdo esperado. |
| STR-GAP-06 | Não localizado | Não localizado | Não localizado neste snapshot | Redirecionar stdin com `< file` (teste mínimo planejado, Fase 3) | Lacuna — teste não localizado | O conteúdo do arquivo é lido como entrada. |
| STR-GAP-07 | Não localizado | Não localizado | Não localizado neste snapshot | Abstrair duplicação de descritores (`dup`, `dup2`) e combinação equivalente a `2>&1` (plano, Fase 3) | Lacuna — teste não localizado | O descritor duplicado referencia o fluxo esperado e a saída combinada é produzida corretamente. |
| STR-GAP-08 | Não localizado | Não localizado | Não localizado neste snapshot | Modelar `/dev/null` como stream nulo (plano, Fase 3) | Lacuna — teste não localizado | Escritas são aceitas sem erro e o conteúdo é descartado. |

As linhas `STR-GAP-*` indicam somente que não foi possível localizar testes correspondentes neste snapshot. `STR-CONTRACT-01` cobre somente a interface e seu probe, não o comportamento das futuras implementações de streams. Ao implementar ou localizar cada teste, registre o ID real, o nome exato da suíte e do procedimento, o arquivo e a categoria, mantendo explícitas as lacunas ainda abertas.

## Referências

- `IMPLEMENTATION_PLAN.md`, Fase 3 — Streams, critérios de conclusão e testes mínimos.
- `docs/test-reports/phase-streams/execution-log.md` — status atual das evidências de execução.
