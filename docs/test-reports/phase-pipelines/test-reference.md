# Pipelines — Referência de testes

**Fase do projeto:** Pipes (Fase 9 em `IMPLEMENTATION_PLAN.md`), identificada neste relatório como Pipelines.

Os requisitos abaixo vêm dos critérios e testes mínimos documentados no plano. O snapshot consultado não contém os procedimentos de teste nem seus nomes; por isso, as entradas são lacunas de rastreabilidade, não testes implementados. A execução informada pelo desenvolvedor menciona 13 suítes, mas não identifica individualmente nenhuma delas.

| ID | Suíte | Procedimento | Arquivo | Requisito | Categoria | Resultado esperado |
|---|---|---|---|---|---|---|
| PIPE-GAP-01 | Não localizado | Não localizado | Não localizado neste snapshot | Executar corretamente `cmd1 \| cmd2` (plano, Fase 9) | Lacuna — teste não localizado | O pipeline executa os comandos e produz a saída correta. |
| PIPE-GAP-02 | Não localizado | Não localizado | Não localizado neste snapshot | Conectar stdout de um processo ao stdin do seguinte (plano, Fase 9) | Lacuna — teste não localizado | Os dados emitidos por uma etapa são recebidos pela etapa seguinte. |
| PIPE-GAP-03 | Não localizado | Não localizado | Não localizado neste snapshot | Monitorar os processos de um pipeline como um único job (plano, Fase 9) | Lacuna — teste não localizado | O pipeline é representado e acompanhado como um único job. |
| PIPE-GAP-04 | Não localizado | Não localizado | Não localizado neste snapshot | Manter corretamente o status do pipeline (plano, Fase 9) | Lacuna — teste não localizado | O status final corresponde ao comportamento definido para o pipeline. |
| PIPE-GAP-05 | Não localizado | Não localizado | Não localizado neste snapshot | Suportar `\|&` em nível mínimo, se a arquitetura permitir (plano, Fase 9) | Lacuna — teste não localizado | Quando suportado, stdout e stderr são encaminhados conforme a semântica de `\|&`. |
| PIPE-GAP-06 | Não localizado | Não localizado | Não localizado neste snapshot | Executar pipeline de múltiplas etapas e manter saída e status consistentes (teste mínimo planejado, Fase 9) | Lacuna — teste não localizado | Um pipeline `a \| b \| c` produz a saída composta e o status esperado. |
| PIPE-GAP-07 | Não localizado | Não localizado | Não localizado neste snapshot | Manter status consistente quando um comando do pipeline falha (teste mínimo planejado, Fase 9) | Lacuna — teste não localizado | Uma etapa com falha não deixa o status final em estado inconsistente. |
| PIPE-GAP-08 | Não localizado | Não localizado | Não localizado neste snapshot | Registrar pipeline em background como job sem bloquear o shell (teste mínimo planejado, Fase 9) | Lacuna — teste não localizado | O pipeline aparece na tabela de jobs e o shell permanece responsivo. |

As linhas `PIPE-GAP-*` identificam requisitos sem teste localizável neste snapshot; não afirmam que nenhum teste exista em outra branch ou artefato. Para fechar cada lacuna, substituir/acompanhar a entrada com o ID real, suíte, procedimento e arquivo, preservando a descrição da cobertura.

## Referências

- `IMPLEMENTATION_PLAN.md`, Fase 9 — Pipes, critérios de conclusão e testes mínimos.
- `docs/test-reports/phase-pipelines/execution-log.md` — relato agregado de execução recebido; não fornece a relação individual dos testes.
