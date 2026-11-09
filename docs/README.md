# VBAsh — Documentação

Este diretório reúne a documentação de arquitetura, planejamento, governança e evidências. Comece por esta página para localizar a fonte correta antes de alterar código ou atualizar um status.

## Documentos de referência

| Documento | Finalidade | Quando atualizar |
|---|---|---|
| [Estado atual do projeto](PROJECT_STATUS.md) | Fotografia do que existe, o que foi testado e o que permanece pendente | Ao integrar uma fase ou receber novo relatório de execução |
| [Estado das branches](BRANCH_STATUS.md) | Inventário datado das branches e decisões de preservação/limpeza | Após merges, criação, arquivamento ou exclusão de branches |
| [Fluxo de desenvolvimento](DEVELOPMENT_WORKFLOW.md) | Regras para Copilot, branches, testes, commits, PRs e rastreabilidade | Quando o processo de trabalho mudar |
| [Plano de implementação](../IMPLEMENTATION_PLAN.md) | Roadmap e critérios por fase | Ao alterar escopo, dependências ou critérios |
| [Referências e relatórios de testes](test-reports/README.md) | Índice de cobertura e evidências de execução | Sempre que uma suíte, requisito ou execução mudar |

## Regra de precedência

- O código-fonte do branch/commit consultado é a evidência do que está implementado naquele snapshot.
- O relatório de execução é a evidência do que foi efetivamente executado.
- A referência de testes liga requisitos aos procedimentos existentes.
- O plano define o que se pretende construir, mas não prova que esteja implementado.
- PROJECT_STATUS resume o estado atual e aponta para as evidências; não substitui os documentos de origem.

Se houver divergência, não escolha silenciosamente uma versão: registre a divergência, identifique o snapshot e reconcilie os documentos.
