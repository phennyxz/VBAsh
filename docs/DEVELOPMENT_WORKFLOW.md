# VBAsh — Fluxo de desenvolvimento

Este fluxo existe para evitar divergência entre o que o Copilot altera, o que está no GitHub, o que foi importado no Excel e o que os relatórios dizem que passou.

## 1. Antes de cada tarefa

No clone local:

```bash
git fetch --all --prune
git status --short --branch
git branch -vv
git rev-parse HEAD
git log -1 --oneline --decorate
```

- Confirmar a branch-base e o SHA exato.
- Se houver alterações locais, identificá-las antes de trocar de branch.
- Não usar uma branch Copilot antiga como base sem comparar seu histórico.
- Ler o estado do projeto e a fase correspondente em docs/PROJECT_STATUS.md, docs/BRANCH_STATUS.md e IMPLEMENTATION_PLAN.md.

## 2. Abrir uma tarefa isolada

Padrão de nome: feature/<fase>-<assunto> ou fix/<assunto>; documentação isolada pode usar docs/<assunto>.

Exemplo para a próxima etapa: feature/phase-5-lexer-audit.

A branch deve nascer da ponta atual de feature/phase-3-integration. Registrar o SHA-base no início da tarefa. Não criar a branch até o clone local e a ponta remota terem sido reconciliados.

## 3. Contrato para o Copilot

Todo prompt de implementação deve conter:

1. Repositório e branch-base exatos.
2. SHA-base conhecido.
3. Objetivo e comportamento esperado.
4. Arquivos permitidos para alteração.
5. Arquivos/áreas proibidos.
6. Contratos públicos que devem ser preservados.
7. Casos de teste e critérios de aceitação.
8. Instrução explícita para não declarar teste Excel PASS sem execução real.
9. Entregáveis: resumo, lista de arquivos, decisões, riscos, testes executados/não executados e commit SHA.

O Copilot não deve receber instrução vaga como “implemente a Fase 5” quando a fase já contém código. Primeiro fornecer a auditoria e a matriz de lacunas; só depois autorizar correções.

## 4. Implementação e revisão

- Manter cada tarefa pequena e com escopo explícito.
- Não misturar refatoração estrutural, mudança de comportamento e documentação sem necessidade.
- Preservar API pública ou descrever migração e atualizar consumidores.
- Não criar classes auxiliares apenas por preferência estética; justificar pelo acoplamento, responsabilidade ou teste.
- Revisar diff completo, incluindo arquivos gerados, testes, instalador e documentação.
- Executar verificações estáticas disponíveis, sem confundi-las com compilação VBA.

## 5. Testes no Excel/VBA

A validação final deve ocorrer no ambiente Excel/VBA quando a mudança depender dele.

O relatório precisa registrar:
- data/hora e fuso, quando disponível;
- nome do workbook e versão do Excel/Windows, quando disponíveis;
- branch, SHA exato e estado limpo/modificado do checkout;
- revisão do runner e revisão de cada módulo relevante carregado;
- resultado por suíte, assertions, falhas, erros e duração;
- saída original do runner, sem reescrever o resultado.

O checkout e os módulos carregados são identidades diferentes. Se o runner informar UNKNOWN ou UNVERIFIED, manter essa ressalva. Não atribuir o SHA remoto mais próximo como se fosse o SHA testado.

## 6. Referência e log de testes

- test-reference.md: requisitos, IDs, procedimento exato, arquivo, cobertura e lacunas.
- execution-log.md: histórico cronológico de execuções; acrescentar, nunca apagar resultados antigos.
- Um teste presente no código não está automaticamente executado.
- NOT RUN, BLOCKED, FAIL e ERROR nunca podem ser convertidos em PASS por inspeção estática.
- Atualizar a referência quando o comportamento ou a suíte mudar.
- Atualizar o índice em docs/test-reports/README.md e o resumo em docs/PROJECT_STATUS.md após uma execução relevante.

## 7. Pull request e integração

Antes do PR:
- revisar o diff contra a branch-base;
- confirmar que o branch não está baseado em snapshot antigo;
- verificar que o relatório de testes corresponde às revisões carregadas;
- listar limitações e testes não executados.

O PR deve declarar escopo, evidências, ressalvas e itens explicitamente fora do escopo. O destino padrão das tarefas será feature/phase-3-integration; main só recebe marcos revisados.

## 8. Fechamento de tarefa

Ao concluir:
1. registrar commit e PR;
2. atualizar IMPLEMENTATION_PLAN.md se o status/escopo da fase mudou;
3. atualizar docs/PROJECT_STATUS.md;
4. atualizar referências/logs de teste;
5. atualizar docs/BRANCH_STATUS.md quando branches forem criadas, integradas ou removidas;
6. deixar explícito o próximo passo e a branch-base correta.

