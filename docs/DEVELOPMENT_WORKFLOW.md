# VBAsh — Fluxo de desenvolvimento

O VBAsh segue uma única linha evolutiva. A branch canônica atual é feature/phase-3-integration; main é uma referência histórica atrasada e não deve ser usada como base ou destino de integração.

## 1. Antes de cada tarefa

No clone local:

```bash
git fetch --all --prune
git status --short --branch
git branch -vv
git rev-parse HEAD
git rev-parse origin/feature/phase-3-integration
git log -1 --oneline --decorate
```

- Confirmar que o clone está na linha canônica.
- Identificar alterações locais antes de atualizar ou trocar de branch.
- Não usar main antiga nem branches Copilot antigas como base.
- Ler docs/README.md, docs/PROJECT_STATUS.md e a seção correspondente de IMPLEMENTATION_PLAN.md.

## 2. Modelo de trabalho

- Trabalhar diretamente na branch canônica para tarefas normais.
- Fazer commits pequenos, cada um com objetivo claro.
- Não criar uma branch permanente para cada fase.
- Criar branch temporária somente quando houver necessidade real de isolamento ou risco de uma experiência.
- Se uma branch temporária for criada, integrar ou descartar ao fim da tarefa e removê-la; não deixá-la como linha paralela esquecida.
- Não abrir PR para o próprio fluxo rotineiro da linha única.
- A branch main histórica não deve receber merges da linha canônica enquanto não for normalizada/renomeada.

## 3. Contrato para o Copilot

Todo prompt de implementação deve conter:

1. Repositório e branch canônica exatos.
2. SHA-base conhecido.
3. Objetivo e comportamento esperado.
4. Arquivos permitidos para alteração.
5. Arquivos/áreas proibidos.
6. Contratos públicos que devem ser preservados.
7. Casos de teste e critérios de aceitação.
8. Instrução explícita para não declarar teste Excel PASS sem execução real.
9. Entregáveis: resumo, lista de arquivos, decisões, riscos, testes executados/não executados e commit SHA.

Quando o código já existe, o Copilot deve primeiro produzir auditoria e matriz de lacunas. Não pedir “implemente a fase” como se ela estivesse vazia.

## 4. Implementação e revisão

- Manter cada alteração pequena e com escopo explícito.
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
- Atualizar docs/test-reports/README.md e docs/PROJECT_STATUS.md após uma execução relevante.

## 7. Fechamento de tarefa

Ao concluir:
1. registrar commit e SHA;
2. atualizar IMPLEMENTATION_PLAN.md se o status/escopo da fase mudou;
3. atualizar docs/PROJECT_STATUS.md;
4. atualizar referências/logs de teste;
5. atualizar docs/BRANCH_STATUS.md somente quando a estratégia ou o estado das referências mudar;
6. deixar explícito o próximo passo na mesma linha canônica.

