# VBAsh — Inventário de branches

**Data da consulta:** 2026-11-09  
**Repositório:** phennyxz/VBAsh  
**main na consulta:** 0d4115f4ced6f3772fa38182a71a34fdfee387ce

Esta é uma fotografia remota. Status de branch pode mudar; repetir a comparação antes de qualquer exclusão ou merge. “Atrás” e “à frente” são contagens da comparação com main, não uma avaliação de qualidade.

## Branch de trabalho preservada

| Branch | Comparação com main | Leitura |
|---|---:|---|
| feature/phase-3-integration | 147 à frente, 0 atrás | Branch ativa de referência. Contém o estado mais recente usado para o r14 e o trabalho das Fases 3/4. Preservar. |

A comparação apontou 28 arquivos alterados em relação ao ponto de divergência, incluindo tokenizer, parser, sessão, registry, aliases, streams e documentação. O SHA do HEAD desta branch não foi exposto pelo resultado da comparação; consultar localmente com git rev-parse HEAD antes de criar uma branch de trabalho ou registrar execução.

## Branches cujo conteúdo está contido em main

| Branch | Estado relativo a main | Evidência de histórico | Decisão |
|---|---|---|---|
| work | 0 à frente, 34 atrás; ponta identificada pelo merge-base 56be9084a7a97b36bf052acda139c8af188a078b | PR #2 integrado | Candidata a exclusão após confirmar que não há checkout local dependente |
| feature/vbash-pipelines | 0 à frente, 6 atrás; ponta 2b9ea25f4e93c7a481b63a1cee7bed3c669a6bb6 | PR #11 integrado | Candidata a exclusão após confirmar checkout local |
| feature/descriptor-table | 0 à frente, 1 atrás; ponta 4e50829898bef8592c5bb61573059afb40676105 | PR #12 integrado | Candidata a exclusão após confirmar checkout local |
| copilot/featurevbash-pipelines-again | 0 à frente, 15 atrás; merge-base cde73a117c82267d8e953654c81b1f2539caa99d | Sem delta exclusivo observado na comparação com main | Candidata a exclusão após conferência |
| copilot/featurevbash-nullstream | 0 à frente, 16 atrás; ponta 12ddf1de2e54119f50198975c2703e6f993a284e | PR #10 integrado | Candidata a exclusão após confirmar checkout local |
| copilot/featurefile-stream | 0 à frente, 15 atrás; merge-base cde73a117c82267d8e953654c81b1f2539caa99d | Sem delta exclusivo observado na comparação com main | Candidata a exclusão após conferência |
| copilot/copilotsetup-audit-vbash-repo | 0 à frente, 34 atrás; merge-base 56be9084a7a97b36bf052acda139c8af188a078b | Relacionada ao histórico antigo de work | Candidata a exclusão após conferência |

“Contido em main” significa que a comparação não encontrou commits exclusivos à branch; não significa que o trabalho local de alguém esteja salvo ou que a branch deva ser apagada automaticamente.

## Branches que exigem preservação/conferência

| Branch | Resultado | Decisão |
|---|---|---|
| copilot/featurevbash-pipelines | Divergente: 1 commit exclusivo e 15 commits atrás de main; a comparação isolada encontrou apenas alteração em VBAshTests.bas (2 adições/2 remoções) | Não apagar até inspecionar o commit exclusivo e confirmar se a alteração foi superada pelo runner r14 |
| copilot/vbash-read-only-audit | A branch apareceu na listagem, mas a comparação retornou HTTP 404 | Estado não resolvido; não apagar nem usar como base sem nova consulta |

## Política para o próximo ciclo

1. Manter main como branch estável; não enviar trabalho experimental diretamente para ela.
2. Usar feature/phase-3-integration como base de integração enquanto esse for o branch de trabalho ativo. O nome é histórico e já não descreve todo o conteúdo.
3. Criar uma branch curta por tarefa, sempre a partir da ponta atual da branch de integração.
4. Integrar a tarefa de volta por PR para feature/phase-3-integration, após revisão e teste.
5. Só promover para main em um marco coerente, com PR que descreva o conjunto de mudanças e as evidências.
6. Branches Copilot devem ser temporárias e específicas de uma tarefa; não são branches de integração.
7. Antes de excluir uma branch: comparar com a base, verificar commits exclusivos, conferir PR associado e confirmar que não há trabalho local não publicado.
8. Não fazer force-push nem apagar branches como parte de uma simples limpeza documental.

## Comandos locais para reconciliação

Executar no clone local antes de iniciar o próximo trabalho:

```bash
git fetch --all --prune
git status --short --branch
git branch -vv
git log --oneline --decorate --graph --all -40
git rev-parse HEAD
git rev-parse feature/phase-3-integration
git rev-parse main
```

Guardar a saída junto ao registro da tarefa. Se o clone local não tiver a branch remota atualizada, não presumir que o HEAD local representa o remoto.
