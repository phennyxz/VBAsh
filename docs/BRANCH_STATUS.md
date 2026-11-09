# VBAsh — Estratégia de branches

**Data da decisão:** 2026-11-09  
**Repositório:** phennyxz/VBAsh  
**Linha evolutiva canônica atual:** feature/phase-3-integration

## Decisão de projeto

O VBAsh tem uma única linha evolutiva de desenvolvimento. A branch feature/phase-3-integration contém o produto em construção e é a referência efetiva para implementação, testes e documentação.

A branch main não representa uma versão funcional alternativa nem uma linha estável deliberadamente mantida. É uma referência histórica anterior, atualmente atrasada em relação ao desenvolvimento real. Não usar main como base, destino de PR ou critério para decidir qual implementação é mais atual.

Não manter branches permanentes por fase, subsistema ou ferramenta. Branches antigas integradas ou cujo trabalho foi superado não têm função operacional. O histórico dos commits permanece no repositório mesmo após excluir referências de branches.

## Inventário observado em 2026-11-09

| Branch | Situação observada | Decisão de manutenção |
|---|---|---|
| feature/phase-3-integration | Linha evolutiva mais recente; contém o estado de referência do projeto e o relatório r14 | **Preservar como linha canônica até eventual renomeação para main** |
| main | Referência histórica atrasada; não é produto funcional independente | Não usar para desenvolvimento. Recomenda-se renomear a linha canônica para main e remover a referência antiga depois de ajustar a branch padrão |
| work | Histórico antigo integrado pelo PR #2; sem papel na linha atual | Excluir |
| feature/vbash-pipelines | Histórico integrado pelo PR #11; sem papel na linha atual | Excluir |
| feature/descriptor-table | Histórico integrado pelo PR #12; sem papel na linha atual | Excluir |
| copilot/featurevbash-pipelines | Divergência exclusiva em VBAshTests.bas; alteração não é necessária para a linha canônica conforme decisão atual do projeto | Excluir |
| copilot/featurevbash-pipelines-again | Trabalho antigo/superado; sem papel na linha atual | Excluir |
| copilot/featurevbash-nullstream | Histórico integrado pelo PR #10; sem papel na linha atual | Excluir |
| copilot/featurefile-stream | Trabalho antigo/superado; sem papel na linha atual | Excluir |
| copilot/copilotsetup-audit-vbash-repo | Trabalho antigo de configuração/auditoria; sem papel na linha atual | Excluir |
| copilot/vbash-read-only-audit | Consulta retornou 404; referência não resolvida | Excluir se ainda existir e após confirmar o nome exato |

A decisão de excluir branches não apaga commits do histórico Git. Antes da exclusão remota, sincronizar o clone local e confirmar que não há trabalho local não publicado que se queira preservar. A decisão do projeto é que essas linhas não precisam ser mantidas.

## Modelo operacional desejado

1. Uma única branch de desenvolvimento: a linha evolutiva canônica.
2. Commits pequenos, descritivos e frequentes diretamente nessa linha.
3. Não abrir PR para integrar trabalho rotineiro dentro do mesmo repositório.
4. Branch temporária só quando uma experiência realmente precisar ser isolada; ao terminar, integrar ou descartar e apagar a branch.
5. Copilot deve trabalhar na branch canônica, quando possível, e não deixar branches automáticas como linhas permanentes.
6. Não usar main antiga como referência de implementação. A branch canônica atual é a fonte de verdade.
7. Manter tags/releases para marcos versionados, em vez de conservar branches antigas.
8. Testes devem registrar SHA exato do código e revisões dos módulos efetivamente carregados no Excel.

## Transição técnica recomendada

A normalização ideal é renomear feature/phase-3-integration para main, tornando a linha evolutiva atual a branch principal convencional. Isso exige atualizar a branch padrão do GitHub e verificar referências locais, links, automações e configurações antes da renomeação.

A integração GitHub disponível nesta sessão permite ler branches e atualizar arquivos, mas não expõe operações administrativas de renomear/excluir branches nem alterar a branch padrão. Portanto, esta documentação registra a decisão; não afirma que as referências remotas já foram removidas ou renomeadas.

## Reconciliação local

Antes de limpar as referências remotas:

```bash
git fetch --all --prune
git status --short --branch
git branch -vv
git log --oneline --decorate --graph --all -50
git rev-parse HEAD
git rev-parse origin/feature/phase-3-integration
git rev-parse origin/main
```

Depois de confirmar que o trabalho local está preservado, a limpeza de referências locais pode ser feita com:

```bash
git branch -D work feature/vbash-pipelines feature/descriptor-table
git branch -D copilot/featurevbash-pipelines copilot/featurevbash-pipelines-again
git branch -D copilot/featurevbash-nullstream copilot/featurefile-stream
git branch -D copilot/copilotsetup-audit-vbash-repo
```

Não executar esses comandos cegamente: primeiro verificar se alguma branch local contém commits não publicados que o usuário deseje guardar. A remoção remota e a troca da branch padrão devem ser feitas nas configurações/controles administrativos do GitHub, após a linha canônica ser estabelecida como padrão.
