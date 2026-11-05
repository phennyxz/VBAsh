# VBAsh — Referências e relatórios de testes

Esta estrutura mantém a rastreabilidade entre requisitos, testes implementados, execuções reais e a versão exata do código submetida a teste. Ela também torna explícitas as lacunas de cobertura e os limites das evidências disponíveis.

## Fases e nomenclatura

Use `phase-<slug>` para o diretório de cada fase, com o slug baseado no nome adotado pelo projeto. Neste repositório, `phase-pipelines` representa a fase chamada “Pipes” no plano de implementação (Fase 9), e `phase-streams` representa “Streams” (Fase 3). Não reutilize diretórios de fases diferentes.

| Fase | Referência | Execução | Estado |
|---|---|---|---|
| Pipelines | `phase-pipelines/test-reference.md` | `phase-pipelines/execution-log.md` | Aguardando registro formal |
| Streams | `phase-streams/test-reference.md` | `phase-streams/execution-log.md` | Execução informada (17/17 PASS); fase ainda incompleta devido às lacunas abertas |

Uma fase só deve ser declarada aprovada quando houver evidência de execução e rastreabilidade suficientes. Um relato fornecido por uma pessoa deve ser identificado como tal, sem apresentá-lo como execução verificada independentemente.

## Referência de testes e log de execução

- **`test-reference.md`** descreve o que deve ser testado e onde cada teste está implementado. Registre o identificador, suíte, nome exato do procedimento, caminho do arquivo, requisito, categoria e resultado esperado. Use caminhos relativos à raiz do repositório. Se não houver teste identificado para um requisito, registre explicitamente a lacuna; não invente nome de procedimento ou arquivo.
- **`execution-log.md`** registra cada execução realizada. Acrescente uma nova entrada em vez de substituir entradas anteriores. Registre a data e hora disponíveis, SHA exato testado, branch, ambiente, contagens e resultado de cada suíte. Preserve a saída original do executor quando disponível e identifique claramente dados ausentes, informados por terceiros ou ainda não confirmados.

Uma referência de teste não prova que o teste foi executado. Um log de execução não substitui a referência que permite localizar seu teste.

## Como localizar os testes

Use o caminho registrado na referência para abrir o arquivo e localize o procedimento pelo nome exato. Em módulos VBA, registre o nome do módulo/procedimento e o arquivo exportado correspondente; em outros formatos, registre o arquivo e o símbolo executável. Se o código ou o procedimento não puder ser localizado no snapshot consultado, indique isso e não o apresente como cobertura confirmada.

## Estados de execução

- **PASS** — executado e aprovado.
- **FAIL** — executado, mas uma ou mais asserções falharam.
- **ERROR** — ocorreu erro inesperado durante a execução.
- **NOT RUN** — não executado.
- **BLOCKED** — a execução não foi possível por limitação do ambiente.

Use o estado para descrever o que de fato ocorreu. Validação estática, teste planejado, contagem relatada sem detalhes ou ausência de resultados não equivalem a aprovação funcional.

## Adicionar uma fase sem perder histórico

1. Crie `phase-<slug>/test-reference.md` e `phase-<slug>/execution-log.md`, seguindo a nomenclatura do plano/projeto.
2. Registre todos os requisitos funcionais conhecidos na referência, vinculando cada um a um teste localizado ou marcando a lacuna.
3. Adicione uma linha à tabela deste índice.
4. No log, acrescente uma entrada por execução; não reescreva nem remova relatórios anteriores, mesmo quando houver falhas.
5. Após qualquer alteração de testes ou comportamento, atualize a referência junto com a implementação e registre o SHA exato usado na execução.

## Escopo e evidência atual

As referências desta estrutura descrevem o que foi possível identificar no snapshot consultado. Os relatórios recebidos sem arquivos de teste, nomes de suítes ou SHA são preservados como relatos informados, não completados por inferência.

## Rastreabilidade de versão nos relatórios

A partir da revisão `VBAshTests-r1`, o relatório TXT gerado por `VBAsh_RunTests` (`VBAsh TEST REPORT`) separa três coisas:

1. **Checkout (fonte em disco).** Branch e SHA completo são lidos, somente leitura, de `.git/HEAD` (e `packed-refs`) procurando uma pasta `.git` a partir da pasta do workbook. Se não existir, o campo é `UNKNOWN`. O estado do checkout (limpo/modificado) não é detectável por VBA e é sempre `UNKNOWN`; registre-o manualmente (`git status`). Isso **não** prova quais módulos o Excel carregou.
2. **Módulos carregados.** Cada módulo relevante (`VBAshTests`, `IStream`, `clsFileStream`, `clsMemoryStream`, `clsNullStream`) tem a tag `' @VBAshRevision <id>` na primeira linha. O runner lê essa tag do código efetivamente presente no projeto VBA via `VBProject.VBComponents(...).CodeModule` (exige "Confiar no acesso ao modelo de objeto de projeto do VBA"). As três classes de stream também expõem `ModuleRevision`, lido de uma instância carregada; tag e propriedade devem coincidir (`MISMATCH` caso contrário). Os identificadores são revisões explícitas (`<módulo>-r<N>`), não SHAs; incremente `N` a cada alteração do módulo.
3. **Integridade da execução.** A correspondência checkout/módulos é sempre `UNVERIFIED`: o runner nunca presume que o Git é igual ao projeto VBA.

### Significado dos valores

- `UNKNOWN`: a informação não está disponível (sem `.git`, estado do checkout).
- `UNVERIFIED`: a informação não pôde ser confirmada (sem acesso ao VBProject, módulo sem tag, correspondência checkout/módulos).
- `MISMATCH`: tag no código e propriedade da instância divergem — o projeto VBA tem módulos misturados; reimporte todos os arquivos do mesmo commit.

### Interpretação de divergências

Compare o SHA do checkout com o commit em que cada revisão `<módulo>-rN` foi introduzida (`git log -S"<módulo>-rN"`). Se as revisões carregadas diferirem das do commit, o Excel não executou aquele commit e o relatório deve ser descartado ou reexecutado após reimportar. Para falhas, o relatório traz suíte, nº da assertion (relativo à suíte e válido apenas para a revisão do runner informada), descrição, esperado/obtido, erro, source e etapa; em `ERROR` informa-se a última assertion aprovada, pois a etapa exata não é registrada.

### Execuções antigas

Relatórios anteriores (sem esses campos) permanecem como estão e não recebem versões retroativas; registre-os como "versão não identificada". Novos registros em `execution-log.md` devem copiar os campos Execution/Checkout/Loaded module versions/Execution integrity do relatório TXT.
