# VBAsh — Referências e relatórios de testes

Esta estrutura mantém a rastreabilidade entre requisitos, testes implementados, execuções reais e a versão exata do código submetida a teste. Ela também torna explícitas as lacunas de cobertura e os limites das evidências disponíveis.

## Fases e nomenclatura

Use `phase-<slug>` para o diretório de cada fase, com o slug baseado no nome adotado pelo projeto. Neste repositório, `phase-pipelines` representa a fase chamada “Pipes” no plano de implementação (Fase 9), e `phase-streams` representa “Streams” (Fase 3). Não reutilize diretórios de fases diferentes.

| Fase | Referência | Execução | Estado |
|---|---|---|---|
| Pipelines | `phase-pipelines/test-reference.md` | `phase-pipelines/execution-log.md` | Aguardando registro formal |
| Streams | `phase-streams/test-reference.md` | `phase-streams/execution-log.md` | Pendente |

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
