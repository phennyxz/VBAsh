# Streams — Log de execução

Registro anterior a esta execução: nenhum teste da fase Streams havia sido executado neste marco. `StreamContract` / `TestStreamContract` e `MemoryStream` / `TestMemoryStream` estavam escritos em `VBAshTests.bas`, mas permaneciam NOT RUN porque o ambiente Copilot não disponibiliza Excel/VBA. O resultado abaixo é um relato de execução realizada fora deste ambiente e não foi verificado independentemente nesta sessão.

Acrescente uma nova seção a cada execução real. Preserve entradas anteriores, inclusive as que terminarem em FAIL, ERROR ou BLOCKED. Não preencha campos com estimativas.

## Modelo de registro

| Campo | Registro |
|---|---|
| Data e hora | Pendente |
| Versão do VBAsh | Pendente |
| Commit SHA testado | Pendente — registrar o SHA exato |
| Branch testada | Pendente |
| Sistema operacional | Pendente |
| Versão do Excel | Pendente |
| Ambiente VBA utilizado | Pendente |
| Total de suítes executadas | Pendente |
| Total de testes aprovados | Pendente |
| Total de falhas | Pendente |
| Total de erros inesperados | Pendente |
| Total de asserções | Pendente |
| Duração total | Pendente |
| Saída original do executor | Anexar/copiar sem alterações, se disponível |

### Resultado individual das suítes

| Suíte | Procedimento | Estado (PASS/FAIL/ERROR/NOT RUN/BLOCKED) | Asserções | Duração | Observações |
|---|---|---|---:|---|---|
| Pendente | Pendente | Pendente | Pendente | Pendente | Preencher após execução real |

## Execução real informada — Marco 2.1

Os dados abaixo foram informados como resultado de uma execução real no Excel/VBA. A saída original do executor não foi fornecida nesta sessão, portanto o relato não foi verificado independentemente.

| Campo | Registro |
|---|---|
| Data | 03/11/2026 |
| Início | 15:06:35 — fuso horário não informado |
| Término | 15:06:35 — fuso horário não informado |
| Duração | Não informada |
| Versão do VBAsh | Não informada |
| Commit SHA testado | Não informado; não inferido a partir do estado atual do repositório |
| Branch testada | Não informada |
| Sistema operacional / versão do Excel / arquitetura | Não informados |
| Ambiente VBA | Excel/VBA, conforme relato |
| Suítes executadas | 15/15 |
| PASS / FAIL / ERROR | 15 / 0 / 0 |
| Total de asserções | 540 |
| Saída original do executor | Não fornecida |

### Resultados relevantes informados

| Suíte | Procedimento | Estado informado | Asserções | Observações |
|---|---|---|---:|---|
| StreamContract | TestStreamContract | PASS | 41 | Resultado informado; não verificado independentemente nesta sessão |
| MemoryStream | TestMemoryStream | PASS | 112 | Resultado informado; não verificado independentemente nesta sessão |

## Alteração de `clsFileStream` — 2026-11-05

| Campo | Registro |
|---|---|
| Data | 2026-11-05 |
| Branch | `feature/vbash-pipelines` |
| Commit SHA testado | Não aplicável — alteração ainda não commitada durante esta verificação |
| Ambiente | Linux/Copilot; Excel/VBA indisponível |
| Suítes VBA executadas | 0 |
| Estado de `FileStream` / `TestFileStream` | NOT RUN |
| Verificações estáticas | PASS: `git diff --check`; 17 grupos registrados e mapeados; leitura em bloco sem concatenação caractere a caractere; escopo de seis arquivos e limites de linha verificados. |
| Resultado | Implementação e casos de teste adicionados ao código-fonte. A execução, compilação e confirmação de comportamento de `Scripting.FileSystemObject.TextStream` permanecem pendentes no Excel/VBA. |
| Saída original do executor | Não disponível; nenhuma execução VBA ocorreu |

Este registro não declara aprovação de testes no Excel. As verificações estáticas não substituem compilação nem execução do runtime VBA.

## Relatório Excel informado — 2026-11-05

Os números a seguir foram fornecidos pelo usuário a partir de uma execução Excel/VBA; o workbook, os módulos VBA efetivamente carregados e a saída original não estão disponíveis nesta sessão para confirmação independente.

| Campo | Resultado informado |
|---|---|
| Suítes | 17 |
| PASS / FAIL / ERROR | 15 / 1 / 1 |
| Asserções | 620 |
| `NullStream` | FAIL: expectativa da escrita multiline era 11; o valor correto é 10 (`5 + 2 + 3`). |
| `FileStream` | ERROR `-2147220303` (`VBAsh_STREAM_ERROR_CLOSED`), atribuído à verificação externa do arquivo após reabrir, escrever e fechar a instância. |

### Investigação e alterações neste snapshot

- Corrigida somente a expectativa numérica do teste `NullStream`; `clsNullStream` permaneceu inalterada.
- Na fonte atual, o erro 1201 é lançado por `clsFileStream` somente pelos guards de leitura/escrita quando `mIsClosed` é verdadeiro. A função `ReadFileStreamTestFile` usa um `TextStream` independente do `FileSystemObject` e não chama métodos da instância `clsFileStream`; portanto, a origem reportada não é reproduzível pelo fluxo estático da fonte atual. Não há evidência suficiente para atribuir a causa ao stream, ao helper ou a uma divergência dos módulos carregados no Excel.
- Ampliada `TestFileStream` para validar escrita→reabertura para leitura e leitura→reabertura para escrita, capacidades, arquivo/tamanho após close e persistência. A verificação externa agora informa etapa, número/descrição/origem do erro, existência/tamanho do arquivo e estado da instância se falhar.
- Essas alterações adicionam oito assertions à fonte de teste em relação ao caso reportado. Nenhuma execução ou compilação Excel/VBA foi feita nesta sessão; o status desses casos permanece NOT RUN, e o diagnóstico final do 1201 depende de uma nova execução da suíte atualizada.
- Verificações estáticas desta revisão: `git diff --check` passou; a expectativa `NullStream` é 10; as 17 suítes do runner estão registradas e mapeadas; foram confirmadas oito assertions adicionais em `TestFileStream`; `clsFileStream.cls` e os demais módulos fora de teste/documentação não foram alterados.

## Auditoria de continuidade — 2026-11-05

| Campo | Resultado |
|---|---|
| Branch / HEAD de origem | `feature/vbash-pipelines` / `25bb9cd3b5e4e985d6ab7ef3d986d6ce2c246b8d` |
| Worktree inicial | Limpo; branch alinhada com `origin/feature/vbash-pipelines` |
| Presença do commit informado | Confirmada como HEAD e ancestral da branch |
| Execução Excel/VBA nesta sessão | Não executada |

O erro 1201 reportado não é reproduzível apenas pela fonte disponível: `ReadFileStreamTestFile` abre um `TextStream` independente e só consulta propriedades de leitura da interface em seu caminho de diagnóstico; os únicos `Err.Raise VBAsh_STREAM_ERROR_CLOSED` em `clsFileStream` são os guards das operações públicas de leitura/escrita, que o helper não invoca. A sequência atual fecha o objeto COM anterior em `CloseStream`, define a referência como `Nothing`, cria e atribui o objeto novo em `OpenForWrite`, escreve e fecha antes da leitura externa. Não foi encontrado defeito comprovado em `clsFileStream`, e a classe/contrato não foram alterados.

Há uma divergência verificável no número de assertion: a suíte atual contém 74 assertions; a #66 atual é “read limit equal to file size returns all text” e a verificação final de `"next"` é a #74. O relato anterior chamou a verificação final de #66, portanto foi produzido com uma revisão anterior de `VBAshTests.bas` (antes das oito assertions adicionadas pelo commit `25bb9cd`). Isso comprova que o relatório citado não usa a numeração do teste presente no HEAD atual, mas não identifica quais módulos estavam carregados no Excel nem a origem do 1201. A expectativa NullStream de 10 e as oito assertions de regressão/diagnóstico estão presentes no HEAD.

Nenhuma correção de código foi necessária nesta auditoria. As verificações estáticas confirmam que `clsFileStream` implementa todos os membros de `IStream`, mantém as cinco assinaturas públicas legadas, as 17 suítes do runner têm dispatch correspondente e a leitura em blocos continua usando `TextStream.Read(maxChars)`. O próximo teste deve ser executado com `VBAshTests.bas` e `clsFileStream.cls` importados do mesmo checkout/commit; se falhar, preservar a linha completa `ERROR | FileStream` com número, descrição e source, além dos dados de existência/tamanho do arquivo e estado do stream incluídos no diagnóstico.

## Auditoria de compilação VBA — 2026-11-05

| Campo | Resultado |
|---|---|
| Branch / HEAD de origem | `feature/vbash-pipelines` / `db662eaffd9e5ff77c4ab24fe3bfa395a5e242a7` |
| Worktree inicial | Limpo; branch alinhada com `origin/feature/vbash-pipelines` |
| Presença do commit informado | Confirmada como HEAD e ancestral da branch |
| Compilador VBA / Excel disponível | Não |

O trecho reportado (`Private Sub Class_Initialize(): mIsClosed = True`) ocorre em `clsFileStream.cls:17-19`. A variável está declarada na própria classe em `clsFileStream.cls:13`; ela também está declarada nas outras classes que a utilizam (`clsMemoryStream.cls:14`, `clsNullStream.cls:12`). Portanto, no snapshot `db662ea`, não há identificador `mIsClosed` sem declaração nem renomeação/remoção acidental demonstrável. A falha indica que o componente efetivamente compilado no Excel não corresponde ao texto atual de `clsFileStream.cls`, ou que a mensagem foi associada a outro componente/revisão; isso precisa ser confirmado no VBE inspecionando o nome e o conteúdo do módulo que apresenta o erro. Não se pode atribuir uma causa mais específica sem esse projeto VBA.

Auditoria estática: os cinco módulos de Streams sob revisão e `VBAshTests.bas` têm `Option Explicit`; as tags `@VBAshRevision` ficam em comentário antes de `Option Explicit` e não há declarações `Attribute` nesses arquivos. `IStream` contém dez membros e `clsFileStream`, `clsMemoryStream`, `clsNullStream` e `clsStreamContractProbe` implementam nomes, tipos de retorno e parâmetros correspondentes. `clsFileDescriptorTable` usa os membros `CreateObject("Scripting.Dictionary")`, `Exists` e o item padrão do Dictionary por late binding, não adicionando uma referência de compilação VBA. O runner registra e despacha 17 suítes, conforme `EXPECTED_TEST_COUNT`. Não foram encontrados outros problemas de declaração ou discrepâncias de interface nos módulos inspecionados.

Nenhum fonte foi alterado nesta auditoria: não havia erro comprovado para corrigir, então não houve incremento de revisões. `git diff --check` e as verificações estáticas de declarações de `mIsClosed`, opções, tags, membros de `IStream` e registro/contagem de suítes passaram. Isso não equivale a compilar o projeto. A compilação real, incluindo referências do projeto e os módulos realmente importados, permanece pendente no Excel/VBE. Para a próxima tentativa, reimporte os módulos do commit `db662ea`, execute **Debug > Compile VBAProject** e, se persistir, anote o nome exato do componente, a linha destacada e confirme se o campo privado `mIsClosed As Boolean` consta no módulo compilado.

## Rastreabilidade de versão — runner `VBAshTests-r1`

Base: `feature/vbash-pipelines` / `2adbc0d796e44877e6c186d10cc94d7289630b73`. Adicionados metadados de revisão (`' @VBAshRevision <módulo>-r1` na linha 1 de `VBAshTests`, `IStream`, `clsFileStream`, `clsMemoryStream`, `clsNullStream`; propriedade `ModuleRevision` nas três classes de stream) e o novo formato de relatório (ver `docs/test-reports/README.md`). Não há alteração funcional nos streams; `TestFileStream` continua com 74 assertions e 17 suítes. Execuções anteriores não foram reatribuídas a versões. Verificações estáticas apenas (`git diff --check`, registro/contagem de suítes, contagem de assertions); **nenhuma execução Excel/VBA**, então o novo relatório está NOT RUN.
