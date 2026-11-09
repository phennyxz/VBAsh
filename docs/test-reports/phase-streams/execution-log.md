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

## Correção estrutural de módulos VBA — 2026-11-05

Uma nova auditoria, ampliada para todos os 35 módulos fonte VBA (`.bas` e `.cls`; o repositório não contém `.frm` ou outros formatos de módulo), identificou a causa do erro de compilação reportado: em `clsFileStream.cls`, a propriedade `ModuleRevision` antecedia as declarações privadas de módulo. O mesmo erro estrutural existia em `clsMemoryStream.cls` e `clsNullStream.cls`. Em VBA, declarações de módulo após o início de qualquer procedimento não são permitidas. A auditoria anterior descrita acima estava incorreta ao concluir que não havia esse defeito; sua busca não detectou declarações fora da seção de declarações. A ocorrência de `mIsClosed` dentro de `Class_Initialize` era válida no fonte; o erro concreto era a posição das declarações privadas após `ModuleRevision`.

As declarações privadas foram movidas para antes da primeira propriedade/procedimento nas três classes, sem alterações de nomes, tipos, APIs ou lógica funcional. Revisões atualizadas em cada módulo e propriedade correspondente: `clsFileStream-r2`, `clsMemoryStream-r2` e `clsNullStream-r2`. Os demais módulos não foram modificados. Não houve alteração de resultados históricos de testes.

Verificações estáticas: `git diff --check`; inventário estrutural de todos os 35 módulos, sem declarações de módulo após o começo da seção de procedimentos; campos privados `m*` declarados; `Option Explicit` e `Implements` em posições apropriadas; tags e propriedades `ModuleRevision` coincidentes nas três classes; dez membros de `IStream` implementados com assinaturas correspondentes nas quatro classes implementadoras; 17 suítes registradas, despachadas e esperadas. Não foi encontrada ferramenta de compilação VBA/Excel neste ambiente. Portanto, esta correção está verificada estaticamente, mas ainda requer **Debug > Compile VBAProject** e execução dos testes no Excel.

## Colisão de identificadores no relatório de testes — 2026-11-05

Corrigida a colisão case-insensitive em `VBAshTests.bas`: a variável local `checkoutReport` de `BuildTestReport` ocultava a função `CheckoutReport`, tornando a chamada ambígua para o compilador VBA. A variável foi renomeada para `checkoutText`; a função e sua assinatura, a chamada e o conteúdo composto do relatório foram preservados. Como a política de rastreabilidade exige incrementar a revisão do módulo a cada alteração, a tag e `RUNNER_REVISION` foram atualizadas em conjunto de `VBAshTests-r1` para `VBAshTests-r2`.

Auditoria complementar percorreu os 35 arquivos `.bas` e `.cls`; não há `.frm` no repositório. A busca case-insensitive verificou nomes de procedimentos, variáveis/parâmetros locais e chamadas no escopo, além de variáveis de módulo e nomes de procedimentos. Nenhuma outra colisão que crie ambiguidade de chamada ou impeça compilação foi identificada; nomes compartilhados em escopos independentes não foram alterados. `CheckoutReport(warnings)` permanece a chamada à função após a renomeação.

Verificações estáticas: `git diff --check`; ausência de `checkoutReport` em `BuildTestReport`; revisão de runner/tag `VBAshTests-r2`; chamada e declaração da função preservadas; auditoria case-insensitive sem outros casos acionáveis. Não foi executada compilação Excel/VBA nesta sessão; validar posteriormente com **Debug > Compile VBAProject** no Excel. Nenhum resultado histórico foi alterado.

## ERROR após a assertion #74 de FileStream — 2026-11-05

A origem do ERROR reportado é identificável no fluxo atual de `TestFileStream`: o código de erro era guardado em `errorNumber` para assertions negativas esperadas (incluindo a leitura/escrita após fechamento, que deixa `1201` nessa variável), mas não era zerado no caminho de sucesso antes de `GoTo Cleanup`. Depois de a assertion #74 passar, `CleanupFileStreamTest` terminava normalmente; em seguida, a condição `If errorNumber <> 0 Then` interpretava o código antigo da assertion esperada como uma falha não tratada e executava `Err.Raise errorNumber, errorSource, errorDescription`. Como `errorSource` e `errorDescription` só são preenchidos no handler de falha real, ambos estavam vazios. Portanto, o `1201` observado depois de 74 assertions não era evidência de falha no `CloseStream`, no helper de leitura externa ou na exclusão da pasta: era uma propagação final de um valor residual. Essa conclusão é baseada no caminho de dados estático; o Excel não foi executado nesta investigação.

Corrigido o caminho sem remover assertions ou alterar streams: `TestFileStream` agora mantém um marcador separado de falha não tratada (`testFailed`) e limpa os campos de erro residual no caminho normal antes da limpeza. Um erro de teste continua preservado; uma falha de limpeza torna-se erro quando não havia erro anterior, ou é acrescentada como erro secundário sem substituir o primeiro. A rotina registra a sequência: entrada na limpeza, segunda chamada de `CloseStream`, verificação da pasta, exclusão, retorno da limpeza e propagação final a `RunTest`. O erro propagado inclui número/descrição/source primários, erro secundário quando existir, última assertion/estágio e estado `IsClosed`, `CanRead`, `CanWrite`, `HasError`, `LastErrorNumber`. `RUNNER_REVISION` e tag avançaram para `VBAshTests-r3`; as revisões das classes de stream não mudaram.

Verificações estáticas: fluxo de dados do valor residual até `Err.Raise`, preservação do primeiro erro e registro explícito de cada estágio; `git diff --check`. Não houve execução Excel/VBA, então a instrumentação e a correção ainda precisam de confirmação real no Excel. Reexecutar `VBAsh_RunTests` com os módulos da mesma revisão e conferir que `FileStream` termina sem um ERROR artificial após a assertion #74; se surgir falha real, o relatório deve trazer estágio e estado do stream.

## Rastreabilidade de versão — runner `VBAshTests-r1`

Base: `feature/vbash-pipelines` / `2adbc0d796e44877e6c186d10cc94d7289630b73`. Adicionados metadados de revisão (`' @VBAshRevision <módulo>-r1` na linha 1 de `VBAshTests`, `IStream`, `clsFileStream`, `clsMemoryStream`, `clsNullStream`; propriedade `ModuleRevision` nas três classes de stream) e o novo formato de relatório (ver `docs/test-reports/README.md`). Não há alteração funcional nos streams; `TestFileStream` continua com 74 assertions e 17 suítes. Execuções anteriores não foram reatribuídas a versões. Verificações estáticas apenas (`git diff --check`, registro/contagem de suítes, contagem de assertions); **nenhuma execução Excel/VBA**, então o novo relatório está NOT RUN.

## Execução Excel informada — 2026-11-05, 16:31:20

Resultado real informado pelo usuário, executado no Excel com `VBAshTests-r3`. Esta execução não foi reproduzida nem verificada independentemente nesta sessão. O relatório do runner não conseguiu identificar o checkout Git associado ao workbook; portanto, não se atribui branch ou SHA a estes resultados e a correspondência entre checkout e módulos carregados é `UNVERIFIED`.

| Campo | Resultado informado |
|---|---|
| Data/hora | 2026-11-05 16:31:20 (fuso não informado) |
| Ambiente | Excel/VBA; versão do Excel e sistema operacional não informados |
| Suítes | 17 |
| PASS / FAIL / ERROR | 17 / 0 / 0 |
| Assertions | 656 |
| Branch / commit do checkout | `UNKNOWN` |
| Checkout/módulos carregados | `UNVERIFIED` |
| Runner | `VBAshTests-r3` |
| Módulos identificados como carregados | `VBAshTests-r3`, `IStream-r1`, `clsFileStream-r2`, `clsMemoryStream-r2`, `clsNullStream-r2` |

### Resultados individuais conhecidos

| Suíte | Estado informado | Assertions |
|---|---|---:|
| FileStream | PASS | 74 |
| NullStream | PASS | 42 |
| Redirection | PASS | 74 |
| Demais 14 suítes | PASS | Contagens individuais não fornecidas |

O usuário informou que todas as demais suítes passaram. Nenhuma saída original do executor foi anexada a este registro. O resultado é uma execução real relatada pelo usuário, não uma verificação independente dos arquivos deste checkout; os históricos de execuções anteriores permanecem inalterados.

## Implementação isolada de `clsFileDescriptorTable` — 2026-11-06

| Campo | Resultado |
|---|---|
| Branch / base | `feature/descriptor-table`, baseada em `main` no commit `d9dde542db555e68191cf65a518a9f00c7d1a458` |
| Execução Excel/VBA | Não executada; Excel/VBA não está disponível neste ambiente |
| Suítes / assertions executadas nesta sessão | 0 / 0 |
| Estado de `TestFileDescriptorTable` | Adicionada ao runner como a 18ª suíte; NOT RUN |
| Implementação | A tabela agora inicializa explicitamente FD 0/1/2 com `IStream` fornecidos, consulta/associa/duplica/fecha descritores e encontra o menor FD livre desde 3. Duplicatas preservam identidade de objeto; ownership é por identidade e permanece ativo enquanto houver descritor associado. Streams emprestados não são fechados pela tabela. |
| Erros e substituição | Números negativos, descritores ausentes, operações antes de inicializar e associação a `Nothing` são rejeitados. Associação inválida mantém o mapeamento anterior. Fechamento final de stream próprio ocorre antes de remover o FD; se falhar, o erro original é propagado com contexto e a associação fica disponível para retry. |
| Testes adicionados | Inicialização padrão, capacidades/identidade, consulta, busca de FD livre, aliasing e escrita compartilhada, fechamento de referência emprestada/própria, substituição, validação, stream previamente fechado e retry após falha de `CloseStream`. |
| Verificações estáticas | PASS: `git diff --check`; auditoria estrutural de 35 módulos VBA; tags e propriedades de revisão consistentes quando presentes; 18 suítes registradas, despachadas e esperadas; 62 assertions presentes no fonte da nova suíte. Nenhuma compilação VBA é alegada. |

A tabela não cria streams padrão implicitamente e ainda não está integrada a `clsSession`, executor, redirecionamentos ou pipes. Os resultados históricos de execução não foram alterados; em particular, o relatório informado de `VBAshTests-r3` não valida esta nova suíte.

## Correção de compilação em `TestFileDescriptorTable` — 2026-11-06

Erro de compilação reportado no Excel: `replacementStream.CloseStream` em `VBAshTests.bas`. `clsMemoryStream` implementa o membro como `Private Sub IStream_CloseStream()`, acessível apenas por referência `IStream`. O teste agora atribui `replacementStream` à variável `stream` (`IStream`) e chama `CloseStream` por ela; a visibilidade em `clsMemoryStream` não foi alterada. Inspeção estática das demais chamadas a membros de `IStream` em variáveis de classes concretas: nenhuma outra incompatível (`clsFileStream` expõe `ReadText`/`WriteText`/`CloseStream` públicos legados). Runner avançou para `VBAshTests-r5`. Verificação apenas estática; compilação e execução no Excel continuam pendentes.

## Execução Excel informada — 2026-11-06

O usuário informou execução real em Excel/VBA de `VBAshTests-r5`: 18/18 suítes PASS, 0 FAIL, 0 ERROR e 718 assertions, incluindo 62 PASS em `FileDescriptorTable`. Módulos identificados: `VBAshTests-r5`, `IStream-r1`, `clsFileStream-r2`, `clsMemoryStream-r2`, `clsNullStream-r2` e `clsFileDescriptorTable-r1`. O relatório não identificou o checkout Git: commit e branch UNKNOWN e correspondência entre checkout e módulos UNVERIFIED. O resultado não foi verificado independentemente nesta sessão.

## Limpeza explícita da tabela — `CloseAll`

`clsFileDescriptorTable-r2` acrescenta limpeza explícita solicitada ao chamador: streams próprios são fechados uma vez por identidade; FDs emprestados são removidos sem fechamento. Um erro de fechamento não impede os demais: FDs e ownership do stream falho permanecem, os erros originais são agregados com todos os FDs iniciais e o chamador pode repetir a operação. A tabela permanece inicializada e reutilizável, mas não permite outra `Initialize`; não há limpeza automática em `Class_Terminate`. Ownership é local à instância e não rastreia referências externas.

`VBAshTests-r6` adiciona a suíte `FileDescriptorTableCloseAll` e mantém as suítes anteriores. Ela cobre aliases, emprestados, erros múltiplos e originais, sucesso parcial, stream fechado antes de falhar, retry, limpeza repetida, reentrada e reutilização. Verificações estáticas: `git diff --check`; estrutura de 35 módulos; assinaturas dos 10 membros de `IStream` nos quatro implementadores; 19 suítes registradas e despachadas; 41 assertions no fonte da suíte nova e 62 na suíte anterior, preservada sem alteração. **NOT RUN:** não houve compilação ou execução Excel/VBA desta nova revisão nesta sessão. O relato de `r5` acima não valida `r6` ou `clsFileDescriptorTable-r2`.


## Integração da Fase 3 — fonte preparada para validação Excel/VBA

| Campo | Registro |
|---|---|
| Branch | `feature/phase-3-integration` |
| Runner na fonte | `VBAshTests-r7` |
| Suítes registradas | 23 |
| Execução Excel/VBA nesta revisão | NOT RUN |
| Compilação VBA nesta revisão | NOT RUN |
| HEAD da branch no momento desta revisão | `590ee89d89fc19d13a37c097c43774c1f64654a6` |
| Commit SHA testado | Não aplicável — ainda não houve compilação nem execução Excel/VBA |
| Resultado funcional | Pendente |

A fonte inclui escopos de descritores, gerenciador de redirecionamentos, adaptadores de captura para stdout/stderr, stream nulo para stdin, buffer de pipe e integração com a sessão. Também inclui testes de escopos, redirecionamento integrado, console e pipe. Foram feitas apenas inspeções estáticas nesta sessão; nenhuma execução VBA foi realizada e nenhum resultado PASS é declarado.

**Atenção ao importar:** use uma cópia limpa do workbook e todos os arquivos `.bas`/ `.cls` da mesma revisão da branch. O instalador enumera os módulos na pasta raiz e inclui os novos módulos automaticamente. Depois de importar, compile o projeto VBA antes de executar `VBAsh_RunTests`. O relatório deve identificar `VBAshTests-r7`, `clsFileDescriptorTable-r3`, `clsDescriptorScope-r1`, `clsRedirectionManager-r1`, `clsConsoleStream-r1` e `clsPipeStream-r1`. Registre qualquer divergência como ERROR/FAIL; não substitua o resultado por uma inspeção estática.
