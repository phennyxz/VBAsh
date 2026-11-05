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
