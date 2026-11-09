# Streams — Referência de testes

**Fase do projeto:** Streams (Fase 3 em `IMPLEMENTATION_PLAN.md`).

Os requisitos abaixo são derivados dos critérios de conclusão da fase no plano. `STR-CONTRACT-01` verifica a interface VBA e seu dispatch por `Implements` usando um probe de teste. `STR-MEMORY-01` cobre `clsMemoryStream`. `STR-FILE-01` cobre a implementação `clsFileStream`. Há execução real informada pelo usuário para NullStream, FileStream e Redirection em 2026-11-05; os resultados não foram verificados independentemente contra o workbook ou o checkout Git. Consulte `execution-log.md`.

| ID | Suíte | Procedimento | Arquivo | Requisito | Categoria | Resultado esperado |
|---|---|---|---|---|---|---|
| STR-CONTRACT-01 | StreamContract | TestStreamContract | `VBAshTests.bas` | Contrato textual `IStream`: leitura, escrita, fechamento idempotente, capacidades, EOF significativo somente em stream aberto e legível, propriedades independentes `CanRead`/`IsClosed`/`HasError` e códigos distintos para falha não suportada, fechada, leitura, escrita e abertura | Execução informada — PASS (41 asserções; não verificada independentemente) | Probe implementa o contrato por `Implements`; `AtEndOfStream` é False se não legível ou fechado; `CanRead`, `IsClosed` e `HasError` representam apenas capacidade, estado de fechamento e ocorrência de falha; operações não suportadas/em stream fechado preenchem erro sem alterar as outras propriedades; fechamentos repetidos não falham. |
| STR-MEMORY-01 | MemoryStream | TestMemoryStream | `VBAshTests.bas` | Implementação `IStream` em memória: capacidades configuráveis, conteúdo inicial, leitura integral e parcial, EOF/vazio, posição, escrita, erros de operações não suportadas/após fechamento, fechamento idempotente e reinicialização | Execução informada — PASS (112 asserções; não verificada independentemente) | Escritas acrescentam ao buffer sem mover a posição de leitura; inicialização substitui conteúdo e restaura capacidades, posição 1, estado aberto e estado de erro limpo; limites de leitura não positivos retornam texto vazio sem alterar a posição. |
| STR-FD-TABLE-01 | FileDescriptorTable | TestFileDescriptorTable | `VBAshTests.bas` | Inicialização explícita dos FDs 0/1/2 com `IStream`; associação/consulta, menor FD livre desde 3, duplicação por identidade, ownership por objeto, fechamento e substituição seguros, validação de erros e retry após falha ao fechar | Execução `VBAshTests-r5` informada — PASS (62 assertions; checkout UNKNOWN/UNVERIFIED) | Descritores duplicados compartilham o mesmo objeto. Streams emprestados nunca são fechados pela tabela; ownership adquirido por identidade persiste até a última referência administrada. Falha ao fechar a última referência mantém FD e ownership para retry. |
| STR-FD-CLOSE-ALL-01 | FileDescriptorTableCloseAll | TestFileDescriptorTableCloseAll | `VBAshTests.bas` | Limpeza explícita de aliases próprios e emprestados, erros múltiplos, falha parcial, retry e reutilização | Teste adicionado em `VBAshTests-r6` — NOT RUN | Fecha uma vez por identidade própria, nunca fecha emprestados, remove FDs bem-sucedidos, agrega erros e preserva os aliases falhos para retry. |
| STR-NULL-01 | NullStream | TestNullStream | `VBAshTests.bas` | Implementação `IStream` equivalente a `/dev/null`: leitura vazia, escrita descartada com retorno do número de caracteres, EOF True enquanto aberto e legível, fechamento idempotente, erros após fechamento/não suportados e limpeza de erro após sucesso | Execução informada pelo usuário — PASS (42 assertions; 2026-11-05 16:31:20; não verificada independentemente) | Nada é armazenado; `AtEndOfStream` é False após fechamento; leitura/escrita após fechamento levantam `VBAsh_STREAM_ERROR_CLOSED`; operações bem-sucedidas e `CloseStream` limpam o erro. |
| STR-FILE-01 | FileStream | TestFileStream | `VBAshTests.bas` | Implementação de `IStream` sobre arquivo texto: leitura em blocos e sequencial, EOF, arquivo vazio, read-only, overwrite, append, retorno de caracteres, erros de capacidade/fechamento/abertura, fechamento idempotente e reabertura da mesma instância nos dois sentidos (write→read e read→write); validação externa somente após fechamento e temporários removidos também após falha | Execução informada pelo usuário — PASS (74 assertions; 2026-11-05 16:31:20; não verificada independentemente) | `TextStream.Read(maxChars)` não faz leitura antecipada; leitura não ultrapassa o conteúdo, EOF é falso quando fechado/não legível, capacidades persistem após close e erros seguem os códigos de `IStream`. Falhas da verificação externa incluem etapa, erro COM e estado do stream para distinguir implementação de teste/ambiente. A integração com descritores/redirecionamento não é coberta por esta suíte. |
| STR-GAP-01 | FileDescriptorTable | TestFileDescriptorTable | `VBAshTests.bas` | Tabela de descritores inicializada explicitamente com 0, 1 e 2 associados a streams fornecidos pelo host | Coberto por `STR-FD-TABLE-01`; PASS informado, checkout UNKNOWN/UNVERIFIED | `Initialize` associa stdin/stdout/stderr nos descritores 0/1/2; não cria streams implicitamente. |
| STR-GAP-02 | Não localizado | Não localizado | Não localizado neste snapshot | Ler stdin do console (plano, Fase 3) | Lacuna — teste não localizado | A entrada lida corresponde ao conteúdo fornecido pelo console. |
| STR-GAP-03 | Não localizado | Não localizado | Não localizado neste snapshot | Escrever stdout no console (plano, Fase 3) | Lacuna — teste não localizado | O texto escrito em stdout chega ao console. |
| STR-GAP-04 | Não localizado | Não localizado | Não localizado neste snapshot | Separar stderr de stdout e permitir redirecionar stderr para arquivo ou buffer (plano, Fase 3) | Lacuna — teste não localizado | stderr pode ser capturado/redirecionado sem se misturar indevidamente com stdout. |
| STR-GAP-05 | Não localizado | Não localizado | Não localizado neste snapshot | Redirecionar stdout com `> file` (teste mínimo planejado, Fase 3) | Lacuna — teste não localizado | O arquivo de saída é criado ou sobrescrito com o conteúdo esperado. |
| STR-GAP-06 | Não localizado | Não localizado | Não localizado neste snapshot | Redirecionar stdin com `< file` (teste mínimo planejado, Fase 3) | Lacuna — teste não localizado | O conteúdo do arquivo é lido como entrada. |
| STR-GAP-07 | Parcial — FileDescriptorTable | TestFileDescriptorTable | `VBAshTests.bas` | Duplicação de descritor por identidade coberta por `STR-FD-TABLE-01`; parsing e aplicação de `2>&1` no shell continuam pendentes e fora desta etapa | Parcial — duplicação da tabela implementada; integração `2>&1` não implementada | Dois FDs apontam ao mesmo `IStream`; o encaminhamento de redirecionamentos ainda não usa a tabela. |
| STR-GAP-08 | NullStream | TestNullStream | `VBAshTests.bas` | Modelar `/dev/null` como stream nulo (plano, Fase 3) — coberto por `STR-NULL-01` (`clsNullStream`); o redirecionamento `> /dev/null` não usa esta classe | Coberto por `STR-NULL-01` — execução informada PASS (42 assertions; não verificada independentemente) | Escritas são aceitas sem erro e o conteúdo é descartado. |

As linhas `STR-GAP-*` indicam cobertura ausente ou parcial conforme descrito em cada requisito. `STR-CONTRACT-01` cobre a interface por meio do probe; `STR-MEMORY-01`, `STR-NULL-01` e `STR-FILE-01` cobrem implementações de streams; `STR-FD-TABLE-01` cobre a tabela isolada. A suíte `CloseAll` requer execução no Excel/VBA. Console, integração com redirecionamentos e pipes permanecem fora desta etapa.

`CloseAll` exige inicialização, considera os FDs existentes na entrada e agrupa por identidade. A tabela não permite mutações ou uma chamada aninhada de `CloseAll` durante a limpeza. Ao final de falhas de fechamento, lança `vbObjectError + 1255`, com source `clsFileDescriptorTable.CloseAll` e descrição contendo os FDs e o número, source e descrição originais de cada falha. O responsável deve chamar `CloseAll` e tratar o erro agregado antes de liberar a tabela; não há limpeza automática no finalizador. Ownership é local à instância, sem contagem de referências externas. Após sucesso ou falha parcial, a instância permanece inicializada e aceita novas associações; `Initialize` continua sendo de uso único.

## Referências

- `IMPLEMENTATION_PLAN.md`, Fase 3 — Streams, critérios de conclusão e testes mínimos.
- `docs/test-reports/phase-streams/execution-log.md` — status atual das evidências de execução.

## Identificação de versão

Todo relatório gerado pelo runner identifica o checkout (quando disponível), as revisões dos módulos carregados e inclui detalhes de falha com número da assertion por suíte. Numeração de assertions só é comparável dentro da mesma revisão de `VBAshTests`; a execução informada da tabela usou `VBAshTests-r5`, enquanto a suíte nova está no runner `VBAshTests-r6`. Isso não prova qual checkout originou módulos carregados em execuções anteriores. Detalhes em `docs/test-reports/README.md`.


## Integração por descritores — revisão em desenvolvimento

O runner `VBAshTests-r7` registra 23 suítes. As novas suítes abaixo foram adicionadas à fonte, mas ainda não foram compiladas nem executadas no Excel/VBA. Não confundir sua presença no runner com resultado PASS.

| ID | Suíte | Procedimento | Cobertura adicionada | Estado |
|---|---|---|---|---|
| STR-SCOPE-01 | DescriptorScopes | TestDescriptorScopes | Snapshot de bindings, pin por identidade, restauração LIFO, remoção de FDs criados no escopo e bloqueio de CloseAll enquanto há escopos ativos | NOT RUN |
| STR-REDIR-01 | DescriptorRedirection | TestDescriptorRedirection | Integração real de sessão/tabela para `<`, `>`, `2>`, `2>&1`, `1>&2`, `n>&m` e `n>&-`; valida arquivo, FD 0, canais visíveis, ordem e restauração | NOT RUN |
| STR-CONSOLE-01 | ConsoleStream | TestConsoleStream | Adaptador de captura textual usado nos FDs padrão do host, reset e fechamento | NOT RUN |
| STR-PIPE-01 | PipeStream | TestPipeStream | Buffer textual, consumo parcial, fechamento do writer, drenagem e EOF | NOT RUN |

A suíte de integração usa streams padrão explícitos e chama `ConfigureDescriptorTable`; as suítes antigas de redirecionamento continuam exercitando o caminho de compatibilidade da sessão sem tabela configurada. A correspondência entre os módulos instalados e a revisão testada deve ser confirmada pelo relatório r7.

Limitações que permanecem fora da cobertura funcional desta fase:
- stdin do host é um stream nulo em EOF; não há leitura interativa do formulário.
- `< arquivo` associa o stream de leitura ao FD 0, mas os builtins atuais não leem esse FD.
- `clsPipeStream` é um buffer textual, sem scheduler, bloqueio, endpoint/processo independente ou execução de pipeline.
- O retorno dos builtins como strings impede saída incremental e reconstrução da ordem temporal de escrita entre stdout e stderr.
