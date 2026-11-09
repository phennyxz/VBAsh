# IMPLEMENTATION_PLAN.md

Este documento define um plano de implementação incremental para o projeto VBAsh, com foco em uma base funcional e testável antes de abordar recursos avançados de compatibilidade com Bash 5.x.

A estratégia é priorizar a criação de um shell mínimo e confiável, depois evoluir em módulos: terminal, sessão, streams, registry de comandos, lexer, parser, builtins, gestão de processos, pipes e emulação POSIX.

## Visão geral da estratégia

O projeto será implementado em marcos pequenos e verificáveis. Cada fase produz algo funcional e testável, com acoplamento mínimo entre módulos. A ordem foi escolhida para respeitar a dependência real entre as camadas:

1. Terminal
2. Session
3. Streams
4. Command Registry
5. Lexer
6. Parser
7. Builtins
8. Process Manager
9. Pipes
10. POSIX Layer

As fases posteriores dependem de uma base sólida em entrada de usuário, estado da sessão, fluxos de I/O e execução de comandos simples.

---

## Fase 1 — Terminal

### Objetivo
Criar a interface mínima do shell dentro do Excel/VBA, permitindo que o usuário insira um comando e receba resposta em um console básico.

### Módulos envolvidos
- CSession
- CStream
- CConsoleAdapter
- CInputBuffer
- COutputBuffer
- VBAshHost

### Dependências
- Ambiente do Excel/VBA com planilha ou UserForm para entrada/saída
- Capacidade de capturar teclado e exibir texto
- Estrutura de sessão com estado inicial

### Critérios de conclusão
- O usuário consegue digitar texto em um campo de entrada.
- O shell exibe prompt inicial.
- O texto digitado pode ser enviado para o motor do shell.
- A saída do shell pode aparecer em console, planilha ou buffer de texto.
- Existe uma rotina de loop principal de leitura/execução/impressão.

### Testes mínimos
- Teste 1: prompt aparece sem erro.
- Teste 2: texto digitado é capturado corretamente.
- Teste 3: console imprime texto de retorno.
- Teste 4: entrada vazia não quebra o loop.
- Teste 5: comando simples sem execução ainda é aceito e encaminhado para o próximo módulo.

---

## Fase 2 — Session

### Objetivo
Definir o estado da sessão do shell: ambiente, diretório atual, variáveis, histórico e flags de execução.

### Módulos envolvidos
- CSession
- CEnv
- CHistory
- CJobTable
- COptionSet
- CFunctionRegistry
- CAliasRegistry

### Dependências
- Fase 1 concluída
- Estrutura mínima de console funcionando

### Critérios de conclusão
- Existe um objeto de sessão única por shell.
- O diretório atual pode ser lido e alterado.
- Variáveis de ambiente podem ser definidas e consultadas.
- Histórico de comandos é armazenado.
- A sessão mantém status de último comando executado.
- O shell consegue distinguir estados como inicial, ativo, terminado e erro.

### Testes mínimos
- Teste 1: criação da sessão sem erros.
- Teste 2: define e consulta variável simples (`FOO=bar`).
- Teste 3: altera diretório atual com `cd` em memória.
- Teste 4: histórico salva comandos em ordem.
- Teste 5: último status de execução é mantido corretamente.

---

## Fase 3 — Streams

### Objetivo
Implementar a camada de entrada/saída que representa stdin, stdout e stderr, além de redirecionamentos básicos.

### Módulos envolvidos
- `IStream`
- `clsFileDescriptorTable`
- `clsDescriptorScope` / `clsDescriptorScopeEntry`
- `clsConsoleStream`
- `clsFileStream`
- `clsPipeStream`
- `clsNullStream`
- `clsRedirectionManager`

### Dependências
- Fase 1: terminal funcional
- Fase 2: sessão e ambiente existindo

### Critérios de conclusão
- O shell possui um fd table com 0, 1 e 2 definidos.
- stdin lê do console.
- stdout escreve no console.
- stderr pode ser separado de stdout.
- Há suporte a `>` e `<` em nível de abstração.
- Há suporte a duplicação de fd (`dup`, `dup2` em forma mínima).
- `/dev/null` pode ser modelado como stream nulo.

### Testes mínimos
- Teste 1: stdout escreve no console.
- Teste 2: stderr pode ser redirecionado para arquivo ou buffer.
- Teste 3: `> file` cria ou sobrescreve arquivo de saída.
- Teste 4: `< file` lê conteúdo de arquivo para entrada.
- Teste 5: `2>&1` ou equivalente em abstração produz saída combinada corretamente.

### Implementação neste marco — `clsFileStream`
- `clsFileStream` implementa `IStream` e mantém os métodos legados `OpenForRead`, `OpenForWrite`, `ReadText`, `WriteText` e `CloseStream`.
- A leitura usa `Scripting.FileSystemObject.TextStream.Read(maxChars)` em blocos, com consulta de EOF sem avançar a posição.
- Abertura read-only, overwrite e append configuram capacidades distintas; fechamento é idempotente, preserva essas capacidades e permite reabertura.
- Falhas de abertura, leitura, escrita e fechamento são mapeadas aos códigos disponíveis em `IStream`; operações incompatíveis ou após fechamento usam os erros de contrato.
- `TestFileStream` cobre os modos, leitura, escrita, EOF, erros e lifecycle com arquivos temporários e limpeza após falhas. Os casos de regressão também cobrem transições da mesma instância escrita→leitura e leitura→escrita, overwrite após reabertura e leitura externa somente depois de fechar.
- `TestNullStream` espera 10 caracteres para `"hello" & vbCrLf & "abc"` (5 + 2 + 3).
- Execução Excel/VBA de `VBAshTests-r3` informada pelo usuário em 2026-11-05: 17/17 suítes PASS, 656 assertions; `FileStream` (74), `NullStream` (42) e `Redirection` (74) PASS. O checkout Git não foi identificado pelo relatório e a correspondência com os módulos carregados permanece UNVERIFIED; detalhes em `docs/test-reports/phase-streams/execution-log.md`.
- `clsFileDescriptorTable` administra associações `FD -> IStream` após inicialização explícita dos FDs 0/1/2 com objetos fornecidos pelo host. Ownership por identidade é local à tabela; `ownsStream=True` transfere a responsabilidade de fechamento. O responsável deve chamar `CloseAll` explicitamente: não há fechamento automático em `Class_Terminate`. A operação remove FDs emprestados sem fechar seus streams, fecha cada stream próprio uma vez e remove seus aliases; se algum fechamento falhar, tenta os demais, preserva os aliases e ownership dos que falharam e retorna erros agregados para retry. A instância continua inicializada e reutilizável, mas `Initialize` não pode ser chamada novamente. Esta implementação isolada não conecta a tabela a `clsSession`, executor ou redirecionamentos.
- O usuário informou execução Excel/VBA de `VBAshTests-r5` em 2026-11-06: 18/18 suítes PASS, 718 assertions, incluindo 62 PASS em `FileDescriptorTable`; checkout/commit não identificado e correspondência UNVERIFIED. A nova suíte `FileDescriptorTableCloseAll` em `VBAshTests-r6` não foi executada no Excel/VBA.

### Integração da Fase 3 — implementação em revisão

A branch `feature/phase-3-integration` conecta a tabela de descritores à sessão e ao executor de comandos simples:
- `clsDescriptorScope` e `clsDescriptorScopeEntry` guardam snapshots de bindings por identidade de stream; os escopos são LIFO e mantêm referências aos streams originais durante a execução.
- `clsRedirectionManager` aplica redirecionamentos em ordem lexical, abrindo `clsFileStream` para entrada/saída e usando `clsNullStream` para `/dev/null`.
- A sessão prepara o escopo antes do builtin, encaminha stdout/stderr pelos FDs atuais e tenta restaurar os bindings depois da execução.
- O host inicializa FD 0 com stream nulo em EOF e FDs 1/2 com adaptadores de captura `clsConsoleStream`; o formulário continua consumindo as strings de resultado após `ProcessInput`.
- O lexer/parser preservam operadores de duplicação/fechamento `2>&1`, `2>&-`, `1>&2`, `n>&m` e `n>&-`.
- `clsPipeStream` fornece apenas o buffer textual de base. Não há ainda scheduler, bloqueio, endpoints independentes por processo ou execução de pipeline.

**Limites conhecidos desta integração:** os builtins ainda retornam strings, portanto a saída não é incremental e a ordem temporal entre chamadas independentes a stdout/stderr não pode ser reconstruída; a sessão escreve primeiro o resultado stdout e depois stderr. A entrada redirecionada está disponível no FD 0, mas os builtins atuais ainda não consomem esse stream. O host ainda usa stdin nulo, não entrada interativa. A execução de `cmd1 | cmd2` continua deliberadamente desabilitada e pertence à Fase 9.

O runner de desenvolvimento foi atualizado para 23 suítes, incluindo escopos, redirecionamento integrado, adaptador de console e stream de pipe. **Esta revisão ainda precisa ser importada, compilada e executada no Excel/VBA.** Até que o relatório dessa execução confirme as 23 suítes e as revisões carregadas, a Fase 3 permanece em validação e não deve ser marcada como concluída.

---

## Fase 4 — Command Registry

### Objetivo
Criar o registrador de comandos que identifica o tipo de comando: builtin, função, alias, externo, comando composto.

### Módulos envolvidos
- CCommand
- CCommandResolver
- CBuiltinRegistry
- CExternalCommandResolver
- CFunctionRegistry
- CAliasRegistry

### Dependências
- Fase 2: sessão funcional
- Fase 3: streams e redireções funcionando

### Implementação incremental — contrato inicial (Fase 4A)

- `clsResolvedCommand` é o resultado da resolução e permanece separado de `clsCommand`, que representa nós da AST.
- O contrato contém `Name`, `CommandType`, `Args`, `Redirections`, `Context`, `SourceNode` e a referência `Builtin` quando aplicável.
- Tipos aceitos: `ALIAS`, `FUNCTION`, `BUILTIN`, `EXTERNAL` e `COMPOUND`.
- `Init` rejeita nome vazio e tipo desconhecido; argumentos e redirecionamentos são coleções independentes por instância.
- Esta entrega define apenas o contrato. Não resolve aliases/funções/executáveis, não altera a sessão e não executa comandos.
- Testes isolados adicionados ao runner `VBAshTests-r8`; a compilação e execução no Excel ainda dependem da validação do usuário.

### Implementação incremental — registry de builtins (Fase 4B)

- `clsBuiltinRegistry` mantém os builtins padrão e agora permite extensões por `RegisterBuiltin(commandName, builtin, Optional replaceExisting=False)`.
- Nomes são consultados sem distinção entre maiúsculas/minúsculas; nomes vazios ou com espaço/tabulação são rejeitados.
- Duplicatas são rejeitadas por padrão. Substituição exige `replaceExisting=True`, evitando sobrescrita silenciosa.
- `HasBuiltin` consulta presença e `UnregisterBuiltin` remove um registro, retornando Boolean para indicar se havia item.
- `Resolve(commandName) As IBuiltin` permanece compatível com o chamador atual. A API ainda não foi exposta por `clsCommandResolver` nem por `clsSession`; a integração do resolver será uma etapa posterior.
- A revisão `clsBuiltinRegistry-r1` e a suíte `BuiltinRegistry` foram adicionadas ao runner `VBAshTests-r9`. A execução no Excel ainda depende da validação do usuário.


### Integração do resolver — Fase 4C

- `clsCommandResolver.ResolveCommand(sourceNode, sessionContext)` recebe somente nós AST `SIMPLE_COMMAND` já expandidos pela sessão; não repete expansão nem executa comandos.
- O resolver classifica builtins registrados como `BUILTIN`, preserva a API legada `Resolve(commandName) As IBuiltin` e retorna `UNRESOLVED` para nomes ainda não registrados.
- O objeto resolvido mantém referência ao nó AST original, copia os argumentos já expandidos e preserva a identidade dos objetos de redirecionamento.
- `clsSession` passa a consumir o resultado do resolver para executar builtins; comandos `UNRESOLVED` retornam status 127. Nenhum processo externo é iniciado.
- Aliases, funções e descoberta de executáveis continuam pendentes; a prioridade completa de resolução ainda não está implementada.
- A suíte `CommandResolver` foi adicionada ao runner `VBAshTests-r10`; a validação de compilação e execução no Excel/VBA depende da importação e do relatório do usuário.

### Integração de aliases — Fase 4D

- clsAliasRegistry mantém nomes e textos de substituição em um dicionário case-sensitive; duplicatas exigem substituição explícita.
- clsAliasExpander opera sobre tokens antes do parser, reutiliza clsTokenizer e não modifica a coleção de entrada nem a AST.
- A expansão ocorre somente em palavras de comando não citadas, incluindo posições após ;, |, &&, || e após palavras de atribuição NOME=valor.
- Palavras de argumentos e destinos de redirecionamento não são expandidos. Redirecionamentos anteriores ao nome do comando são aceitos pelo parser para permitir > arquivo alias.
- O texto de substituição pode introduzir várias palavras; os argumentos originais permanecem depois dos tokens introduzidos.
- A recursão é protegida por uma pilha de aliases ativos e limite de profundidade. Auto-referências e ciclos terminam emitindo literalmente o alias já ativo, evitando expansão infinita.
- A sessão expõe RegisterAlias, HasAlias, GetAlias, RemoveAlias, GetAliasNames e ClearAliases. Os builtins `alias` e `unalias` são registrados por padrão: `alias NOME=VALOR` cria/substitui, `alias NOME` consulta, `alias` lista, `unalias NOME` remove e `unalias -a` limpa a tabela. Aliases com valor terminado em espaço, funções, comandos externos e execução de pipelines continuam fora desta entrega.
- A expansão lexical de NOME=valor permite reconhecer o próximo nome de comando para alias, mas a semântica de atribuições prefixadas ainda não foi implementada no parser/executor.
- A suíte AliasExpansion foi adicionada ao runner VBAshTests-r11 e cobre registry, expansão em posições de comando, argumentos, operadores, redirecionamentos, aspas, atribuições prefixadas, aliases vazios, auto-referência/ciclos e integração com a sessão. A suíte AliasBuiltins foi adicionada ao runner VBAshTests-r12 e cobre cadastro, consulta, listagem, substituição, expansão executada pelo terminal, remoção e limpeza. A validação no Excel/VBA depende da importação e execução do runner r12.

### Critérios de conclusão
- O shell consegue classificar builtins e manter um resultado de resolução separado da AST.
- A expansão lexical de aliases ocorre antes da AST e, portanto, antes da resolução de comandos.
- Os aliases podem ser criados, consultados, listados e removidos pelos builtins `alias` e `unalias`.
- A resolução pós-AST deverá respeitar a prioridade entre função, builtin e executável externo quando os registries de funções e o resolvedor de executáveis forem introduzidos nas fases que fornecem essas capacidades.
- Comandos sem nome não são aceitos.
- O objeto CCommand armazena nome, Args, redireções, contexto e tipo.
- A resolução é separada da execução.

### Testes mínimos
- Teste 1: comando externo simples (`dir`, `echo`, `cmd /c echo`) é reconhecido como externo.
- Teste 2: builtin registrado é reconhecido corretamente.
- Teste 3: função registrada é resolvida antes de executável.
- Teste 4: alias registrado é expandido antes da AST e o nome resultante segue para a resolução pós-AST.
- Teste 5: `alias NOME=VALOR`, `alias NOME`, `alias`, `unalias NOME` e `unalias -a` operam sobre o registry da sessão.
- Teste 6: comando vazio retorna erro sem quebrar shell.

---

### Validação Excel — Fase 4D (2026-11-09)\n\n- Runner: `VBAshTests-r11`.\n- Resultado: 27/27 suítes PASS; 933 asserções; 0 FAIL; 0 ERROR.\n- `AliasExpansion`: 41 asserções PASS.\n- Módulos de sessão/parser identificados como `clsSession-r5` e `clsParser-r3`; registry/expansor novos não foram listados no relatório.\n- Checkout/commit UNKNOWN e correspondência checkout/módulos UNVERIFIED.\n- Conclusão: expansão lexical e integração de aliases funcionalmente validadas; interface de terminal `alias`/`unalias` ainda não estava implementada nesta validação.\n\n## Fase 5 — Lexer

### Objetivo
Implementar a análise léxica do shell, convertendo texto em tokens compreensíveis para o parser.

### Módulos envolvidos
- CTokenizer
- CToken
- CTokenType
- CStringScanner
- CQuoteState
- CEscapeHandler

### Dependências
- Fases 1 a 4 concluídas
- O shell já deve receber texto e produzir um comando bruto

### Critérios de conclusão
- O lexer reconhece tokens básicos: palavras, operadores, pipes, redirecionadores, parênteses e comentários.
- Aspas simples e duplas são tratadas corretamente.
- Escape de caracteres funciona.
- Tokens de variáveis e expansão são preservados corretamente para a etapa de parser/expander.
- O lexer identifica `|`, `>`, `<`, `&&`, `||`, `;`, `(`, `)`, `{`, `}`.

### Testes mínimos
- Teste 1: `echo hello` produz 3 tokens.
- Teste 2: `echo "hello world"` produz 2 tokens, com palavra preservada com espaço.
- Teste 3: `echo 'a b'` mantém o valor literal.
- Teste 4: `echo hello\ world` trata escape corretamente.
- Teste 5: `echo foo | grep foo` produz tokens de pipe corretamente.
- Teste 6: `echo hi # comment` ignora comentário corretamente.

---

## Fase 6 — Parser

### Objetivo
Transformar tokens em uma árvore sintática de comando (AST) que possa ser executada em etapas.

### Módulos envolvidos
- CParser
- CNode
- CCommandNode
- CPipelineNode
- CListNode
- CIfNode
- CForNode
- CWhileNode
- CFunctionDefNode
- CRedirectNode
- CWordNode

### Dependências
- Fase 5: lexer concluído
- Fase 4: command registry pronto

### Critérios de conclusão
- O parser consegue montar AST para comandos simples.
- Suporta listas sequenciais (`;`).
- Suporta pipes (`|`).
- Suporta redireções em comando ou pipeline.
- Suporta blocos `(...)` e `{ ...; }` como nós compostos.
- A árvore pode ser executada em ordem em um runner de execução.

### Testes mínimos
- Teste 1: `echo hi` gera AST de comando simples.
- Teste 2: `echo a | sort` gera AST de pipeline.
- Teste 3: `echo a ; echo b` gera AST de lista sequencial.
- Teste 4: `echo hi > out.txt` inclui redireção no comando.
- Teste 5: `(echo a)` gera nó de subshell/comando composto.
- Teste 6: `{ echo a; echo b; }` gera AST de bloco.

---

## Fase 7 — Builtins

### Objetivo
Implementar os builtins essenciais do shell, priorizando aqueles que permitem um shell útil e funcional.

### Módulos envolvidos
- CBuiltinRegistry
- CBuiltin_echo
- CBuiltin_cd
- CBuiltin_pwd
- CBuiltin_export
- CBuiltin_set
- CBuiltin_unset
- CBuiltin_exit
- CBuiltin_alias
- CBuiltin_read
- CBuiltin_test
- CBuiltin_history

### Dependências
- Fase 2: sessão funcional
- Fase 4: command registry
- Fase 6: parser funcional

### Critérios de conclusão
- Os builtins principais do shell estão registrados.
- Cada builtin executa em contexto de sessão com acesso ao ambiente.
- O resultado do builtin muda corretamente o estado da sessão.
- O shell retorna status de saída do builtin.
- Builtins podem ler e gravar streams corretamente.

### Testes mínimos
- Teste 1: `echo hello` imprime `hello`.
- Teste 2: `cd` altera o diretório atual da sessão.
- Teste 3: `pwd` retorna o diretório atual.
- Teste 4: `export FOO=bar` adiciona variável ao ambiente.
- Teste 5: `unset FOO` remove a variável.
- Teste 6: `exit 5` finaliza a sessão com status 5.
- Teste 7: `read` captura entrada e salva em variável.
- Teste 8: `test -d .` retorna status correto.

---

## Fase 8 — Process Manager

### Objetivo
Implementar a criação e gerenciamento de processos externos e internos, garantindo execução correta de comandos com status de saída.

### Módulos envolvidos
- CExecutableProcess
- CProcessManager
- CJob
- CJobTable
- CWin32ProcessAdapter
- CProcessStatus
- CProcessGroup

### Dependências
- Fase 3: streams
- Fase 4: command registry
- Fase 6: parser
- Fase 7: builtins essenciais

### Critérios de conclusão
- O shell pode iniciar um processo Windows associado a um comando externo.
- Cada processo tem PID e status.
- O shell coleta exit code corretamente.
- O processo pode depender de stdin, stdout e stderr configurados.
- Comandos externos podem ser executados sem quebrar o shell.
- Existe uma tabela de jobs com status inicial.

### Testes mínimos
- Teste 1: executar `echo hello` como processo externo retorna saída correta.
- Teste 2: comando inexistente retorna status de erro.
- Teste 3: processo com saída longa é capturado no buffer de stdout.
- Teste 4: processo com stderr separado é capturado sem contaminar stdout.
- Teste 5: processo finalizado aparece em tabela de jobs com status correto.

---

## Fase 9 — Pipes

### Objetivo
Adicionar suporte funcional a pipelines, respeitando o fluxo de dados entre processos.

### Módulos envolvidos
- CPipeline
- CPipeStream
- CProcessManager
- CRedirectionManager
- CJobTable

### Dependências
- Fase 3: streams
- Fase 8: process manager
- Fase 6: parser com AST pipeline

### Critérios de conclusão
- O shell executa `cmd1 | cmd2` corretamente.
- O stdout de um processo vira stdin do seguinte.
- Processos em pipeline podem ser monitorados como job único.
- O shell mantém status do pipeline corretamente.
- `|&` pode ser implementado em nível mínimo, se a arquitetura permitir.

### Testes mínimos
- Teste 1: `echo hello | more` ou `echo hello | findstr hello` funciona com saída correta.
- Teste 2: `dir | findstr .` imprime o resultado do filtro sem quebrar a execução.
- Teste 3: pipeline de múltiplos passos (`a | b | c`) produz saída composta e status adequado.
- Teste 4: pipeline com comando falhando ainda produz status final consistente.
- Teste 5: pipeline em background pode ser registrado em jobs sem travar o shell.

---

## Fase 10 — POSIX Layer

### Objetivo
Implementar a camada de compatibilidade POSIX necessária para scripts e comandos Bash-like mais reais, sem depender de um ambiente POSIX nativo.

### Módulos envolvidos
- CPosixEmulator
- CPathResolver
- CFileSystemAdapter
- CSignalMap
- CJobControlEmulator
- CTtyEmulator
- CDevNodeRegistry

### Dependências
- Fase 8: process manager
- Fase 9: pipes
- Fase 7: builtins essenciais
- Fase 2: sessão funcional

### Critérios de conclusão
- O shell suporta diretórios, `pwd`, `cd`, `PATH` e ambiente de execução em memória.
- `/dev/null` existe como stream nulo.
- `jobs`, `wait`, `fg`, `bg` têm comportamento funcional de shell em contexto de sessão.
- Sinais simples são mapeados e tratados.
- Variáveis e estados Bash-like são preservados para scripts básicos.
- O shell pode executar scripts simples com `source` ou chamada direta.

### Testes mínimos
- Teste 1: `pwd` retorna o diretório atual da sessão.
- Teste 2: `cd /tmp` ou equivalente em diretório virtual funciona.
- Teste 3: `/dev/null` aceita escrita sem erro e descarta conteúdo.
- Teste 4: `jobs` lista job em execução ou concluído.
- Teste 5: `wait` bloqueia e depois retorna quando processo conclui.
- Teste 6: script simples com `source` ou execução direta roda sem erro.
- Teste 7: `kill` e `trap` em nível mínimo são mapeados corretamente.

---

## Marcos de integração

### Marco 1 — Shell funcional mínimo
- Terminal
- Session
- Streams
- Command Registry
- Lexer
- Parser
- Builtins essenciais
- Execução de comando simples

Objetivo: conseguir executar `echo hello` e comandos básicos no Excel/VBA.

### Marco 2 — Shell com execução real
- Process Manager
- Jobs básico
- Redireções
- Pipes simples

Objetivo: executar comandos externos e pipelines básicos sem quebrar a sessão.

### Marco 3 — Bash-like em ambiente não-POSIX
- POSIX Layer
- TTY emulada
- Jobs e sinais
- Scripts básicos

Objetivo: tornar a shell funcional para a maioria dos scripts Bash simples.

---

## Regra de implementação incremental

1. Não implementar recursos avançados antes do shell funcional mínimo.
2. Todo módulo deve ter testes mínimos antes de ser considerado concluído.
3. Cada fase deve deixar o sistema em estado estável.
4. A arquitetura do parser e do executor deve ser testada assim que existir AST e execução.
5. A compatibilidade com Bash 5.x é um objetivo de longo prazo; o primeiro alvo é um shell funcional e previsível.

---

## Ordem recomendada de execução real

1. Terminal
2. Session
3. Streams
4. Command Registry
5. Lexer
6. Parser
7. Builtins básicos
8. Process Manager
9. Pipes
10. POSIX Layer

Essa ordem reduz risco de refatoração e mantém cada módulo com dependências claras e bem delimitadas.

---

## Observações finais

- O componente mais crítico do projeto não é o shell “per se”, e sim a camada de semântica do comando e do processo.
- Para ser compatível com Bash, a implementação deve priorizar comportamento correto em vez de número de features.
- Em um ambiente Excel/VBA, a compatibilidade da sessão, dos streams, dos jobs e da emulação POSIX será o que realmente permitirá scripts Bash simples e previsíveis.

Este plano foi desenhado para evoluir de forma incremental, sem overengineering prematuro e com foco em um shell funcional, verificável e extensível.
### Validação Excel — Fase 4C (2026-11-09)

- Runner: `VBAshTests-r10`.
- Resultado: 26/26 suítes PASS; 892 asserções; 0 FAIL; 0 ERROR.
- Módulos carregados identificados: `clsSession-r4`, `clsResolvedCommand-r1`, `clsBuiltinRegistry-r1`, `clsCommandResolver-r1`; tags e propriedades concordam onde reportado.
- `CommandResolver`: 13 asserções PASS.
- Regressão: AST, execução, redirecionamentos, streams e demais suítes anteriores também PASS.
- Integridade: checkout/commit UNKNOWN e correspondência checkout/módulos UNVERIFIED. A validação comprova as revisões identificadas no projeto VBA carregado, mas não vincula o workbook a um commit Git específico.
- Conclusão: Fase 4C funcionalmente validada no Excel, com a ressalva de identidade do checkout.

\n\n### Implementação de interface de aliases — Fase 4E (em validação)\n\n- `clsBuiltinAlias` e `clsBuiltinUnalias` conectam a sintaxe de terminal ao registry da sessão.\n- `alias NOME=VALOR` cria ou substitui; `alias NOME` consulta; `alias` lista. `unalias NOME` remove e `unalias -a` limpa todos.\n- O registry e a sessão expõem enumeração e limpeza; o instalador inclui os dois novos módulos na lista de limpeza legada.\n- O runner `VBAshTests-r12` acrescenta `AliasBuiltins`, incluindo fluxo real por `ProcessInput`.\n- A implementação ainda precisa ser importada/compilada e executada no Excel/VBA; só então a Fase 4 poderá ser fechada funcionalmente.\n