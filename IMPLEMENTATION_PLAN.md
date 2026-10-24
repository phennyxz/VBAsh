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
- CStream
- CFileDescriptorTable
- CConsoleStream
- CFileStream
- CPipeStream
- CNullStream
- CRedirectionManager

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

### Critérios de conclusão
- O shell consegue classificar um comando em uma de várias categorias.
- O registry resolve a prioridade correta:
  1. alias
  2. função
  3. builtin
  4. executável externo
- Comandos sem nome não são aceitos.
- O objeto CCommand armazena nome, Args, redireções, contexto e tipo.
- A resolução é separada da execução.

### Testes mínimos
- Teste 1: comando externo simples (`dir`, `echo`, `cmd /c echo`) é reconhecido como externo.
- Teste 2: builtin registrado é reconhecido corretamente.
- Teste 3: função registrada é resolvida antes de executável.
- Teste 4: alias registrado é resolvido corretamente.
- Teste 5: comando vazio retorna erro sem quebrar shell.

---

## Fase 5 — Lexer

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