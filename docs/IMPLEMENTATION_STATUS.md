# IMPLEMENTATION_STATUS.md

Status de implementação do VBAsh seguindo `IMPLEMENTATION_PLAN.md`.

## Fase 1 — Terminal

### Objetivo
Criar interface mínima de shell no VBA/Excel com loop REPL básico.

### Implementado
- `VBAsh.bas` com `VBAsh_Start`.
- Inicialização de sessão (`clsSession`), console (`clsConsole`) e tabela de FDs (`clsFileDescriptorTable`).
- Loop de leitura por `InputBox`, encaminhamento para `ProcessInput` e saída via `Debug.Print`.

### Critérios atendidos
- [x] Prompt inicial.
- [x] Captura de entrada de usuário.
- [x] Saída básica em console (`Debug.Print`).
- [x] Loop principal de leitura/execução/impressão.

### Pendências
- [ ] Adaptador de console avançado (UserForm/Worksheet).
- [ ] Testes automatizados formais (atualmente manuais).

## Fase 2 — Session

### Implementado
- `clsSession` com:
  - Diretório atual
  - Ambiente em `Scripting.Dictionary`
  - Histórico de comandos
  - `LastExitCode`
  - `Prompt`
  - `ProcessInput` minimal para `pwd`, `help`, `exit` e atribuição `KEY=VALUE`

### Critérios atendidos parcialmente
- [x] Sessão única por shell.
- [x] Variáveis simples (set/get).
- [x] Histórico básico.
- [x] Status de saída básico.
- [ ] Estados de ciclo de vida explícitos (inicial/ativo/terminado/erro).

## Fase 3 — Streams

### Implementado
- `clsStream` com buffer em memória.
- `clsConsole` com streams de fd 0/1/2.
- `clsFileDescriptorTable` com inicialização default e operações `SetFD/GetFD/CloseFD`.

### Critérios atendidos parcialmente
- [x] fd table básica com 0,1,2.
- [x] stdout/stderr modelados.
- [ ] Redirecionamentos (`>`, `<`, `2>&1`) ainda não executados.
- [ ] `dup/dup2` não implementado.

## Fase 4+ (resumo)
- `clsCommandResolver` criado como scaffold.
- `clsCommand` corrigido para comportamento válido no VBA.
- `clsParser` criado como parser mínimo de comando simples.
- `clsTokenizer` já existente no repositório (base inicial).

## Próximo PR recomendado
1. Integrar `Tokenizer -> Parser -> CommandResolver` dentro de `ProcessInput`.
2. Implementar builtin `echo` real com args.
3. Implementar redirecionamento de saída básico (`>` para arquivo).
4. Adicionar testes manuais documentados por fase.
