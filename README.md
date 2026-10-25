# VBAsh

Um shell mínimo implementado em VBA. A versão atual executa builtins fundamentais e mantém diretório atual, variáveis de ambiente e histórico na sessão.

## Builtins disponíveis

- `echo [-n] [texto ...]`
- `pwd`
- `cd [diretório]`, incluindo `cd -` e `~`
- `export [NOME[=valor] ...]`
- `setenv NOME valor`
- `unset NOME [NOME ...]`
- `help`
- `exit [status]`

Também é possível definir uma variável com `NOME=valor`. Nomes de variáveis seguem o formato usual de identificador ASCII, e são comparados sem distinção entre maiúsculas e minúsculas.

Aspas simples e duplas preservam espaços em argumentos; escapes com barra invertida são aceitos. A interface atual usa `InputBox` para entrada e `MsgBox` para saída.

## Limitações atuais

Este é um shell em estágio inicial: comandos externos, expansão de variáveis, redirecionamentos, pipes, execução de listas e scripts ainda não são suportados. Operadores reconhecidos pelo tokenizador são rejeitados pelo parser até que essas funcionalidades sejam implementadas.

## Teste manual rápido

Inicie `VBAsh_Start` no VBA e experimente:

```text
echo "hello world"
pwd
export DEMO=ok
unset DEMO
cd ~
cd -
exit 3
```
