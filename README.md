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

## Testes

No editor VBA, importe os módulos do projeto e execute a macro `VBAsh_RunTests`. Os testes cobrem tokenização, builtins, estado da sessão, erros, códigos de saída e redirecionamento de saída.

## Redirecionamento inicial

`>` cria ou sobrescreve um arquivo; `>>` acrescenta saída ao arquivo. Caminhos relativos são resolvidos a partir do diretório atual da sessão. O parser reconhece outras classes de operadores, mas pipes, entrada redirecionada (`<`) e redirecionamentos múltiplos ainda não são executados.

## Limitações atuais

Este é um shell em estágio inicial: comandos externos, expansão de variáveis, pipes, execução de listas e scripts ainda não são suportados. Cada builtin é implementado em seu próprio módulo e registrado por `clsBuiltinRegistry`; o estado da sessão e a interface de entrada/saída continuam mínimos.

## Teste manual rápido

Inicie `VBAsh_Start` no VBA e experimente:

```text
echo "hello world"
pwd
export DEMO=ok
echo "salvo em arquivo" > output.txt
echo "mais saída" >> output.txt
unset DEMO
cd ~
cd -
exit 3
```
