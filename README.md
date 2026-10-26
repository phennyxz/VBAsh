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

## Instalação no VBAsh.xlsm

Importe `VBAshInstaller.bas` para o projeto VBA do workbook e salve como `.xlsm`. No Excel, habilite o acesso confiável ao modelo de objeto do projeto VBA em Central de Confiabilidade > Configurações de Macro. Execute `VBAsh_Install`, selecione a pasta raiz contendo os arquivos deste repositório e confirme a importação; o instalador cria o UserForm `frmVBAsh` com visual de terminal e pergunta antes de iniciá-lo.

O instalador só importa os nomes de módulos listados nele, recusa sobrescrever módulos existentes e remove os módulos adicionados nesta tentativa caso a importação falhe. Para remover os módulos VBAsh importados, execute `VBAsh_Uninstall` e confirme; a rotina só remove componentes se encontrar o marcador criado após a instalação. O módulo `modVBAshInstaller` é mantido para que a remoção seja explícita e auditável.

Depois de instalar e salvar o workbook, o macro `Auto_Open` abre o terminal ao abrir o arquivo com macros habilitadas. A janela do workbook fica oculta enquanto o formulário está aberto e volta a aparecer quando o terminal fecha. O Excel continua em execução; para remover VBAsh, reabra o arquivo e execute `VBAsh_Uninstall`.

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
