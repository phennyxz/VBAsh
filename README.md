# VBAsh

Um shell mínimo implementado em VBA. A versão atual executa builtins fundamentais e mantém diretório atual, variáveis de ambiente e histórico na sessão.

## Builtins disponíveis

- `echo [-n] [texto ...]`
- `pwd`
- `cd [diretório]`, incluindo `cd -` e `~`
- `ls [caminho]` para listar diretórios ou exibir o nome de um arquivo
- `mkdir caminho` para criar um diretório
- `clear` para limpar a área de saída do terminal
- `export [NOME[=valor] ...]`
- `setenv NOME valor`
- `unset NOME [NOME ...]`
- `test` ou `[ expressão ]` com testes (`-n`, `-z`, `-d`, `-f`, `-e`, `=`, `==`, `!=`, `-eq`, `-ne`, `-gt`, `-ge`, `-lt`, `-le`)
- `if comando; then comando; [else comando;] fi` em uma única linha; condições aceitam comandos ou testes simples encadeados com `&&` e `||`, avaliados da esquerda para a direita
- `for NOME in valores; do comando; done` em uma única linha, com expansão básica de `$NOME` e intervalos numéricos como `{1..7}`
- `help`
- `exit [status]`

Também é possível definir uma variável com `NOME=valor`. Nomes de variáveis seguem o formato usual de identificador ASCII, e são comparados sem distinção entre maiúsculas e minúsculas.

Aspas simples e duplas preservam espaços em argumentos; escapes com barra invertida são aceitos. O instalador cria um UserForm com aparência de terminal: a saída fica em uma área rolável e copiável, e uma linha de entrada fixa na parte inferior recebe os comandos. Enter executa o comando e a saída rola automaticamente até o final. Mensagens de instalação e erros são exibidas em caixas de diálogo.

## Testes

No editor VBA, importe os módulos do projeto e execute a macro `VBAsh_RunTests`. Os testes cobrem tokenização, builtins, operações de diretório, condicionais, estado da sessão, erros, códigos de saída e redirecionamento de saída.

## Instalação no VBAsh.xlsm

Importe `VBAshInstaller.bas` para o projeto VBA do workbook e salve como `.xlsm`. No Excel, habilite o acesso confiável ao modelo de objeto do projeto VBA em Central de Confiabilidade > Configurações de Macro. Execute `VBAsh_Install` e selecione a pasta raiz contendo os arquivos do projeto. O instalador descobre os arquivos `.bas` e `.cls` nessa pasta automaticamente (exceto `VBAshInstaller.bas`), registra os módulos importados para permitir a desinstalação e cria o UserForm `frmVBAsh`; confirme a importação e, se desejar, o início do terminal.

O instalador recusa sobrescrever módulos existentes e remove os módulos adicionados nesta tentativa caso a importação falhe. Para desinstalar, execute `VBAsh_Uninstall` e confirme; a rotina usa a lista exata de módulos registrada na instalação (ou a lista compatível com instalações antigas). O módulo `modVBAshInstaller` é mantido para que a remoção seja explícita e auditável.

Depois de instalar e salvar o workbook, o macro `Auto_Open` abre o terminal ao abrir o arquivo com macros habilitadas. A janela do workbook fica oculta enquanto o formulário está aberto e volta a aparecer quando o terminal fecha. O Excel continua em execução; para remover VBAsh, reabra o arquivo e execute `VBAsh_Uninstall`.

## Redirecionamento inicial

`>` cria ou sobrescreve um arquivo; `>>` acrescenta saída ao arquivo. `2>` e `2>>` fazem o mesmo para mensagens de erro (`stderr`), e `2>&1` encaminha erros para o destino atual da saída padrão (`stdout`); a ordem é respeitada, como em `comando > arquivo 2>&1`. A gravação é síncrona: o stream do arquivo é fechado antes de `ProcessInput` retornar. Os testes de redirecionamento verificam existência, tamanho e conteúdo do arquivo após esse retorno. `> /dev/null`, `> NUL` e `>&-` descartam `stdout`; `2>&-` descarta `stderr`, sem alterar o código de saída do comando. Caminhos relativos são resolvidos a partir do diretório atual da sessão. Pipes e entrada redirecionada (`<`) ainda não são executados.

No `if`, `comando; then` executa um comando e decide pela cláusula com base no código de saída. A forma `[ ... ]` é um teste (equivalente a `test`), não uma forma de executar comandos; por isso `[ ls mapas ]` tenta interpretar `ls` como operador de teste e é inválido. Para comparar texto, use `=` ou `==`; para comparar números, use operadores como `-eq` e `-gt` (`eq` sem hífen não é um operador de `test`). As condições encadeadas com `&&`/`||` ainda precisam ser comandos ou testes simples. O formato exige `;` antes de `then`, `else` e `fi`. O `for` aceita uma lista explícita de palavras, intervalos inteiros inclusivos `{início..fim}` ascendentes ou descendentes e um único comando simples no corpo; a expansão de variáveis suporta `$NOME` e `${NOME}`, respeita `$NOME` em aspas duplas e mantém literais aspas simples/escapadas. Ainda não há suporte a `elif`, blocos aninhados, loops aninhados, `for` sobre globbing ou listas posicionais, nem comandos compostos dentro de cláusulas. `&&` e `||` fora de `if` ainda não são listas de comandos.

## Limitações atuais

Este é um shell em estágio inicial: comandos externos, expansão avançada de variáveis, pipes, execução de listas e scripts ainda não são suportados. Cada builtin é implementado em seu próprio módulo e registrado por `clsBuiltinRegistry`; o estado da sessão e a interface de entrada/saída continuam mínimos.

## Teste manual rápido

Inicie `VBAsh_Start` no VBA e experimente:

```text
echo "hello world"
pwd
ls
mkdir nova-pasta
export DEMO=ok
echo "salvo em arquivo" > output.txt
echo "mais saída" >> output.txt
unset DEMO
cd ~
cd -
exit 3
```
