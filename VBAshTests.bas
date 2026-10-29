Option Explicit

Private mAssertions As Long

Public Sub VBAsh_RunTests()
    mAssertions = 0
    TestTokenizer
    TestEcho
    TestEnvironment
    TestDirectory
    TestFilesystem
    TestClear
    TestConditionals
    TestForLoops
    TestErrorsAndExitStatus
    TestRedirection
    Debug.Print "VBAsh: " & mAssertions & " assertions passed."
End Sub

Private Sub TestTokenizer()
    Dim tokenizer As clsTokenizer
    Dim tokens As Collection
    Dim token As clsToken

    Set tokenizer = New clsTokenizer
    Set tokens = tokenizer.Tokenize("echo ""hello world""")
    AssertEqual 3, tokens.Count, "token count includes EOF"
    Set token = tokens(2)
    AssertEqual "hello world", token.Value, "double-quoted argument"

    Set tokens = tokenizer.Tokenize("echo 'a b' # comment")
    Set token = tokens(2)
    AssertEqual "a b", token.Value, "single-quoted argument"
    AssertEqual 3, tokens.Count, "comment ignored"

    Set tokens = tokenizer.Tokenize("echo hello\ world")
    Set token = tokens(2)
    AssertEqual "hello world", token.Value, "escaped whitespace"

    Set tokens = tokenizer.Tokenize("echo ""C:\Program Files\VBAsh""")
    Set token = tokens(2)
    AssertEqual "C:\Program Files\VBAsh", token.Value, "quoted Windows path keeps backslashes"

    Set tokens = tokenizer.Tokenize("echo a && echo b || echo c")
    AssertEqual 9, tokens.Count, "logical operators are separate tokens"
    Set token = tokens(3)
    AssertEqual "AND_IF", token.TokenType, "and-if token"
    Set token = tokens(6)
    AssertEqual "OR_IF", token.TokenType, "or-if token"

    Set tokens = tokenizer.Tokenize("echo ""a && b""")
    Set token = tokens(2)
    AssertEqual "a && b", token.Value, "logical operators remain literal inside quotes"

    Set tokens = tokenizer.Tokenize("echo hidden>&-")
    Set token = tokens(3)
    AssertEqual "REDIR_DISCARD", token.TokenType, "close-output redirection token"

    Set tokens = tokenizer.Tokenize("echo hi 2> err.log 2>&1")
    Set token = tokens(3)
    AssertEqual "REDIR_ERR", token.TokenType, "stderr redirection token"
    Set token = tokens(5)
    AssertEqual "REDIR_ERR_TO_OUT", token.TokenType, "stderr-to-stdout token"
    Set tokens = tokenizer.Tokenize("pwd 2>&-")
    Set token = tokens(2)
    AssertEqual "REDIR_ERR_DISCARD", token.TokenType, "close-stderr redirection token"

    Set tokens = tokenizer.Tokenize("echo '$a'")
    Set token = tokens(2)
    AssertEqual "00", token.ExpansionMask, "single quotes disable variable expansion"
    Set tokens = tokenizer.Tokenize("echo ""$a""")
    Set token = tokens(2)
    AssertEqual "11", token.ExpansionMask, "double quotes allow variable expansion"
    Set tokens = tokenizer.Tokenize("echo \$a")
    Set token = tokens(2)
    AssertEqual "01", token.ExpansionMask, "escaped dollar disables variable expansion"
    AssertTrue Not token.BraceExpansionAllowed, "escaped words disable brace expansion"
    Set tokens = tokenizer.Tokenize("echo ""{1..3}""")
    Set token = tokens(2)
    AssertTrue Not token.BraceExpansionAllowed, "quoted words disable brace expansion"
End Sub

Private Sub TestForLoops()
    Dim session As clsSession

    Set session = New clsSession
    AssertEqual "1" & vbCrLf & "2" & vbCrLf & "3", _
        session.ProcessInput("for a in 1 2 3; do echo $a; done"), _
        "for iterates over values and expands loop variable"
    AssertTrue session.HasOutput, "for reports its body output to the terminal"
    AssertEqual "1" & vbCrLf & "2" & vbCrLf & "3" & vbCrLf & "4" & vbCrLf & _
        "5" & vbCrLf & "6" & vbCrLf & "7", _
        session.ProcessInput("for a in {1..7}; do echo $a; done"), _
        "for expands ascending brace ranges"
    AssertEqual "7", session.GetEnv("a"), "for leaves loop variable at final range value"
    AssertEqual "value=7", session.ProcessInput("echo ""value=$a"""), _
        "double-quoted variable expansion"
    AssertEqual "value=7", session.ProcessInput("echo value=${a}"), _
        "braced variable expansion"
    AssertEqual "$a", session.ProcessInput("echo '$a'"), _
        "single-quoted variable remains literal"
    AssertEqual "$a", session.ProcessInput("echo \$a"), _
        "escaped variable remains literal"
    AssertEqual "3" & vbCrLf & "2" & vbCrLf & "1", _
        session.ProcessInput("for a in {3..1}; do echo $a; done"), _
        "for expands descending brace ranges"
    AssertEqual "{1..3}", session.ProcessInput( _
        "for a in ""{1..3}""; do echo $a; done"), _
        "for does not expand quoted brace ranges"
    session.ProcessInput "for a in {1..1001}; do echo invalid; done"
    AssertEqual 2, session.LastExitCode, "for rejects excessively large brace ranges"
    AssertEqual "", session.ProcessInput("for item in; do echo not-run; done"), _
        "for with no values skips its body"
    AssertEqual 0, session.LastExitCode, "empty for loop status"

    session.ProcessInput "for item in 1 2; do test $item -eq 1; done"
    AssertEqual 1, session.LastExitCode, "for returns the final body command status"

    session.ProcessInput "for 1invalid in a; do echo invalid; done"
    AssertEqual 2, session.LastExitCode, "for rejects invalid variable names"
    session.ProcessInput "for item in a do echo invalid; done"
    AssertEqual 2, session.LastExitCode, "for requires separator before do"
    session.ProcessInput "for item in a; do echo invalid done"
    AssertEqual 2, session.LastExitCode, "for requires separator before done"
End Sub

Private Sub TestEcho()
    Dim session As clsSession
    Set session = New clsSession

    AssertEqual "hello world", session.ProcessInput("echo hello world"), "echo joins arguments"
    AssertTrue session.HasOutput, "echo marks output"
    AssertEqual 0, session.LastExitCode, "echo exit status"

    AssertEqual "no newline", session.ProcessInput("echo -n ""no newline"""), "echo -n prints text"
    AssertTrue session.SuppressNewline, "echo -n suppresses newline"

    AssertEqual "hello world", session.ProcessInput("echo ""hello world"""), "quoted echo"
End Sub

Private Sub TestEnvironment()
    Dim session As clsSession
    Set session = New clsSession

    AssertEqual "", session.ProcessInput("FOO=bar"), "assignment has no output"
    AssertEqual "bar", session.GetEnv("FOO"), "assignment sets variable"
    AssertEqual "", session.ProcessInput("export FOO=baz"), "export has no output"
    AssertEqual "baz", session.GetEnv("foo"), "environment is case insensitive"
    AssertEqual "", session.ProcessInput("setenv PATH ""a value"""), "setenv has no output"
    AssertEqual "a value", session.GetEnv("PATH"), "setenv preserves quoted value"
    AssertEqual "", session.ProcessInput("unset FOO"), "unset has no output"
    AssertEqual "", session.GetEnv("FOO"), "unset removes variable"
    AssertEqual 0, session.LastExitCode, "environment builtin exit status"
    AssertEqual 4, session.History.Count, "session records non-empty commands"
    AssertEqual "FOO=bar", session.History(1), "session history preserves command order"

    session.ProcessInput "1INVALID=value"
    AssertEqual 2, session.LastExitCode, "invalid environment name status"
End Sub

Private Sub TestDirectory()
    Dim session As clsSession
    Dim originalDirectory As String
    Dim parentDirectory As String
    Dim fileSystem As Object

    Set session = New clsSession
    originalDirectory = session.CurrentDirectory
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    parentDirectory = fileSystem.GetParentFolderName(originalDirectory)
    AssertTrue Len(parentDirectory) > 0, "test parent directory exists"

    session.ProcessInput "cd .."
    AssertEqual parentDirectory, session.CurrentDirectory, "cd updates session directory"
    AssertEqual parentDirectory & "> ", session.Prompt, "cd updates prompt"

    AssertEqual originalDirectory, session.ProcessInput("cd -"), "cd - prints previous directory"
    AssertEqual originalDirectory, session.CurrentDirectory, "cd - restores directory"
    AssertEqual 0, session.LastExitCode, "cd exit status"
End Sub

Private Sub TestFilesystem()
    Dim session As clsSession
    Dim fileSystem As Object
    Dim textStream As Object
    Dim tempRoot As String
    Dim testRoot As String
    Dim folderPath As String
    Dim filePath As String
    Dim output As String

    Set session = New clsSession
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    tempRoot = fileSystem.GetSpecialFolder(2)
    testRoot = fileSystem.BuildPath(tempRoot, "VBAshTests_" & fileSystem.GetTempName)
    fileSystem.CreateFolder testRoot
    folderPath = fileSystem.BuildPath(testRoot, "adir")
    fileSystem.CreateFolder folderPath
    filePath = fileSystem.BuildPath(testRoot, "bfile.txt")
    Set textStream = fileSystem.CreateTextFile(filePath, True)
    textStream.Close

    output = session.ProcessInput("ls """ & testRoot & """")
    AssertTrue InStr(1, output, "adir", vbBinaryCompare) > 0, "ls lists directories"
    AssertTrue InStr(1, output, "bfile.txt", vbBinaryCompare) > 0, "ls lists files"
    AssertEqual 0, session.LastExitCode, "ls exit status"

    AssertEqual "bfile.txt", session.ProcessInput("ls """ & filePath & """"), _
        "ls accepts a file path"
    AssertTrue session.HasOutput, "ls marks output"

    session.ProcessInput "cd """ & testRoot & """"
    output = session.ProcessInput("ls")
    AssertTrue InStr(1, output, "adir", vbBinaryCompare) > 0, "ls defaults to session directory"

    output = fileSystem.BuildPath(testRoot, "created")
    session.ProcessInput "mkdir """ & output & """"
    AssertTrue fileSystem.FolderExists(output), "mkdir creates a directory"
    AssertTrue Not session.HasOutput, "mkdir has no output"
    AssertEqual 0, session.LastExitCode, "mkdir exit status"

    session.ProcessInput "mkdir created-relative"
    AssertTrue fileSystem.FolderExists(fileSystem.BuildPath(testRoot, "created-relative")), _
        "mkdir resolves paths against session directory"

    session.ProcessInput "mkdir"
    AssertEqual 2, session.LastExitCode, "mkdir rejects missing path"

    fileSystem.DeleteFolder testRoot, True
End Sub

Private Sub TestClear()
    Dim session As clsSession

    Set session = New clsSession
    AssertEqual "", session.ProcessInput("clear"), "clear has no output"
    AssertTrue session.ClearRequested, "clear requests terminal clearing"
    AssertTrue Not session.HasOutput, "clear suppresses output"
    AssertEqual 0, session.LastExitCode, "clear exit status"

    session.ProcessInput "pwd"
    AssertTrue Not session.ClearRequested, "clear request resets for next command"

    session.ProcessInput "clear extra"
    AssertEqual 2, session.LastExitCode, "clear rejects arguments"
End Sub

Private Sub TestConditionals()
    Dim session As clsSession

    Set session = New clsSession
    AssertEqual "yes", session.ProcessInput( _
        "if test -d .; then echo yes; else echo no; fi"), "if selects true branch"
    AssertEqual 0, session.LastExitCode, "true conditional exit status"
    AssertTrue session.HasOutput, "conditional branch reports output"

    AssertEqual "no", session.ProcessInput( _
        "if test -e __vbash_missing_path__; then echo yes; else echo no; fi"), _
        "if selects false branch"
    AssertEqual 0, session.LastExitCode, "false conditional else status"

    AssertEqual "empty", session.ProcessInput( _
        "if test -z """"; then echo empty; fi"), "if supports no else branch"
    AssertEqual 0, session.LastExitCode, "false conditional without else succeeds"

    session.ProcessInput "if test -d . then echo invalid; fi"
    AssertEqual 2, session.LastExitCode, "if rejects missing separator"

    session.ProcessInput "test -d ."
    AssertEqual 0, session.LastExitCode, "test true status"
    session.ProcessInput "test -e __vbash_missing_path__"
    AssertEqual 1, session.LastExitCode, "test false status"
    session.ProcessInput "test ""same value"" = ""same value"""
    AssertEqual 0, session.LastExitCode, "test compares quoted strings"
    session.ProcessInput "test 7 -eq 7"
    AssertEqual 0, session.LastExitCode, "test supports numeric equality"
    session.ProcessInput "test 7 -ne 3"
    AssertEqual 0, session.LastExitCode, "test supports numeric inequality"
    session.ProcessInput "test 7 -gt 3"
    AssertEqual 0, session.LastExitCode, "test supports numeric greater-than"
    session.ProcessInput "test 7 -ge 7"
    AssertEqual 0, session.LastExitCode, "test supports numeric greater-or-equal"
    session.ProcessInput "test 3 -lt 7"
    AssertEqual 0, session.LastExitCode, "test supports numeric less-than"
    session.ProcessInput "test 7 -le 3"
    AssertEqual 1, session.LastExitCode, "test reports false numeric comparison"
    session.ProcessInput "test value -eq 7"
    AssertEqual 2, session.LastExitCode, "test rejects nonnumeric comparison operands"
    AssertEqual 14, session.History.Count, "conditional internals do not duplicate history"

    AssertEqual "bracket", session.ProcessInput( _
        "if [ -d . ] && [ -n ok ]; then echo bracket; else echo no; fi"), _
        "if accepts bracket test with and-if"
    AssertEqual "fallback", session.ProcessInput( _
        "if [ -e __vbash_missing_path__ ] || [ -n ok ]; then echo fallback; fi"), _
        "if accepts bracket test with or-if"
    AssertEqual "chain", session.ProcessInput( _
        "if test -e __vbash_missing_path__ && test -d . || [ -n ok ]; then echo chain; fi"), _
        "if evaluates and-or operators left to right"
    AssertEqual "short-circuit", session.ProcessInput( _
        "if [ -d . ] || [ -q x ]; then echo short-circuit; fi"), _
        "or-if skips a true condition's right side"
    AssertEqual "short-circuit", session.ProcessInput( _
        "if [ -e __vbash_missing_path__ ] && [ -q x ]; then echo no; else echo short-circuit; fi"), _
        "and-if skips a false condition's right side"
    AssertEqual session.CurrentDirectory & vbCrLf & "command", session.ProcessInput( _
        "if pwd; then echo command; fi"), "if evaluates a command and preserves its output"
    AssertTrue Right$(session.ProcessInput("if ls .; then echo command; fi"), _
        Len(vbCrLf & "command")) = vbCrLf & "command", _
        "if accepts ls as a command condition"
    AssertEqual "fallback", session.ProcessInput( _
        "if command-does-not-exist; then echo no; else echo fallback; fi"), _
        "if selects else for a command with nonzero status"
    AssertTrue session.HasErrorOutput, "if preserves condition stderr"
    AssertEqual "Command not found: command-does-not-exist", session.ErrorOutput, _
        "if exposes command failure on stderr"
    AssertEqual "ok", session.ProcessInput( _
        "if ls . >&-; then echo ok; else echo no; fi"), _
        "if preserves command success when stdout is closed"
    AssertEqual "ok", session.ProcessInput( _
        "if ls . > /dev/null; then echo ok; else echo no; fi"), _
        "if redirects output to the null device without changing status"

    session.ProcessInput "if [ -d .; then echo invalid; fi"
    AssertEqual 2, session.LastExitCode, "if rejects missing closing bracket"
    session.ProcessInput "if test -d . &&; then echo invalid; fi"
    AssertEqual 2, session.LastExitCode, "if rejects an incomplete logical condition"

    session.ProcessInput "[ -d . ]"
    AssertEqual 0, session.LastExitCode, "bracket test builtin status"
End Sub

Private Sub TestErrorsAndExitStatus()
    Dim session As clsSession

    Set session = New clsSession
    session.ProcessInput "command-does-not-exist"
    AssertEqual 127, session.LastExitCode, "unknown command status"
    AssertTrue session.HasErrorOutput, "unknown command reports stderr"
    AssertEqual "Command not found: command-does-not-exist", session.ErrorOutput, _
        "unknown command error text"

    session.ProcessInput "pwd extra"
    AssertEqual 2, session.LastExitCode, "invalid builtin usage status"
    AssertTrue session.HasErrorOutput, "invalid builtin usage reports stderr"

    session.ProcessInput "echo ""unterminated"
    AssertEqual 2, session.LastExitCode, "lexer error status"

    session.ProcessInput "echo hi | echo"
    AssertEqual 2, session.LastExitCode, "unsupported operator status"

    session.ProcessInput "echo hi >"
    AssertEqual 2, session.LastExitCode, "missing redirection target status"

    session.ProcessInput "echo hi < input.txt"
    AssertEqual 2, session.LastExitCode, "unsupported input redirection status"

    session.ProcessInput "ls . >&-"
    AssertEqual 0, session.LastExitCode, "close-output redirection preserves command status"
    AssertTrue Not session.HasOutput, "close-output redirection discards command output"

    session.ProcessInput "echo null-device > /dev/null"
    AssertEqual 0, session.LastExitCode, "null-device redirection succeeds"
    AssertTrue Not session.HasOutput, "null-device redirection suppresses output"

    Set session = New clsSession
    session.ProcessInput "exit 5"
    AssertTrue session.ShouldExit, "exit marks session terminated"
    AssertEqual 5, session.LastExitCode, "exit status argument"

    Set session = New clsSession
    session.ProcessInput "exit 256"
    AssertEqual 2, session.LastExitCode, "invalid exit status"
    AssertTrue Not session.ShouldExit, "invalid exit does not terminate"
End Sub

Private Sub TestRedirection()
    Dim session As clsSession
    Dim fileSystem As Object
    Dim path As String
    Dim errorPath As String
    Dim mergePath As String
    Dim textStream As Object
    Dim contents As String

    Set session = New clsSession
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    path = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)
    errorPath = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)
    mergePath = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)

    session.ProcessInput "echo first > """ & path & """"
    AssertEqual 0, session.LastExitCode, "overwrite redirection status"
    AssertTrue Not session.HasOutput, "redirected output not sent to console"
    Set textStream = fileSystem.OpenTextFile(path, 1)
    contents = textStream.ReadAll
    textStream.Close
    AssertEqual "first" & vbCrLf, contents, "overwrite redirection content"

    session.ProcessInput "echo second >> """ & path & """"
    Set textStream = fileSystem.OpenTextFile(path, 1)
    contents = textStream.ReadAll
    textStream.Close
    AssertEqual "first" & vbCrLf & "second" & vbCrLf, contents, "append redirection content"

    fileSystem.DeleteFile path, True

    AssertEqual session.CurrentDirectory, session.ProcessInput("pwd 2> """ & errorPath & """"), _
        "stderr redirection preserves successful stdout"
    AssertEqual 0, session.LastExitCode, "successful command status with stderr redirection"
    AssertTrue session.HasOutput, "successful command keeps stdout visible"
    AssertTrue Not session.HasErrorOutput, "successful command emits no stderr"
    Set textStream = fileSystem.OpenTextFile(errorPath, 1)
    contents = textStream.ReadAll
    textStream.Close
    AssertEqual "", contents, "successful stderr redirection creates an empty file"
    AssertEqual session.CurrentDirectory, session.ProcessInput("pwd 2>&1"), _
        "stderr-to-stdout leaves successful stdout visible"
    AssertTrue session.HasOutput, "successful merged stdout remains visible"
    AssertTrue Not session.HasErrorOutput, "successful merged command emits no stderr"

    session.ProcessInput "pwd extra 2> """ & errorPath & """"
    AssertEqual 2, session.LastExitCode, "failed command status with stderr redirection"
    AssertTrue Not session.HasOutput, "failed pwd produces no stdout"
    AssertTrue Not session.HasErrorOutput, "stderr redirection suppresses terminal error"
    Set textStream = fileSystem.OpenTextFile(errorPath, 1)
    contents = textStream.ReadAll
    textStream.Close
    AssertEqual "VBAsh: Uso: pwd" & vbCrLf, contents, "stderr overwrite content"

    session.ProcessInput "pwd extra 2>> """ & errorPath & """"
    Set textStream = fileSystem.OpenTextFile(errorPath, 1)
    contents = textStream.ReadAll
    textStream.Close
    AssertEqual "VBAsh: Uso: pwd" & vbCrLf & "VBAsh: Uso: pwd" & vbCrLf, _
        contents, "stderr append content"

    session.ProcessInput "pwd extra > """ & mergePath & """ 2>&1"
    AssertEqual 2, session.LastExitCode, "merged stderr preserves failed command status"
    AssertTrue Not session.HasErrorOutput, "2>&1 routes stderr away from terminal"
    Set textStream = fileSystem.OpenTextFile(mergePath, 1)
    contents = textStream.ReadAll
    textStream.Close
    AssertEqual "VBAsh: Uso: pwd" & vbCrLf, contents, "stdout then stderr merge content"

    session.ProcessInput "pwd extra 2>&1 > """ & mergePath & """"
    AssertEqual 2, session.LastExitCode, "ordered redirection preserves failed command status"
    AssertTrue session.HasErrorOutput, "stderr-to-stdout uses the stdout destination at that point"
    AssertEqual "VBAsh: Uso: pwd", session.ErrorOutput, "ordered stderr remains on console"
    Set textStream = fileSystem.OpenTextFile(mergePath, 1)
    contents = textStream.ReadAll
    textStream.Close
    AssertEqual "", contents, "later stdout redirection does not move earlier stderr"

    session.ProcessInput "missing-command 2> """ & errorPath & """"
    AssertEqual 127, session.LastExitCode, "unknown command status survives stderr redirection"
    Set textStream = fileSystem.OpenTextFile(errorPath, 1)
    contents = textStream.ReadAll
    textStream.Close
    AssertEqual "Command not found: missing-command" & vbCrLf, _
        contents, "unknown command message is stderr"
    session.ProcessInput "pwd extra 2>&-"
    AssertEqual 2, session.LastExitCode, "discarding stderr preserves the command status"
    AssertTrue Not session.HasErrorOutput, "2>&- discards stderr"

    fileSystem.DeleteFile errorPath, True
    fileSystem.DeleteFile mergePath, True
End Sub

Private Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal description As String)
    mAssertions = mAssertions + 1
    If CStr(expected) <> CStr(actual) Then
        Err.Raise vbObjectError + 1100, "VBAshTests", _
            description & ": expected [" & CStr(expected) & "], got [" & CStr(actual) & "]"
    End If
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal description As String)
    mAssertions = mAssertions + 1
    If Not condition Then Err.Raise vbObjectError + 1101, "VBAshTests", description
End Sub
