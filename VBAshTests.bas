Option Explicit

Private mAssertions As Long

Public Sub VBAsh_RunTests()
    mAssertions = 0
    TestTokenizer
    TestEcho
    TestEnvironment
    TestDirectory
    TestConditionals
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
    AssertEqual originalDirectory & "> ", session.Prompt, "cd updates prompt"

    AssertEqual originalDirectory, session.ProcessInput("cd -"), "cd - prints previous directory"
    AssertEqual originalDirectory, session.CurrentDirectory, "cd - restores directory"
    AssertEqual 0, session.LastExitCode, "cd exit status"
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

    session.ProcessInput "if test -d . then echo invalid; fi"
    AssertEqual 2, session.LastExitCode, "if rejects missing separator"

    session.ProcessInput "test -d ."
    AssertEqual 0, session.LastExitCode, "test true status"
    session.ProcessInput "test -e __vbash_missing_path__"
    AssertEqual 1, session.LastExitCode, "test false status"
    session.ProcessInput "test ""same value"" = ""same value"""
    AssertEqual 0, session.LastExitCode, "test compares quoted strings"
    AssertEqual 6, session.History.Count, "conditional internals do not duplicate history"
End Sub

Private Sub TestErrorsAndExitStatus()
    Dim session As clsSession

    Set session = New clsSession
    session.ProcessInput "command-does-not-exist"
    AssertEqual 127, session.LastExitCode, "unknown command status"
    AssertTrue session.HasOutput, "unknown command reports error"

    session.ProcessInput "pwd extra"
    AssertEqual 2, session.LastExitCode, "invalid builtin usage status"

    session.ProcessInput "echo ""unterminated"
    AssertEqual 2, session.LastExitCode, "lexer error status"

    session.ProcessInput "echo hi | echo"
    AssertEqual 2, session.LastExitCode, "unsupported operator status"

    session.ProcessInput "echo hi >"
    AssertEqual 2, session.LastExitCode, "missing redirection target status"

    session.ProcessInput "echo hi < input.txt"
    AssertEqual 2, session.LastExitCode, "unsupported input redirection status"

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
    Dim textStream As Object
    Dim contents As String

    Set session = New clsSession
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    path = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)

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
