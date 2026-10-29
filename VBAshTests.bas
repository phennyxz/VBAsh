Option Explicit

Private mAssertions As Long
Private mExecutedTests As Long
Private mPassedTests As Long
Private mFailedTests As Long
Private mErrorTests As Long
Private mTestResults As Collection

Private Const ASSERT_EQUAL_ERROR As Long = vbObjectError + 1100
Private Const ASSERT_TRUE_ERROR As Long = vbObjectError + 1101
Private Const EXPECTED_TEST_COUNT As Long = 13

Public Sub VBAsh_RunTests()
    Dim startedAt As Date
    Dim finishedAt As Date
    Dim reportText As String
    Dim reportPath As String
    Dim reportError As String
    Dim summary As String
    Dim message As String

    mAssertions = 0
    mExecutedTests = 0
    mPassedTests = 0
    mFailedTests = 0
    mErrorTests = 0
    Set mTestResults = New Collection
    startedAt = Now

    RunTest "Tokenizer"
    RunTest "TokenizerWordSegments"
    RunTest "ParserAst"
    RunTest "AstExecution"
    RunTest "Echo"
    RunTest "Environment"
    RunTest "Directory"
    RunTest "Filesystem"
    RunTest "Clear"
    RunTest "Conditionals"
    RunTest "ForLoops"
    RunTest "ErrorsAndExitStatus"
    RunTest "Redirection"

    finishedAt = Now
    summary = TestRunSummary()
    reportText = BuildTestReport(startedAt, finishedAt, summary)
    If SaveTestReport(reportText, reportPath, reportError) Then
        message = summary & vbCrLf & vbCrLf & "Relatório: " & reportPath
    Else
        message = summary & vbCrLf & vbCrLf & _
                  "Não foi possível salvar o relatório TXT." & vbCrLf & reportError & _
                  vbCrLf & "Alternativa: abra o workbook em uma pasta gravável e execute a suíte novamente."
    End If

    If mFailedTests + mErrorTests = 0 And mExecutedTests = EXPECTED_TEST_COUNT Then
        MsgBox message, vbInformation, "VBAsh — resultados dos testes"
    Else
        MsgBox message, vbExclamation, "VBAsh — resultados dos testes"
    End If
End Sub

Private Sub RunTest(ByVal testName As String)
    Dim startedAt As Double
    Dim elapsed As Double
    Dim assertionStart As Long
    Dim resultStatus As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorSource As String

    startedAt = Timer
    assertionStart = mAssertions
    On Error GoTo TestError

    Select Case testName
        Case "Tokenizer": TestTokenizer
        Case "TokenizerWordSegments": TestTokenizerWordSegments
        Case "ParserAst": TestParserAst
        Case "AstExecution": TestAstExecution
        Case "Echo": TestEcho
        Case "Environment": TestEnvironment
        Case "Directory": TestDirectory
        Case "Filesystem": TestFilesystem
        Case "Clear": TestClear
        Case "Conditionals": TestConditionals
        Case "ForLoops": TestForLoops
        Case "ErrorsAndExitStatus": TestErrorsAndExitStatus
        Case "Redirection": TestRedirection
        Case Else
            Err.Raise vbObjectError + 1102, "VBAshTests.RunTest", _
                      "Teste não registrado: " & testName
    End Select

    resultStatus = "PASS"
    GoTo RecordResult

TestError:
    errorNumber = Err.Number
    errorDescription = Err.Description
    errorSource = Err.Source
    If errorNumber = ASSERT_EQUAL_ERROR Or errorNumber = ASSERT_TRUE_ERROR Then
        resultStatus = "FAIL"
    Else
        resultStatus = "ERROR"
    End If

RecordResult:
    elapsed = Timer - startedAt
    If elapsed < 0 Then elapsed = elapsed + 86400
    mExecutedTests = mExecutedTests + 1
    Select Case resultStatus
        Case "PASS": mPassedTests = mPassedTests + 1
        Case "FAIL": mFailedTests = mFailedTests + 1
        Case "ERROR": mErrorTests = mErrorTests + 1
    End Select

    If resultStatus = "PASS" Then
        mTestResults.Add resultStatus & " | " & testName & _
                         " | assertions: " & CStr(mAssertions - assertionStart) & _
                         " | duration: " & Format$(elapsed, "0.000") & " s"
    Else
        mTestResults.Add resultStatus & " | " & testName & _
                         " | error " & CStr(errorNumber) & ": " & errorDescription & _
                         " | source: " & errorSource & _
                         " | assertions: " & CStr(mAssertions - assertionStart) & _
                         " | duration: " & Format$(elapsed, "0.000") & " s"
    End If
End Sub

Private Function TestRunSummary() As String
    Dim completion As String

    If mExecutedTests <> EXPECTED_TEST_COUNT Then
        completion = "INCOMPLETO"
    ElseIf mFailedTests + mErrorTests > 0 Then
        completion = "FALHAS"
    Else
        completion = "COMPLETO"
    End If

    TestRunSummary = "Estado: " & completion & vbCrLf & _
                     "Testes executados: " & CStr(mExecutedTests) & "/" & _
                     CStr(EXPECTED_TEST_COUNT) & vbCrLf & _
                     "Aprovados (PASS): " & CStr(mPassedTests) & vbCrLf & _
                     "Falhas de asserção (FAIL): " & CStr(mFailedTests) & vbCrLf & _
                     "Erros inesperados (ERROR): " & CStr(mErrorTests) & vbCrLf & _
                     "Asserções executadas: " & CStr(mAssertions)
End Function

Private Function BuildTestReport(ByVal startedAt As Date, ByVal finishedAt As Date, _
                                 ByVal summary As String) As String
    Dim result As Variant
    Dim report As String

    report = "VBAsh — relatório de testes" & vbCrLf & _
             "Início: " & Format$(startedAt, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
             "Término: " & Format$(finishedAt, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
             summary & vbCrLf & vbCrLf & _
             "Resultados individuais:" & vbCrLf
    For Each result In mTestResults
        report = report & CStr(result) & vbCrLf
    Next result
    BuildTestReport = report
End Function

Private Function SaveTestReport(ByVal reportText As String, ByRef reportPath As String, _
                                ByRef failureDetails As String) As Boolean
    Dim fileSystem As Object
    Dim folders As Collection
    Dim folder As Variant
    Dim folderError As String

    On Error GoTo SaveError
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folders = New Collection
    If Len(Environ$("APPDATA")) > 0 Then folders.Add Environ$("APPDATA")
    If Len(Environ$("TEMP")) > 0 Then folders.Add Environ$("TEMP")
    If Len(ThisWorkbook.Path) > 0 Then folders.Add ThisWorkbook.Path

    For Each folder In folders
        folderError = ""
        If TryWriteTestReport(fileSystem, CStr(folder), reportText, _
                              reportPath, folderError) Then
            SaveTestReport = True
            Exit Function
        End If
        If Len(failureDetails) > 0 Then failureDetails = failureDetails & vbCrLf
        failureDetails = failureDetails & CStr(folder) & ": " & folderError
    Next folder

    If folders.Count = 0 Then
        failureDetails = "APPDATA, TEMP e a pasta do workbook não estão disponíveis."
    End If
    Exit Function

SaveError:
    failureDetails = "Erro " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function TryWriteTestReport(ByVal fileSystem As Object, ByVal baseFolder As String, _
                                    ByVal reportText As String, ByRef reportPath As String, _
                                    ByRef failureDescription As String) As Boolean
    Dim reportFolder As String
    Dim stream As Object

    On Error GoTo WriteError
    reportFolder = fileSystem.BuildPath(baseFolder, "VBAsh")
    EnsureReportFolder fileSystem, reportFolder
    reportFolder = fileSystem.BuildPath(reportFolder, "TestReports")
    EnsureReportFolder fileSystem, reportFolder
    reportPath = fileSystem.BuildPath(reportFolder, _
                 "VBAsh_TestResults_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt")
    Set stream = fileSystem.CreateTextFile(reportPath, True, True)
    stream.Write reportText
    stream.Close
    Set stream = Nothing
    TryWriteTestReport = True
    Exit Function

WriteError:
    failureDescription = "Erro " & CStr(Err.Number) & ": " & Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    On Error GoTo 0
End Function

Private Sub EnsureReportFolder(ByVal fileSystem As Object, ByVal folderPath As String)
    Dim parentFolder As String

    If fileSystem.FolderExists(folderPath) Then Exit Sub
    parentFolder = fileSystem.GetParentFolderName(folderPath)
    If Len(parentFolder) = 0 Or StrComp(parentFolder, folderPath, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1103, "VBAshTests.EnsureReportFolder", _
                  "Não foi possível determinar a pasta pai de " & folderPath
    End If
    If Not fileSystem.FolderExists(parentFolder) Then
        EnsureReportFolder fileSystem, parentFolder
    End If
    fileSystem.CreateFolder folderPath
End Sub

Private Sub TestParserAst()
    Dim tokenizer As clsTokenizer
    Dim parser As clsParser
    Dim tokens As Collection
    Dim node As clsCommand
    Dim child As clsCommand
    Dim token As clsToken
    Dim sourceToken As clsToken
    Dim wordSegment As clsWordSegment
    Dim commandLine As String
    Dim legacyCommand As clsCommand

    Set tokenizer = New clsTokenizer
    Set parser = New clsParser

    commandLine = "echo ab""cd ef""gh"
    Set tokens = tokenizer.Tokenize(commandLine)
    Set node = parser.ParseAst(tokens, commandLine)
    AssertEqual "SIMPLE_COMMAND", node.NodeType, "simple command AST node"
    AssertEqual 2, node.Words.Count, "simple command preserves all word tokens"
    AssertEqual "echo", node.Name, "legacy command name remains available"
    AssertEqual 1, node.Args.Count, "legacy arguments remain available"
    Set token = node.Words(2)
    Set sourceToken = tokens(2)
    AssertTrue token Is sourceToken, "AST retains original word-token object"
    AssertEqual "abcd efgh", token.Value, "AST token keeps flattened compatibility value"
    AssertEqual 3, token.Segments.Count, "AST token retains lexical segments"
    AssertTrue Not token.BraceExpansionAllowed, "AST retains brace-expansion metadata"
    AssertEqual sourceToken.SourceStart, token.SourceStart, "AST retains original word start"
    AssertEqual sourceToken.SourceEnd, token.SourceEnd, "AST retains original word end"
    Set wordSegment = token.Segments(2)
    AssertEqual "DOUBLE", wordSegment.QuoteContext, "AST retains quoted segment context"
    AssertEqual token.SourceStart, node.SourceStart + 5, "simple node starts at first token"
    AssertEqual token.SourceEnd, node.SourceEnd, "simple node ends at final word"
    AssertEqual commandLine, _
        Mid$(commandLine, node.SourceStart, node.SourceEnd - node.SourceStart + 1), _
        "simple node span covers the complete command"

    Set node = parser.ParseAst(tokenizer.Tokenize("echo '$NOME'"), "echo '$NOME'")
    Set token = node.Words(2)
    AssertEqual "00000", token.ExpansionMask, "AST retains expansion mask"
    Set wordSegment = token.Segments(1)
    AssertEqual "SINGLE", wordSegment.QuoteContext, "AST retains single-quote metadata"
    AssertEqual "$NOME", wordSegment.Text, "AST retains source word segment text"

    Set legacyCommand = parser.Parse(tokenizer.Tokenize("echo hi"), "echo hi")
    AssertEqual "SIMPLE_COMMAND", legacyCommand.NodeType, "legacy Parse returns simple AST node"
    AssertEqual "echo", legacyCommand.Name, "legacy Parse preserves command name"
    AssertEqual "hi", legacyCommand.Args(1), "legacy Parse preserves arguments"

    Set node = parser.ParseAst(tokenizer.Tokenize("echo a; pwd; echo b"), "echo a; pwd; echo b")
    AssertEqual "SEQUENCE", node.NodeType, "semicolon creates sequence node"
    AssertEqual 3, node.Children.Count, "sequence preserves all commands"
    Set child = node.Children(1)
    AssertEqual "echo", child.Name, "sequence preserves first command order"
    Set child = node.Children(2)
    AssertEqual "pwd", child.Name, "sequence preserves second command order"
    Set child = node.Children(3)
    AssertEqual "echo", child.Name, "sequence preserves third command order"
    AssertEqual "b", child.Args(1), "sequence retains final command arguments"
    AssertEqual "echo a; pwd; echo b", _
        Mid$("echo a; pwd; echo b", node.SourceStart, node.SourceEnd - node.SourceStart + 1), _
        "sequence span covers all child commands"

    Set node = parser.ParseAst(tokenizer.Tokenize("a && b"), "a && b")
    AssertEqual "AND_IF", node.NodeType, "and-if node type"
    AssertEqual 2, node.Children.Count, "and-if has two operands"
    AssertEqual "a", node.Children(1).Name, "and-if left operand"
    AssertEqual "b", node.Children(2).Name, "and-if right operand"

    Set node = parser.ParseAst(tokenizer.Tokenize("a || b"), "a || b")
    AssertEqual "OR_IF", node.NodeType, "or-if node type"
    AssertEqual "a", node.Children(1).Name, "or-if left operand"
    AssertEqual "b", node.Children(2).Name, "or-if right operand"

    Set node = parser.ParseAst(tokenizer.Tokenize("a && b || c"), "a && b || c")
    AssertEqual "OR_IF", node.NodeType, "and-or operators share precedence"
    Set child = node.Children(1)
    AssertEqual "AND_IF", child.NodeType, "and-or operators associate left"
    AssertEqual "c", node.Children(2).Name, "or-if has final right operand"

    Set node = parser.ParseAst(tokenizer.Tokenize("a || b && c"), "a || b && c")
    AssertEqual "AND_IF", node.NodeType, "or-and operators share precedence"
    Set child = node.Children(1)
    AssertEqual "OR_IF", child.NodeType, "or-and operators associate left"
    AssertEqual "c", node.Children(2).Name, "and-if has final right operand"

    Set node = parser.ParseAst(tokenizer.Tokenize("a && b && c"), "a && b && c")
    AssertEqual "AND_IF", node.NodeType, "repeated and-if root"
    AssertEqual "AND_IF", node.Children(1).NodeType, "and-if is left associative"
    Set node = parser.ParseAst(tokenizer.Tokenize("a || b || c"), "a || b || c")
    AssertEqual "OR_IF", node.NodeType, "repeated or-if root"
    AssertEqual "OR_IF", node.Children(1).NodeType, "or-if is left associative"

    Set node = parser.ParseAst(tokenizer.Tokenize("a && b; c || d"), "a && b; c || d")
    AssertEqual "SEQUENCE", node.NodeType, "semicolon has lower precedence"
    AssertEqual "AND_IF", node.Children(1).NodeType, "first sequence item keeps logical tree"
    AssertEqual "OR_IF", node.Children(2).NodeType, "second sequence item keeps logical tree"

    Set node = parser.ParseAst(tokenizer.Tokenize("echo hi > out"), "echo hi > out")
    AssertEqual "SIMPLE_COMMAND", node.NodeType, "redirection remains part of simple command"
    AssertEqual 1, node.Redirections.Count, "existing redirection parsing is preserved"
    AssertEqual "REDIR_OUT", node.Redirections(1).Operator, "redirection operator preserved"
    AssertEqual "out", node.Redirections(1).Target, "redirection target preserved"

    AssertAstSyntaxError tokenizer, parser, "", "empty input is rejected"
    AssertAstSyntaxError tokenizer, parser, "; a", "leading sequence separator is rejected"
    AssertAstSyntaxError tokenizer, parser, "a ;; b", "consecutive sequence separators are rejected"
    AssertAstSyntaxError tokenizer, parser, "a &&", "incomplete and-if is rejected"
    AssertAstSyntaxError tokenizer, parser, "a ||", "incomplete or-if is rejected"
    AssertAstSyntaxError tokenizer, parser, "a && || b", "consecutive logical operators are rejected"
    AssertAstSyntaxError tokenizer, parser, "a | b", "out-of-scope pipeline operator is rejected"
    AssertLegacyParseError tokenizer, parser, "echo a;", "legacy Parse continues rejecting separators"
    AssertLegacyParseError tokenizer, parser, "a && b", "legacy Parse continues rejecting logical lists"

End Sub

Private Sub AssertAstSyntaxError(ByVal tokenizer As clsTokenizer, ByVal parser As clsParser, _
                                 ByVal commandLine As String, ByVal description As String)
    Dim tokens As Collection
    Dim node As clsCommand
    Dim errorNumber As Long
    Dim errorSource As String

    Set tokens = tokenizer.Tokenize(commandLine)
    On Error GoTo ParseError
    Set node = parser.ParseAst(tokens, commandLine)
    On Error GoTo 0
    AssertTrue False, description
    Exit Sub

ParseError:
    errorNumber = Err.Number
    errorSource = Err.Source
    Err.Clear
    On Error GoTo 0
    AssertTrue StrComp(errorSource, "clsParser.ParseAst", vbBinaryCompare) = 0 And _
               IsParserSyntaxErrorNumber(errorNumber), _
        description & " (unexpected error from " & errorSource & ": " & CStr(errorNumber) & ")"
End Sub

Private Sub AssertLegacyParseError(ByVal tokenizer As clsTokenizer, ByVal parser As clsParser, _
                                   ByVal commandLine As String, ByVal description As String)
    Dim tokens As Collection
    Dim command As clsCommand
    Dim errorNumber As Long
    Dim errorSource As String

    Set tokens = tokenizer.Tokenize(commandLine)
    On Error GoTo ParseError
    Set command = parser.Parse(tokens, commandLine)
    On Error GoTo 0
    AssertTrue False, description
    Exit Sub

ParseError:
    errorNumber = Err.Number
    errorSource = Err.Source
    Err.Clear
    On Error GoTo 0
    AssertTrue StrComp(errorSource, "clsParser.Parse", vbBinaryCompare) = 0 And _
               errorNumber = vbObjectError + 1001, _
        description & " (unexpected error from " & errorSource & ": " & CStr(errorNumber) & ")"
End Sub

Private Function IsParserSyntaxErrorNumber(ByVal errorNumber As Long) As Boolean
    Select Case errorNumber
        Case vbObjectError + 1001, vbObjectError + 1002, vbObjectError + 1003, _
             vbObjectError + 1005, vbObjectError + 1083
            IsParserSyntaxErrorNumber = True
    End Select
End Function

Private Sub TestTokenizer()
    Dim tokenizer As clsTokenizer
    Dim tokens As Collection
    Dim token As clsToken

    Set tokenizer = New clsTokenizer
    Set tokens = tokenizer.Tokenize("echo hello")
    AssertEqual 3, tokens.Count, "plain command token count includes EOF"
    Set token = tokens(2)
    AssertEqual "hello", token.Value, "plain command word"

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

Private Sub TestTokenizerWordSegments()
    Dim tokenizer As clsTokenizer
    Dim tokens As Collection
    Dim token As clsToken
    Dim segment As clsWordSegment
    Dim commandLine As String

    Set tokenizer = New clsTokenizer

    Set tokens = tokenizer.Tokenize("echo '$NOME'")
    Set token = tokens(2)
    AssertEqual 1, token.Segments.Count, "single-quoted word has one lexical segment"
    AssertEqual "00000", token.ExpansionMask, "single-quoted variable text is non-expandable"
    Set segment = token.Segments(1)
    AssertEqual "SINGLE", segment.QuoteContext, "single-quoted context is preserved"
    AssertEqual "$NOME", segment.Text, "single-quoted text is preserved"
    AssertTrue Not segment.IsEscaped, "single-quoted text is not marked escaped"

    Set tokens = tokenizer.Tokenize("echo ""$NOME""")
    Set token = tokens(2)
    AssertEqual "11111", token.ExpansionMask, "double-quoted variable text remains expandable"
    Set segment = token.Segments(1)
    AssertEqual "DOUBLE", segment.QuoteContext, "double-quoted context is preserved"
    AssertEqual "$NOME", segment.Text, "double-quoted text is preserved"
    AssertTrue Not segment.IsEscaped, "unescaped double-quoted text is not marked escaped"

    Set tokens = tokenizer.Tokenize("echo ""texto com espaços""")
    Set token = tokens(2)
    Set segment = token.Segments(1)
    AssertEqual "texto com espaços", segment.Text, "double-quoted spaces remain in one segment"
    AssertEqual "DOUBLE", segment.QuoteContext, "spaced text retains double-quote context"

    Set tokens = tokenizer.Tokenize("echo ""\$NOME""")
    Set token = tokens(2)
    AssertEqual "01111", token.ExpansionMask, "escaped dollar remains non-expandable"
    AssertEqual 2, token.Segments.Count, "escaped character splits its quoted segment"
    Set segment = token.Segments(1)
    AssertEqual "DOUBLE", segment.QuoteContext, "escaped dollar retains double-quote context"
    AssertTrue segment.IsEscaped, "escaped dollar is marked escaped"

    commandLine = "echo ab""cd ef""gh"
    Set tokens = tokenizer.Tokenize(commandLine)
    Set token = tokens(2)
    AssertEqual "abcd efgh", token.Value, "adjacent quoted fragments remain one word"
    AssertEqual 3, token.Segments.Count, "adjacent fragments retain their boundaries"
    Set segment = token.Segments(1)
    AssertEqual "ab", segment.Text, "first unquoted fragment is retained"
    AssertEqual "UNQUOTED", segment.QuoteContext, "first fragment context"
    Set segment = token.Segments(2)
    AssertEqual "cd ef", segment.Text, "quoted fragment containing spaces is retained"
    AssertEqual "DOUBLE", segment.QuoteContext, "middle fragment context"
    Set segment = token.Segments(3)
    AssertEqual "gh", segment.Text, "last unquoted fragment is retained"
    AssertEqual "UNQUOTED", segment.QuoteContext, "last fragment context"
    AssertEqual "ab""cd ef""gh", _
        Mid$(commandLine, token.SourceStart, token.SourceEnd - token.SourceStart + 1), _
        "word source span includes its quote syntax"

    Set tokens = tokenizer.Tokenize("echo a\ b")
    Set token = tokens(2)
    AssertEqual "a b", token.Value, "escaped space remains within a word"
    AssertEqual 3, token.Segments.Count, "escaped space creates a segment boundary"
    Set segment = token.Segments(2)
    AssertEqual " ", segment.Text, "escaped segment retains the whitespace"
    AssertEqual "UNQUOTED", segment.QuoteContext, "escaped space quote context"
    AssertTrue segment.IsEscaped, "escaped space is marked escaped"
    AssertEqual 7, segment.SourceStart, "escaped segment starts at its backslash"
    AssertEqual 8, segment.SourceEnd, "escaped segment ends at its character"

    Set tokens = tokenizer.Tokenize("echo """"")
    Set token = tokens(2)
    AssertEqual "", token.Value, "empty quoted word retains its value"
    AssertEqual 1, token.Segments.Count, "empty quoted word retains a segment"
    Set segment = token.Segments(1)
    AssertEqual "", segment.Text, "empty quoted segment has empty text"
    AssertEqual "DOUBLE", segment.QuoteContext, "empty quoted segment context"

    Set tokens = tokenizer.Tokenize("echo ""a&&b|c;d>e"" x\|y; echo z")
    AssertEqual 7, tokens.Count, "quotes and escapes preserve operator boundaries"
    Set token = tokens(2)
    AssertEqual "a&&b|c;d>e", token.Value, "operators inside quotes remain word text"
    Set token = tokens(3)
    AssertEqual "x|y", token.Value, "escaped operator remains within a word"
    AssertEqual 3, token.Segments.Count, "escaped operator has its own segment"
    Set segment = token.Segments(2)
    AssertTrue segment.IsEscaped, "escaped operator metadata is preserved"
    Set token = tokens(4)
    AssertEqual "SEMI", token.TokenType, "unquoted command separator remains an operator"
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

Private Sub TestAstExecution()
    Dim session As clsSession
    Dim result As String
    Dim historyCount As Long

    Set session = New clsSession

    historyCount = session.History.Count
    result = session.ProcessInput("echo first; echo second")
    AssertEqual "first" & vbCrLf & "second", result, "sequence combines stdout in order"
    AssertEqual 0, session.LastExitCode, "sequence returns final command status"
    AssertTrue session.HasOutput, "sequence reports combined stdout"
    AssertEqual historyCount + 1, session.History.Count, "sequence is recorded once in history"

    result = session.ProcessInput("clear; echo after-clear")
    AssertEqual "after-clear", result, "command after clear keeps its stdout"
    AssertTrue Not session.ClearRequested, "later sequence command clears prior clear request"

    result = session.ProcessInput("VBAsh_AST_SEQUENCE=ready; echo $VBAsh_AST_SEQUENCE")
    AssertEqual "ready", result, "sequence shares session state between commands"
    AssertEqual "ready", session.GetEnv("VBAsh_AST_SEQUENCE"), "sequence assignment persists"

    result = session.ProcessInput("echo first; pwd extra")
    AssertEqual "first", result, "sequence retains stdout before a failing command"
    AssertEqual 2, session.LastExitCode, "sequence preserves final failing status"
    AssertTrue session.HasOutput, "sequence retains earlier stdout when final command fails"
    AssertTrue session.HasErrorOutput, "sequence retains stderr from final command"
    AssertEqual "VBAsh: Uso: pwd", session.ErrorOutput, "sequence reports builtin error"

    session.ProcessInput "pwd extra && export VBAsh_AST_AND=unexpected"
    AssertEqual 2, session.LastExitCode, "and-if keeps left failure status when right is skipped"
    AssertTrue Not session.Env.Exists("VBAsh_AST_AND"), "and-if short-circuit skips assignment"
    AssertTrue session.HasErrorOutput, "and-if preserves left stderr"

    session.ProcessInput "pwd extra && echo VBAsh_AST_AND_UNEXPECTED"
    AssertEqual 2, session.LastExitCode, "and-if skipped command preserves left status"
    AssertTrue Not session.HasOutput, "and-if short-circuit suppresses skipped stdout"

    result = session.ProcessInput("echo visible || export VBAsh_AST_OR=unexpected")
    AssertEqual "visible", result, "or-if preserves left stdout when right is skipped"
    AssertEqual 0, session.LastExitCode, "or-if keeps left success status when right is skipped"
    AssertTrue Not session.Env.Exists("VBAsh_AST_OR"), "or-if short-circuit skips assignment"

    result = session.ProcessInput("echo visible || pwd extra")
    AssertEqual "visible", result, "or-if does not execute skipped error command"
    AssertEqual 0, session.LastExitCode, "or-if skipped error command keeps left status"
    AssertTrue Not session.HasErrorOutput, "or-if short-circuit suppresses skipped stderr"

    result = session.ProcessInput("echo left && echo right")
    AssertEqual "left" & vbCrLf & "right", result, "and-if executes right after success"
    AssertEqual 0, session.LastExitCode, "and-if returns executed right status"

    session.ProcessInput "echo left && pwd extra"
    AssertEqual 2, session.LastExitCode, "and-if returns failing right status"
    AssertTrue session.HasErrorOutput, "and-if retains stderr from failing right command"

    session.ProcessInput "pwd extra || export VBAsh_AST_FALLBACK=ran"
    AssertEqual "ran", session.GetEnv("VBAsh_AST_FALLBACK"), "or-if executes right after failure"

    result = session.ProcessInput("pwd extra || echo recovered")
    AssertEqual "recovered", result, "or-if executes right after failure"
    AssertEqual 0, session.LastExitCode, "or-if returns executed right status"
    AssertTrue session.HasOutput, "or-if reports stdout from recovery command"
    AssertTrue session.HasErrorOutput, "or-if retains stderr emitted by left command"

    session.ProcessInput "pwd extra && export VBAsh_AST_NESTED_AND=bad || export VBAsh_AST_NESTED_FALLBACK=ran"
    AssertTrue Not session.Env.Exists("VBAsh_AST_NESTED_AND"), _
               "a && b || c short-circuits the middle command"
    AssertEqual "ran", session.GetEnv("VBAsh_AST_NESTED_FALLBACK"), _
                "a && b || c executes fallback according to AST associativity"
    AssertEqual 0, session.LastExitCode, "a && b || c returns final executed status"

    result = session.ProcessInput("pwd || export VBAsh_AST_NESTED_OR=bad && export VBAsh_AST_NESTED_AND=ran")
    AssertTrue Not session.Env.Exists("VBAsh_AST_NESTED_OR"), _
               "a || b && c skips the middle command after success"
    AssertEqual "ran", session.GetEnv("VBAsh_AST_NESTED_AND"), _
                "a || b && c executes final command according to left associativity"
    AssertEqual 0, session.LastExitCode, "a || b && c returns final executed status"
    AssertTrue Len(result) > 0, "nested logical expression preserves executed stdout"
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
    Dim contents As String

    Set session = New clsSession
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    path = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)
    errorPath = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)
    mergePath = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)

    session.ProcessInput "echo first > """ & path & """"
    AssertEqual 0, session.LastExitCode, "overwrite redirection status"
    AssertTrue Not session.HasOutput, "redirected output not sent to console"
    contents = ReadRedirectedFile(fileSystem, path, "first" & vbCrLf, "stdout overwrite")
    AssertEqual "first" & vbCrLf, contents, "overwrite redirection content"

    session.ProcessInput "echo second >> """ & path & """"
    contents = ReadRedirectedFile(fileSystem, path, _
        "first" & vbCrLf & "second" & vbCrLf, "stdout append")
    AssertEqual "first" & vbCrLf & "second" & vbCrLf, contents, "append redirection content"

    fileSystem.DeleteFile path, True

    AssertEqual session.CurrentDirectory, session.ProcessInput("pwd 2> """ & errorPath & """"), _
        "stderr redirection preserves successful stdout"
    AssertEqual 0, session.LastExitCode, "successful command status with stderr redirection"
    AssertTrue session.HasOutput, "successful command keeps stdout visible"
    AssertTrue Not session.HasErrorOutput, "successful command emits no stderr"
    contents = ReadRedirectedFile(fileSystem, errorPath, "", "empty stderr overwrite")
    AssertEqual "", contents, "successful stderr redirection creates an empty file"
    AssertEqual session.CurrentDirectory, session.ProcessInput("pwd 2>&1"), _
        "stderr-to-stdout leaves successful stdout visible"
    AssertTrue session.HasOutput, "successful merged stdout remains visible"
    AssertTrue Not session.HasErrorOutput, "successful merged command emits no stderr"

    session.ProcessInput "pwd extra 2> """ & errorPath & """"
    AssertEqual 2, session.LastExitCode, "failed command status with stderr redirection"
    AssertTrue Not session.HasOutput, "failed pwd produces no stdout"
    AssertTrue Not session.HasErrorOutput, "stderr redirection suppresses terminal error"
    contents = ReadRedirectedFile(fileSystem, errorPath, _
        "VBAsh: Uso: pwd" & vbCrLf, "stderr overwrite")
    AssertEqual "VBAsh: Uso: pwd" & vbCrLf, contents, "stderr overwrite content"

    session.ProcessInput "pwd extra 2>> """ & errorPath & """"
    contents = ReadRedirectedFile(fileSystem, errorPath, _
        "VBAsh: Uso: pwd" & vbCrLf & "VBAsh: Uso: pwd" & vbCrLf, "stderr append")
    AssertEqual "VBAsh: Uso: pwd" & vbCrLf & "VBAsh: Uso: pwd" & vbCrLf, _
        contents, "stderr append content"

    session.ProcessInput "pwd extra > """ & mergePath & """ 2>&1"
    AssertEqual 2, session.LastExitCode, "merged stderr preserves failed command status"
    AssertTrue Not session.HasErrorOutput, "2>&1 routes stderr away from terminal"
    contents = ReadRedirectedFile(fileSystem, mergePath, _
        "VBAsh: Uso: pwd" & vbCrLf, "stdout and stderr merge")
    AssertEqual "VBAsh: Uso: pwd" & vbCrLf, contents, "stdout then stderr merge content"

    session.ProcessInput "pwd extra 2>&1 > """ & mergePath & """"
    AssertEqual 2, session.LastExitCode, "ordered redirection preserves failed command status"
    AssertTrue session.HasErrorOutput, "stderr-to-stdout uses the stdout destination at that point"
    AssertEqual "VBAsh: Uso: pwd", session.ErrorOutput, "ordered stderr remains on console"
    contents = ReadRedirectedFile(fileSystem, mergePath, "", "ordered stderr then stdout redirect")
    AssertEqual "", contents, "later stdout redirection does not move earlier stderr"

    session.ProcessInput "missing-command 2> """ & errorPath & """"
    AssertEqual 127, session.LastExitCode, "unknown command status survives stderr redirection"
    contents = ReadRedirectedFile(fileSystem, errorPath, _
        "Command not found: missing-command" & vbCrLf, "unknown-command stderr")
    AssertEqual "Command not found: missing-command" & vbCrLf, _
        contents, "unknown command message is stderr"
    session.ProcessInput "pwd extra 2>&-"
    AssertEqual 2, session.LastExitCode, "discarding stderr preserves the command status"
    AssertTrue Not session.HasErrorOutput, "2>&- discards stderr"

    fileSystem.DeleteFile errorPath, True
    fileSystem.DeleteFile mergePath, True
End Sub

Private Function ReadRedirectedFile(ByVal fileSystem As Object, ByVal path As String, _
                                    ByVal expectedContents As String, ByVal description As String) As String
    Dim textStream As Object
    Dim file As Object
    Dim resolvedPath As String
    Dim fileSize As Double
    Dim fileExists As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    resolvedPath = fileSystem.GetAbsolutePathName(path)
    AssertEqual path, resolvedPath, description & ": resolved output path"
    fileExists = fileSystem.FileExists(resolvedPath)
    AssertTrue fileExists, _
        description & ": output file exists at " & resolvedPath
    Set file = fileSystem.GetFile(resolvedPath)
    fileSize = CDbl(file.Size)
    AssertEqual (Len(expectedContents) = 0), (fileSize = 0), _
        description & ": output file is empty at " & resolvedPath

    On Error GoTo ReadError
    Set textStream = fileSystem.OpenTextFile(resolvedPath, 1)
    AssertEqual (fileSize = 0), CBool(textStream.AtEndOfStream), _
        description & ": stream EOF state for " & resolvedPath
    If fileSize = 0 Then
        ReadRedirectedFile = ""
        textStream.Close
        Set textStream = Nothing
        Exit Function
    End If

    ReadRedirectedFile = textStream.ReadAll
    textStream.Close
    Set textStream = Nothing
    Exit Function

ReadError:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    If Not textStream Is Nothing Then textStream.Close
    Set textStream = Nothing
    On Error GoTo 0
    Err.Raise errorNumber, "TestRedirection.ReadRedirectedFile", _
        errorDescription & " (case: " & description & "; path: " & resolvedPath & _
        "; exists: " & CStr(fileExists) & _
        "; size: " & CStr(fileSize) & " bytes; original source: " & errorSource & ")"
End Function

Private Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal description As String)
    mAssertions = mAssertions + 1
    If CStr(expected) <> CStr(actual) Then
        Err.Raise ASSERT_EQUAL_ERROR, "VBAshTests", _
            description & ": expected [" & CStr(expected) & "], got [" & CStr(actual) & "]"
    End If
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal description As String)
    mAssertions = mAssertions + 1
    If Not condition Then Err.Raise ASSERT_TRUE_ERROR, "VBAshTests", description
End Sub
