' @VBAshRevision VBAshTests-r7
Option Explicit

Private mAssertions As Long
Private mExecutedTests As Long
Private mPassedTests As Long
Private mFailedTests As Long
Private mErrorTests As Long
Private mTestResults As Collection
Private mFailureDetails As Collection
Private mSuiteStart As Long
Private mLastAssertNumber As Long
Private mLastAssertDescription As String
Private mFailExpected As String
Private mFailActual As String
Private Const RUNNER_REVISION As String = "VBAshTests-r7"

Private Const ASSERT_EQUAL_ERROR As Long = vbObjectError + 1100
Private Const ASSERT_TRUE_ERROR As Long = vbObjectError + 1101
Private Const EXPECTED_TEST_COUNT As Long = 23

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
    Set mFailureDetails = New Collection
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
    RunTest "StreamContract"
    RunTest "MemoryStream"
    RunTest "NullStream"
    RunTest "FileStream"
    RunTest "FileDescriptorTable"
    RunTest "FileDescriptorTableCloseAll"
    RunTest "DescriptorScopes"
    RunTest "DescriptorRedirection"
    RunTest "ConsoleStream"
    RunTest "PipeStream"

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
    mSuiteStart = mAssertions
    mLastAssertNumber = 0
    mLastAssertDescription = ""
    mFailExpected = ""
    mFailActual = ""
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
        Case "StreamContract": TestStreamContract
        Case "MemoryStream": TestMemoryStream
        Case "NullStream": TestNullStream
        Case "FileStream": TestFileStream
        Case "FileDescriptorTable": TestFileDescriptorTable
        Case "FileDescriptorTableCloseAll": TestFileDescriptorTableCloseAll
        Case "DescriptorScopes": TestDescriptorScopes
        Case "DescriptorRedirection": TestDescriptorRedirection
        Case "ConsoleStream": TestConsoleStream
        Case "PipeStream": TestPipeStream
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

    mTestResults.Add resultStatus & " | " & testName & _
                     " | suite revision: " & RUNNER_REVISION & _
                     " | assertions executed: " & CStr(mAssertions - assertionStart) & _
                     " | duration: " & Format$(elapsed, "0.000") & " s"

    If resultStatus = "FAIL" Then
        mFailureDetails.Add "FAIL | suite: " & testName & _
                            " | assertion #" & CStr(mLastAssertNumber) & _
                            " (of " & CStr(mAssertions - assertionStart) & " executed; numbering valid only for " & RUNNER_REVISION & ")" & _
                            " | description: " & mLastAssertDescription & _
                            " | expected: [" & mFailExpected & "]" & _
                            " | actual: [" & mFailActual & "]" & _
                            " | error " & CStr(errorNumber) & ": " & errorDescription & _
                            " | source: " & errorSource & _
                            " | stage: " & mLastAssertDescription
    ElseIf resultStatus = "ERROR" Then
        mFailureDetails.Add "ERROR | suite: " & testName & _
                            " | last assertion passed: #" & CStr(mLastAssertNumber) & _
                            " (" & mLastAssertDescription & ")" & _
                            " | error " & CStr(errorNumber) & ": " & errorDescription & _
                            " | source: " & errorSource & _
                            " | stage: after assertion #" & CStr(mLastAssertNumber) & _
                            " (precise stage not recorded)"
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
    Dim warnings As String
    Dim moduleReport As String
    Dim checkoutText As String
    Dim unverifiedCount As Long
    Dim mismatchCount As Long

    moduleReport = LoadedModuleReport(warnings, unverifiedCount, mismatchCount)
    checkoutText = CheckoutReport(warnings)

    report = "VBAsh TEST REPORT" & vbCrLf & vbCrLf & _
             "Execution:" & vbCrLf & _
             "- Date: " & Format$(startedAt, "yyyy-mm-dd") & vbCrLf & _
             "- Time: " & Format$(startedAt, "hh:nn:ss") & " (end " & _
                          Format$(finishedAt, "hh:nn:ss") & ")" & vbCrLf & _
             "- Runner version: " & RUNNER_REVISION & vbCrLf & _
             "- Workbook: " & WorkbookIdentity() & vbCrLf & vbCrLf & _
             "Checkout (source tree, NOT proof of loaded modules):" & vbCrLf & _
             checkoutText & vbCrLf & _
             "Loaded module versions (read from the running VBA project):" & vbCrLf & _
             moduleReport & vbCrLf & _
             "Results:" & vbCrLf & summary & vbCrLf & vbCrLf & _
             "Suite details:" & vbCrLf
    For Each result In mTestResults
        report = report & "- " & CStr(result) & vbCrLf
    Next result

    report = report & vbCrLf & "Failure details:" & vbCrLf
    If mFailureDetails.Count = 0 Then
        report = report & "- none" & vbCrLf
    Else
        For Each result In mFailureDetails
            report = report & "- " & CStr(result) & vbCrLf
        Next result
    End If

    report = report & vbCrLf & "Execution integrity:" & vbCrLf & _
             "- Checkout/module correspondence: UNVERIFIED (the runner never assumes the " & _
             "Git checkout equals the modules loaded in Excel; compare revisions manually)" & vbCrLf
    If mismatchCount > 0 Then
        report = report & "- Version identification status: MISMATCH (" & CStr(mismatchCount) & _
                 " module(s) report conflicting revisions)" & vbCrLf
    ElseIf unverifiedCount > 0 Then
        report = report & "- Version identification status: PARTIAL (" & CStr(unverifiedCount) & _
                 " module revision(s) UNVERIFIED)" & vbCrLf
    Else
        report = report & "- Version identification status: IDENTIFIED (revision tags read from loaded modules)" & vbCrLf
    End If
    If Len(warnings) = 0 Then
        report = report & "- Warnings: none" & vbCrLf
    Else
        report = report & "- Warnings:" & vbCrLf & warnings
    End If
    BuildTestReport = report
End Function

Private Function WorkbookIdentity() As String
    On Error Resume Next
    WorkbookIdentity = ThisWorkbook.FullName
    If Err.Number <> 0 Or Len(WorkbookIdentity) = 0 Then WorkbookIdentity = "UNKNOWN"
End Function

' Reads .git/HEAD above the workbook folder, if present. This describes the
' checkout on disk only; it says nothing about the modules loaded in Excel.
Private Function CheckoutReport(ByRef warnings As String) As String
    Dim fileSystem As Object
    Dim folder As String
    Dim head As String
    Dim branch As String
    Dim sha As String
    Dim gitDir As String
    Dim depth As Long

    branch = "UNKNOWN"
    sha = "UNKNOWN"
    On Error GoTo Unavailable
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    folder = ThisWorkbook.Path
    Do While Len(folder) > 0 And depth < 6
        If fileSystem.FolderExists(folder & "\.git") Then
            gitDir = folder & "\.git"
            Exit Do
        End If
        If fileSystem.GetParentFolderName(folder) = folder Then Exit Do
        folder = fileSystem.GetParentFolderName(folder)
        depth = depth + 1
    Loop
    If Len(gitDir) = 0 Then GoTo Unavailable

    head = Trim$(ReadSmallTextFile(fileSystem, gitDir & "\HEAD"))
    If Left$(head, 5) = "ref: " Then
        branch = Mid$(head, 6)
        If Left$(branch, 11) = "refs/heads/" Then branch = Mid$(branch, 12)
        sha = ResolveGitRef(fileSystem, gitDir, Mid$(head, 6))
    ElseIf Len(head) = 40 Then
        branch = "DETACHED"
        sha = head
    End If
    CheckoutReport = "- Branch: " & branch & vbCrLf & _
                     "- Commit: " & sha & vbCrLf & _
                     "- Checkout status: UNKNOWN (working-tree cleanliness is not detectable from VBA)" & vbCrLf & _
                     "- Source: " & gitDir & " (read-only, UNVERIFIED relation to loaded modules)" & vbCrLf
    Exit Function

Unavailable:
    CheckoutReport = "- Branch: UNKNOWN" & vbCrLf & "- Commit: UNKNOWN" & vbCrLf & _
                     "- Checkout status: UNKNOWN" & vbCrLf & _
                     "- Source: no .git folder found near the workbook" & vbCrLf
    warnings = warnings & "  * Checkout identity unavailable; record the commit manually." & vbCrLf
End Function

Private Function ReadSmallTextFile(ByVal fileSystem As Object, ByVal path As String) As String
    Dim stream As Object
    If Not fileSystem.FileExists(path) Then Exit Function
    Set stream = fileSystem.OpenTextFile(path, 1)
    If Not stream.AtEndOfStream Then ReadSmallTextFile = stream.ReadAll
    stream.Close
End Function

Private Function ResolveGitRef(ByVal fileSystem As Object, ByVal gitDir As String, _
                               ByVal refName As String) As String
    Dim text As String
    Dim lines As Variant
    Dim line As Variant

    text = Trim$(ReadSmallTextFile(fileSystem, gitDir & "\" & Replace(refName, "/", "\")))
    If Len(text) = 40 Then
        ResolveGitRef = text
        Exit Function
    End If
    text = ReadSmallTextFile(fileSystem, gitDir & "\packed-refs")
    lines = Split(Replace(text, vbCr, ""), vbLf)
    For Each line In lines
        If Len(line) > 41 And Right$(line, Len(refName)) = refName Then
            ResolveGitRef = Left$(line, 40)
            Exit Function
        End If
    Next line
    ResolveGitRef = "UNKNOWN"
End Function

Private Function LoadedModuleReport(ByRef warnings As String, ByRef unverifiedCount As Long, _
                                    ByRef mismatchCount As Long) As String
    Dim names As Variant
    Dim index As Long
    Dim tag As String
    Dim prop As String
    Dim line As String
    Dim note As String

    names = Array("VBAshTests", "IStream", "clsFileStream", "clsMemoryStream", _
                  "clsNullStream", "clsFileDescriptorTable", "clsDescriptorScope", _
                  "clsDescriptorScopeEntry", "clsRedirectionManager", "clsConsoleStream", "clsPipeStream")
    For index = LBound(names) To UBound(names)
        tag = ReadModuleTag(CStr(names(index)))
        prop = ReadModuleProperty(CStr(names(index)))
        note = ""
        If Len(prop) = 0 Then
            line = tag
        ElseIf tag = "UNVERIFIED" Then
            line = prop
            note = " [instance property only; code tag UNVERIFIED]"
        ElseIf tag = prop Then
            line = tag
            note = " [code tag and instance property agree]"
        Else
            line = "MISMATCH tag=" & tag & " property=" & prop
            mismatchCount = mismatchCount + 1
            warnings = warnings & "  * " & CStr(names(index)) & ": tag and property differ." & vbCrLf
        End If
        If line = "UNVERIFIED" Then
            unverifiedCount = unverifiedCount + 1
            warnings = warnings & "  * " & CStr(names(index)) & ": revision UNVERIFIED " & _
                       "(enable Trust access to the VBA project object model, or module lacks tag)." & vbCrLf
        End If
        LoadedModuleReport = LoadedModuleReport & "- " & CStr(names(index)) & ": " & line & note & vbCrLf
    Next index
End Function

' Reads the "' @VBAshRevision <id>" tag from the loaded code module via VBIDE.
Private Function ReadModuleTag(ByVal moduleName As String) As String
    Dim codeModule As Object
    Dim lineIndex As Long
    Dim text As String
    Dim marker As String

    marker = "@VBAshRevision "
    ReadModuleTag = "UNVERIFIED"
    On Error GoTo Done
    Set codeModule = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    For lineIndex = 1 To IIf(codeModule.CountOfLines < 10, codeModule.CountOfLines, 10)
        text = codeModule.Lines(lineIndex, 1)
        If InStr(text, marker) > 0 Then
            ReadModuleTag = Trim$(Mid$(text, InStr(text, marker) + Len(marker)))
            Exit Function
        End If
    Next lineIndex
Done:
End Function

' Reads ModuleRevision from a live instance; IStream and VBAshTests have none.
Private Function ReadModuleProperty(ByVal moduleName As String) As String
    On Error GoTo Done
    Select Case moduleName
        Case "clsFileStream"
            Dim fileStream As clsFileStream
            Set fileStream = New clsFileStream
            ReadModuleProperty = fileStream.ModuleRevision
        Case "clsMemoryStream"
            Dim memoryStream As clsMemoryStream
            Set memoryStream = New clsMemoryStream
            ReadModuleProperty = memoryStream.ModuleRevision
        Case "clsNullStream"
            Dim nullStream As clsNullStream
            Set nullStream = New clsNullStream
            ReadModuleProperty = nullStream.ModuleRevision
        Case "clsFileDescriptorTable"
            Dim descriptorTable As clsFileDescriptorTable
            Set descriptorTable = New clsFileDescriptorTable
            ReadModuleProperty = descriptorTable.ModuleRevision
        Case "clsDescriptorScope"
            Dim descriptorScope As clsDescriptorScope
            Set descriptorScope = New clsDescriptorScope
            ReadModuleProperty = descriptorScope.ModuleRevision
        Case "clsRedirectionManager"
            Dim redirectionManager As clsRedirectionManager
            Set redirectionManager = New clsRedirectionManager
            ReadModuleProperty = redirectionManager.ModuleRevision
        Case "clsConsoleStream"
            Dim consoleStream As clsConsoleStream
            Set consoleStream = New clsConsoleStream
            ReadModuleProperty = consoleStream.ModuleRevision
        Case "clsPipeStream"
            Dim pipeStream As clsPipeStream
            Set pipeStream = New clsPipeStream
            ReadModuleProperty = pipeStream.ModuleRevision
    End Select
Done:
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

    Set node = parser.ParseAst(tokenizer.Tokenize("a | b"), "a | b")
    AssertEqual "PIPELINE", node.NodeType, "pipe creates pipeline node"
    AssertEqual 2, node.Children.Count, "pipeline preserves two stages"
    Set child = node.Children(1)
    AssertEqual "SIMPLE_COMMAND", child.NodeType, "first pipeline stage is simple command"
    AssertEqual "a", child.Name, "first pipeline stage order"
    Set child = node.Children(2)
    AssertEqual "SIMPLE_COMMAND", child.NodeType, "second pipeline stage is simple command"
    AssertEqual "b", child.Name, "second pipeline stage order"
    AssertEqual 1, node.SourceStart, "pipeline starts at first stage"
    AssertEqual 5, node.SourceEnd, "pipeline ends at final stage"
    AssertEqual "a | b", Mid$("a | b", node.SourceStart, node.SourceEnd - node.SourceStart + 1), _
                "pipeline source span covers its stages"

    Set node = parser.ParseAst(tokenizer.Tokenize("a | b | c"), "a | b | c")
    AssertEqual "PIPELINE", node.NodeType, "three-stage pipeline node"
    AssertEqual 3, node.Children.Count, "pipeline preserves three stages"
    AssertEqual "a", node.Children(1).Name, "three-stage pipeline first stage order"
    AssertEqual "b", node.Children(2).Name, "three-stage pipeline second stage order"
    AssertEqual "c", node.Children(3).Name, "three-stage pipeline final stage order"

    Set node = parser.ParseAst(tokenizer.Tokenize("a | b && c | d; e"), "a | b && c | d; e")
    AssertEqual "SEQUENCE", node.NodeType, "pipeline expression followed by semicolon is sequence"
    AssertEqual "AND_IF", node.Children(1).NodeType, "and-if is above sequence"
    AssertEqual "PIPELINE", node.Children(1).Children(1).NodeType, "pipeline is above and-if"
    AssertEqual "PIPELINE", node.Children(1).Children(2).NodeType, "right pipeline is above and-if"
    AssertEqual "SIMPLE_COMMAND", node.Children(2).NodeType, "final sequence child is simple command"
    AssertEqual "e", node.Children(2).Name, "final sequence command remains ordered"

    Set node = parser.ParseAst(tokenizer.Tokenize("echo ""a | b"""), "echo ""a | b""")
    AssertEqual "SIMPLE_COMMAND", node.NodeType, "double-quoted pipe is not an operator"
    AssertEqual "a | b", node.Args(1), "double-quoted pipe remains in word"
    Set node = parser.ParseAst(tokenizer.Tokenize("echo 'a | b'"), "echo 'a | b'")
    AssertEqual "SIMPLE_COMMAND", node.NodeType, "single-quoted pipe is not an operator"
    AssertEqual "a | b", node.Args(1), "single-quoted pipe remains in word"

    Set node = parser.ParseAst(tokenizer.Tokenize("a > out | b"), "a > out | b")
    AssertEqual "PIPELINE", node.NodeType, "redirection before pipe remains in pipeline"
    AssertEqual 1, node.Children(1).Redirections.Count, "first stage retains redirection"
    AssertEqual "out", node.Children(1).Redirections(1).Target, "first-stage redirection target"
    AssertEqual 0, node.Children(2).Redirections.Count, "second stage has no first-stage redirection"
    Set node = parser.ParseAst(tokenizer.Tokenize("a | b 2> err"), "a | b 2> err")
    AssertEqual 0, node.Children(1).Redirections.Count, "first stage has no second-stage redirection"
    AssertEqual 1, node.Children(2).Redirections.Count, "second stage retains redirection"
    AssertEqual "REDIR_ERR", node.Children(2).Redirections(1).Operator, _
                "second-stage redirection operator"
    AssertEqual "err", node.Children(2).Redirections(1).Target, "second-stage redirection target"

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
    AssertAstSyntaxError tokenizer, parser, "| a", "leading pipeline operator is rejected"
    AssertAstSyntaxError tokenizer, parser, "a |", "trailing pipeline operator is rejected"
    AssertAstSyntaxError tokenizer, parser, "a | | b", "consecutive pipeline operators are rejected"
    AssertAstSyntaxError tokenizer, parser, "a ||| b", "mixed logical and pipeline operators are rejected"
    AssertLegacyParseError tokenizer, parser, "echo a;", "legacy Parse continues rejecting separators"
    AssertLegacyParseError tokenizer, parser, "a && b", "legacy Parse continues rejecting logical lists"
    AssertLegacyParseError tokenizer, parser, "a | b", "legacy Parse continues rejecting pipelines"

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

    result = session.ProcessInput("export VBAsh_AST_PIPELINE=unexpected | echo ignored")
    AssertEqual "", result, "unsupported pipeline emits no partial output"
    AssertEqual 2, session.LastExitCode, "unsupported pipeline returns syntax-class error status"
    AssertEqual "VBAsh: Execução de pipelines ainda não é suportada.", _
                session.ErrorOutput, "unsupported pipeline reports explicit error"
    AssertTrue session.HasErrorOutput, "unsupported pipeline reports stderr"
    AssertTrue Not session.HasOutput, "unsupported pipeline reports no stdout"
    AssertTrue Not session.Env.Exists("VBAsh_AST_PIPELINE"), _
               "unsupported pipeline does not execute its first stage"
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
    Dim variablePath As String
    Dim sequenceVariablePath As String
    Dim contents As String

    Set session = New clsSession
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    path = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)
    errorPath = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)
    mergePath = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)
    variablePath = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)
    sequenceVariablePath = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)

    session.ProcessInput "echo first > """ & path & """"
    AssertEqual 0, session.LastExitCode, "overwrite redirection status"
    AssertTrue Not session.HasOutput, "redirected output not sent to console"
    contents = ReadRedirectedFile(fileSystem, path, "first" & vbCrLf, "stdout overwrite")
    AssertEqual "first" & vbCrLf, contents, "overwrite redirection content"

    session.SetEnv "VBAsh_AST_OUT", variablePath
    session.ProcessInput "echo variable > ""$VBAsh_AST_OUT"""
    AssertEqual 0, session.LastExitCode, "expanded redirection target status"
    contents = ReadRedirectedFile(fileSystem, variablePath, _
        "variable" & vbCrLf, "quoted variable stdout target")
    AssertEqual "variable" & vbCrLf, contents, "quoted variable redirection target"

    session.ProcessInput "VBAsh_AST_SEQUENCE_OUT=""" & sequenceVariablePath & _
        """; echo sequence > ""$VBAsh_AST_SEQUENCE_OUT"""
    AssertEqual 0, session.LastExitCode, "sequence-expanded redirection target status"
    contents = ReadRedirectedFile(fileSystem, sequenceVariablePath, _
        "sequence" & vbCrLf, "sequence variable stdout target")
    AssertEqual "sequence" & vbCrLf, contents, "sequence expands redirection target at execution"

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

    session.ProcessInput "echo invalid > """""
    AssertEqual 2, session.LastExitCode, "empty redirection destination is rejected"
    AssertTrue session.HasErrorOutput, "empty redirection destination reports an error"

    fileSystem.DeleteFile errorPath, True
    fileSystem.DeleteFile mergePath, True
    fileSystem.DeleteFile variablePath, True
    fileSystem.DeleteFile sequenceVariablePath, True
End Sub

Private Sub TestStreamContract()
    Dim probe As clsStreamContractProbe
    Dim stream As IStream
    Dim errorNumber As Long

    AssertTrue VBAsh_STREAM_ERROR_NOT_SUPPORTED <> VBAsh_STREAM_ERROR_CLOSED, _
               "unsupported and closed stream errors are distinct"
    AssertTrue VBAsh_STREAM_ERROR_READ_FAILURE <> VBAsh_STREAM_ERROR_WRITE_FAILURE, _
               "read and write failures are distinct"
    AssertTrue VBAsh_STREAM_ERROR_OPEN_FAILURE <> VBAsh_STREAM_ERROR_READ_FAILURE, _
               "open and read failures are distinct"

    Set probe = New clsStreamContractProbe
    Set stream = probe
    AssertTrue Not stream.IsClosed, "new stream starts open"
    AssertTrue stream.CanRead, "stream contract reports readable capability"
    AssertTrue stream.CanWrite, "stream contract reports writable capability"
    AssertTrue stream.AtEndOfStream, "empty readable stream starts at EOF"
    AssertEqual 3, stream.WriteText("abc"), "stream contract reports written character count"
    AssertTrue Not stream.AtEndOfStream, "writing data clears EOF state"
    AssertEqual "ab", stream.ReadText(2), "stream contract reads a bounded chunk"
    AssertTrue Not stream.AtEndOfStream, "partial read is not EOF"
    AssertEqual "c", stream.ReadText(2), "stream contract reads remaining data"
    AssertTrue stream.AtEndOfStream, "stream contract reports EOF after final read"
    AssertEqual "", stream.ReadText(0), "zero-length read returns empty text"

    probe.ConfigureCapabilities False, True
    AssertTrue Not stream.AtEndOfStream, _
               "non-readable stream is not reported as exhausted"
    On Error Resume Next
    stream.ReadText 1
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual VBAsh_STREAM_ERROR_NOT_SUPPORTED, errorNumber, _
                "unsupported read raises the contract error"
    AssertTrue Not stream.CanRead, "read error does not change read capability"
    AssertTrue Not stream.IsClosed, "read error does not change closed state"
    AssertTrue Not stream.AtEndOfStream, _
               "EOF is false when reading is unsupported, even after consuming data"
    AssertTrue stream.HasError, "unsupported read records error state"
    AssertEqual VBAsh_STREAM_ERROR_NOT_SUPPORTED, stream.LastErrorNumber, _
                "stream exposes error number"
    AssertTrue Len(stream.LastErrorDescription) > 0, "stream exposes error description"

    probe.ConfigureCapabilities True, True
    AssertEqual "", stream.ReadText(0), "non-positive read limit returns empty text"
    AssertTrue Not stream.HasError, "successful operation clears the error state"

    probe.ConfigureCapabilities True, False
    On Error Resume Next
    stream.WriteText "x"
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual VBAsh_STREAM_ERROR_NOT_SUPPORTED, errorNumber, _
                "unsupported write raises the contract error"
    AssertTrue Not stream.CanWrite, "write error does not change write capability"
    AssertTrue Not stream.IsClosed, "write error does not change closed state"
    AssertTrue stream.AtEndOfStream, _
               "write error does not change EOF on an exhausted readable stream"
    AssertTrue stream.HasError, "unsupported write records error state"

    stream.CloseStream
    AssertTrue stream.IsClosed, "stream contract reports closed state"
    AssertTrue stream.CanRead, "closing does not rewrite read capability"
    AssertTrue Not stream.AtEndOfStream, "closed state is distinct from EOF"
    AssertTrue Not stream.HasError, "successful close clears error state"
    On Error Resume Next
    stream.CloseStream
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual 0, errorNumber, "closing an already closed stream is idempotent"
    AssertTrue stream.IsClosed, "repeated close keeps the stream closed"
    On Error Resume Next
    stream.ReadText 1
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, errorNumber, _
                "operation on a closed stream raises the closed-stream error"
    AssertTrue stream.IsClosed, "closed-stream error does not change closed state"
    AssertTrue stream.CanRead, "closed-stream error does not change read capability"
    AssertTrue Not stream.AtEndOfStream, "closed stream remains outside EOF semantics"
    AssertTrue stream.HasError, "operation on a closed stream records error state"
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, stream.LastErrorNumber, _
                "closed-stream error has its distinct code"
End Sub

Private Sub TestConsoleStream()
    Dim consoleStream As clsConsoleStream
    Dim stream As IStream

    Set consoleStream = New clsConsoleStream
    consoleStream.Initialize False, True
    Set stream = consoleStream
    AssertTrue Not stream.CanRead And stream.CanWrite, "console output adapter is write-only"
    AssertEqual 5, stream.WriteText("hello"), "console adapter accepts text characters"
    AssertEqual "hello", consoleStream.Contents, "console adapter captures host output"
    consoleStream.Reset
    AssertEqual "", consoleStream.Contents, "console adapter reset clears prior command output"
    AssertTrue Not stream.IsClosed, "reset leaves console adapter open"
    stream.CloseStream
    AssertTrue stream.IsClosed, "console adapter close is observable through IStream"
End Sub

Private Sub TestPipeStream()
    Dim pipe As clsPipeStream
    Dim stream As IStream

    Set pipe = New clsPipeStream
    Set stream = pipe
    AssertTrue stream.CanRead And stream.CanWrite, "pipe starts with both endpoints open"
    AssertTrue Not stream.AtEndOfStream, "open pipe writer prevents EOF before data"
    AssertEqual 5, stream.WriteText("hello"), "pipe accepts text from writer"
    AssertEqual "he", stream.ReadText(2), "pipe reader consumes buffered text"
    AssertEqual 3, pipe.BufferedLength, "pipe reports unread buffered character count"
    pipe.CloseWriter
    AssertEqual "llo", stream.ReadText(10), "pipe drains remaining buffered text after writer closes"
    AssertTrue stream.AtEndOfStream, "pipe reports EOF after buffer drains and writer closes"
    pipe.CloseReader
    AssertTrue stream.IsClosed, "pipe is closed when both endpoints are closed"
End Sub

Private Sub TestDescriptorScopes()
    Dim table As clsFileDescriptorTable
    Dim stdinStream As clsMemoryStream
    Dim stdoutStream As clsMemoryStream
    Dim stderrStream As clsMemoryStream
    Dim originalStream As clsMemoryStream
    Dim temporaryStream As clsMemoryStream
    Dim scope As clsDescriptorScope
    Dim nestedScope As clsDescriptorScope
    Dim stream As IStream
    Dim errorNumber As Long

    Set table = New clsFileDescriptorTable
    Set stdinStream = New clsMemoryStream
    stdinStream.Initialize "", True, False
    Set stdoutStream = New clsMemoryStream
    stdoutStream.Initialize "", False, True
    Set stderrStream = New clsMemoryStream
    stderrStream.Initialize "", False, True
    table.Initialize stdinStream, stdoutStream, stderrStream

    Set originalStream = New clsMemoryStream
    originalStream.Initialize "original", True, True
    table.SetDescriptor 3, originalStream, True
    Set scope = table.BeginDescriptorScope

    Set temporaryStream = New clsMemoryStream
    temporaryStream.Initialize "temporary", True, True
    table.SetDescriptor 3, temporaryStream, True
    table.SetDescriptor 4, temporaryStream, True

    On Error Resume Next
    table.CloseAll
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertTrue errorNumber <> 0, "CloseAll rejects active descriptor scopes"
    AssertTrue Not DescriptorStreamIsClosed(originalStream), "snapshot pins preserve the original owned stream"

    Set nestedScope = table.BeginDescriptorScope
    On Error Resume Next
    table.RestoreDescriptorScope scope
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertTrue errorNumber <> 0, "outer scope cannot be restored before inner scope"

    table.RestoreDescriptorScope nestedScope
    table.RestoreDescriptorScope scope
    Set stream = table.GetDescriptor(3)
    AssertTrue stream Is originalStream, "scope restores original descriptor identity"
    AssertTrue stream.CanRead, "restored original stream remains usable"
    AssertTrue DescriptorStreamIsClosed(temporaryStream), "owned temporary stream closes after restoration"
    AssertTrue Not table.HasDescriptor(4), "descriptor created inside scope is removed"
End Sub

Private Sub TestDescriptorRedirection()
    Dim session As clsSession
    Dim table As clsFileDescriptorTable
    Dim stdinStream As clsNullStream
    Dim stdoutStream As clsConsoleStream
    Dim stderrStream As clsConsoleStream
    Dim fileSystem As Object
    Dim path As String
    Dim result As String
    Dim contents As String

    Set session = New clsSession
    Set table = New clsFileDescriptorTable
    Set stdinStream = New clsNullStream
    stdinStream.Initialize True, False
    Set stdoutStream = New clsConsoleStream
    stdoutStream.Initialize False, True
    Set stderrStream = New clsConsoleStream
    stderrStream.Initialize False, True
    table.Initialize stdinStream, stdoutStream, stderrStream, True
    session.ConfigureDescriptorTable table, stdoutStream, stderrStream

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    path = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), fileSystem.GetTempName)
    result = session.ProcessInput("echo descriptor > """ & path & """")
    AssertEqual 0, session.LastExitCode, "descriptor redirection executes builtin successfully"
    AssertTrue Not session.HasOutput, "file redirection suppresses console output"
    contents = ReadRedirectedFile(fileSystem, path, "descriptor" & vbCrLf, _
                                  "descriptor-based stdout")
    AssertEqual "descriptor" & vbCrLf, contents, "descriptor manager writes stdout to file"

    result = session.ProcessInput("pwd extra 2> """ & path & """")
    AssertEqual 2, session.LastExitCode, "descriptor stderr redirection preserves builtin status"
    AssertTrue Not session.HasErrorOutput, "redirected stderr is not returned to terminal"
    contents = ReadRedirectedFile(fileSystem, path, "VBAsh: Uso: pwd" & vbCrLf, _
                                  "descriptor-based stderr")
    AssertEqual "VBAsh: Uso: pwd" & vbCrLf, contents, "descriptor manager writes stderr to file"
    result = session.ProcessInput("echo routed 1>&2")
    AssertEqual 0, session.LastExitCode, "stdout-to-stderr duplication preserves command status"
    AssertTrue Not session.HasOutput, "stdout-to-stderr duplication removes normal output channel"
    AssertTrue session.HasErrorOutput, "stdout-to-stderr duplication is visible on stderr channel"
    AssertEqual "routed", session.ErrorOutput, "stdout-to-stderr duplication preserves output text"

    fileSystem.DeleteFile path, True
    table.CloseAll
End Sub

Private Sub TestFileDescriptorTable()
    Dim table As clsFileDescriptorTable
    Dim ownedDefaults As clsFileDescriptorTable
    Dim stdinStream As clsMemoryStream
    Dim stdoutStream As clsMemoryStream
    Dim stderrStream As clsMemoryStream
    Dim ownedStream As clsMemoryStream
    Dim replacementStream As clsMemoryStream
    Dim anotherStream As clsMemoryStream
    Dim probe As clsStreamContractProbe
    Dim unassignedStream As IStream
    Dim stream As IStream
    Dim streamAlias As IStream
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    Set table = New clsFileDescriptorTable
    On Error Resume Next
    table.HasDescriptor 0
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    AssertTrue errorNumber <> 0, "table operations require explicit standard-stream initialization"
    AssertEqual "clsFileDescriptorTable.HasDescriptor", errorSource, _
                "uninitialized table error identifies the operation"

    Set stdinStream = New clsMemoryStream
    stdinStream.Initialize "stdin", True, False
    Set stdoutStream = New clsMemoryStream
    stdoutStream.Initialize "", False, True
    Set stderrStream = New clsMemoryStream
    stderrStream.Initialize "", False, True
    table.Initialize stdinStream, stdoutStream, stderrStream

    AssertTrue table.HasDescriptor(0), "initialization associates stdin with descriptor zero"
    AssertTrue table.HasDescriptor(1), "initialization associates stdout with descriptor one"
    AssertTrue table.HasDescriptor(2), "initialization associates stderr with descriptor two"
    Set stream = table.GetDescriptor(0)
    AssertTrue stream Is stdinStream, "descriptor zero returns the supplied stdin object"
    AssertTrue stream.CanRead And Not stream.CanWrite, "stdin capabilities are preserved"
    Set stream = table.GetDescriptor(1)
    AssertTrue stream Is stdoutStream, "descriptor one returns the supplied stdout object"
    AssertTrue Not stream.CanRead And stream.CanWrite, "stdout capabilities are preserved"
    Set stream = table.GetDescriptor(2)
    AssertTrue stream Is stderrStream, "descriptor two returns the supplied stderr object"
    AssertTrue Not stream.IsClosed, "standard streams are borrowed and open by default"
    AssertEqual 3, table.FindFreeDescriptor(), "first free descriptor starts at three"

    AssertTrue Not table.HasDescriptor(3), "unassigned descriptor is absent"
    On Error Resume Next
    Set stream = table.GetDescriptor(3)
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    AssertTrue errorNumber <> 0, "retrieving an absent descriptor raises an error"
    AssertEqual "clsFileDescriptorTable.GetDescriptor", errorSource, _
                "absent-descriptor error identifies the operation"
    AssertTrue InStr(1, errorDescription, "does not exist", vbTextCompare) > 0, _
               "absent-descriptor error explains the missing association"

    Set ownedStream = New clsMemoryStream
    ownedStream.Initialize "", False, True
    table.SetDescriptor 5, ownedStream
    table.SetDescriptor 4, ownedStream
    AssertEqual 3, table.FindFreeDescriptor(), "free descriptor search finds the lowest hole"
    table.SetDescriptor 3, stdoutStream
    AssertEqual 6, table.FindFreeDescriptor(), "free descriptor search skips occupied numbers"

    table.DuplicateDescriptor 3, 6
    Set stream = table.GetDescriptor(3)
    Set streamAlias = table.GetDescriptor(6)
    AssertTrue stream Is streamAlias, "duplicate descriptors preserve object identity"
    AssertTrue streamAlias Is stdoutStream, "duplicate points at the original stream object"
    AssertEqual 4, streamAlias.WriteText("same"), "write through duplicate reaches the shared stream"
    AssertEqual "same", stdoutStream.Contents, "shared stream contains duplicate write"
    table.CloseDescriptor 3
    AssertTrue Not DescriptorStreamIsClosed(stdoutStream), _
               "closing one borrowed alias does not close its stream"
    AssertTrue table.HasDescriptor(6), "closing one alias leaves the other descriptor present"
    Set stream = table.GetDescriptor(6)
    AssertEqual 1, stream.WriteText("!"), "remaining alias stays functional"
    AssertEqual "same!", stdoutStream.Contents, "remaining alias writes to the same destination"
    table.CloseDescriptor 6
    AssertTrue Not DescriptorStreamIsClosed(stdoutStream), _
               "closing the last borrowed alias does not close the stream"

    Set ownedStream = New clsMemoryStream
    ownedStream.Initialize "", False, True
    table.SetDescriptor 7, ownedStream, True
    table.DuplicateDescriptor 7, 8
    Set stream = table.GetDescriptor(8)
    AssertTrue stream Is ownedStream, "owned duplicate shares the owned object"
    table.CloseDescriptor 7
    AssertTrue Not DescriptorStreamIsClosed(ownedStream), _
               "owned stream stays open while a duplicate refers to it"
    table.CloseDescriptor 8
    AssertTrue DescriptorStreamIsClosed(ownedStream), "last owned descriptor closes the stream"
    AssertTrue Not table.HasDescriptor(8), "successful final close removes the descriptor"

    Set ownedStream = New clsMemoryStream
    ownedStream.Initialize "", False, True
    Set replacementStream = New clsMemoryStream
    replacementStream.Initialize "", False, True
    table.SetDescriptor 9, ownedStream, True
    table.SetDescriptor 10, ownedStream, False
    table.SetDescriptor 9, replacementStream, False
    AssertTrue Not DescriptorStreamIsClosed(ownedStream), _
               "replacing one alias does not close an owned stream still referenced elsewhere"
    table.CloseDescriptor 10
    AssertTrue DescriptorStreamIsClosed(ownedStream), _
               "ownership persists after a later borrowed association and owner alias replacement"
    Set stream = table.GetDescriptor(9)
    AssertTrue stream Is replacementStream, "replacement associates the new stream"

    Set ownedStream = New clsMemoryStream
    ownedStream.Initialize "", False, True
    Set anotherStream = New clsMemoryStream
    anotherStream.Initialize "", False, True
    table.SetDescriptor 11, ownedStream, True
    table.SetDescriptor 11, anotherStream, False
    AssertTrue DescriptorStreamIsClosed(ownedStream), _
               "replacing the last owned descriptor closes its old stream"
    Set stream = table.GetDescriptor(11)
    AssertTrue stream Is anotherStream, "occupied descriptor now returns the replacement object"
    AssertTrue Not DescriptorStreamIsClosed(anotherStream), "borrowed replacement remains open"

    Set unassignedStream = Nothing
    On Error Resume Next
    table.SetDescriptor 11, unassignedStream, True
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertTrue errorNumber <> 0, "associating Nothing is rejected"
    Set stream = table.GetDescriptor(11)
    AssertTrue stream Is anotherStream, "invalid replacement preserves the previous association"
    AssertTrue Not DescriptorStreamIsClosed(anotherStream), _
               "invalid replacement leaves the old stream usable"

    On Error Resume Next
    table.SetDescriptor -1, anotherStream
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertTrue errorNumber <> 0, "negative descriptors are rejected"
    AssertTrue table.HasDescriptor(11), "invalid descriptor operation preserves valid associations"

    On Error Resume Next
    table.DuplicateDescriptor 99, 12
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertTrue errorNumber <> 0, "duplicating an absent source descriptor is rejected"
    AssertTrue Not table.HasDescriptor(12), "failed duplication does not create its destination"

    Set replacementStream = New clsMemoryStream
    replacementStream.Initialize "closed", True, False
    Set stream = replacementStream
    stream.CloseStream
    table.SetDescriptor 13, replacementStream
    AssertTrue table.HasDescriptor(13), "a descriptor can reference an already closed stream"
    Set stream = table.GetDescriptor(13)
    AssertTrue stream Is replacementStream And stream.IsClosed, _
               "existing closed stream is distinguishable from an absent descriptor"
    table.CloseDescriptor 13
    AssertTrue DescriptorStreamIsClosed(replacementStream), _
               "closing borrowed closed-stream association does not change it"

    Set probe = New clsStreamContractProbe
    probe.ConfigureCloseFailure True
    table.SetDescriptor 14, probe, True
    On Error Resume Next
    table.CloseDescriptor 14
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 1260, errorNumber, _
                "failed owned close preserves the original error number"
    AssertEqual "clsFileDescriptorTable.CloseDescriptor", errorSource, _
                "failed owned close identifies the table operation"
    AssertTrue InStr(1, errorDescription, "Configured close failure", vbTextCompare) > 0, _
               "failed owned close preserves the original error description"
    AssertTrue table.HasDescriptor(14), "failed final close preserves the descriptor association"
    Set stream = table.GetDescriptor(14)
    AssertTrue stream Is probe And Not stream.IsClosed, _
               "failed close retains the original open stream"
    AssertEqual 1, probe.CloseCount, "first close failure attempts the stream exactly once"

    probe.ConfigureCloseFailure False
    table.CloseDescriptor 14
    AssertTrue Not table.HasDescriptor(14), "retry after close failure removes the descriptor"
    AssertTrue DescriptorStreamIsClosed(probe), "successful retry closes the owned stream"
    AssertEqual 2, probe.CloseCount, "retry invokes CloseStream again"

    Set probe = New clsStreamContractProbe
    probe.ConfigureCloseFailure True
    table.SetDescriptor 15, probe, True
    On Error Resume Next
    table.SetDescriptor 15, anotherStream
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 1260, errorNumber, _
                "failed owned close prevents descriptor replacement"
    Set stream = table.GetDescriptor(15)
    AssertTrue stream Is probe And Not stream.IsClosed, _
               "failed replacement preserves the previous descriptor mapping"
    AssertTrue Not DescriptorStreamIsClosed(anotherStream), _
               "failed replacement leaves the incoming stream untouched"
    probe.ConfigureCloseFailure False
    table.SetDescriptor 15, anotherStream
    Set stream = table.GetDescriptor(15)
    AssertTrue stream Is anotherStream And DescriptorStreamIsClosed(probe), _
               "replacement succeeds after the old stream can be closed"
    table.CloseDescriptor 15

    Set stdinStream = New clsMemoryStream
    stdinStream.Initialize "", True, False
    Set stdoutStream = New clsMemoryStream
    stdoutStream.Initialize "", False, True
    Set stderrStream = New clsMemoryStream
    stderrStream.Initialize "", False, True
    Set ownedDefaults = New clsFileDescriptorTable
    ownedDefaults.Initialize stdinStream, stdoutStream, stderrStream, True
    ownedDefaults.CloseDescriptor 0
    AssertTrue DescriptorStreamIsClosed(stdinStream), _
               "explicit standard-stream ownership closes the final owned descriptor"
    AssertTrue Not DescriptorStreamIsClosed(stdoutStream), _
               "closing owned stdin does not close another standard stream"
    ownedDefaults.CloseDescriptor 1
    ownedDefaults.CloseDescriptor 2

    Set table = Nothing
    Set ownedDefaults = Nothing
End Sub

Private Sub TestFileDescriptorTableCloseAll()
    Dim table As clsFileDescriptorTable
    Dim borrowed As clsStreamContractProbe
    Dim failed As clsStreamContractProbe
    Dim partial As clsStreamContractProbe
    Dim succeeded As clsStreamContractProbe
    Dim stream As IStream
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    Set table = New clsFileDescriptorTable
    On Error Resume Next
    table.CloseAll
    errorNumber = Err.Number
    errorSource = Err.Source
    Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 1250, errorNumber, "CloseAll requires initialization"
    AssertEqual "clsFileDescriptorTable.CloseAll", errorSource, _
                "uninitialized cleanup identifies its operation"

    Set borrowed = New clsStreamContractProbe
    Set failed = New clsStreamContractProbe
    Set partial = New clsStreamContractProbe
    Set succeeded = New clsStreamContractProbe
    failed.ConfigureCloseFailureDetails vbObjectError + 1280, "failure.one", "first failure", False
    partial.ConfigureCloseFailureDetails vbObjectError + 1281, "failure.two", "partial failure", True
    table.Initialize borrowed, borrowed, borrowed
    table.SetDescriptor 3, failed, True
    table.DuplicateDescriptor 3, 4
    table.SetDescriptor 5, partial, True
    table.DuplicateDescriptor 5, 6
    table.SetDescriptor 7, succeeded, True
    table.DuplicateDescriptor 7, 8
    table.SetDescriptor 9, borrowed

    On Error Resume Next
    table.CloseAll
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 1255, errorNumber, "multiple close failures raise aggregate error"
    AssertEqual "clsFileDescriptorTable.CloseAll", errorSource, "aggregate identifies CloseAll"
    AssertTrue InStr(1, errorDescription, "FDs [3, 4]: error " & CStr(vbObjectError + 1280) & _
               "; source [failure.one]; description [first failure]", vbBinaryCompare) > 0 Or _
               InStr(1, errorDescription, "FDs [4, 3]: error " & CStr(vbObjectError + 1280) & _
               "; source [failure.one]; description [first failure]", vbBinaryCompare) > 0, _
               "aggregate includes all aliases and original first error"
    AssertTrue InStr(1, errorDescription, "FDs [5, 6]: error " & CStr(vbObjectError + 1281) & _
               "; source [failure.two]; description [partial failure]", vbBinaryCompare) > 0 Or _
               InStr(1, errorDescription, "FDs [6, 5]: error " & CStr(vbObjectError + 1281) & _
               "; source [failure.two]; description [partial failure]", vbBinaryCompare) > 0, _
               "aggregate includes all aliases and original partial-close error"
    AssertEqual 1, failed.CloseCount, "failed owned aliases are closed only once per attempt"
    AssertEqual 1, partial.CloseCount, "partially closed aliases are attempted once"
    AssertEqual 1, succeeded.CloseCount, "later distinct owned stream is attempted after failures"
    AssertEqual 0, borrowed.CloseCount, "borrowed aliases are never closed"
    AssertTrue DescriptorStreamIsClosed(partial), "partial failure can leave the stream closed"
    AssertTrue DescriptorStreamIsClosed(succeeded), "successful owned stream is closed"
    AssertTrue Not table.HasDescriptor(0) And Not table.HasDescriptor(1) And _
               Not table.HasDescriptor(2) And Not table.HasDescriptor(9), _
               "all borrowed descriptors are removed"
    AssertTrue Not table.HasDescriptor(7) And Not table.HasDescriptor(8), _
               "all aliases of a successful owned stream are removed"
    AssertTrue table.HasDescriptor(3) And table.HasDescriptor(4) And _
               table.HasDescriptor(5) And table.HasDescriptor(6), _
               "all aliases of failed owned streams remain"
    Set stream = table.GetDescriptor(3)
    AssertTrue stream Is failed, "first failed mapping preserves stream identity"
    Set stream = table.GetDescriptor(4)
    AssertTrue stream Is failed, "first failed alias preserves stream identity"
    Set stream = table.GetDescriptor(5)
    AssertTrue stream Is partial, "partial failure preserves stream identity"
    Set stream = table.GetDescriptor(6)
    AssertTrue stream Is partial, "partial failure preserves alias identity"
    AssertEqual 7, table.FindFreeDescriptor(), "successful cleanup releases occupied descriptor numbers"

    failed.ConfigureCloseFailure False
    partial.ConfigureCloseFailure False
    table.CloseAll
    AssertEqual 2, failed.CloseCount, "retry closes the first retained owned stream"
    AssertEqual 2, partial.CloseCount, "retry closes a stream that closed before reporting failure"
    AssertEqual 1, succeeded.CloseCount, "retry does not close previously removed streams"
    AssertEqual 0, borrowed.CloseCount, "retry does not close previously removed borrowed streams"
    AssertTrue Not table.HasDescriptor(3) And Not table.HasDescriptor(4) And _
               Not table.HasDescriptor(5) And Not table.HasDescriptor(6), _
               "successful retry removes every retained alias"
    AssertTrue DescriptorStreamIsClosed(failed), "retry closes the previously failed stream"
    AssertEqual 3, table.FindFreeDescriptor(), "empty table returns the first free descriptor"
    table.CloseAll
    AssertEqual 2, failed.CloseCount, "repeated cleanup on an empty table does not close again"
    AssertEqual 2, partial.CloseCount, "repeated cleanup does not retry removed partial stream"

    table.SetDescriptor 0, borrowed
    table.SetDescriptor 3, succeeded
    AssertTrue table.HasDescriptor(0) And table.HasDescriptor(3), _
               "the initialized table can be reused after cleanup"
    AssertEqual 4, table.FindFreeDescriptor(), "free descriptor search reflects reuse"
    On Error Resume Next
    table.Initialize borrowed, borrowed, borrowed
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 1254, errorNumber, "Initialize remains single-use after cleanup"
    Set stream = table.GetDescriptor(0)
    AssertTrue stream Is borrowed, "rejected initialization preserves reused descriptors"
    table.CloseAll
    AssertEqual 0, borrowed.CloseCount, "cleanup after reuse does not close borrowed streams"
    AssertEqual 1, succeeded.CloseCount, "reused borrowed closed stream is not closed again"
    AssertTrue Not table.HasDescriptor(0) And Not table.HasDescriptor(3), _
               "cleanup removes descriptors added after reuse"

    Set failed = New clsStreamContractProbe
    Set succeeded = New clsStreamContractProbe
    failed.ConfigureCleanupReentry table
    table.SetDescriptor 3, failed, True
    table.SetDescriptor 4, succeeded, True
    On Error Resume Next
    table.CloseAll
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 1255, errorNumber, "reentrant cleanup becomes an aggregate close error"
    AssertTrue InStr(1, errorDescription, "error " & CStr(vbObjectError + 1256), vbBinaryCompare) > 0, _
               "aggregate identifies the rejected reentrant cleanup"
    AssertTrue table.HasDescriptor(3) And Not table.HasDescriptor(4), _
               "reentrant failure preserves its FD while other streams are closed"
    AssertEqual 1, succeeded.CloseCount, "reentrant cleanup does not stop another stream"
    failed.ConfigureCleanupReentry Nothing
    table.CloseAll
    AssertEqual 2, failed.CloseCount, "retry succeeds after reentry is disabled"
End Sub

Private Function DescriptorStreamIsClosed(ByVal streamObject As Object) As Boolean
    Dim stream As IStream

    Set stream = streamObject
    DescriptorStreamIsClosed = stream.IsClosed
End Function

Private Sub TestMemoryStream()
    Dim memory As clsMemoryStream
    Dim stream As IStream
    Dim errorNumber As Long

    Set memory = New clsMemoryStream
    memory.Initialize "abcdef", True, True
    Set stream = memory
    AssertTrue stream.CanRead, "memory stream initializes as readable"
    AssertTrue stream.CanWrite, "memory stream initializes as writable"
    AssertTrue Not stream.IsClosed, "memory stream initializes open"
    AssertEqual "abcdef", memory.Contents, "initial content is retained"
    AssertEqual 1, memory.Position, "initial read position is one"
    AssertTrue Not stream.AtEndOfStream, "nonempty stream does not start at EOF"

    AssertEqual "abcdef", stream.ReadText(20), "read limit larger than content returns all text"
    AssertEqual 7, memory.Position, "full read advances position past the last character"
    AssertTrue stream.AtEndOfStream, "full read reaches EOF"
    AssertEqual "", stream.ReadText(1), "read at EOF returns empty text"
    AssertEqual 7, memory.Position, "read at EOF does not advance position"

    memory.Initialize "abcdef", True, True
    AssertEqual "ab", stream.ReadText(2), "first partial read returns bounded text"
    AssertEqual 3, memory.Position, "first partial read advances position"
    AssertEqual "cd", stream.ReadText(2), "second partial read continues from current position"
    AssertEqual 5, memory.Position, "second partial read advances position"
    AssertEqual "ef", stream.ReadText(20), "last partial read returns remaining text"
    AssertEqual 7, memory.Position, "last partial read advances position to EOF"
    AssertTrue stream.AtEndOfStream, "multiple partial reads reach EOF"
    AssertEqual "", stream.ReadText(0), "zero read limit returns empty text"
    AssertEqual "", stream.ReadText(-1), "negative read limit returns empty text"
    AssertEqual 7, memory.Position, "non-positive read limits preserve position"

    memory.Initialize "seed", True, True
    AssertEqual "se", stream.ReadText(2), "read cursor can advance before writing"
    AssertEqual 4, stream.WriteText("tail"), "write returns accepted character count"
    AssertEqual "seedtail", memory.Contents, "write appends to existing buffer"
    AssertEqual 3, memory.Position, "write preserves current read position"
    AssertEqual "edtail", stream.ReadText(20), "appended content follows unread text"
    AssertEqual 9, memory.Position, "reading appended content advances position"

    memory.Initialize "input", True, False
    AssertTrue stream.CanRead, "read-only initialization enables reading"
    AssertTrue Not stream.CanWrite, "read-only initialization disables writing"
    AssertEqual "input", stream.ReadText(20), "read-only stream returns initial content"
    On Error Resume Next
    stream.WriteText "x"
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual VBAsh_STREAM_ERROR_NOT_SUPPORTED, errorNumber, _
                "write on read-only stream raises unsupported error"
    AssertTrue stream.HasError, "unsupported write records error state"
    AssertEqual VBAsh_STREAM_ERROR_NOT_SUPPORTED, stream.LastErrorNumber, _
                "unsupported write exposes its error code"
    AssertTrue Len(stream.LastErrorDescription) > 0, _
               "unsupported write exposes its error description"
    AssertTrue stream.AtEndOfStream, "unsupported write preserves EOF after consuming all readable text"
    AssertEqual "input", memory.Contents, "unsupported write preserves buffer contents"
    AssertEqual 6, memory.Position, "unsupported write preserves read position"
    AssertTrue stream.CanRead, "unsupported write preserves read capability"
    AssertTrue Not stream.CanWrite, "unsupported write preserves write capability"

    memory.Initialize "", False, True
    AssertTrue Not stream.CanRead, "write-only initialization disables reading"
    AssertTrue stream.CanWrite, "write-only initialization enables writing"
    AssertTrue Not stream.AtEndOfStream, "write-only stream is not reported at EOF"
    On Error Resume Next
    stream.ReadText 1
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual VBAsh_STREAM_ERROR_NOT_SUPPORTED, errorNumber, _
                "read on write-only stream raises unsupported error"
    AssertEqual "", memory.Contents, "unsupported read preserves buffer contents"
    AssertEqual 1, memory.Position, "unsupported read preserves read position"
    AssertTrue Not stream.CanRead, "unsupported read preserves read capability"
    AssertTrue stream.CanWrite, "unsupported read preserves write capability"
    AssertEqual 3, stream.WriteText("new"), "write-only stream accepts text"
    AssertEqual "new", memory.Contents, "write-only stream stores written text"
    AssertEqual 1, memory.Position, "write-only stream writes do not move read position"
    AssertTrue Not stream.HasError, "successful write clears prior error state"

    memory.Initialize
    AssertTrue stream.CanRead, "empty stream is configured as readable"
    AssertTrue stream.CanWrite, "empty stream is configured as writable by default"
    AssertEqual "", memory.Contents, "readable stream starts with empty content"
    AssertEqual 1, memory.Position, "empty readable stream starts at position one"
    AssertTrue stream.AtEndOfStream, "empty readable stream starts at EOF"
    AssertEqual "", stream.ReadText(1), "read from initially empty stream returns empty text"
    AssertEqual 1, memory.Position, "read from empty stream preserves position"

    memory.Initialize "end", True, True
    AssertEqual "end", stream.ReadText(10), "read reaches the end before append"
    AssertEqual 4, memory.Position, "position after reaching EOF is preserved as the next index"
    AssertTrue stream.AtEndOfStream, "stream reaches EOF before append"
    AssertEqual 2, stream.WriteText("XY"), "write after EOF reports appended character count"
    AssertEqual "endXY", memory.Contents, "write after EOF appends new content"
    AssertEqual 4, memory.Position, "write after EOF preserves the prior position"
    AssertTrue Not stream.AtEndOfStream, "appended content makes the stream non-EOF"
    AssertEqual "XY", stream.ReadText(10), "newly appended content is readable"
    AssertEqual 6, memory.Position, "reading appended content advances position"
    AssertTrue stream.AtEndOfStream, "reading appended content reaches EOF again"

    memory.Initialize "closed", True, True
    stream.CloseStream
    AssertTrue stream.IsClosed, "close marks memory stream closed"
    AssertTrue stream.CanRead, "close does not change read capability"
    AssertTrue stream.CanWrite, "close does not change write capability"
    AssertTrue Not stream.AtEndOfStream, "closed stream is not at EOF"
    stream.CloseStream
    AssertTrue stream.IsClosed, "repeated close keeps stream closed"
    AssertTrue Not stream.HasError, "repeated close leaves no error"
    On Error Resume Next
    stream.ReadText 1
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, errorNumber, _
                "read after close raises closed-stream error"
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, stream.LastErrorNumber, _
                "read after close records closed-stream error"
    AssertTrue stream.IsClosed, "closed-stream error preserves closed state"
    AssertTrue stream.CanRead, "closed-stream error preserves read capability"
    AssertTrue Not stream.AtEndOfStream, "closed-stream error does not imply EOF"
    AssertEqual "closed", memory.Contents, "closed-stream read failure preserves buffer contents"
    AssertEqual 1, memory.Position, "closed-stream read failure preserves position"
    On Error Resume Next
    stream.WriteText "x"
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, errorNumber, _
                "write after close raises closed-stream error"
    AssertTrue stream.HasError, "write after close records error state"
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, stream.LastErrorNumber, _
                "write after close records closed-stream error code"
    AssertTrue Len(stream.LastErrorDescription) > 0, _
               "write after close records error description"
    AssertEqual "closed", memory.Contents, "closed-stream write failure preserves buffer contents"
    AssertEqual 1, memory.Position, "closed-stream write failure preserves position"
    AssertTrue stream.IsClosed, "closed-stream write failure preserves closed state"
    AssertTrue stream.CanRead, "closed-stream write failure preserves read capability"
    AssertTrue stream.CanWrite, "closed-stream write failure preserves write capability"

    memory.Initialize "fresh", False, True
    AssertTrue Not stream.IsClosed, "reinitialization reopens the stream"
    AssertTrue Not stream.CanRead, "reinitialization replaces read capability"
    AssertTrue stream.CanWrite, "reinitialization replaces write capability"
    AssertEqual "fresh", memory.Contents, "reinitialization replaces old content"
    AssertEqual 1, memory.Position, "reinitialization resets read position"
    AssertTrue Not stream.HasError, "reinitialization clears error state"
    AssertTrue Not stream.AtEndOfStream, "reinitialized write-only stream is not at EOF"
    AssertEqual 1, stream.WriteText("!"), "reinitialized stream accepts writes"
    AssertEqual "fresh!", memory.Contents, "write after reinitialization appends to new content"

    memory.Initialize "blocked", False, False
    AssertTrue Not stream.CanRead, "disabled stream reports no read capability"
    AssertTrue Not stream.CanWrite, "disabled stream reports no write capability"
    AssertTrue Not stream.AtEndOfStream, "disabled stream is not reported at EOF"
    On Error Resume Next
    stream.ReadText 1
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual VBAsh_STREAM_ERROR_NOT_SUPPORTED, errorNumber, _
                "read on disabled stream raises unsupported error"
    AssertEqual "blocked", memory.Contents, "unsupported read preserves disabled-stream content"
    AssertEqual 1, memory.Position, "unsupported read preserves disabled-stream position"
    AssertTrue Not stream.CanRead, "unsupported read preserves disabled-stream read capability"
    AssertTrue Not stream.CanWrite, "unsupported read preserves disabled-stream write capability"
    On Error Resume Next
    stream.WriteText "x"
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual VBAsh_STREAM_ERROR_NOT_SUPPORTED, errorNumber, _
                "write on disabled stream raises unsupported error"
    AssertEqual "blocked", memory.Contents, "unsupported write preserves disabled-stream content"
    AssertEqual 1, memory.Position, "unsupported write preserves disabled-stream position"
    AssertTrue Not stream.CanRead, "unsupported write preserves disabled-stream read capability"
    AssertTrue Not stream.CanWrite, "unsupported write preserves disabled-stream write capability"
End Sub

Private Sub TestNullStream()
    Dim nullStream As clsNullStream
    Dim stream As IStream
    Dim errorNumber As Long

    Set nullStream = New clsNullStream
    Set stream = nullStream
    AssertTrue stream.CanRead, "null stream is readable by default"
    AssertTrue stream.CanWrite, "null stream is writable by default"
    AssertTrue Not stream.IsClosed, "null stream starts open"
    AssertTrue Not stream.HasError, "null stream starts without error"
    AssertEqual 0, stream.LastErrorNumber, "initial error number is zero"
    AssertEqual "", stream.LastErrorDescription, "initial error description is empty"
    AssertTrue stream.AtEndOfStream, "open null stream is at EOF"

    AssertEqual "", stream.ReadText(10), "read returns empty text"
    AssertEqual "", stream.ReadText(0), "zero-limit read returns empty text"
    AssertEqual "", stream.ReadText(-1), "negative-limit read returns empty text"
    AssertTrue stream.AtEndOfStream, "read keeps EOF"

    AssertEqual 0, stream.WriteText(""), "empty write returns zero"
    AssertEqual 1, stream.WriteText("a"), "one-char write returns one"
    AssertEqual 10, stream.WriteText("hello" & vbCrLf & "abc"), "multi-line write counts all characters"
    AssertEqual 100000, stream.WriteText(String$(100000, "x")), "large write returns full length"
    AssertEqual "", stream.ReadText(100), "written content is not stored"
    AssertTrue stream.AtEndOfStream, "write does not change EOF"
    AssertTrue Not stream.HasError, "successful operations leave no error"

    stream.CloseStream
    AssertTrue stream.IsClosed, "close marks null stream closed"
    AssertTrue stream.CanRead, "close keeps read capability"
    AssertTrue stream.CanWrite, "close keeps write capability"
    AssertTrue Not stream.AtEndOfStream, "closed null stream is not at EOF"
    AssertTrue Not stream.HasError, "close leaves no error"
    stream.CloseStream
    AssertTrue stream.IsClosed, "repeated close keeps stream closed"
    AssertTrue Not stream.HasError, "repeated close leaves no error"

    On Error Resume Next
    stream.ReadText 1
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, errorNumber, "read after close raises closed error"
    AssertTrue stream.HasError, "read after close records error"
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, stream.LastErrorNumber, "read after close exposes code"
    AssertTrue Len(stream.LastErrorDescription) > 0, "read after close exposes description"

    On Error Resume Next
    stream.WriteText "x"
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, errorNumber, "write after close raises closed error"
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, stream.LastErrorNumber, "write after close exposes code"
    stream.CloseStream
    AssertTrue Not stream.HasError, "close clears prior error"

    nullStream.Initialize
    AssertTrue Not stream.IsClosed, "Initialize reopens the stream"
    AssertEqual 3, stream.WriteText("abc"), "write works after reinitialize"

    nullStream.Initialize True, False
    AssertTrue stream.CanRead And Not stream.CanWrite, "read-only configuration"
    On Error Resume Next
    stream.WriteText "x"
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual VBAsh_STREAM_ERROR_NOT_SUPPORTED, errorNumber, "write on non-writable raises unsupported"
    AssertEqual "", stream.ReadText(1), "successful read after failure"
    AssertTrue Not stream.HasError, "successful read clears prior error"

    nullStream.Initialize False, True
    AssertTrue Not stream.AtEndOfStream, "non-readable null stream is not at EOF"
    On Error Resume Next
    stream.ReadText 1
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEqual VBAsh_STREAM_ERROR_NOT_SUPPORTED, errorNumber, "read on non-readable raises unsupported"
    AssertEqual 2, stream.WriteText("ok"), "successful write after failure"
    AssertTrue Not stream.HasError, "successful write clears prior error"
End Sub

Private Sub TestFileStream()
    Dim fileSystem As Object
    Dim testFile As Object
    Dim seedStream As Object
    Dim fileStream As clsFileStream
    Dim stream As IStream
    Dim testFolder As String
    Dim testFolderCreated As Boolean
    Dim readPath As String
    Dim emptyPath As String
    Dim writePath As String
    Dim appendPath As String
    Dim createPath As String
    Dim missingPath As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim testFailed As Boolean
    Dim failureStage As String
    Dim stageTrace As String

    On Error GoTo TestFailure
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    testFolder = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2), _
                                      "VBAshFileStream_" & fileSystem.GetTempName)
    fileSystem.CreateFolder testFolder
    testFolderCreated = True
    readPath = fileSystem.BuildPath(testFolder, "read.txt")
    emptyPath = fileSystem.BuildPath(testFolder, "empty.txt")
    writePath = fileSystem.BuildPath(testFolder, "overwrite.txt")
    appendPath = fileSystem.BuildPath(testFolder, "append.txt")
    createPath = fileSystem.BuildPath(testFolder, "created.txt")
    missingPath = fileSystem.BuildPath(testFolder, "missing.txt")

    Set seedStream = fileSystem.CreateTextFile(readPath, True)
    seedStream.Write "abcdefg"
    seedStream.Close
    Set seedStream = fileSystem.CreateTextFile(emptyPath, True)
    seedStream.Close
    Set seedStream = fileSystem.CreateTextFile(writePath, True)
    seedStream.Write "old contents"
    seedStream.Close
    Set seedStream = fileSystem.CreateTextFile(appendPath, True)
    seedStream.Write "base"
    seedStream.Close

    Set fileStream = New clsFileStream
    Set stream = fileStream
    AssertTrue stream.IsClosed, "new file stream starts closed"
    AssertTrue Not stream.CanRead And Not stream.CanWrite, _
               "new file stream has no configured capabilities"
    AssertTrue Not stream.AtEndOfStream, "closed file stream is not at EOF"

    fileStream.OpenForRead readPath
    AssertTrue stream.CanRead, "file stream opens for reading"
    AssertTrue Not stream.CanWrite, "read-only open disables writing"
    AssertTrue Not stream.IsClosed, "read open marks stream open"
    AssertTrue Not stream.AtEndOfStream, "nonempty file starts before EOF"
    AssertEqual "ab", fileStream.ReadText(2), "legacy read returns a bounded first chunk"
    AssertEqual "cde", stream.ReadText(3), "interface read continues sequentially"
    AssertTrue Not stream.AtEndOfStream, "partial reads do not report EOF"
    AssertEqual "fg", stream.ReadText(20), "oversized read returns only remaining content"
    AssertTrue stream.AtEndOfStream, "reading final characters reaches EOF"
    AssertEqual "", stream.ReadText(2), "read after EOF returns empty text"
    AssertTrue stream.AtEndOfStream, "read after EOF keeps EOF true"

    fileStream.CloseStream
    AssertTrue stream.IsClosed, "close marks file stream closed"
    AssertTrue stream.CanRead And Not stream.CanWrite, _
               "close preserves the configured capabilities"
    AssertTrue Not stream.AtEndOfStream, "closed file stream is not at EOF"
    fileStream.CloseStream
    AssertTrue stream.IsClosed, "repeated close is idempotent"
    AssertTrue Not stream.HasError, "successful repeated close leaves no error"

    fileStream.OpenForRead emptyPath
    AssertTrue stream.CanRead And Not stream.CanWrite, "empty file opens read-only"
    AssertTrue stream.AtEndOfStream, "open empty file is at EOF"
    AssertEqual "", stream.ReadText(10), "reading empty file returns empty text"
    AssertTrue stream.AtEndOfStream, "empty file remains at EOF after read"
    fileStream.CloseStream

    fileStream.OpenForRead readPath
    AssertEqual "", stream.ReadText(0), "zero read limit returns empty text"
    AssertEqual "", stream.ReadText(-1), "negative read limit returns empty text"
    AssertEqual "a", stream.ReadText(1), "non-positive reads do not advance the file"
    fileStream.CloseStream

    fileStream.OpenForWrite createPath, False
    AssertTrue Not stream.CanRead And stream.CanWrite, "write open configures write-only access"
    AssertTrue Not stream.AtEndOfStream, "write-only file stream is not at EOF"
    AssertEqual 3, fileStream.WriteText("one"), "legacy write returns written character count"
    AssertEqual 3, stream.WriteText("two"), "interface write returns written character count"
    On Error Resume Next
    stream.ReadText 1
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo TestFailure
    AssertEqual VBAsh_STREAM_ERROR_NOT_SUPPORTED, errorNumber, _
                "read on write-only file stream is unsupported"
    AssertTrue stream.HasError, "unsupported file read records an error"
    AssertEqual 1, stream.WriteText("!"), "successful write clears the prior error"
    AssertTrue Not stream.HasError, "successful file write clears error state"
    fileStream.CloseStream
    AssertTrue Not stream.CanRead And stream.CanWrite, _
               "closing a write stream preserves its configured capabilities"
    AssertTrue stream.IsClosed, "write stream is closed before external file verification"
    AssertTrue fileSystem.FileExists(createPath), "closed write stream leaves the file available externally"
    Set testFile = fileSystem.GetFile(createPath)
    AssertEqual 7, CDbl(testFile.Size), "closed write stream persists all written characters"
    Set testFile = Nothing
    AssertEqual "onetwo!", ReadFileStreamTestFile(fileSystem, createPath, stream), _
                "writes create a file and preserve successive text"

    fileStream.OpenForRead createPath
    AssertTrue stream.CanRead And Not stream.CanWrite, _
               "reopening a written file for reading replaces capabilities"
    AssertEqual "onetwo!", stream.ReadText(20), _
                "reopened read stream sees content persisted by the prior close"
    fileStream.CloseStream

    fileStream.OpenForWrite writePath, False
    AssertEqual 5, stream.WriteText("fresh"), "overwrite stream accepts new contents"
    fileStream.CloseStream
    AssertEqual "fresh", ReadFileStreamTestFile(fileSystem, writePath, stream), _
                "non-append open overwrites existing content"

    fileStream.OpenForWrite appendPath, True
    AssertTrue Not stream.CanRead And stream.CanWrite, "append stream is write-only"
    AssertEqual 3, stream.WriteText("-xy"), "append write reports character count"
    fileStream.CloseStream
    AssertEqual "base-xy", ReadFileStreamTestFile(fileSystem, appendPath, stream), _
                "append open preserves existing content"

    On Error Resume Next
    fileStream.OpenForRead missingPath
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo TestFailure
    AssertEqual VBAsh_STREAM_ERROR_OPEN_FAILURE, errorNumber, _
                "opening a missing file for reading reports open failure"
    AssertTrue stream.IsClosed, "failed open leaves stream closed"
    AssertTrue Not stream.CanRead And Not stream.CanWrite, _
               "failed open does not retain stale capabilities"
    AssertTrue stream.HasError, "failed open records error state"
    AssertEqual VBAsh_STREAM_ERROR_OPEN_FAILURE, stream.LastErrorNumber, _
                "failed open exposes the contract error code"
    AssertTrue Len(stream.LastErrorDescription) > 0, "failed open exposes a description"

    fileStream.OpenForRead readPath
    AssertTrue Not stream.HasError, "successful reopen clears a previous open error"
    On Error Resume Next
    stream.WriteText "x"
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo TestFailure
    AssertEqual VBAsh_STREAM_ERROR_NOT_SUPPORTED, errorNumber, _
                "write on read-only file stream is unsupported"
    AssertTrue stream.HasError, "unsupported file write records an error"
    AssertEqual "abc", stream.ReadText(3), "successful read clears the prior write error"
    AssertTrue Not stream.HasError, "successful file read clears error state"

    fileStream.CloseStream
    On Error Resume Next
    stream.ReadText 1
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo TestFailure
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, errorNumber, _
                "read after close raises the closed-stream error"
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, stream.LastErrorNumber, _
                "read after close records its contract error"
    AssertTrue Not stream.AtEndOfStream, "EOF query on a closed stream is false"
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, stream.LastErrorNumber, _
                "closed-stream EOF query does not mask the prior error"
    On Error Resume Next
    stream.WriteText "x"
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo TestFailure
    AssertEqual VBAsh_STREAM_ERROR_CLOSED, errorNumber, _
                "write after close raises the closed-stream error"
    AssertTrue stream.CanRead And Not stream.CanWrite, _
               "closed-stream operations preserve read capability"
    AssertTrue stream.IsClosed, "closed-stream operations preserve closed state"
    fileStream.CloseStream
    AssertTrue Not stream.HasError, "successful close clears the prior error"

    fileStream.OpenForRead readPath
    AssertEqual "abcdefg", stream.ReadText(7), "read limit equal to file size returns all text"
    AssertTrue stream.AtEndOfStream, "exact-size read reaches EOF"
    fileStream.CloseStream

    fileStream.OpenForWrite createPath, False
    AssertTrue Not stream.IsClosed, "a closed instance can be reopened"
    AssertTrue Not stream.CanRead And stream.CanWrite, _
               "reopen replaces the previous read capability"
    AssertEqual 4, stream.WriteText("next"), "reopened instance accepts writes"
    fileStream.CloseStream
    AssertTrue stream.IsClosed, "reopened write stream is closed before external verification"
    AssertTrue fileSystem.FileExists(createPath), "reopened write leaves the file available externally"
    Set testFile = fileSystem.GetFile(createPath)
    AssertEqual 4, CDbl(testFile.Size), "reopened write persists the replacement content"
    Set testFile = Nothing
    AssertEqual "next", ReadFileStreamTestFile(fileSystem, createPath, stream), _
                "reopen releases the old file and writes to the new stream"
    failureStage = "after assertion #74; normal cleanup path"
    GoTo Cleanup

TestFailure:
    testFailed = True
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    If Len(failureStage) = 0 Then
        failureStage = "after assertion #" & CStr(mLastAssertNumber) & _
                       " (" & mLastAssertDescription & ")"
    End If
    Err.Clear

Cleanup:
    If Not testFailed Then
        errorNumber = 0
        errorSource = ""
        errorDescription = ""
        failureStage = ""
    End If
    stageTrace = AppendFileStreamStage(stageTrace, "entering Cleanup")
    CleanupFileStreamTest fileStream, stream, fileSystem, testFolder, testFolderCreated, _
                          errorNumber, errorSource, errorDescription, failureStage, stageTrace
    stageTrace = AppendFileStreamStage(stageTrace, "return from CleanupFileStreamTest")
    If errorNumber <> 0 Then
        stageTrace = AppendFileStreamStage(stageTrace, "final propagation to RunTest")
        If Len(errorSource) = 0 Then errorSource = "VBAshTests.TestFileStream"
        errorDescription = "FileStream test error; first error at [" & failureStage & _
                           "]; stage trace: " & stageTrace & _
                           "; original error " & CStr(errorNumber) & _
                           " [" & errorDescription & "]; source [" & errorSource & _
                           "]; stream state: " & FileStreamState(stream)
        On Error GoTo 0
        Err.Raise errorNumber, errorSource, errorDescription
    End If
End Sub

Private Function ReadFileStreamTestFile(ByVal fileSystem As Object, ByVal path As String, _
                                        ByVal stream As IStream) As String
    Dim textStream As Object
    Dim file As Object
    Dim stage As String
    Dim fileExists As Boolean
    Dim fileSize As Double
    Dim originalErrorNumber As Long
    Dim originalErrorSource As String
    Dim originalErrorDescription As String

    On Error GoTo ReadFailure
    stage = "check file existence"
    fileExists = fileSystem.FileExists(path)
    If fileExists Then
        stage = "get file size"
        Set file = fileSystem.GetFile(path)
        fileSize = CDbl(file.Size)
        Set file = Nothing
    End If
    stage = "open independent FileSystemObject TextStream"
    Set textStream = fileSystem.OpenTextFile(path, 1, False)
    stage = "read independent FileSystemObject TextStream"
    ReadFileStreamTestFile = textStream.ReadAll
    stage = "close independent FileSystemObject TextStream"
    textStream.Close
    Set textStream = Nothing
    Exit Function

ReadFailure:
    originalErrorNumber = Err.Number
    originalErrorSource = Err.Source
    originalErrorDescription = Err.Description
    On Error GoTo 0
    Err.Raise originalErrorNumber, originalErrorSource, _
        "ReadFileStreamTestFile failed during " & stage & _
        " (error " & CStr(originalErrorNumber) & ": " & originalErrorDescription & _
        "; source: " & originalErrorSource & "; path: " & path & _
        "; exists: " & CStr(fileExists) & "; size: " & CStr(fileSize) & _
        " bytes; IStream closed: " & CStr(stream.IsClosed) & _
        "; CanRead: " & CStr(stream.CanRead) & _
        "; CanWrite: " & CStr(stream.CanWrite) & _
        "; HasError: " & CStr(stream.HasError) & _
        "; LastErrorNumber: " & CStr(stream.LastErrorNumber) & ")"
End Function

Private Sub CleanupFileStreamTest(ByVal fileStream As clsFileStream, _
                                  ByVal stream As IStream, ByVal fileSystem As Object, _
                                  ByVal folderPath As String, _
                                  ByVal folderCreated As Boolean, _
                                  ByRef errorNumber As Long, ByRef errorSource As String, _
                                  ByRef errorDescription As String, _
                                  ByRef failureStage As String, ByRef stageTrace As String)
    Dim cleanupErrorNumber As Long
    Dim cleanupErrorSource As String
    Dim cleanupErrorDescription As String
    Dim cleanupErrorStage As String
    Dim cleanupErrorDetails As String
    Dim folderExists As Boolean
    Dim firstErrorNumber As Long
    Dim firstErrorSource As String
    Dim firstErrorDescription As String
    Dim firstErrorStage As String

    firstErrorNumber = errorNumber
    firstErrorSource = errorSource
    firstErrorDescription = errorDescription
    firstErrorStage = failureStage

    stageTrace = AppendFileStreamStage(stageTrace, "second call to CloseStream")
    If Not fileStream Is Nothing Then
        On Error Resume Next
        Err.Clear
        fileStream.CloseStream
        If Err.Number <> 0 Then
            cleanupErrorNumber = Err.Number
            cleanupErrorSource = Err.Source
            cleanupErrorDescription = Err.Description
            cleanupErrorStage = "second call to CloseStream"
            cleanupErrorDetails = AppendFileStreamCleanupError(cleanupErrorDetails, _
                cleanupErrorNumber, cleanupErrorSource, cleanupErrorDescription, cleanupErrorStage)
            Err.Clear
        End If
        On Error GoTo 0
    End If

    If folderCreated And Not fileSystem Is Nothing Then
        stageTrace = AppendFileStreamStage(stageTrace, "verify temporary folder exists")
        On Error Resume Next
        Err.Clear
        folderExists = fileSystem.FolderExists(folderPath)
        If Err.Number <> 0 Then
            If cleanupErrorNumber = 0 Then
                cleanupErrorNumber = Err.Number
                cleanupErrorSource = Err.Source
                cleanupErrorDescription = Err.Description
                cleanupErrorStage = "verify temporary folder exists"
            End If
            cleanupErrorDetails = AppendFileStreamCleanupError(cleanupErrorDetails, _
                Err.Number, Err.Source, Err.Description, "verify temporary folder exists")
            Err.Clear
        End If
        On Error GoTo 0

        If folderExists Then
            stageTrace = AppendFileStreamStage(stageTrace, "delete temporary folder")
            On Error Resume Next
            Err.Clear
            fileSystem.DeleteFolder folderPath, True
            If Err.Number <> 0 Then
                If cleanupErrorNumber = 0 Then
                    cleanupErrorNumber = Err.Number
                    cleanupErrorSource = Err.Source
                    cleanupErrorDescription = Err.Description
                    cleanupErrorStage = "delete temporary folder"
                End If
                cleanupErrorDetails = AppendFileStreamCleanupError(cleanupErrorDetails, _
                    Err.Number, Err.Source, Err.Description, "delete temporary folder")
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End If

    If firstErrorNumber <> 0 Then
        errorNumber = firstErrorNumber
        errorSource = firstErrorSource
        errorDescription = firstErrorDescription
        failureStage = firstErrorStage
        If cleanupErrorNumber <> 0 Then
            errorDescription = errorDescription & "; cleanup errors: " & cleanupErrorDetails
        End If
    ElseIf cleanupErrorNumber <> 0 Then
        errorNumber = cleanupErrorNumber
        errorSource = cleanupErrorSource
        errorDescription = "Cleanup errors: " & cleanupErrorDetails
        failureStage = cleanupErrorStage
    End If
    stageTrace = AppendFileStreamStage(stageTrace, "return from cleanup")
End Sub

Private Function AppendFileStreamCleanupError(ByVal details As String, _
                                              ByVal errorNumber As Long, _
                                              ByVal errorSource As String, _
                                              ByVal errorDescription As String, _
                                              ByVal stage As String) As String
    If Len(details) > 0 Then details = details & "; "
    AppendFileStreamCleanupError = details & "error " & CStr(errorNumber) & _
                                   " [" & errorDescription & "]; source [" & errorSource & _
                                   "]; stage [" & stage & "]"
End Function

Private Function AppendFileStreamStage(ByVal stageTrace As String, ByVal stage As String) As String
    If Len(stageTrace) > 0 Then stageTrace = stageTrace & " > "
    AppendFileStreamStage = stageTrace & stage
End Function

Private Function FileStreamState(ByVal stream As IStream) As String
    If stream Is Nothing Then
        FileStreamState = "instance unavailable"
    Else
        FileStreamState = "IsClosed=" & CStr(stream.IsClosed) & _
                          ", CanRead=" & CStr(stream.CanRead) & _
                          ", CanWrite=" & CStr(stream.CanWrite) & _
                          ", HasError=" & CStr(stream.HasError) & _
                          ", LastErrorNumber=" & CStr(stream.LastErrorNumber)
    End If
End Function

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
    mLastAssertNumber = mAssertions - mSuiteStart
    mLastAssertDescription = description
    If CStr(expected) <> CStr(actual) Then
        mFailExpected = CStr(expected)
        mFailActual = CStr(actual)
        Err.Raise ASSERT_EQUAL_ERROR, "VBAshTests", _
            description & ": expected [" & CStr(expected) & "], got [" & CStr(actual) & "]"
    End If
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal description As String)
    mAssertions = mAssertions + 1
    mLastAssertNumber = mAssertions - mSuiteStart
    mLastAssertDescription = description
    If Not condition Then
        mFailExpected = "True"
        mFailActual = "False"
        Err.Raise ASSERT_TRUE_ERROR, "VBAshTests", description
    End If
End Sub
