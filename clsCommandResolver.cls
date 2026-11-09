' @VBAshRevision clsCommandResolver-r3
Option Explicit

Private mBuiltinRegistry As clsBuiltinRegistry
Private mFunctionRegistry As clsFunctionRegistry
Private mExternalResolver As clsExternalCommandResolver

Public Property Get ModuleRevision() As String
    ModuleRevision = "clsCommandResolver-r3"
End Property

Private Sub Class_Initialize()
    Set mBuiltinRegistry = New clsBuiltinRegistry
    Set mFunctionRegistry = New clsFunctionRegistry
    Set mExternalResolver = New clsExternalCommandResolver
End Sub

' Compatibility API retained for existing callers.
Public Function Resolve(ByVal commandName As String) As IBuiltin
    Set Resolve = mBuiltinRegistry.Resolve(commandName)
End Function

Public Sub RegisterBuiltin(ByVal commandName As String, ByVal builtin As IBuiltin, _
                           Optional ByVal replaceExisting As Boolean = False)
    mBuiltinRegistry.RegisterBuiltin commandName, builtin, replaceExisting
End Sub

Public Function HasBuiltin(ByVal commandName As String) As Boolean
    HasBuiltin = mBuiltinRegistry.HasBuiltin(commandName)
End Function

Public Function UnregisterBuiltin(ByVal commandName As String) As Boolean
    UnregisterBuiltin = mBuiltinRegistry.UnregisterBuiltin(commandName)
End Function

Public Sub RegisterFunction(ByVal functionName As String, ByVal functionBody As clsCommand, _
                             Optional ByVal replaceExisting As Boolean = False)
    mFunctionRegistry.RegisterFunction functionName, functionBody, replaceExisting
End Sub

Public Function HasFunction(ByVal functionName As String) As Boolean
    HasFunction = mFunctionRegistry.HasFunction(functionName)
End Function

Public Function GetFunction(ByVal functionName As String) As clsCommand
    Set GetFunction = mFunctionRegistry.Resolve(functionName)
End Function

Public Function RemoveFunction(ByVal functionName As String) As Boolean
    RemoveFunction = mFunctionRegistry.RemoveFunction(functionName)
End Function

' Resolves an already-expanded AST node. Alias expansion is lexical and has
' already happened; function names take precedence over builtins and PATH.
Public Function ResolveCommand(ByVal sourceNode As clsCommand, _
                               ByVal sessionContext As clsSession) As clsResolvedCommand
    Dim resolved As clsResolvedCommand
    Dim argument As Variant
    Dim redirection As clsRedirection
    Dim builtin As IBuiltin
    Dim functionBody As clsCommand
    Dim commandName As String
    Dim executablePath As String

    If sourceNode Is Nothing Then
        Err.Raise vbObjectError + 1330, "clsCommandResolver.ResolveCommand", _
                  "Nó AST ausente."
    End If

    If sourceNode.NodeType <> "SIMPLE_COMMAND" Then
        Set resolved = New clsResolvedCommand
        resolved.Init sourceNode.NodeType, "COMPOUND", sessionContext
        Set resolved.SourceNode = sourceNode
        Set ResolveCommand = resolved
        Exit Function
    End If

    commandName = Trim$(sourceNode.Name)
    If Len(commandName) = 0 Then
        Err.Raise vbObjectError + 1332, "clsCommandResolver.ResolveCommand", _
                  "O nome do comando não pode ser vazio."
    End If

    Set functionBody = mFunctionRegistry.Resolve(commandName)
    Set builtin = mBuiltinRegistry.Resolve(commandName)
    Set resolved = New clsResolvedCommand

    If Not functionBody Is Nothing Then
        resolved.Init commandName, "FUNCTION", sessionContext
        Set resolved.FunctionBody = functionBody
    ElseIf Not builtin Is Nothing Then
        resolved.Init commandName, "BUILTIN", sessionContext
        Set resolved.Builtin = builtin
    Else
        executablePath = mExternalResolver.ResolveExecutable(commandName, sessionContext)
        If Len(executablePath) > 0 Then
            resolved.Init commandName, "EXTERNAL", sessionContext
            resolved.ExecutablePath = executablePath
        Else
            resolved.Init commandName, "UNRESOLVED", sessionContext
        End If
    End If

    Set resolved.SourceNode = sourceNode
    For Each argument In sourceNode.Args
        resolved.AddArgument CStr(argument)
    Next argument
    For Each redirection In sourceNode.Redirections
        resolved.AddRedirection redirection
    Next redirection
    Set ResolveCommand = resolved
End Function
