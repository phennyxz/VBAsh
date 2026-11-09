' @VBAshRevision clsCommandResolver-r1
Option Explicit

Private mBuiltinRegistry As clsBuiltinRegistry

Private Sub Class_Initialize()
    Set mBuiltinRegistry = New clsBuiltinRegistry
End Sub

' Compatibility API retained for existing callers.
Public Function Resolve(ByVal commandName As String) As IBuiltin
    Set Resolve = mBuiltinRegistry.Resolve(commandName)
End Function

' Resolves an already-expanded SIMPLE_COMMAND AST node. This method classifies;
' it never executes a builtin or starts an external process.
Public Function ResolveCommand(ByVal sourceNode As clsCommand, _
                               ByVal sessionContext As clsSession) As clsResolvedCommand
    Dim resolved As clsResolvedCommand
    Dim argument As Variant
    Dim redirection As clsRedirection
    Dim builtin As IBuiltin
    Dim commandName As String

    If sourceNode Is Nothing Then
        Err.Raise vbObjectError + 1330, "clsCommandResolver.ResolveCommand", _
                  "Nó AST ausente."
    End If
    If sourceNode.NodeType <> "SIMPLE_COMMAND" Then
        Err.Raise vbObjectError + 1331, "clsCommandResolver.ResolveCommand", _
                  "A resolução de comando simples exige um nó SIMPLE_COMMAND."
    End If

    commandName = Trim$(sourceNode.Name)
    If Len(commandName) = 0 Then
        Err.Raise vbObjectError + 1332, "clsCommandResolver.ResolveCommand", _
                  "O nome do comando não pode ser vazio."
    End If

    Set builtin = mBuiltinRegistry.Resolve(commandName)
    Set resolved = New clsResolvedCommand
    If builtin Is Nothing Then
        resolved.Init commandName, "UNRESOLVED", sessionContext
    Else
        resolved.Init commandName, "BUILTIN", sessionContext
        Set resolved.Builtin = builtin
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
