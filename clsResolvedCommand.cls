' @VBAshRevision clsResolvedCommand-r1
Option Explicit

' A resolved command is execution metadata, not an AST node.
' CommandType values: ALIAS, FUNCTION, BUILTIN, EXTERNAL, COMPOUND.

Public Name As String
Public CommandType As String
Public Args As Collection
Public Redirections As Collection
Public Context As clsSession
Public SourceNode As clsCommand
Public Builtin As IBuiltin

Public Property Get ModuleRevision() As String
    ModuleRevision = "clsResolvedCommand-r1"
End Property

Private Sub Class_Initialize()
    Set Args = New Collection
    Set Redirections = New Collection
End Sub

Public Sub Init(ByVal commandName As String, ByVal resolvedType As String, _
                ByVal sessionContext As clsSession)
    Dim normalizedType As String

    If Len(Trim$(commandName)) = 0 Then
        Err.Raise vbObjectError + 1300, "clsResolvedCommand.Init", _
                  "O nome do comando não pode ser vazio."
    End If

    normalizedType = UCase$(Trim$(resolvedType))
    Select Case normalizedType
        Case "ALIAS", "FUNCTION", "BUILTIN", "EXTERNAL", "COMPOUND"
            ' Valid resolved command type.
        Case Else
            Err.Raise vbObjectError + 1301, "clsResolvedCommand.Init", _
                      "Tipo de comando resolvido inválido: " & resolvedType
    End Select

    Name = commandName
    CommandType = normalizedType
    Set Context = sessionContext
    Set SourceNode = Nothing
    Set Builtin = Nothing
    Set Args = New Collection
    Set Redirections = New Collection
End Sub

Public Sub AddArgument(ByVal argument As String)
    If Len(CommandType) = 0 Then
        Err.Raise vbObjectError + 1302, "clsResolvedCommand.AddArgument", _
                  "O comando precisa ser inicializado antes de receber argumentos."
    End If
    Args.Add argument
End Sub

Public Sub AddRedirection(ByVal redirection As clsRedirection)
    If Len(CommandType) = 0 Then
        Err.Raise vbObjectError + 1303, "clsResolvedCommand.AddRedirection", _
                  "O comando precisa ser inicializado antes de receber redirecionamentos."
    End If
    If redirection Is Nothing Then
        Err.Raise vbObjectError + 1304, "clsResolvedCommand.AddRedirection", _
                  "O redirecionamento não pode ser Nothing."
    End If
    Redirections.Add redirection
End Sub
